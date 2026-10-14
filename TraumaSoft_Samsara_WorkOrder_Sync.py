"""
TraumaSoft -> Samsara Work Order Sync

Reads vehicle maintenance work orders from the TraumaSoft reporting database
(table `vm_work_orders`, joined to `sched_vehicles` / `cost_centers`) and
mirrors them into Samsara Maintenance Work Orders:

  * new TraumaSoft work orders  -> POST /maintenance/work-orders
  * changed TraumaSoft work orders (status, description, odometer, dates)
                                -> PATCH /maintenance/work-orders

A local JSON ledger maps TraumaSoft work order ids to Samsara work order ids so
the job is idempotent across runs.  Every Samsara description is prefixed with
`[TS-WO:<id>]` so the ledger can be rebuilt from Samsara if it is ever lost
(`--rebuild-ledger`).

Configuration is via environment variables / .env (see README).

Usage:
    python TraumaSoft_Samsara_WorkOrder_Sync.py                # normal run
    python TraumaSoft_Samsara_WorkOrder_Sync.py --dry-run      # no writes to Samsara
    python TraumaSoft_Samsara_WorkOrder_Sync.py --discover     # print vm_work_orders schema + samples
    python TraumaSoft_Samsara_WorkOrder_Sync.py --rebuild-ledger
    python TraumaSoft_Samsara_WorkOrder_Sync.py --since 2026-01-01 --limit 50
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import smtplib
import ssl
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo

from samsara_workorders_api import (
    DEFAULT_BASE_URL,
    SamsaraApiError,
    SamsaraClient,
    WORK_ORDER_CATEGORIES,
    WORK_ORDER_PRIORITIES,
    WORK_ORDER_STATUSES,
)

try:
    from dotenv import load_dotenv

    load_dotenv(r".env")
except ImportError:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

# =============================
# CONFIG
# =============================
def _env_bool(name: str, default: bool = False) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in ("1", "true", "yes", "y")


def _env_json(name: str, default: Any) -> Any:
    raw = os.getenv(name)
    if not raw:
        return default
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise SystemExit(f"{name} is not valid JSON: {e}")


OUTPUT_DIR = os.getenv("OUTPUT_DIR", "Reports")
STATE_DIR = os.getenv("SYNC_STATE_DIR", os.path.join(OUTPUT_DIR, "State"))
LEDGER_FILE = os.getenv("SAMSARA_WO_LEDGER_FILE", os.path.join(STATE_DIR, "samsara_workorder_ledger.json"))
# SQL ledger ("endpoint table"): when SAMSARA_WO_LEDGER_DSN is set the mapping lives in a
# database table instead of the JSON file.  DDL for the table is in sql/.
LEDGER_DSN = os.getenv("SAMSARA_WO_LEDGER_DSN", "")
LEDGER_TABLE = os.getenv("SAMSARA_WO_LEDGER_TABLE", "samsara_workorder_map")
LEDGER_DIALECT = os.getenv("SAMSARA_WO_LEDGER_DIALECT", "mysql").strip().lower()  # mysql | mssql | sqlite

DB_DSN = os.getenv("DB_DSN") or os.getenv("TS_DSN", "")
# Where TraumaSoft work orders come from: db (ODBC DSN -> vm_work_orders), api (TraumaSoft
# third-party REST API via traumasoft_api.py), or csv (an export file).  Auto-detected when unset.
TS_WO_SOURCE = os.getenv("TS_WO_SOURCE", "").strip().lower()
TS_WO_API_PATH = os.getenv("TS_WO_API_PATH", "ThirdParty/Data/WorkOrders")
TS_WO_API_PARAMS = _env_json("TS_WO_API_PARAMS", {})
TS_WO_CSV_FILE = os.getenv("TS_WO_CSV_FILE", "")
TS_TIMEZONE = os.getenv("TS_TIMEZONE", "America/New_York")
TS_WO_TABLE = os.getenv("TS_WO_TABLE", "vm_work_orders")
TS_WO_SQL_FILE = os.getenv("TS_WO_SQL_FILE")  # optional: full override of the source query
TS_WO_LOOKBACK_DAYS = int(os.getenv("TS_WO_LOOKBACK_DAYS", "30"))
TS_WO_EXCLUDE_VEHICLE_IDS = {
    int(x) for x in os.getenv("TS_WO_EXCLUDE_VEHICLE_IDS", "1001,1002,1003").split(",") if x.strip().isdigit()
}

SAMSARA_API_TOKEN = os.getenv("SAMSARA_API_TOKEN", "")
SAMSARA_BASE_URL = os.getenv("SAMSARA_BASE_URL", DEFAULT_BASE_URL)
SAMSARA_DEFAULT_PRIORITY = os.getenv("SAMSARA_DEFAULT_PRIORITY", "Medium")
SAMSARA_SEND_INITIAL_STATUS = _env_bool("SAMSARA_SEND_INITIAL_STATUS", True)
# How to match a TraumaSoft vehicle to a Samsara vehicle: name | externalId | vin
SAMSARA_VEHICLE_MATCH = os.getenv("SAMSARA_VEHICLE_MATCH", "name").strip()
SAMSARA_VEHICLE_EXTERNAL_ID_KEY = os.getenv("SAMSARA_VEHICLE_EXTERNAL_ID_KEY", "traumasoftVehicleId")

TEST_MODE = _env_bool("TEST_MODE", False)
TEST_MODE_RECIPIENT = os.getenv("TEST_MODE_RECIPIENT") or os.getenv("REPORT_TEST_EMAIL") or "reports-test@example.com"
STATUS_EMAIL_RECIPIENT = os.getenv("STATUS_EMAIL_RECIPIENT") or os.getenv("REPORT_STATUS_EMAIL") or "reports-status@example.com"
PROD_RECIPIENTS = [r.strip() for r in os.getenv("SYNC_RECIPIENTS", "").split(",") if r.strip()]
EMAIL_SUBJECT = "TraumaSoft -> Samsara Work Order Sync - {date}{test_suffix}"

TS_TAG_RE = re.compile(r"\[TS-WO:(\d+)\]")

# Candidate column names for each logical field, first match wins.  Override or
# extend with TS_WO_COLUMN_MAP='{"description": ["complaint"], ...}'.
DEFAULT_COLUMN_CANDIDATES: Dict[str, List[str]] = {
    "id": ["id", "work_order_id", "wo_id", "Id", "ID", "WorkOrderId", "WorkOrderID"],
    "number": ["number", "wo_number", "work_order_number", "work_order_no", "WorkOrderNumber", "Number"],
    "vehicle_id": ["vehicle_id", "veh_id", "vehicleId", "VehicleId", "VehicleID"],
    "vehicle_name": ["vehicle_name", "vehicle", "unit", "unit_name", "vehicle_number", "VehicleName", "Vehicle", "Unit"],
    "cost_center_id": ["cost_center_id", "costCenterId", "CostCenterId"],
    "cost_center_name": ["cost_center_name", "cost_center", "station", "CostCenter", "Station"],
    "description": ["description", "complaint", "notes", "details", "problem", "work_description", "summary", "title", "name", "Description", "Notes"],
    "status": ["status", "status_name", "state", "wo_status", "Status", "StatusName"],
    "status_id": ["status_id"],
    "priority": ["priority", "priority_name", "priority_id"],
    "category": ["category", "type", "work_type", "wo_type", "service_type"],
    "created_at": ["date_created", "created_at", "created", "date_opened", "opened_at", "date_added", "date", "timestamp"],
    "updated_at": ["date_updated", "updated_at", "modified", "date_modified", "last_updated", "last_modified"],
    "started_at": ["date_started", "started_at", "start_date", "date_start"],
    "completed_at": ["date_completed", "completed_at", "completed", "date_closed", "closed_at", "close_date", "date_finished"],
    "due_at": ["due_date", "date_due", "due_at"],
    "odometer": ["odometer_started", "odometer", "mileage", "odometer_completed"],
    "odometer_type": ["odometer_type"],
    "engine_hours": ["engine_hours", "hours"],
    "deleted": ["deleted", "is_deleted"],
}

DEFAULT_STATUS_MAP: Dict[str, str] = {
    "open": "Open", "new": "Open", "pending": "Open", "created": "Open", "scheduled": "Open",
    "assigned": "Assigned",
    "in progress": "In Progress", "in_progress": "In Progress", "inprogress": "In Progress",
    "working": "In Progress", "started": "In Progress", "active": "In Progress",
    "on hold": "On Hold", "on_hold": "On Hold", "hold": "On Hold", "waiting": "On Hold", "deferred": "On Hold",
    "pending parts": "Pending Parts", "pending_parts": "Pending Parts", "parts": "Pending Parts",
    "awaiting parts": "Pending Parts", "waiting for parts": "Pending Parts", "waiting on parts": "Pending Parts",
    "pending approval": "Pending Approval", "pending_approval": "Pending Approval", "approval": "Pending Approval",
    "completed": "Completed", "complete": "Completed", "done": "Completed", "finished": "Completed", "resolved": "Completed",
    "closed": "Closed",
    "cancelled": "Cancelled", "canceled": "Cancelled", "void": "Cancelled", "voided": "Cancelled", "deleted": "Cancelled",
}

DEFAULT_PRIORITY_MAP: Dict[str, str] = {
    "low": "Low", "1": "Low", "minor": "Low",
    "medium": "Medium", "normal": "Medium", "2": "Medium", "standard": "Medium", "routine": "Medium",
    "high": "High", "3": "High", "major": "High",
    "urgent": "Urgent", "critical": "Urgent", "emergency": "Urgent", "4": "Urgent", "immediate": "Urgent",
}

DEFAULT_CATEGORY_MAP: Dict[str, str] = {
    "annual": "Annual", "annual inspection": "Annual", "inspection": "Annual",
    "corrective": "Corrective", "repair": "Corrective", "breakdown": "Corrective", "unscheduled": "Corrective",
    "damage": "Damage Repair", "damage repair": "Damage Repair", "accident": "Damage Repair", "body": "Damage Repair",
    "preventive": "Preventive", "preventative": "Preventive", "pm": "Preventive", "scheduled": "Preventive",
    "service": "Preventive", "maintenance": "Preventive",
    "recall": "Recall",
}

TERMINAL_STATUSES = {"Completed", "Closed", "Cancelled"}


# =============================
# DATA MODEL
# =============================
@dataclass
class WorkOrderRecord:
    ts_id: int
    vehicle_id: Optional[int]
    vehicle_name: Optional[str]
    description: str
    status: str                     # Samsara status
    priority: str                   # Samsara priority
    category: Optional[str]         # Samsara category or None
    created_at: Optional[datetime]
    started_at: Optional[datetime]
    completed_at: Optional[datetime]
    due_at: Optional[datetime]
    odometer_meters: Optional[int]
    engine_hours: Optional[int]
    cost_center_name: Optional[str]
    ts_number: Optional[str] = None
    raw_status: Optional[str] = None
    extra: Dict[str, Any] = field(default_factory=dict)

    @property
    def tag(self) -> str:
        return f"[TS-WO:{self.ts_id}]"


# =============================
# PURE HELPERS (unit tested)
# =============================
def resolve_column_map(columns: Iterable[str], overrides: Optional[Dict[str, List[str]]] = None) -> Dict[str, Optional[str]]:
    """Pick the actual column for each logical field from the available columns."""
    def norm(name: str) -> str:  # 'DateUpdated' == 'date_updated' == 'Date Updated'
        return re.sub(r"[^a-z0-9]", "", str(name).lower())

    available: Dict[str, str] = {}
    for c in columns:
        available.setdefault(norm(c), c)
    candidates = {k: list(v) for k, v in DEFAULT_COLUMN_CANDIDATES.items()}
    for key, names in (overrides or {}).items():
        if isinstance(names, str):
            names = [names]
        candidates[key] = list(names) + candidates.get(key, [])
    resolved: Dict[str, Optional[str]] = {}
    for key, names in candidates.items():
        resolved[key] = next((available[norm(n)] for n in names if norm(n) in available), None)
    return resolved


def _norm_key(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower())


def map_status(raw: Any, completed_at: Optional[datetime] = None, status_map: Optional[Dict[str, str]] = None) -> str:
    """Map a TraumaSoft status (text or numeric id) onto a Samsara status."""
    table = dict(DEFAULT_STATUS_MAP)
    table.update({_norm_key(k): v for k, v in (status_map or {}).items()})
    if raw is not None and str(raw).strip() != "":
        key = _norm_key(raw)
        if key in table and table[key] in WORK_ORDER_STATUSES:
            return table[key]
        for s in WORK_ORDER_STATUSES:
            if key == s.lower():
                return s
        logging.warning("Unmapped TraumaSoft work order status %r; defaulting", raw)
    return "Completed" if completed_at else "Open"


def map_priority(raw: Any, default: str = "Medium", priority_map: Optional[Dict[str, str]] = None) -> str:
    table = dict(DEFAULT_PRIORITY_MAP)
    table.update({_norm_key(k): v for k, v in (priority_map or {}).items()})
    if raw is not None and str(raw).strip() != "":
        key = _norm_key(raw)
        if key in table and table[key] in WORK_ORDER_PRIORITIES:
            return table[key]
        for p in WORK_ORDER_PRIORITIES:
            if key == p.lower():
                return p
    return default if default in WORK_ORDER_PRIORITIES else "Medium"


def map_category(raw: Any, category_map: Optional[Dict[str, str]] = None) -> Optional[str]:
    table = dict(DEFAULT_CATEGORY_MAP)
    table.update({_norm_key(k): v for k, v in (category_map or {}).items()})
    if raw is None or str(raw).strip() == "":
        return None
    key = _norm_key(raw)
    if key in table and table[key] in WORK_ORDER_CATEGORIES:
        return table[key]
    for c in WORK_ORDER_CATEGORIES:
        if key == c.lower():
            return c
    return None


def odometer_to_meters(value: Any, odometer_type: Any = None) -> Optional[int]:
    """TraumaSoft odometer (miles unless odometer_type says km) -> integer meters."""
    if value is None or value == "":
        return None
    try:
        num = float(value)
    except (TypeError, ValueError):
        return None
    if num <= 0:
        return None
    unit = _norm_key(odometer_type or "")
    if unit.startswith("k"):  # km / kilometers / kilometres
        return int(round(num * 1000))
    return int(round(num * 1609.344))


def to_rfc3339(value: Any, tz_name: str = TS_TIMEZONE) -> Optional[str]:
    """Naive TraumaSoft timestamps are local time; emit RFC 3339 with offset."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        dt = value
    elif isinstance(value, date):
        dt = datetime.combine(value, datetime.min.time())
    else:
        s = str(value).strip()
        if s in ("0000-00-00", "0000-00-00 00:00:00"):
            return None
        try:
            dt = datetime.fromisoformat(s)
        except ValueError:
            return None
    if dt.year < 1971:  # TraumaSoft uses zero/epoch dates for "not set"
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo(tz_name))
    return dt.isoformat(timespec="seconds")


def build_description(rec: WorkOrderRecord) -> str:
    """`[TS-WO:123] <description>` plus a short context trailer."""
    body = (rec.description or "").strip() or "(no description in TraumaSoft)"
    parts = [f"{rec.tag} {body}"]
    trailer = []
    if rec.ts_number:
        trailer.append(f"TraumaSoft WO# {rec.ts_number}")
    if rec.cost_center_name:
        trailer.append(f"Station: {rec.cost_center_name}")
    if rec.raw_status:
        trailer.append(f"TS status: {rec.raw_status}")
    if trailer:
        parts.append(" | ".join(trailer))
    return "\n".join(parts)


def build_create_payload(rec: WorkOrderRecord, asset_id: str, send_initial_status: bool = True) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "assetId": asset_id,
        "description": build_description(rec),
        "priority": rec.priority,
    }
    if rec.category:
        payload["category"] = rec.category
    if rec.odometer_meters:
        payload["odometerMeters"] = rec.odometer_meters
    if rec.engine_hours:
        payload["engineHours"] = rec.engine_hours
    if rec.due_at:
        payload["dueAtTime"] = to_rfc3339(rec.due_at)
    if send_initial_status:
        # Supported per Samsara changelog "POST maintenance/work-orders: Add support
        # for initial status and historical timestamps on create".
        if rec.status != "Open":
            payload["status"] = rec.status
        if rec.started_at:
            payload["startedAtTime"] = to_rfc3339(rec.started_at)
        if rec.completed_at and rec.status in TERMINAL_STATUSES:
            payload["completedAtTime"] = to_rfc3339(rec.completed_at)
    return payload


def build_update_payload(rec: WorkOrderRecord) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "description": build_description(rec),
        "priority": rec.priority,
        "status": rec.status,
    }
    if rec.category:
        payload["category"] = rec.category
    if rec.odometer_meters:
        payload["odometerMeters"] = rec.odometer_meters
    if rec.engine_hours:
        payload["engineHours"] = rec.engine_hours
    if rec.due_at:
        payload["dueAtTime"] = to_rfc3339(rec.due_at)
    if rec.completed_at and rec.status in TERMINAL_STATUSES:
        payload["completedAtTime"] = to_rfc3339(rec.completed_at)
    return payload


def fingerprint(rec: WorkOrderRecord) -> str:
    """Stable hash of everything the sync would PATCH; unchanged => skip."""
    data = build_update_payload(rec)
    return hashlib.sha256(json.dumps(data, sort_keys=True, default=str).encode("utf-8")).hexdigest()


def normalize_vehicle_name(name: Any) -> str:
    """Case/space/punctuation-insensitive vehicle name key (e.g. 'M-12' == 'm 12' == 'M12')."""
    if name is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


class VehicleIndex:
    """Maps TraumaSoft vehicles onto Samsara vehicle (asset) ids."""

    def __init__(self, vehicles: Iterable[Dict[str, Any]], external_id_key: str = SAMSARA_VEHICLE_EXTERNAL_ID_KEY):
        self.by_name: Dict[str, str] = {}
        self.by_external: Dict[str, str] = {}
        self.by_vin: Dict[str, str] = {}
        self.names: Dict[str, str] = {}
        self.duplicates: set = set()
        for v in vehicles:
            vid = str(v.get("id") or "")
            if not vid:
                continue
            self.names[vid] = v.get("name") or vid
            key = normalize_vehicle_name(v.get("name"))
            if key:
                if key in self.by_name and self.by_name[key] != vid:
                    self.duplicates.add(key)
                self.by_name.setdefault(key, vid)
            ext = (v.get("externalIds") or {}).get(external_id_key)
            if ext:
                self.by_external[str(ext).strip()] = vid
            vin = v.get("vin") or v.get("vehicleVin")
            if vin:
                self.by_vin[str(vin).strip().upper()] = vid

    def resolve(self, rec: WorkOrderRecord, mode: str = "name") -> Optional[str]:
        mode = (mode or "name").lower()
        if mode == "externalid":
            return self.by_external.get(str(rec.vehicle_id)) if rec.vehicle_id is not None else None
        if mode == "vin":
            vin = rec.extra.get("vehicle_vin")
            return self.by_vin.get(str(vin).strip().upper()) if vin else None
        key = normalize_vehicle_name(rec.vehicle_name)
        if not key:
            return None
        if key in self.duplicates:
            logging.warning("Samsara has multiple vehicles named like %r; using first match", rec.vehicle_name)
        return self.by_name.get(key)


def row_to_record(row: Dict[str, Any], colmap: Dict[str, Optional[str]], maps: Optional[Dict[str, Dict[str, str]]] = None,
                  default_priority: str = SAMSARA_DEFAULT_PRIORITY) -> WorkOrderRecord:
    maps = maps or {}

    def get(key: str) -> Any:
        col = colmap.get(key)
        return row.get(col) if col else None

    def as_dt(v: Any) -> Optional[datetime]:
        s = to_rfc3339(v)
        return datetime.fromisoformat(s) if s else None

    completed_at = as_dt(get("completed_at"))
    raw_status = get("status")
    if raw_status is None:
        raw_status = get("status_id")
    status = map_status(raw_status, completed_at, maps.get("status"))
    if status not in TERMINAL_STATUSES and completed_at and colmap.get("status") is None:
        status = "Completed"

    odo_type = get("odometer_type") or row.get("vehicle_odometer_type")
    eh = get("engine_hours")
    try:
        engine_hours = int(float(eh)) if eh not in (None, "") and float(eh) > 0 else None
    except (TypeError, ValueError):
        engine_hours = None

    vid = get("vehicle_id")
    try:
        vehicle_id = int(vid) if vid not in (None, "") else None
    except (TypeError, ValueError):
        vehicle_id = None

    return WorkOrderRecord(
        ts_id=int(get("id")),
        vehicle_id=vehicle_id,
        vehicle_name=row.get("vehicle_name") or get("vehicle_name"),
        description=str(get("description") or "").strip(),
        status=status,
        priority=map_priority(get("priority"), default_priority, maps.get("priority")),
        category=map_category(get("category"), maps.get("category")),
        created_at=as_dt(get("created_at")),
        started_at=as_dt(get("started_at")),
        completed_at=completed_at,
        due_at=as_dt(get("due_at")),
        odometer_meters=odometer_to_meters(get("odometer"), odo_type),
        engine_hours=engine_hours,
        cost_center_name=row.get("cost_center_name") or get("cost_center_name"),
        ts_number=str(get("number")).strip() if get("number") not in (None, "") else None,
        raw_status=str(raw_status).strip() if raw_status not in (None, "") else None,
        extra={"vehicle_vin": row.get("vehicle_vin")},
    )


# =============================
# LEDGER  (JSON file or SQL "endpoint table")
# =============================
LEDGER_COLUMNS = (
    "ts_wo_id", "samsara_wo_id", "samsara_asset_id", "ts_vehicle_id", "vehicle_name",
    "samsara_status", "fingerprint", "last_action", "first_synced_at", "last_synced_at",
)


class Ledger:
    """JSON-file ledger: {"version": 1, "work_orders": {"<ts_id>": {...}}}."""

    def __init__(self, path: str):
        self.path = Path(path)
        self.data: Dict[str, Any] = {"version": 1, "work_orders": {}}
        if self.path.exists():
            with open(self.path, "r", encoding="utf-8") as f:
                self.data = json.load(f)
            self.data.setdefault("work_orders", {})

    def describe(self) -> str:
        return f"json:{self.path}"

    @property
    def work_orders(self) -> Dict[str, Dict[str, Any]]:
        return self.data["work_orders"]

    def get(self, ts_id: int) -> Optional[Dict[str, Any]]:
        return self.work_orders.get(str(ts_id))

    def put(self, ts_id: int, samsara_id: str, fp: Optional[str], status: str, asset_id: Optional[str],
            vehicle_id: Optional[int] = None, vehicle_name: Optional[str] = None, action: Optional[str] = None) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        entry = self.work_orders.setdefault(str(ts_id), {"created_at": now})
        entry.update({"samsara_id": str(samsara_id), "fingerprint": fp, "status": status,
                      "asset_id": asset_id, "vehicle_id": vehicle_id, "vehicle_name": vehicle_name,
                      "last_action": action, "updated_at": now})

    def open_ts_ids(self) -> List[int]:
        return [int(k) for k, v in self.work_orders.items() if v.get("status") not in TERMINAL_STATUSES]

    def max_ts_id(self) -> int:
        return max((int(k) for k in self.work_orders), default=0)

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix="ledger_", suffix=".json", dir=str(self.path.parent))
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, sort_keys=True)
        os.replace(tmp, self.path)


class SqlLedger:
    """
    Database-backed ledger.  `conn` is any DB-API connection using `?` parameters
    (pyodbc, sqlite3).  The whole table is cached in memory at start-up (it is
    small: one row per synced work order) and every put() is written through
    immediately with update-then-insert so it works on MySQL, SQL Server and SQLite.
    """

    def __init__(self, conn, table: str = LEDGER_TABLE, dialect: str = LEDGER_DIALECT):
        self.conn = conn
        self.table = table
        self.dialect = dialect
        self.work_orders: Dict[str, Dict[str, Any]] = {}
        self._load()

    def describe(self) -> str:
        return f"sql:{self.dialect}:{self.table}"

    @staticmethod
    def ddl_path(dialect: str) -> Path:
        return Path(__file__).resolve().parent / "sql" / f"samsara_workorder_map.{dialect}.sql"

    def ensure_table(self) -> None:
        """Create the mapping table if it does not exist, using sql/<dialect>.sql."""
        ddl = self.ddl_path(self.dialect).read_text(encoding="utf-8").replace("samsara_workorder_map", self.table)
        ddl = "\n".join(line for line in ddl.splitlines() if not line.strip().startswith("--"))
        cur = self.conn.cursor()
        for stmt in [x.strip() for x in ddl.split(";") if x.strip()]:
            cur.execute(stmt)
        self.conn.commit()
        cur.close()
        self._load()

    def _load(self) -> None:
        cur = self.conn.cursor()
        try:
            cur.execute(f"SELECT {', '.join(LEDGER_COLUMNS)} FROM {self.table}")
        except Exception as e:  # table missing -> empty ledger; --init-ledger-table creates it
            logging.warning("Ledger table %s not readable (%s); treating as empty", self.table, e)
            self.work_orders = {}
            try:
                self.conn.rollback()
            except Exception:
                pass
            return
        self.work_orders = {}
        for r in cur.fetchall():
            row = dict(zip(LEDGER_COLUMNS, r))
            self.work_orders[str(row["ts_wo_id"])] = {
                "samsara_id": str(row["samsara_wo_id"]), "fingerprint": row["fingerprint"],
                "status": row["samsara_status"], "asset_id": row["samsara_asset_id"],
                "vehicle_id": row["ts_vehicle_id"], "vehicle_name": row["vehicle_name"],
                "last_action": row["last_action"], "created_at": str(row["first_synced_at"]),
                "updated_at": str(row["last_synced_at"]),
            }
        cur.close()

    def get(self, ts_id: int) -> Optional[Dict[str, Any]]:
        return self.work_orders.get(str(ts_id))

    def put(self, ts_id: int, samsara_id: str, fp: Optional[str], status: str, asset_id: Optional[str],
            vehicle_id: Optional[int] = None, vehicle_name: Optional[str] = None, action: Optional[str] = None) -> None:
        now = datetime.now().replace(microsecond=0)
        existing = self.get(ts_id)
        cur = self.conn.cursor()
        if existing:
            cur.execute(
                f"UPDATE {self.table} SET samsara_wo_id=?, samsara_asset_id=?, ts_vehicle_id=?, vehicle_name=?, "
                f"samsara_status=?, fingerprint=?, last_action=?, last_synced_at=? WHERE ts_wo_id=?",
                (str(samsara_id), asset_id, vehicle_id, vehicle_name, status, fp, action, now, int(ts_id)),
            )
            if cur.rowcount == 0:
                existing = None
        if not existing:
            cur.execute(
                f"INSERT INTO {self.table} ({', '.join(LEDGER_COLUMNS)}) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (int(ts_id), str(samsara_id), asset_id, vehicle_id, vehicle_name, status, fp, action, now, now),
            )
        cur.close()
        entry = self.work_orders.setdefault(str(ts_id), {"created_at": now.isoformat()})
        entry.update({"samsara_id": str(samsara_id), "fingerprint": fp, "status": status, "asset_id": asset_id,
                      "vehicle_id": vehicle_id, "vehicle_name": vehicle_name, "last_action": action,
                      "updated_at": now.isoformat()})

    def open_ts_ids(self) -> List[int]:
        return [int(k) for k, v in self.work_orders.items() if v.get("status") not in TERMINAL_STATUSES]

    def max_ts_id(self) -> int:
        return max((int(k) for k in self.work_orders), default=0)

    def save(self) -> None:
        self.conn.commit()


def open_ledger():
    """SQL ledger when SAMSARA_WO_LEDGER_DSN is set, otherwise the JSON file."""
    if LEDGER_DSN:
        import pyodbc

        logging.info("Ledger: SQL table %s via DSN=%s (%s)", LEDGER_TABLE, LEDGER_DSN, LEDGER_DIALECT)
        return SqlLedger(pyodbc.connect(f"DSN={LEDGER_DSN};"), LEDGER_TABLE, LEDGER_DIALECT)
    return Ledger(LEDGER_FILE)


# =============================
# TRAUMASOFT SOURCE
# =============================
def get_db_connection():
    import pyodbc  # imported lazily so unit tests can run without the ODBC driver

    conn_str = f"DSN={DB_DSN};"
    logging.info("Connecting using %s", conn_str)
    return pyodbc.connect(conn_str)


def introspect_columns(conn, table: str) -> List[str]:
    cur = conn.cursor()
    cur.execute(f"SELECT * FROM {table} WHERE 1=0")
    cols = [d[0] for d in cur.description]
    cur.close()
    return cols


def default_source_sql(colmap: Dict[str, Optional[str]], table: str = TS_WO_TABLE) -> Tuple[str, str]:
    """
    Returns (sql_recent, sql_by_ids).  `sql_recent` takes one `?` (lookback start);
    `sql_by_ids` is formatted with a placeholder list at call time.
    Only columns known to exist in the TraumaSoft schema are referenced by name;
    the work order row itself is selected as `wo.*`.
    """
    base = f"""
SELECT
    wo.*,
    v.name AS vehicle_name,
    v.odometer AS vehicle_current_odometer,
    v.odometer_type AS vehicle_odometer_type,
    cc.name AS cost_center_name
FROM {table} wo
LEFT JOIN sched_vehicles v ON v.id = wo.{colmap.get('vehicle_id') or 'vehicle_id'}
LEFT JOIN cost_centers cc ON cc.id = wo.{colmap.get('cost_center_id') or 'cost_center_id'}
"""
    where = []
    if colmap.get("deleted"):
        where.append(f"COALESCE(wo.{colmap['deleted']}, 0) = 0")
    date_cols = [c for c in (colmap.get("updated_at"), colmap.get("created_at"), colmap.get("completed_at")) if c]
    id_col = colmap.get("id") or "id"
    if date_cols:
        expr = f"COALESCE({', '.join('wo.' + c for c in date_cols)})"
        recent_filter = f"{expr} >= ?"
    else:
        # No usable date column: fall back to "id greater than" using the lookback as a dummy param.
        logging.warning("No date column found on %s; falling back to id-based incremental filter", table)
        recent_filter = f"wo.{id_col} > ?"
    sql_recent = base + "WHERE " + " AND ".join(where + [recent_filter]) + f"\nORDER BY wo.{id_col}"
    sql_by_ids = base + "WHERE " + " AND ".join(where + [f"wo.{id_col} IN ({{placeholders}})"]) + f"\nORDER BY wo.{id_col}"
    return sql_recent, sql_by_ids


def _rows(cur) -> List[Dict[str, Any]]:
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def fetch_work_orders(conn, colmap: Dict[str, Optional[str]], since: datetime, extra_ids: List[int],
                      limit: Optional[int] = None, ledger_max_id: int = 0) -> List[Dict[str, Any]]:
    if TS_WO_SQL_FILE:
        sql_recent = Path(TS_WO_SQL_FILE).read_text(encoding="utf-8")
        sql_by_ids = None
        logging.info("Using source SQL from %s", TS_WO_SQL_FILE)
    else:
        sql_recent, sql_by_ids = default_source_sql(colmap)

    cur = conn.cursor()
    has_date = any(colmap.get(k) for k in ("updated_at", "created_at", "completed_at"))
    param = since if (has_date or TS_WO_SQL_FILE) else ledger_max_id
    cur.execute(sql_recent, (param,))
    rows = _rows(cur)
    seen = {r.get(colmap.get("id") or "id") for r in rows}

    missing = [i for i in extra_ids if i not in seen]
    if missing and sql_by_ids:
        for i in range(0, len(missing), 500):
            chunk = missing[i:i + 500]
            cur.execute(sql_by_ids.format(placeholders=",".join("?" * len(chunk))), chunk)
            rows.extend(_rows(cur))
    cur.close()
    if limit:
        rows = rows[:limit]
    return rows


# =============================
# TRAUMASOFT SOURCE: API / CSV
# =============================
TS_API_WO_PATH_CANDIDATES = [
    "ThirdParty/Data/WorkOrders", "ThirdParty/Data/WorkOrder", "ThirdParty/Data/VehicleWorkOrders",
    "ThirdParty/Data/Maintenance", "ThirdParty/Data/VehicleMaintenance", "ThirdParty/Data/FleetMaintenance",
    "ThirdParty/Data/MaintenanceRecords", "ThirdParty/Data/VehicleService", "ThirdParty/Data/ServiceRecords",
    "ThirdParty/Data/Repairs", "ThirdParty/Data/VehicleRepairs", "ThirdParty/Data/FleetWorkOrders",
    "ThirdParty/Data/VehicleStatusLog", "ThirdParty/Data/VehicleStatuses",
]


def open_traumasoft_api():
    """Instantiate the project's TraumasoftAPI client (traumasoft_api.py next to this script)."""
    try:
        from traumasoft_api import TraumasoftAPI  # the existing client in the Reports folder
    except ImportError as e:
        raise SystemExit(
            "TS_WO_SOURCE=api needs traumasoft_api.py (the existing TraumaSoft client) importable "
            f"from the same folder as this script: {e}"
        )
    return TraumasoftAPI()


def _api_rows(api, path: str, params: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Pull every row from a TraumaSoft list endpoint using whatever the client offers."""
    params = dict(params or {})
    if hasattr(api, "paginate"):
        return [r for r in api.paginate(path, params=params) if isinstance(r, dict)]
    payload = api.get(path, params=params)
    extract = getattr(api, "extract_rows", None)
    rows = extract(payload) if extract else payload
    if isinstance(rows, dict):
        rows = rows.get("data") or rows.get("rows") or rows.get("items") or []
    return [r for r in rows if isinstance(r, dict)]


def _first(d: Dict[str, Any], *keys: str) -> Any:
    for k in keys:
        if k in d and d[k] not in (None, ""):
            return d[k]
    return None


def ts_api_vehicle_names(api) -> Dict[str, str]:
    """TraumaSoft vehicle id -> name from the API's vehicle list."""
    names: Dict[str, str] = {}
    try:
        vehicles = api.list_vehicles(include_deleted=True, include_disabled=True)
    except TypeError:
        vehicles = api.list_vehicles()
    except Exception as e:
        logging.warning("Could not list TraumaSoft vehicles via API: %s", e)
        return names
    for v in vehicles or []:
        if not isinstance(v, dict):
            continue
        vid = _first(v, "id", "Id", "ID", "vehicle_id", "VehicleId")
        name = _first(v, "name", "Name", "unit", "Unit", "vehicle_name", "VehicleName", "number", "Number")
        if vid is not None and name:
            names[str(vid)] = str(name)
    return names


def ts_api_cost_center_names(api) -> Dict[str, str]:
    names: Dict[str, str] = {}
    try:
        for cc in api.list_cost_centers() or []:
            if isinstance(cc, dict):
                cid = _first(cc, "id", "Id", "ID", "cost_center_id")
                name = _first(cc, "name", "Name")
                if cid is not None and name:
                    names[str(cid)] = str(name)
    except Exception as e:
        logging.warning("Could not list TraumaSoft cost centers via API: %s", e)
    return names


def enrich_rows(rows: List[Dict[str, Any]], colmap: Dict[str, Optional[str]],
                vehicle_names: Dict[str, str], cost_center_names: Dict[str, str]) -> None:
    """Fill vehicle_name / cost_center_name on API or CSV rows from id lookups."""
    vcol, ccol = colmap.get("vehicle_id"), colmap.get("cost_center_id")
    vname_col, ccname_col = colmap.get("vehicle_name"), colmap.get("cost_center_name")
    for r in rows:
        if not r.get("vehicle_name"):
            name = r.get(vname_col) if vname_col else None
            if not name and vcol and r.get(vcol) is not None:
                name = vehicle_names.get(str(r.get(vcol)))
            if name:
                r["vehicle_name"] = name
        if not r.get("cost_center_name"):
            name = r.get(ccname_col) if ccname_col else None
            if not name and ccol and r.get(ccol) is not None:
                name = cost_center_names.get(str(r.get(ccol)))
            if name:
                r["cost_center_name"] = name


def filter_rows_since(rows: List[Dict[str, Any]], colmap: Dict[str, Optional[str]], since: datetime) -> List[Dict[str, Any]]:
    """Keep rows whose updated/created/completed date is on or after `since` (rows with no parseable date are kept)."""
    date_cols = [c for c in (colmap.get("updated_at"), colmap.get("created_at"), colmap.get("completed_at")) if c]
    if not date_cols:
        return rows
    since_s = to_rfc3339(since)
    since_dt = datetime.fromisoformat(since_s) if since_s else None
    out = []
    for r in rows:
        stamps = [to_rfc3339(r.get(c)) for c in date_cols]
        stamps = [datetime.fromisoformat(x) for x in stamps if x]
        if not stamps or since_dt is None or max(stamps) >= since_dt:
            out.append(r)
    return out


def load_rows_from_api(api, since: datetime, column_overrides: Optional[Dict[str, List[str]]] = None,
                       extra_ids: Optional[List[int]] = None) -> Tuple[List[Dict[str, Any]], Dict[str, Optional[str]]]:
    logging.info("Fetching TraumaSoft work orders via API %s params=%s", TS_WO_API_PATH, TS_WO_API_PARAMS)
    rows = _api_rows(api, TS_WO_API_PATH, TS_WO_API_PARAMS)
    cols = sorted({k for r in rows for k in r.keys()})
    colmap = resolve_column_map(cols, column_overrides)
    enrich_rows(rows, colmap, ts_api_vehicle_names(api), ts_api_cost_center_names(api))
    colmap = resolve_column_map(sorted({k for r in rows for k in r.keys()}), column_overrides)
    keep_ids = {str(i) for i in (extra_ids or [])}
    idcol = colmap.get("id")
    recent = filter_rows_since(rows, colmap, since)
    if keep_ids and idcol:
        seen = {str(r.get(idcol)) for r in recent}
        recent += [r for r in rows if str(r.get(idcol)) in keep_ids and str(r.get(idcol)) not in seen]
    logging.info("TraumaSoft API returned %s work orders, %s in window", len(rows), len(recent))
    return recent, colmap


def load_rows_from_csv(path: str, since: datetime, column_overrides: Optional[Dict[str, List[str]]] = None,
                       extra_ids: Optional[List[int]] = None) -> Tuple[List[Dict[str, Any]], Dict[str, Optional[str]]]:
    import csv

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = [dict(r) for r in csv.DictReader(f)]
    for r in rows:  # treat blank cells as NULL
        for k, v in list(r.items()):
            if isinstance(v, str) and v.strip() == "":
                r[k] = None
    cols = list(rows[0].keys()) if rows else []
    colmap = resolve_column_map(cols, column_overrides)
    enrich_rows(rows, colmap, {}, {})
    colmap = resolve_column_map(sorted({k for r in rows for k in r.keys()}), column_overrides)
    keep_ids = {str(i) for i in (extra_ids or [])}
    idcol = colmap.get("id")
    recent = filter_rows_since(rows, colmap, since)
    if keep_ids and idcol:
        seen = {str(r.get(idcol)) for r in recent}
        recent += [r for r in rows if str(r.get(idcol)) in keep_ids and str(r.get(idcol)) not in seen]
    logging.info("CSV %s: %s rows, %s in window", path, len(rows), len(recent))
    return recent, colmap


def resolve_source(explicit: Optional[str] = None) -> str:
    src = (explicit or TS_WO_SOURCE or "").lower()
    if src in ("db", "api", "csv"):
        return src
    if TS_WO_CSV_FILE:
        return "csv"
    if DB_DSN:
        return "db"
    if os.getenv("TS_API_BASE_URL") or os.getenv("TS_API_KEY"):
        return "api"
    raise SystemExit("No work order source configured: set DB_DSN (db), TS_API_BASE_URL/TS_API_KEY (api) or TS_WO_CSV_FILE (csv), or TS_WO_SOURCE.")


def probe_api(api, extra_paths: Optional[List[str]] = None) -> None:
    """Try likely TraumaSoft work-order endpoints and report which ones answer."""
    paths = list(extra_paths or []) + [p for p in TS_API_WO_PATH_CANDIDATES if p not in (extra_paths or [])]
    print(f"\n== probing {len(paths)} candidate TraumaSoft endpoints ==")
    for path in paths:
        try:
            payload = api.get(path, params={"limit": 1, "page_size": 1})
        except Exception as e:  # TraumasoftAPIError carries status_code
            code = getattr(e, "status_code", None)
            print(f"  {path:45s} -> {code or type(e).__name__}: {str(e)[:120]}")
            continue
        extract = getattr(api, "extract_rows", None)
        rows = extract(payload) if extract else payload
        n = len(rows) if isinstance(rows, list) else "?"
        keys = sorted(rows[0].keys()) if isinstance(rows, list) and rows and isinstance(rows[0], dict) else []
        print(f"  {path:45s} -> OK rows={n} keys={keys[:25]}")
    print("\nSet TS_WO_API_PATH to the endpoint that returned work orders, then run --discover.")


def discover_rows(rows: List[Dict[str, Any]], colmap: Dict[str, Optional[str]], label: str) -> None:
    cols = sorted({k for r in rows for k in r.keys()})
    print(f"\n== {label}: {len(rows)} rows, {len(cols)} columns ==")
    print(", ".join(cols))
    print("\n== resolved logical fields ==")
    for k, v in colmap.items():
        print(f"  {k:17s} -> {v}")
    print("\n== first 5 rows ==")
    for r in rows[:5]:
        print(json.dumps(r, default=str, indent=2))
    for key in ("status", "status_id", "priority", "category"):
        col = colmap.get(key)
        if not col:
            continue
        counts: Dict[str, int] = {}
        for r in rows:
            counts[repr(r.get(col))] = counts.get(repr(r.get(col)), 0) + 1
        print(f"\n== distinct {key} ({col}) ==")
        for val, n in sorted(counts.items(), key=lambda kv: -kv[1])[:25]:
            print(f"  {val}: {n}")


# =============================
# SYNC
# =============================
@dataclass
class SyncResult:
    created: List[Tuple[int, str]] = field(default_factory=list)
    updated: List[Tuple[int, str]] = field(default_factory=list)
    unchanged: int = 0
    skipped_no_vehicle: List[Tuple[int, Optional[str]]] = field(default_factory=list)
    failed: List[Tuple[int, str]] = field(default_factory=list)
    dry_run: bool = False

    def summary_lines(self) -> List[str]:
        lines = [
            f"Created:   {len(self.created)}",
            f"Updated:   {len(self.updated)}",
            f"Unchanged: {self.unchanged}",
            f"Skipped (vehicle not found in Samsara): {len(self.skipped_no_vehicle)}",
            f"Failed:    {len(self.failed)}",
        ]
        if self.dry_run:
            lines.insert(0, "DRY RUN - nothing was written to Samsara")
        if self.skipped_no_vehicle:
            lines.append("")
            lines.append("Vehicles not matched in Samsara (TS WO id: vehicle):")
            for ts_id, name in self.skipped_no_vehicle[:50]:
                lines.append(f"  {ts_id}: {name}")
        if self.failed:
            lines.append("")
            lines.append("Failures (TS WO id: error):")
            for ts_id, err in self.failed[:50]:
                lines.append(f"  {ts_id}: {err}")
        return lines


def sync_records(records: List[WorkOrderRecord], client: Optional[SamsaraClient], vehicles: VehicleIndex,
                 ledger, dry_run: bool = False, match_mode: str = SAMSARA_VEHICLE_MATCH,
                 send_initial_status: bool = SAMSARA_SEND_INITIAL_STATUS) -> SyncResult:
    result = SyncResult(dry_run=dry_run)
    for rec in records:
        try:
            entry = ledger.get(rec.ts_id)
            fp = fingerprint(rec)

            if entry and entry.get("samsara_id"):
                if entry.get("fingerprint") == fp:
                    result.unchanged += 1
                    continue
                payload = build_update_payload(rec)
                logging.info("UPDATE %s -> Samsara WO %s (%s)", rec.tag, entry["samsara_id"], rec.status)
                if not dry_run:
                    client.update_work_order(entry["samsara_id"], payload)
                    ledger.put(rec.ts_id, entry["samsara_id"], fp, rec.status, entry.get("asset_id"),
                               rec.vehicle_id, rec.vehicle_name, "update")
                    ledger.save()
                result.updated.append((rec.ts_id, entry["samsara_id"]))
                continue

            asset_id = vehicles.resolve(rec, match_mode)
            if not asset_id:
                logging.warning("SKIP %s: vehicle %r (TS id %s) not found in Samsara", rec.tag, rec.vehicle_name, rec.vehicle_id)
                result.skipped_no_vehicle.append((rec.ts_id, rec.vehicle_name))
                continue

            payload = build_create_payload(rec, asset_id, send_initial_status)
            logging.info("CREATE %s -> asset %s (%s) status=%s", rec.tag, vehicles.names.get(asset_id, asset_id), asset_id, rec.status)
            if dry_run:
                result.created.append((rec.ts_id, "(dry-run)"))
                continue

            try:
                created = client.create_work_order(payload)
            except SamsaraApiError as e:
                # Older API behaviour: initial status / historical timestamps not accepted on create.
                stripped = {k: v for k, v in payload.items() if k not in ("status", "startedAtTime", "completedAtTime")}
                if e.status_code == 400 and stripped != payload:
                    logging.warning("Create with initial status rejected (%s); retrying without it and PATCHing status", e)
                    created = client.create_work_order(stripped)
                    if rec.status != "Open":
                        client.update_work_order(created["id"], build_update_payload(rec))
                else:
                    raise
            samsara_id = str(created["id"])
            ledger.put(rec.ts_id, samsara_id, fp, rec.status, asset_id, rec.vehicle_id, rec.vehicle_name, "create")
            ledger.save()
            result.created.append((rec.ts_id, samsara_id))
        except Exception as e:  # keep going; one bad row must not stop the batch
            logging.error("FAIL %s: %s", rec.tag, e, exc_info=True)
            result.failed.append((rec.ts_id, str(e)))
    return result


def rebuild_ledger(client: SamsaraClient, ledger) -> int:
    """Scan Samsara work orders for `[TS-WO:<id>]` tags and repopulate the ledger."""
    count = 0
    for wo in client.list_work_orders():
        m = TS_TAG_RE.search(wo.get("description") or "")
        if not m:
            continue
        ts_id = int(m.group(1))
        existing = ledger.get(ts_id)
        if existing and existing.get("samsara_id") == str(wo.get("id")):
            continue
        # fingerprint=None forces one re-sync PATCH on the next run so TS wins.
        ledger.put(ts_id, str(wo["id"]), None, wo.get("status") or "Open", wo.get("assetId"), action="rebuild")
        count += 1
    ledger.save()
    return count


# =============================
# EMAIL / LOGGING
# =============================
def setup_logging(output_dir: str, run_date_str: str) -> str:
    log_dir = Path(output_dir) / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = str(log_dir / f"SamsaraWorkOrderSync_{run_date_str}.log")
    root = logging.getLogger()
    for h in list(root.handlers):
        root.removeHandler(h)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    root.addHandler(fh)
    root.addHandler(sh)
    root.setLevel(logging.INFO)
    return log_path


def send_email(subject: str, body: str, recipients: List[str]) -> None:
    if not recipients:
        logging.info("No recipients configured; skipping email %r", subject)
        return
    smtp_server = os.getenv("SMTP_SERVER", "smtp.example.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASS")
    smtp_from = os.getenv("SMTP_FROM", smtp_user or "reports@example.com")
    ehlo_host = os.getenv("SMTP_EHLO_HOST")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_from
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, smtp_port, timeout=60) as server:
        if ehlo_host:
            server.ehlo(ehlo_host)
        server.starttls(context=context)
        if ehlo_host:
            server.ehlo(ehlo_host)
        if smtp_user and smtp_pass:
            server.login(smtp_user, smtp_pass)
        server.send_message(msg)
    logging.info("Email sent to %s: %s", recipients, subject)


# =============================
# DISCOVER MODE
# =============================
def discover(conn) -> None:
    cols = introspect_columns(conn, TS_WO_TABLE)
    colmap = resolve_column_map(cols, _env_json("TS_WO_COLUMN_MAP", {}))
    print(f"\n== {TS_WO_TABLE} columns ({len(cols)}) ==")
    print(", ".join(cols))
    print("\n== resolved logical fields ==")
    for k, v in colmap.items():
        print(f"  {k:15s} -> {v}")
    cur = conn.cursor()
    id_col = colmap.get("id") or "id"
    cur.execute(f"SELECT * FROM {TS_WO_TABLE} ORDER BY {id_col} DESC LIMIT 5")
    print("\n== 5 most recent rows ==")
    for r in _rows(cur):
        print(json.dumps(r, default=str, indent=2))
    for key in ("status", "status_id", "priority", "category"):
        col = colmap.get(key)
        if not col:
            continue
        cur.execute(f"SELECT {col}, COUNT(*) AS n FROM {TS_WO_TABLE} GROUP BY {col} ORDER BY n DESC LIMIT 25")
        print(f"\n== distinct {key} ({col}) ==")
        for r in _rows(cur):
            print(f"  {r[col]!r}: {r['n']}")
    cur.close()


# =============================
# MAIN
# =============================
def parse_args(argv: List[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Sync TraumaSoft work orders into Samsara.")
    p.add_argument("--dry-run", action="store_true", help="Read TraumaSoft and Samsara but write nothing to Samsara.")
    p.add_argument("--discover", action="store_true", help="Print the work order source's columns, samples and distinct statuses, then exit.")
    p.add_argument("--source", choices=["db", "api", "csv"], help="Override TS_WO_SOURCE for this run.")
    p.add_argument("--probe-api", action="store_true",
                   help="Try likely TraumaSoft API endpoints for work orders and report which respond, then exit.")
    p.add_argument("--probe-path", action="append", help="Extra endpoint path(s) to try with --probe-api (repeatable).")
    p.add_argument("--rebuild-ledger", action="store_true", help="Rebuild the ledger from [TS-WO:id] tags in Samsara, then exit.")
    p.add_argument("--init-ledger-table", action="store_true",
                   help="Create the SQL mapping table (SAMSARA_WO_LEDGER_DSN / _TABLE / _DIALECT) from sql/ if missing, then exit.")
    p.add_argument("--since", type=lambda s: datetime.strptime(s, "%Y-%m-%d"), help="Sync work orders changed since YYYY-MM-DD.")
    p.add_argument("--lookback-days", type=int, default=TS_WO_LOOKBACK_DAYS)
    p.add_argument("--limit", type=int, help="Process at most N work orders (useful for a first test).")
    p.add_argument("--ts-id", type=int, action="append", help="Sync only these TraumaSoft work order ids (repeatable).")
    p.add_argument("--no-email", action="store_true")
    return p.parse_args(argv[1:])


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv or sys.argv)
    tz = ZoneInfo(TS_TIMEZONE)
    now_local = datetime.now(tz)
    run_date_str = now_local.date().isoformat()

    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    log_path = setup_logging(OUTPUT_DIR, run_date_str)
    logging.info("=== TraumaSoft -> Samsara Work Order Sync started ===")
    logging.info("DB_DSN=%s | TEST_MODE=%s | DRY_RUN=%s | LEDGER=%s", DB_DSN, TEST_MODE, args.dry_run,
                 f"sql:{LEDGER_DSN}/{LEDGER_TABLE}" if LEDGER_DSN else LEDGER_FILE)

    if args.init_ledger_table:
        if not LEDGER_DSN:
            raise SystemExit("--init-ledger-table needs SAMSARA_WO_LEDGER_DSN (and optionally _TABLE, _DIALECT)")
        ledger = open_ledger()
        ledger.ensure_table()
        logging.info("Ledger table ready: %s (%s rows)", ledger.describe(), len(ledger.work_orders))
        return 0

    if args.probe_api:
        probe_api(open_traumasoft_api(), args.probe_path)
        return 0

    source = resolve_source(args.source)
    logging.info("Work order source: %s", source)
    column_overrides = _env_json("TS_WO_COLUMN_MAP", {})

    if args.discover:
        if source == "db":
            conn = get_db_connection()
            try:
                discover(conn)
            finally:
                conn.close()
        else:
            far_past = datetime(1970, 1, 2)
            if source == "api":
                rows, colmap = load_rows_from_api(open_traumasoft_api(), far_past, column_overrides)
                discover_rows(rows, colmap, f"TraumaSoft API {TS_WO_API_PATH}")
            else:
                rows, colmap = load_rows_from_csv(TS_WO_CSV_FILE, far_past, column_overrides)
                discover_rows(rows, colmap, f"CSV {TS_WO_CSV_FILE}")
        return 0

    client: Optional[SamsaraClient] = None
    if not args.dry_run or SAMSARA_API_TOKEN:
        client = SamsaraClient(SAMSARA_API_TOKEN, base_url=SAMSARA_BASE_URL)
    ledger = open_ledger()

    if args.rebuild_ledger:
        n = rebuild_ledger(client, ledger)
        logging.info("Ledger rebuilt: %s work orders mapped from Samsara into %s", n, ledger.describe())
        return 0

    # --- Samsara vehicles
    if client:
        vehicles = VehicleIndex(client.list_vehicles(), SAMSARA_VEHICLE_EXTERNAL_ID_KEY)
        logging.info("Samsara vehicles loaded: %s", len(vehicles.names))
    else:
        vehicles = VehicleIndex([])
        logging.warning("No SAMSARA_API_TOKEN; dry-run will report every work order as 'vehicle not found'")

    # --- TraumaSoft work orders
    since = args.since or (now_local.replace(tzinfo=None) - timedelta(days=args.lookback_days))
    extra_ids = list(args.ts_id) if args.ts_id else ledger.open_ts_ids()
    if source == "db":
        conn = get_db_connection()
        try:
            cols = introspect_columns(conn, TS_WO_TABLE)
            colmap = resolve_column_map(cols, column_overrides)
            logging.info("Column map: %s", {k: v for k, v in colmap.items() if v})
            if not colmap.get("id"):
                raise SystemExit(f"Could not find an id column on {TS_WO_TABLE}; set TS_WO_COLUMN_MAP")
            rows = fetch_work_orders(conn, colmap, since, extra_ids, args.limit, ledger.max_ts_id())
        finally:
            conn.close()
    elif source == "api":
        rows, colmap = load_rows_from_api(open_traumasoft_api(), since, column_overrides, extra_ids)
    else:
        rows, colmap = load_rows_from_csv(TS_WO_CSV_FILE, since, column_overrides, extra_ids)
    if source != "db":
        logging.info("Column map: %s", {k: v for k, v in colmap.items() if v})
        if not colmap.get("id"):
            raise SystemExit("Could not find an id column in the work order rows; set TS_WO_COLUMN_MAP")
        if args.limit:
            rows = rows[:args.limit]
    logging.info("TraumaSoft work orders fetched: %s (since %s)", len(rows), since)

    maps = {
        "status": _env_json("TS_WO_STATUS_MAP", {}),
        "priority": _env_json("TS_WO_PRIORITY_MAP", {}),
        "category": _env_json("TS_WO_CATEGORY_MAP", {}),
    }
    records: List[WorkOrderRecord] = []
    for row in rows:
        try:
            rec = row_to_record(row, colmap, maps)
        except Exception as e:
            logging.error("Could not normalise row %s: %s", row.get(colmap.get("id") or "id"), e)
            continue
        if rec.vehicle_id in TS_WO_EXCLUDE_VEHICLE_IDS:
            continue
        if args.ts_id and rec.ts_id not in args.ts_id:
            continue
        records.append(rec)

    result = sync_records(records, client, vehicles, ledger, dry_run=args.dry_run)

    lines = [
        f"TraumaSoft -> Samsara Work Order Sync - {run_date_str}",
        f"Since: {since} | TraumaSoft rows: {len(rows)} | Ledger: {ledger.describe()}",
        f"Log File: {log_path}",
        "",
    ] + result.summary_lines()
    body = "\n".join(lines)
    logging.info("\n%s", body)

    if args.no_email:
        logging.info("--no-email: skipping summary emails")
    else:
        recipients = [TEST_MODE_RECIPIENT] if TEST_MODE else PROD_RECIPIENTS
        test_suffix = " [TEST]" if TEST_MODE else ""
        subject = EMAIL_SUBJECT.format(date=run_date_str, test_suffix=test_suffix)
        if result.failed or result.skipped_no_vehicle:
            subject += " (ATTENTION)"
        email_error = None
        try:
            send_email(subject, body, recipients)
        except Exception as e:
            email_error = str(e)
            logging.error("Main email send failed: %s", e, exc_info=True)
        try:
            status_subject = f"STATUS - {subject}" + (" (EMAIL FAILED)" if email_error else "")
            send_email(status_subject, body + (f"\n\nMAIN EMAIL ERROR:\n{email_error}\n" if email_error else ""), [STATUS_EMAIL_RECIPIENT])
        except Exception as e:
            logging.error("Status email failed: %s", e, exc_info=True)

    logging.info("=== TraumaSoft -> Samsara Work Order Sync completed ===")
    return 1 if result.failed else 0


if __name__ == "__main__":
    sys.exit(main())
