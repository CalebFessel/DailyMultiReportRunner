"""
Report data layer backed by the Traumasoft ThirdParty API.

Produces the same DataFrames the SQL queries in `Script` produce, so the Excel,
append/retention, email, and status_logger layers stay untouched. Each builder
here replaces one `pd.read_sql_query(...)` call.

What the live probes established, and what it forces:

  * Trips backfill. GetTrips returned data 90 days back, so OTP and the UHU
    numerator keep full history and `--date` backfill still works for them.

  * Shifts do not. /Schedule/Shifts ignores every filter and always returns the
    same rolling window anchored to now (751 rows, today-1..today+2, identical
    for requests at -30, +0 and +30 days). Anything derived from scheduled
    hours or crew therefore describes only the present. Historical staffing and
    the UHU denominator cannot be reconstructed after the fact -- they have to
    be captured daily as they pass.

  * Cost center is absent from trips and vehicles. It reaches a trip only via
    shift_name -> the crew on that shift -> employee.cost_center_name. Since
    shifts are present-only, that mapping is accumulated to disk daily and
    reused for historical legs (see CostCenterMap).

  * Vehicle status history does not exist. `oos_since` is likewise accumulated
    locally (see OutOfServiceHistory).
"""

import os
import json
import logging
from pathlib import Path
from collections import Counter, defaultdict
from datetime import datetime, timedelta, date, time

import pandas as pd

log = logging.getLogger(__name__)

# =============================
# CONFIG
# =============================
# State that has to survive between runs because the API cannot reproduce it.
STATE_DIR = os.getenv("TS_STATE_DIR", "state")
COST_CENTER_MAP_FILE = os.path.join(STATE_DIR, "shift_cost_center_map.json")
OOS_HISTORY_FILE = os.path.join(STATE_DIR, "vehicle_oos_history.json")

# Trip timestamp names this tenant emits, in preference order, for the moment a
# unit reached the patient. The current SQL scores against ePCR field 549,
# which this API does not expose; `at_scene` is the closest CAD equivalent and
# the bedside variant is closer still where it is recorded.
ARRIVAL_TIMESTAMP_KEYS = [
    k.strip() for k in os.getenv(
        "TS_ARRIVAL_TIMESTAMP_KEYS", "at_scene: At Patient Bedside,at_scene"
    ).split(",") if k.strip()
]

# Minutes either side of scheduled pickup that still count as on time.
OTP_ON_TIME_WINDOW_MINUTES = int(os.getenv("OTP_ON_TIME_WINDOW_MINUTES", "10"))

OTP_EXCLUDED_COST_CENTERS = ("CPR",)

# Fleet status strings, which replace the status id + status_reason LIKE
# filtering the SQL used. Observed on this tenant: In Service,
# New - Waiting for Delivery, Out of Service, Out of Service - Collision,
# Retired, Waiting for Inspection.
IN_SERVICE_STATUSES = {"In Service"}
OUT_OF_SERVICE_STATUSES = {"Out of Service", "Out of Service - Collision"}
NON_FLEET_STATUSES = {"Retired", "New - Waiting for Delivery", "Waiting for Inspection"}

# Staffing: cost centers excluded by name, and the minimum crew for a row to
# count as a staffed unit (the SQL used HAVING COUNT(DISTINCT user_id) > 2).
STAFFING_EXCLUDED_COST_CENTER_PATTERNS = ("dispatch", "cpr", "training", "admin")
STAFFING_MIN_CREW = int(os.getenv("STAFFING_MIN_CREW", "3"))

# The SQL excluded certification templates 101-104. The API exposes a
# license_level string instead of those ids, so the exclusion is by name and
# must be confirmed against Lists/User/LicenseLevels before trusting counts.
STAFFING_EXCLUDED_LICENSE_LEVELS = {
    lv.strip().lower()
    for lv in os.getenv("STAFFING_EXCLUDED_LICENSE_LEVELS", "").split(",")
    if lv.strip()
}

# Which span of a trip counts as utilized time for UHU. "task" is enroute ->
# clear (the whole committed period); "loaded" is transporting -> at_destination
# (patient on board only). The SQL used an estimated duration, so neither is a
# like-for-like match -- see docs/API_MIGRATION.md.
UHU_SPAN = os.getenv("UHU_SPAN", "task")
UHU_SPANS = {
    "task": ("enroute", "clear"),
    "loaded": ("transporting", "at_destination"),
}

# The UHU SQL excluded schedules named like Dispatch / Comm / Call, which are
# communications rosters rather than transport units and would otherwise sink
# every ratio with scheduled hours that can never be utilized.
UHU_EXCLUDED_PROFILE_PATTERNS = ("dispatch", "comm", "call")

# Trip statuses that are not runs. The SQL restricted loaded hours to
# last_status_id IN (1..7); these are the equivalent status strings observed on
# this tenant for legs that never ran.
UHU_EXCLUDED_TRIP_STATUSES = {"canceled", "cancelled", "disregard", "no transport"}

UNASSIGNED_COST_CENTER = "No Cost Center Assigned"


# =============================
# PARSING HELPERS
# =============================
def parse_ts(value):
    """
    Parse the timestamp shapes this API returns.

    Shifts come back as '2026-08-17T09:00' (no zone); trip timestamps are
    ISO-8601 with an offset. Offsets are dropped to naive local time so the
    two can be compared, which matches how the SQL treated these columns.
    """
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
        else:
            return None
    if parsed.tzinfo is not None:
        parsed = parsed.replace(tzinfo=None)
    return parsed


def timestamp_map(leg):
    """Flatten a leg's `timestamps` array of {name: iso} maps into one dict."""
    flat = {}
    for entry in leg.get("timestamps") or []:
        if isinstance(entry, dict):
            for name, value in entry.items():
                if value:
                    flat[name] = value
    return flat


def arrival_time(leg, keys=None):
    """The first configured arrival stamp present on this leg."""
    stamps = timestamp_map(leg)
    for key in (keys or ARRIVAL_TIMESTAMP_KEYS):
        if stamps.get(key):
            parsed = parse_ts(stamps[key])
            if parsed:
                return parsed
    return None


def span_minutes(leg, start_key, end_key):
    """Minutes between two named stamps on a leg, or None if either is absent."""
    stamps = timestamp_map(leg)
    start = parse_ts(stamps.get(start_key))
    end = parse_ts(stamps.get(end_key))
    if not start or not end or end <= start:
        return None
    return (end - start).total_seconds() / 60.0


def overlap_minutes(start, end, window_start, window_end):
    """Minutes of [start, end) that fall inside [window_start, window_end)."""
    if not start or not end:
        return 0.0
    latest_start = max(start, window_start)
    earliest_end = min(end, window_end)
    if earliest_end <= latest_start:
        return 0.0
    return (earliest_end - latest_start).total_seconds() / 60.0


# =============================
# ACCUMULATED STATE
# =============================
class CostCenterMap:
    """
    shift_name -> cost center, accumulated across runs.

    Cost center is not on trips or vehicles, and the only route to it --
    shift_name -> crew -> employee.cost_center_name -- is available solely for
    the rolling shift window. Historical legs would otherwise be
    unattributable, so every run folds what it can see into a file on disk.

    Crew counts per (shift_name, cost_center) are kept rather than a single
    winner, because a profile picks up a second cost center whenever someone
    from a neighbouring station fills in. The dominant cost center wins and the
    contested ones stay visible via `ambiguous()`.
    """

    def __init__(self, path=COST_CENTER_MAP_FILE):
        self.path = path
        self.counts = defaultdict(Counter)
        self._load()

    def _load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as handle:
                stored = json.load(handle)
        except (FileNotFoundError, ValueError):
            return
        for shift_name, centers in (stored.get("counts") or {}).items():
            self.counts[shift_name] = Counter(centers)

    def save(self):
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "updated": datetime.now().isoformat(timespec="seconds"),
            "counts": {name: dict(counter) for name, counter in self.counts.items()},
        }
        with open(self.path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, sort_keys=True)

    def update(self, shifts, employees):
        """Fold today's visible shift window into the accumulated mapping."""
        employee_cc = {
            str(emp.get("user_id")): emp.get("cost_center_name")
            for emp in employees
            if emp.get("user_id") and emp.get("cost_center_name")
        }
        learned = 0
        for shift in shifts:
            name = shift.get("shift_name")
            cost_center = employee_cc.get(str(shift.get("user_id")))
            if name and cost_center:
                self.counts[name][cost_center] += 1
                learned += 1
        log.info(
            "Cost-center map: folded in %s crew rows, %s profiles known",
            learned,
            len(self.counts),
        )
        return learned

    def resolve(self, shift_name):
        """Dominant cost center for a shift profile, or None if unknown."""
        if not shift_name:
            return None
        counter = self.counts.get(shift_name)
        if not counter:
            return None
        # most_common breaks ties by insertion order; sort for determinism.
        return sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]

    def ambiguous(self):
        """Profiles seen against more than one cost center."""
        return {
            name: dict(counter)
            for name, counter in self.counts.items()
            if len(counter) > 1
        }


class OutOfServiceHistory:
    """
    vehicle id -> the first date it was seen out of service, accumulated daily.

    The API exposes only a vehicle's current status; there is no status log and
    no work-order endpoint, so `oos_since` and days-out have to be observed
    rather than queried. A vehicle returning to service clears its entry, so
    the next spell starts a fresh clock.
    """

    def __init__(self, path=OOS_HISTORY_FILE):
        self.path = path
        self.since = {}
        self._load()

    def _load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as handle:
                stored = json.load(handle)
        except (FileNotFoundError, ValueError):
            return
        self.since = dict(stored.get("since") or {})

    def save(self):
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "updated": datetime.now().isoformat(timespec="seconds"),
            "since": self.since,
        }
        with open(self.path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, sort_keys=True)

    def update(self, vehicles, run_date):
        """Record newly out-of-service vehicles; clear ones that came back."""
        run_date_str = run_date.isoformat() if hasattr(run_date, "isoformat") else str(run_date)
        for vehicle in vehicles:
            key = str(vehicle.get("id"))
            status = vehicle.get("vehicle_status")
            if status in OUT_OF_SERVICE_STATUSES:
                self.since.setdefault(key, run_date_str)
            elif key in self.since:
                del self.since[key]

    def days_out(self, vehicle_id, as_of):
        since = self.since.get(str(vehicle_id))
        if not since:
            return None, None
        since_date = parse_ts(since)
        if not since_date:
            return None, None
        as_of_date = as_of if isinstance(as_of, date) else parse_ts(as_of).date()
        return since, (as_of_date - since_date.date()).days


# =============================
# ON-TIME PERFORMANCE
# =============================
def score_leg(leg, arrival_keys=None, window=OTP_ON_TIME_WINDOW_MINUTES):
    """
    Classify one leg as Early / On Time / Late / Missing Data.

    Mirrors the SQL's CASE: the delta is arrival minus scheduled pickup, and
    anything inside +/- `window` minutes is on time.
    """
    pickup = parse_ts(leg.get("pickup_time"))
    arrived = arrival_time(leg, arrival_keys)
    if not pickup or not arrived:
        return "Missing Data", None
    delta = (arrived - pickup).total_seconds() / 60.0
    if -window <= delta <= window:
        return "On Time", delta
    if delta < -window:
        return "Early", delta
    return "Late", delta


def scored_legs(legs, cost_center_map, arrival_keys=None):
    """
    Score every leg that OTP can actually judge.

    Only legs carrying both a scheduled pickup and an arrival stamp are
    scorable. On this tenant that is exactly the population where shift_name is
    always present, so cost-center attribution is complete for the rows that
    survive this filter; cancellations and disregards, which supply the
    unattributable legs, are dropped here as the SQL dropped them via
    last_status_id.
    """
    rows = []
    for leg in legs:
        status, delta = score_leg(leg, arrival_keys)
        if status == "Missing Data":
            continue
        shift_name = leg.get("shift_name")
        rows.append(
            {
                "leg_id": leg.get("leg_id"),
                "run_number": leg.get("run_number"),
                "shift_name": shift_name,
                "call_type": leg.get("call_type") or "Unknown Call Type",
                "cost_center": cost_center_map.resolve(shift_name) or UNASSIGNED_COST_CENTER,
                "status": status,
                "delta_minutes": round(delta, 2) if delta is not None else None,
            }
        )
    return pd.DataFrame(rows)


def _otp_aggregate(scored, group_cols):
    if scored.empty:
        return pd.DataFrame(
            columns=group_cols
            + ["total_runs", "early_runs", "on_time_runs", "late_runs", "on_time_percentage"]
        )

    keep = ~scored["cost_center"].isin(OTP_EXCLUDED_COST_CENTERS)
    scored = scored[keep]

    grouped = (
        scored.assign(
            early_runs=(scored["status"] == "Early").astype(int),
            on_time_runs=(scored["status"] == "On Time").astype(int),
            late_runs=(scored["status"] == "Late").astype(int),
        )
        .groupby(group_cols, dropna=False)
        .agg(
            total_runs=("status", "size"),
            early_runs=("early_runs", "sum"),
            on_time_runs=("on_time_runs", "sum"),
            late_runs=("late_runs", "sum"),
        )
        .reset_index()
    )
    grouped["on_time_percentage"] = (
        100.0 * (grouped["on_time_runs"] + grouped["early_runs"]) / grouped["total_runs"]
    ).round(2)
    return grouped


def build_otp_by_call_type(scored):
    df = _otp_aggregate(scored, ["cost_center", "call_type"])
    return df.sort_values(["cost_center", "on_time_percentage"], ascending=[True, False])


def build_otp_by_cost_center(scored):
    df = _otp_aggregate(scored, ["cost_center"])
    return df.sort_values("on_time_percentage", ascending=False)


# =============================
# STAFFING
# =============================
def _excluded_cost_center(name):
    if not name:
        return False
    lowered = name.lower()
    return any(pattern in lowered for pattern in STAFFING_EXCLUDED_COST_CENTER_PATTERNS)


def _excluded_uhu_profile(name):
    if not name:
        return False
    lowered = name.lower()
    return any(pattern in lowered for pattern in UHU_EXCLUDED_PROFILE_PATTERNS)


def _staffing_rows(shifts, employees, predicate):
    """
    Collapse crew-level shift rows into one row per staffed unit.

    The API returns one row per crew member per shift, which is the shape the
    SQL grouped over. Rows are keyed on the shift profile and its start/end so
    two crews on the same profile at different times stay separate.
    """
    by_user = {str(emp.get("user_id")): emp for emp in employees if emp.get("user_id")}
    units = defaultdict(
        lambda: {"crew": {}, "cost_centers": Counter(), "start": None, "end": None}
    )

    for shift in shifts:
        start = parse_ts(shift.get("start_time"))
        end = parse_ts(shift.get("end_time"))
        if not start or not end or not predicate(start, end):
            continue
        if shift.get("deleted"):
            continue

        level = (shift.get("license_level") or "").strip().lower()
        if level and level in STAFFING_EXCLUDED_LICENSE_LEVELS:
            continue

        user_id = str(shift.get("user_id") or "")
        if not user_id:
            continue

        key = (shift.get("shift_name"), start, end)
        unit = units[key]
        unit["start"] = start
        unit["end"] = end

        employee = by_user.get(user_id, {})
        first = employee.get("first_name") or ""
        last = employee.get("last_name") or ""
        label = f"{first} {last} (ID {user_id})".strip()
        unit["crew"][user_id] = label
        if employee.get("cost_center_name"):
            unit["cost_centers"][employee["cost_center_name"]] += 1

    rows = []
    for (shift_name, start, end), unit in units.items():
        if len(unit["crew"]) < STAFFING_MIN_CREW:
            continue
        cost_center = None
        if unit["cost_centers"]:
            cost_center = sorted(
                unit["cost_centers"].items(), key=lambda kv: (-kv[1], kv[0])
            )[0][0]
        if _excluded_cost_center(cost_center):
            continue
        rows.append(
            {
                "shift_profile": shift_name,
                "cost_center": cost_center or UNASSIGNED_COST_CENTER,
                "start_time": start,
                "end_time": end,
                "crew_count": len(unit["crew"]),
                "crew_members": "\n".join(sorted(unit["crew"].values())),
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(
            columns=[
                "shift_profile", "cost_center", "start_time",
                "end_time", "crew_count", "crew_members",
            ]
        )
    return df.sort_values(["cost_center", "shift_profile", "start_time"])


def build_staffing_active_now(shifts, employees, as_of):
    """Units staffed at a moment in time."""
    return _staffing_rows(shifts, employees, lambda s, e: s <= as_of <= e)


def build_staffing_for_date(shifts, employees, target_date):
    """Units whose shift starts on a given calendar day."""
    return _staffing_rows(shifts, employees, lambda s, e: s.date() == target_date)


# =============================
# VEHICLES
# =============================
def _fleet_partition(vehicles):
    in_service, out_of_service = [], []
    for vehicle in vehicles:
        status = vehicle.get("vehicle_status")
        if status in NON_FLEET_STATUSES:
            continue
        if status in IN_SERVICE_STATUSES:
            in_service.append(vehicle)
        elif status in OUT_OF_SERVICE_STATUSES:
            out_of_service.append(vehicle)
    return in_service, out_of_service


def used_vehicle_ids(legs):
    """Vehicles that actually ran a leg on the day."""
    return {str(leg.get("vehicle_id")) for leg in legs if leg.get("vehicle_id")}


def build_vehicle_summary(vehicles, legs, run_date):
    in_service, out_of_service = _fleet_partition(vehicles)
    used = used_vehicle_ids(legs)
    in_service_ids = {str(v.get("id")) for v in in_service}
    used_in_service = in_service_ids & used

    rows = [
        {"metric": "Total Fleet (excluding retired/pending)", "value": len(in_service) + len(out_of_service)},
        {"metric": "In Service", "value": len(in_service)},
        {"metric": "Out Of Service", "value": len(out_of_service)},
        {"metric": "Vehicles Used", "value": len(used)},
        {"metric": "In-Service Vehicles Used", "value": len(used_in_service)},
        {"metric": "Unused In-Service Vehicles", "value": len(in_service_ids - used)},
        {"metric": "Report Date", "value": str(run_date)},
    ]
    return pd.DataFrame(rows)


def build_vehicles_in_use(vehicles, legs):
    used = used_vehicle_ids(legs)
    leg_counts = Counter(str(leg.get("vehicle_id")) for leg in legs if leg.get("vehicle_id"))
    rows = [
        {
            "vehicle_id": v.get("id"),
            "vehicle_name": v.get("name"),
            "vehicle_status": v.get("vehicle_status"),
            "shift_name": v.get("shift_name"),
            "legs_run": leg_counts.get(str(v.get("id")), 0),
        }
        for v in vehicles
        if str(v.get("id")) in used
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


def build_vehicles_all_in_service(vehicles):
    in_service, _ = _fleet_partition(vehicles)
    rows = [
        {
            "vehicle_id": v.get("id"),
            "vehicle_name": v.get("name"),
            "vehicle_status": v.get("vehicle_status"),
            "odometer": v.get("odometer"),
            "shift_name": v.get("shift_name"),
        }
        for v in in_service
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


def build_vehicles_unused_in_service(vehicles, legs):
    in_service, _ = _fleet_partition(vehicles)
    used = used_vehicle_ids(legs)
    rows = [
        {
            "vehicle_id": v.get("id"),
            "vehicle_name": v.get("name"),
            "vehicle_status": v.get("vehicle_status"),
            "odometer": v.get("odometer"),
        }
        for v in in_service
        if str(v.get("id")) not in used
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


def build_vehicles_out_of_service(vehicles, oos_history, run_date):
    """
    Out-of-service sheet.

    Four columns the SQL produced have no API equivalent: status_reason, the
    work-order odometers, and the work-order station. `oos_since` and days-out
    come from locally accumulated observations instead of the status log, so
    they are blank for any vehicle first seen out of service before this
    history began.
    """
    _, out_of_service = _fleet_partition(vehicles)
    rows = []
    for vehicle in out_of_service:
        since, days = oos_history.days_out(vehicle.get("id"), run_date)
        rows.append(
            {
                "vehicle_id": vehicle.get("id"),
                "vehicle_name": vehicle.get("name"),
                "vehicle_status": vehicle.get("vehicle_status"),
                "current_odometer": vehicle.get("odometer"),
                "oos_since": since,
                "total_days_out_of_service": days,
            }
        )
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


# =============================
# UNIT-HOUR UTILIZATION
# =============================
def build_uhu(shifts, legs, cost_center_map, window_start, window_end, span=UHU_SPAN):
    """
    Scheduled vs utilized hours per shift profile.

    Scheduled hours are the portion of each shift overlapping the window,
    counted once per unit rather than once per crew member. Utilized hours come
    from real trip timestamps -- a better basis than the estimated durations the
    SQL summed, but not the same number.

    Only meaningful for a window inside the rolling shift feed. For an older
    date the scheduled side is empty and the ratio is meaningless.
    """
    start_key, end_key = UHU_SPANS.get(span, UHU_SPANS["task"])

    scheduled = defaultdict(float)
    seen_units = defaultdict(set)
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = shift.get("shift_name")
        if not name or _excluded_uhu_profile(name):
            continue
        start = parse_ts(shift.get("start_time"))
        end = parse_ts(shift.get("end_time"))
        # One unit's scheduled hours, not one per crew member on it.
        unit_key = (name, start, end)
        if unit_key in seen_units[name]:
            continue
        seen_units[name].add(unit_key)
        scheduled[name] += overlap_minutes(start, end, window_start, window_end) / 60.0

    utilized = defaultdict(float)
    runs = Counter()
    for leg in legs:
        name = leg.get("shift_name")
        if not name or _excluded_uhu_profile(name):
            continue
        # A cancelled or disregarded leg was never a run and must not dilute
        # hours-per-run, even though it still carries a shift assignment.
        status = (leg.get("trip_status") or "").strip().lower()
        if status in UHU_EXCLUDED_TRIP_STATUSES:
            continue
        runs[name] += 1
        minutes = span_minutes(leg, start_key, end_key)
        if minutes:
            utilized[name] += minutes / 60.0

    rows = []
    for name in sorted(set(scheduled) | set(utilized)):
        scheduled_hours = round(scheduled.get(name, 0.0), 2)
        utilized_hours = round(utilized.get(name, 0.0), 2)
        total_runs = runs.get(name, 0)
        rows.append(
            {
                "cost_center_name": cost_center_map.resolve(name) or UNASSIGNED_COST_CENTER,
                "shift_profile_name": name,
                "scheduled_hours": scheduled_hours,
                "utilized_hours": utilized_hours,
                "total_runs": total_runs,
                "hours_per_run": round(utilized_hours / total_runs, 3) if total_runs else 0,
                "uhu_ratio": round(utilized_hours / scheduled_hours, 3) if scheduled_hours else 0,
            }
        )
    return pd.DataFrame(rows)


def build_uhu_by_shift_profile(shifts, legs, cost_center_map, window_start, window_end, span=UHU_SPAN):
    df = build_uhu(shifts, legs, cost_center_map, window_start, window_end, span)
    if df.empty:
        return df
    return df.sort_values("uhu_ratio", ascending=False)


def build_uhu_by_cost_center(shifts, legs, cost_center_map, window_start, window_end, span=UHU_SPAN):
    df = build_uhu(shifts, legs, cost_center_map, window_start, window_end, span)
    if df.empty:
        return pd.DataFrame(
            columns=[
                "cost_center_name", "scheduled_hours", "utilized_hours",
                "total_runs", "hours_per_run", "uhu_ratio",
            ]
        )
    grouped = (
        df.groupby("cost_center_name", dropna=False)
        .agg(
            scheduled_hours=("scheduled_hours", "sum"),
            utilized_hours=("utilized_hours", "sum"),
            total_runs=("total_runs", "sum"),
        )
        .reset_index()
    )
    grouped["hours_per_run"] = grouped.apply(
        lambda r: round(r["utilized_hours"] / r["total_runs"], 3) if r["total_runs"] else 0,
        axis=1,
    )
    grouped["uhu_ratio"] = grouped.apply(
        lambda r: round(r["utilized_hours"] / r["scheduled_hours"], 3) if r["scheduled_hours"] else 0,
        axis=1,
    )
    grouped["scheduled_hours"] = grouped["scheduled_hours"].round(2)
    grouped["utilized_hours"] = grouped["utilized_hours"].round(2)
    return grouped.sort_values("uhu_ratio", ascending=False)


# =============================
# ORCHESTRATION
# =============================
def fetch_day(api, metrics_date):
    """
    One pull of everything a day's reports need.

    Trips are fetched for the target date and honour backfill. Shifts,
    employees and vehicles describe the present regardless of the date asked
    for, which is why the accumulated state files exist.
    """
    if isinstance(metrics_date, str):
        metrics_date = datetime.strptime(metrics_date, "%Y-%m-%d").date()

    log.info("Fetching trips for %s ...", metrics_date)
    legs = api.get_trips(metrics_date, range_days=1)

    log.info("Fetching shifts (rolling window, not %s) ...", metrics_date)
    shifts = api.list_shifts()

    log.info("Fetching employees ...")
    employees = api.list_employees()

    log.info("Fetching vehicles ...")
    vehicles = api.list_vehicles()

    log.info(
        "Fetched %s legs, %s shift rows, %s employees, %s vehicles",
        len(legs), len(shifts), len(employees), len(vehicles),
    )
    return {
        "metrics_date": metrics_date,
        "legs": legs,
        "shifts": shifts,
        "employees": employees,
        "vehicles": vehicles,
    }


def build_all(data, cost_center_map=None, oos_history=None, now=None):
    """Build every report DataFrame from one fetch, refreshing accumulated state."""
    metrics_date = data["metrics_date"]
    legs = data["legs"]
    shifts = data["shifts"]
    employees = data["employees"]
    vehicles = data["vehicles"]
    now = now or datetime.now()

    cost_center_map = cost_center_map or CostCenterMap()
    oos_history = oos_history or OutOfServiceHistory()

    # Learn from the present before attributing the past.
    cost_center_map.update(shifts, employees)
    oos_history.update(vehicles, metrics_date)

    scored = scored_legs(legs, cost_center_map)
    window_start = datetime.combine(metrics_date, time(0, 0))
    window_end = window_start + timedelta(days=1)

    reports = {
        "otp_by_call_type": build_otp_by_call_type(scored),
        "otp_by_cost_center": build_otp_by_cost_center(scored),
        "staffing_active_now": build_staffing_active_now(shifts, employees, now),
        "staffing_tomorrow": build_staffing_for_date(
            shifts, employees, now.date() + timedelta(days=1)
        ),
        "vehicle_summary": build_vehicle_summary(vehicles, legs, metrics_date),
        "vehicles_in_use": build_vehicles_in_use(vehicles, legs),
        "vehicles_unused_in_service": build_vehicles_unused_in_service(vehicles, legs),
        "vehicles_all_in_service": build_vehicles_all_in_service(vehicles),
        "vehicles_out_of_service": build_vehicles_out_of_service(
            vehicles, oos_history, metrics_date
        ),
        "uhu_by_cost_center": build_uhu_by_cost_center(
            shifts, legs, cost_center_map, window_start, window_end
        ),
        "uhu_by_shift_profile": build_uhu_by_shift_profile(
            shifts, legs, cost_center_map, window_start, window_end
        ),
    }

    cost_center_map.save()
    oos_history.save()
    return reports
