"""Unit tests for the TraumaSoft -> Samsara work order sync (no DB / network needed)."""
import json
import os
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import TraumaSoft_Samsara_WorkOrder_Sync as sync  # noqa: E402
from samsara_workorders_api import SamsaraApiError, SamsaraClient  # noqa: E402


def make_row(**over):
    row = {
        "id": 123, "vehicle_id": 55, "cost_center_id": 7, "description": "Brake squeal front",
        "status": "In Progress", "priority": "High", "category": "Repair",
        "date_created": datetime(2026, 10, 1, 8, 30), "date_updated": datetime(2026, 10, 2, 9, 0),
        "date_completed": None, "odometer_started": 120345, "odometer_type": "miles",
        "vehicle_name": "M-12", "cost_center_name": "Station 4", "vehicle_odometer_type": "miles",
    }
    row.update(over)
    return row


class ColumnMapTests(unittest.TestCase):
    def test_resolves_first_candidate_case_insensitively(self):
        cm = sync.resolve_column_map(["ID", "Vehicle_Id", "complaint", "date_completed", "Status"])
        self.assertEqual(cm["id"], "ID")
        self.assertEqual(cm["vehicle_id"], "Vehicle_Id")
        self.assertEqual(cm["description"], "complaint")
        self.assertEqual(cm["completed_at"], "date_completed")
        self.assertEqual(cm["status"], "Status")
        self.assertIsNone(cm["priority"])

    def test_override_takes_precedence(self):
        cm = sync.resolve_column_map(["id", "description", "problem_text"], {"description": ["problem_text"]})
        self.assertEqual(cm["description"], "problem_text")


class MappingTests(unittest.TestCase):
    def test_status_mapping(self):
        self.assertEqual(sync.map_status("open"), "Open")
        self.assertEqual(sync.map_status("IN PROGRESS"), "In Progress")
        self.assertEqual(sync.map_status("Waiting for parts"), "Pending Parts")
        self.assertEqual(sync.map_status("canceled"), "Cancelled")
        self.assertEqual(sync.map_status("Completed"), "Completed")
        self.assertEqual(sync.map_status(None, completed_at=datetime(2026, 1, 1)), "Completed")
        self.assertEqual(sync.map_status(None), "Open")
        self.assertEqual(sync.map_status("weird-status"), "Open")
        self.assertEqual(sync.map_status(3, status_map={"3": "On Hold"}), "On Hold")

    def test_priority_mapping(self):
        self.assertEqual(sync.map_priority("normal"), "Medium")
        self.assertEqual(sync.map_priority("EMERGENCY"), "Urgent")
        self.assertEqual(sync.map_priority(None, default="Low"), "Low")
        self.assertEqual(sync.map_priority("??", default="Bogus"), "Medium")

    def test_category_mapping(self):
        self.assertEqual(sync.map_category("PM"), "Preventive")
        self.assertEqual(sync.map_category("Damage Repair"), "Damage Repair")
        self.assertIsNone(sync.map_category("unknown"))
        self.assertIsNone(sync.map_category(None))

    def test_odometer_conversion(self):
        self.assertEqual(sync.odometer_to_meters(100, "miles"), 160934)
        self.assertEqual(sync.odometer_to_meters(100, "km"), 100000)
        self.assertEqual(sync.odometer_to_meters("100.0", None), 160934)
        self.assertIsNone(sync.odometer_to_meters(0))
        self.assertIsNone(sync.odometer_to_meters("n/a"))

    def test_rfc3339_uses_traumasoft_timezone(self):
        self.assertEqual(sync.to_rfc3339(datetime(2026, 7, 4, 12, 0), "America/New_York"), "2026-07-04T12:00:00-04:00")
        self.assertEqual(sync.to_rfc3339("2026-01-15 09:30:00", "America/New_York"), "2026-01-15T09:30:00-05:00")
        self.assertIsNone(sync.to_rfc3339("0000-00-00 00:00:00"))
        self.assertIsNone(sync.to_rfc3339(None))

    def test_vehicle_name_normalisation(self):
        self.assertEqual(sync.normalize_vehicle_name("M-12"), sync.normalize_vehicle_name("m 12"))
        self.assertEqual(sync.normalize_vehicle_name("M-12"), sync.normalize_vehicle_name("M12"))


class RecordAndPayloadTests(unittest.TestCase):
    def setUp(self):
        self.colmap = sync.resolve_column_map(make_row().keys())
        self.rec = sync.row_to_record(make_row(), self.colmap)

    def test_row_to_record(self):
        r = self.rec
        self.assertEqual(r.ts_id, "123")
        self.assertEqual(r.vehicle_name, "M-12")
        self.assertEqual(r.status, "In Progress")
        self.assertEqual(r.priority, "High")
        self.assertEqual(r.category, "Corrective")
        self.assertEqual(r.odometer_meters, int(round(120345 * 1609.344)))
        self.assertEqual(r.cost_center_name, "Station 4")

    def test_completed_date_without_status_column_marks_completed(self):
        row = make_row(date_completed=datetime(2026, 10, 3, 10, 0))
        del row["status"]
        cm = sync.resolve_column_map(row.keys())
        self.assertEqual(sync.row_to_record(row, cm).status, "Completed")

    def test_description_carries_ts_tag(self):
        desc = sync.build_description(self.rec)
        self.assertTrue(desc.startswith("[TS-WO:123] Brake squeal front"))
        self.assertIn("Station: Station 4", desc)
        self.assertEqual(sync.TS_TAG_RE.search(desc).group(1), "123")

    def test_create_payload(self):
        p = sync.build_create_payload(self.rec, "asset-1")
        self.assertEqual(p["assetId"], "asset-1")
        self.assertEqual(p["priority"], "High")
        self.assertEqual(p["category"], "Corrective")
        self.assertEqual(p["status"], "In Progress")
        self.assertNotIn("completedAtTime", p)
        self.assertNotIn("dueAtTime", p)
        p2 = sync.build_create_payload(self.rec, "asset-1", send_initial_status=False)
        self.assertNotIn("status", p2)

    def test_create_payload_open_omits_status(self):
        rec = sync.row_to_record(make_row(status="open"), self.colmap)
        self.assertNotIn("status", sync.build_create_payload(rec, "a"))

    def test_update_payload_completed(self):
        rec = sync.row_to_record(make_row(status="completed", date_completed=datetime(2026, 10, 3, 10, 0)), self.colmap)
        p = sync.build_update_payload(rec)
        self.assertEqual(p["status"], "Completed")
        self.assertEqual(p["completedAtTime"], "2026-10-03T10:00:00-04:00")

    def test_fingerprint_changes_with_status(self):
        a = sync.fingerprint(self.rec)
        b = sync.fingerprint(sync.row_to_record(make_row(status="completed"), self.colmap))
        self.assertNotEqual(a, b)
        self.assertEqual(a, sync.fingerprint(sync.row_to_record(make_row(), self.colmap)))


class VehicleIndexTests(unittest.TestCase):
    vehicles = [
        {"id": "1", "name": "M-12", "vin": "1ABC", "externalIds": {"traumasoftVehicleId": "55"}},
        {"id": "2", "name": "M 13", "vin": "2DEF"},
    ]

    def test_resolve_by_name_external_vin(self):
        idx = sync.VehicleIndex(self.vehicles)
        rec = sync.row_to_record(make_row(), sync.resolve_column_map(make_row().keys()))
        self.assertEqual(idx.resolve(rec, "name"), "1")
        self.assertEqual(idx.resolve(rec, "externalId"), "1")
        self.assertEqual(idx.resolve(rec), "1")  # default vin mode falls back to name
        rec.extra["vehicle_vin"] = "2def"
        self.assertEqual(idx.resolve(rec, "vin"), "2")
        rec.vehicle_name = "M-99"
        self.assertEqual(idx.resolve(rec, "name"), "2")  # falls back to the VIN set above
        rec.extra["vehicle_vin"] = None
        rec.vehicle_id = 999
        self.assertIsNone(idx.resolve(rec, "name"))


class SyncLoopTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ledger = sync.Ledger(os.path.join(self.tmp.name, "ledger.json"))
        self.idx = sync.VehicleIndex([{"id": "A1", "name": "M-12"}])
        self.colmap = sync.resolve_column_map(make_row().keys())

    def tearDown(self):
        self.tmp.cleanup()

    def test_create_then_unchanged_then_update(self):
        client = mock.Mock()
        client.create_work_order.return_value = {"id": "WO-9"}
        rec = sync.row_to_record(make_row(), self.colmap)

        r1 = sync.sync_records([rec], client, self.idx, self.ledger)
        self.assertEqual(r1.created, [("123", "WO-9")])
        self.assertEqual(client.create_work_order.call_args[0][0]["assetId"], "A1")
        self.assertEqual(self.ledger.get("123")["samsara_id"], "WO-9")

        r2 = sync.sync_records([rec], client, self.idx, self.ledger)
        self.assertEqual(r2.unchanged, 1)
        client.update_work_order.assert_not_called()

        rec2 = sync.row_to_record(make_row(status="completed", date_completed=datetime(2026, 10, 3)), self.colmap)
        r3 = sync.sync_records([rec2], client, self.idx, self.ledger)
        self.assertEqual(r3.updated, [("123", "WO-9")])
        client.update_work_order.assert_called_once()
        self.assertEqual(client.update_work_order.call_args[0][0], "WO-9")
        self.assertEqual(client.update_work_order.call_args[0][1]["status"], "Completed")
        self.assertEqual(self.ledger.get(123)["status"], "Completed")
        self.assertEqual(self.ledger.open_ts_ids(), [])

        with open(self.ledger.path) as f:
            self.assertIn("WO-9", json.load(f)["work_orders"]["123"]["samsara_id"])

    def test_vehicle_not_found_is_skipped(self):
        client = mock.Mock()
        rec = sync.row_to_record(make_row(vehicle_name="Unknown"), self.colmap)
        r = sync.sync_records([rec], client, self.idx, self.ledger)
        self.assertEqual(r.skipped_no_vehicle, [("123", "Unknown")])
        client.create_work_order.assert_not_called()

    def test_dry_run_writes_nothing(self):
        rec = sync.row_to_record(make_row(), self.colmap)
        r = sync.sync_records([rec], None, self.idx, self.ledger, dry_run=True)
        self.assertEqual(r.created, [("123", "(dry-run)")])
        self.assertIsNone(self.ledger.get(123))

    def test_create_falls_back_when_initial_status_rejected(self):
        client = mock.Mock()
        client.create_work_order.side_effect = [SamsaraApiError(400, "status not allowed"), {"id": "WO-2"}]
        rec = sync.row_to_record(make_row(), self.colmap)
        r = sync.sync_records([rec], client, self.idx, self.ledger)
        self.assertEqual(r.created, [("123", "WO-2")])
        self.assertNotIn("status", client.create_work_order.call_args_list[1][0][0])
        client.update_work_order.assert_called_once_with("WO-2", mock.ANY)

    def test_failure_does_not_stop_batch(self):
        client = mock.Mock()
        client.create_work_order.side_effect = [SamsaraApiError(500, "boom"), {"id": "WO-3"}]
        recs = [sync.row_to_record(make_row(id=1), self.colmap), sync.row_to_record(make_row(id=2), self.colmap)]
        r = sync.sync_records(recs, client, self.idx, self.ledger)
        self.assertEqual([f[0] for f in r.failed], ["1"])
        self.assertEqual(r.created, [("2", "WO-3")])

    def test_rebuild_ledger_from_tags(self):
        client = mock.Mock()
        client.list_work_orders.return_value = [
            {"id": "WO-1", "description": "[TS-WO:10] thing", "status": "Open", "assetId": "A1"},
            {"id": "WO-2", "description": "no tag", "status": "Open"},
        ]
        n = sync.rebuild_ledger(client, self.ledger)
        self.assertEqual(n, 1)
        self.assertEqual(self.ledger.get(10)["samsara_id"], "WO-1")
        self.assertIsNone(self.ledger.get(10)["fingerprint"])


class FakeResponse:
    def __init__(self, status, body=None, headers=None):
        self.status_code = status
        self._body = body
        self.headers = headers or {}
        self.content = b"x" if body is not None else b""
        self.text = json.dumps(body) if body is not None else ""

    def json(self):
        return self._body


class SamsaraClientTests(unittest.TestCase):
    def _client(self, responses):
        session = mock.Mock()
        session.headers = {}
        session.request.side_effect = responses
        c = SamsaraClient("tok", session=session, write_min_interval_s=0)
        return c, session

    def test_pagination(self):
        c, s = self._client([
            FakeResponse(200, {"data": [{"id": "1"}], "pagination": {"hasNextPage": True, "endCursor": "c1"}}),
            FakeResponse(200, {"data": [{"id": "2"}], "pagination": {"hasNextPage": False, "endCursor": ""}}),
        ])
        self.assertEqual([v["id"] for v in c.list_vehicles()], ["1", "2"])
        self.assertEqual(s.request.call_args_list[1][1]["params"]["after"], "c1")
        self.assertEqual(s.headers["Authorization"], "Bearer tok")

    @mock.patch("samsara_workorders_api.time.sleep")
    def test_retry_on_429_then_create(self, _sleep):
        c, s = self._client([
            FakeResponse(429, {"message": "slow down"}, {"Retry-After": "1"}),
            FakeResponse(200, {"data": {"id": "WO-1", "status": "Open"}}),
        ])
        wo = c.create_work_order({"assetId": "A", "description": "d", "priority": None})
        self.assertEqual(wo["id"], "WO-1")
        call = s.request.call_args_list[1]
        self.assertEqual(call[0][0], "POST")
        self.assertTrue(call[0][1].endswith("/maintenance/work-orders"))
        self.assertEqual(call[1]["json"], {"assetId": "A", "description": "d"})  # None dropped

    def test_update_puts_id_in_body_and_raises_on_4xx(self):
        c, s = self._client([FakeResponse(200, {"data": {"id": "WO-1"}}), FakeResponse(404, {"message": "nope"})])
        c.update_work_order("WO-1", {"status": "Closed"})
        self.assertEqual(s.request.call_args[0][0], "PATCH")
        self.assertEqual(s.request.call_args[1]["json"], {"status": "Closed", "id": "WO-1"})
        with self.assertRaises(SamsaraApiError) as ctx:
            c.update_work_order("WO-1", {"status": "Closed"})
        self.assertEqual(ctx.exception.status_code, 404)


if __name__ == "__main__":
    unittest.main()


class SqlLedgerTests(unittest.TestCase):
    """Exercises the SQL 'endpoint table' ledger against an in-memory SQLite DB."""

    def setUp(self):
        import sqlite3
        self.conn = sqlite3.connect(":memory:")
        self.ledger = sync.SqlLedger(self.conn, "samsara_workorder_map", "sqlite")

    def tearDown(self):
        self.conn.close()

    def test_missing_table_reads_as_empty_then_ensure_table_creates_it(self):
        self.assertEqual(self.ledger.work_orders, {})
        self.ledger.ensure_table()
        cols = [r[1] for r in self.conn.execute("PRAGMA table_info(samsara_workorder_map)")]
        self.assertEqual(tuple(cols), sync.LEDGER_COLUMNS)
        self.ledger.ensure_table()  # idempotent

    def test_put_inserts_then_updates_and_reloads(self):
        self.ledger.ensure_table()
        self.ledger.put(123, "WO-9", "fp1", "Open", "A1", 55, "M-12", "create")
        self.ledger.save()
        self.ledger.put(123, "WO-9", "fp2", "Completed", "A1", 55, "M-12", "update")
        self.ledger.put(124, "WO-10", None, "In Progress", "A2", 56, "M-13", "create")
        self.ledger.save()

        rows = self.conn.execute("SELECT ts_wo_id, samsara_wo_id, samsara_status, fingerprint, last_action, vehicle_name "
                                 "FROM samsara_workorder_map ORDER BY ts_wo_id").fetchall()
        self.assertEqual(rows, [("123", "WO-9", "Completed", "fp2", "update", "M-12"),
                                ("124", "WO-10", "In Progress", None, "create", "M-13")])

        fresh = sync.SqlLedger(self.conn, "samsara_workorder_map", "sqlite")
        self.assertEqual(fresh.get(123)["samsara_id"], "WO-9")
        self.assertEqual(fresh.get(123)["status"], "Completed")
        self.assertEqual(fresh.open_ts_ids(), ["124"])
        self.assertEqual(fresh.max_ts_id(), 124)
        self.assertEqual(fresh.describe(), "sql:sqlite:samsara_workorder_map")

    def test_sync_loop_with_sql_ledger(self):
        self.ledger.ensure_table()
        client = mock.Mock()
        client.create_work_order.return_value = {"id": "WO-1"}
        idx = sync.VehicleIndex([{"id": "A1", "name": "M-12"}])
        colmap = sync.resolve_column_map(make_row().keys())
        rec = sync.row_to_record(make_row(), colmap)

        r1 = sync.sync_records([rec], client, idx, self.ledger)
        self.assertEqual(r1.created, [("123", "WO-1")])
        r2 = sync.sync_records([rec], client, idx, self.ledger)
        self.assertEqual(r2.unchanged, 1)
        row = self.conn.execute("SELECT samsara_wo_id, samsara_asset_id, ts_vehicle_id, last_action "
                                "FROM samsara_workorder_map WHERE ts_wo_id=123").fetchone()
        self.assertEqual(row, ("WO-1", "A1", 55, "create"))


class FakeTraumasoftAPI:
    """Stand-in for the project's traumasoft_api.TraumasoftAPI."""
    ROWS = [
        {"Id": 501, "VehicleId": 55, "CostCenterId": 7, "Description": "Check engine light", "Status": "Open",
         "Priority": "High", "DateCreated": "2026-10-10 08:00:00", "DateUpdated": "2026-10-11 09:00:00", "DateCompleted": None},
        {"Id": 400, "VehicleId": 56, "CostCenterId": 7, "Description": "Old PM", "Status": "Completed",
         "Priority": "Normal", "DateCreated": "2026-01-01 08:00:00", "DateUpdated": "2026-01-03 09:00:00", "DateCompleted": "2026-01-03 09:00:00"},
    ]

    def __init__(self):
        self.calls = []

    def paginate(self, path, params=None, page_size=100, max_pages=1000):
        self.calls.append((path, params))
        if path != "ThirdParty/Data/WorkOrders":
            raise RuntimeError("404")
        return list(self.ROWS)

    def get(self, path, params=None, legacy_hmac=False):
        if path == "ThirdParty/Data/WorkOrders":
            return {"data": list(self.ROWS)}
        raise RuntimeError(f"404 {path}")

    @staticmethod
    def extract_rows(payload):
        return payload.get("data", [])

    def list_vehicles(self, include_deleted=False, include_disabled=False, fields=None):
        return [{"Id": 55, "Name": "M-12"}, {"Id": 56, "Name": "M-13"}]

    def list_cost_centers(self):
        return [{"Id": 7, "Name": "Station 4"}]


class ApiAndCsvSourceTests(unittest.TestCase):
    def test_api_source_resolves_columns_enriches_and_filters_by_since(self):
        api = FakeTraumasoftAPI()
        rows, colmap = sync.load_rows_from_api(api, datetime(2026, 10, 1))
        self.assertEqual(colmap["id"], "Id")
        self.assertEqual(colmap["vehicle_id"], "VehicleId")
        self.assertEqual(colmap["description"], "Description")
        self.assertEqual(colmap["status"], "Status")
        self.assertEqual(colmap["updated_at"], "DateUpdated")
        self.assertEqual([r["Id"] for r in rows], [501])
        self.assertEqual(rows[0]["vehicle_name"], "M-12")
        self.assertEqual(rows[0]["cost_center_name"], "Station 4")
        rec = sync.row_to_record(rows[0], colmap)
        self.assertEqual((rec.ts_id, rec.vehicle_name, rec.status, rec.priority), ("501", "M-12", "Open", "High"))

    def test_api_source_keeps_open_ledger_ids_outside_window(self):
        rows, _ = sync.load_rows_from_api(FakeTraumasoftAPI(), datetime(2026, 10, 1), extra_ids=[400])
        self.assertEqual(sorted(r["Id"] for r in rows), [400, 501])

    def test_csv_source(self):
        tmp = tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False, newline="")
        tmp.write("id,vehicle,station,description,status,date_updated\n"
                  "9,M-12,Station 4,Flat tire,open,2026-10-12 10:00\n"
                  "8,M-13,,Old thing,closed,2026-01-01 10:00\n")
        tmp.close()
        try:
            rows, colmap = sync.load_rows_from_csv(tmp.name, datetime(2026, 10, 1))
            self.assertEqual(colmap["vehicle_name"], "vehicle_name")  # enriched from the "vehicle" column
            self.assertEqual(colmap["cost_center_name"], "cost_center_name")  # enriched from "station"
            self.assertEqual([r["id"] for r in rows], ["9"])
            self.assertEqual(rows[0]["vehicle_name"], "M-12")
            rec = sync.row_to_record(rows[0], colmap)
            self.assertEqual((rec.ts_id, rec.vehicle_name, rec.cost_center_name, rec.status), ("9", "M-12", "Station 4", "Open"))
        finally:
            os.unlink(tmp.name)

    def test_resolve_source(self):
        self.assertEqual(sync.resolve_source("api"), "api")
        with mock.patch.object(sync, "TS_WO_CSV_FILE", "x.csv"), mock.patch.object(sync, "TS_WO_SOURCE", ""):
            self.assertEqual(sync.resolve_source(None), "csv")
        with mock.patch.object(sync, "TS_WO_CSV_FILE", ""), mock.patch.object(sync, "DB_DSN", ""), \
             mock.patch.object(sync, "TS_WO_SOURCE", ""), mock.patch.dict(os.environ, {"TS_API_BASE_URL": "https://x"}):
            self.assertEqual(sync.resolve_source(None), "vehicles")
        with mock.patch.object(sync, "TS_WO_CSV_FILE", ""), mock.patch.object(sync, "DB_DSN", ""), \
             mock.patch.object(sync, "TS_WO_SOURCE", ""), mock.patch.dict(os.environ, {"TS_API_BASE_URL": "", "TS_API_KEY": ""}):
            with self.assertRaises(SystemExit):
                sync.resolve_source(None)

    def test_probe_api_reports_hits_and_misses(self):
        import io, contextlib
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            sync.probe_api(FakeTraumasoftAPI(), ["ThirdParty/Data/Nope"])
        out = buf.getvalue()
        self.assertIn("ThirdParty/Data/Nope", out)
        self.assertIn("ThirdParty/Data/WorkOrders", out)
        self.assertIn("OK rows=2", out)


class FakeFleetAPI:
    """TraumaSoft client stand-in for the vehicles source (Fleet/Vehicles + Fleet/CustomStatus)."""
    def __init__(self):
        self.vehicles = [
            {"id": 55, "name": "M-12", "vehicle_status": "In Service", "vin": "1ABC", "odometer": 120000, "disabled": False, "deleted": False},
            {"id": 56, "name": "M-13", "vehicle_status": "Out of Service", "vin": "2DEF", "odometer": 88000, "disabled": False, "deleted": False},
            {"id": 57, "name": "M-14", "vehicle_status": "Out of Service", "vin": "3GHI", "odometer": 5, "disabled": False, "deleted": True},
        ]
    def list_vehicles(self, include_deleted=False, include_disabled=False, fields=None):
        return list(self.vehicles)
    def list_custom_statuses(self, include_deleted=False):
        return {"custom_statuses": [{"id": 1, "name": "In Service", "notify_group": "vehicles"},
                                    {"id": 2, "name": "Out of Service", "notify_group": "vehicles"},
                                    {"id": 3, "name": "In Shop", "notify_group": "vehicles"}]}


class VehicleEpisodeSourceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.state_path = os.path.join(self.tmp.name, "state.json")
        self.api = FakeFleetAPI()

    def tearDown(self):
        self.tmp.cleanup()

    def test_is_oos_status(self):
        self.assertTrue(sync.is_oos_status("Out of Service"))
        self.assertTrue(sync.is_oos_status("In Shop - Mechanical"))
        self.assertFalse(sync.is_oos_status("In Service"))
        self.assertTrue(sync.is_oos_status("Down", oos_statuses=["down"]))
        self.assertFalse(sync.is_oos_status("Out of Service", oos_statuses=["In Shop"]))

    def test_episode_lifecycle_open_then_completed(self):
        t0 = datetime(2026, 10, 14, 9, 0)
        state = sync.VehicleStatusState(self.state_path)
        rows, colmap = sync.load_rows_from_vehicles(self.api, state, now=t0)
        self.assertEqual([r["id"] for r in rows], ["V56-202610140900"])  # deleted M-14 ignored, M-12 in service
        self.assertEqual(rows[0]["status"], "Open")
        self.assertEqual(rows[0]["vehicle_vin"], "2DEF")
        rec = sync.row_to_record(rows[0], colmap)
        self.assertEqual((rec.ts_id, rec.vehicle_name, rec.status, rec.category), ("V56-202610140900", "M-13", "Open", "Corrective"))
        self.assertEqual(rec.odometer_meters, int(round(88000 * 1609.344)))
        self.assertTrue(os.path.exists(self.state_path))

        # same status next run: same episode id, still open, status history unchanged
        state = sync.VehicleStatusState(self.state_path)
        rows, _ = sync.load_rows_from_vehicles(self.api, state, now=t0 + timedelta(hours=1))
        self.assertEqual([r["id"] for r in rows], ["V56-202610140900"])

        # moves to another OOS status: still the same episode, history grows
        self.api.vehicles[1]["vehicle_status"] = "In Shop"
        rows, _ = sync.load_rows_from_vehicles(self.api, state, now=t0 + timedelta(hours=2))
        self.assertIn("Out of Service -> In Shop", rows[0]["description"])

        # back in service: episode completed, emitted once more with Completed
        self.api.vehicles[1]["vehicle_status"] = "In Service"
        t3 = t0 + timedelta(days=1)
        rows, colmap = sync.load_rows_from_vehicles(self.api, state, now=t3)
        self.assertEqual([(r["id"], r["status"]) for r in rows], [("V56-202610140900", "Completed")])
        rec = sync.row_to_record(rows[0], colmap)
        self.assertEqual(rec.status, "Completed")
        self.assertEqual(rec.completed_at.replace(tzinfo=None), t3)

        # OOS again later => a new episode; the old one is still emitted within retention
        self.api.vehicles[1]["vehicle_status"] = "Out of Service"
        rows, _ = sync.load_rows_from_vehicles(self.api, state, now=t3 + timedelta(days=2))
        self.assertEqual(sorted(r["id"] for r in rows), ["V56-202610140900", "V56-202610170900"])
        # ...and dropped after retention
        rows, _ = sync.load_rows_from_vehicles(self.api, state, now=t3 + timedelta(days=40))
        self.assertEqual([r["id"] for r in rows], ["V56-202610170900"])

    def test_dry_run_does_not_persist_state(self):
        state = sync.VehicleStatusState(self.state_path)
        sync.load_rows_from_vehicles(self.api, state, now=datetime(2026, 10, 14, 9, 0), persist=False)
        self.assertFalse(os.path.exists(self.state_path))

    def test_end_to_end_episode_to_samsara(self):
        state = sync.VehicleStatusState(self.state_path)
        rows, colmap = sync.load_rows_from_vehicles(self.api, state, now=datetime(2026, 10, 14, 9, 0))
        recs = [sync.row_to_record(r, colmap) for r in rows]
        idx = sync.VehicleIndex([{"id": "A2", "name": "Medic 13", "vin": "2def"}])  # name differs, VIN matches
        client = mock.Mock(); client.create_work_order.return_value = {"id": "WO-77"}
        ledger = sync.Ledger(os.path.join(self.tmp.name, "ledger.json"))
        res = sync.sync_records(recs, client, idx, ledger)
        self.assertEqual(res.created, [("V56-202610140900", "WO-77")])
        payload = client.create_work_order.call_args[0][0]
        self.assertEqual(payload["assetId"], "A2")
        self.assertTrue(payload["description"].startswith("[TS-WO:V56-202610140900] Vehicle M-13 marked 'Out of Service'"))
        self.assertNotIn("status", payload)
        self.assertEqual(sync.TS_TAG_RE.search(payload["description"]).group(1), "V56-202610140900")

    def test_resolve_source_prefers_vehicles_for_api_credentials(self):
        with mock.patch.object(sync, "TS_WO_CSV_FILE", ""), mock.patch.object(sync, "DB_DSN", ""), \
             mock.patch.object(sync, "TS_WO_SOURCE", ""), mock.patch.dict(os.environ, {"TS_API_KEY": "k"}):
            self.assertEqual(sync.resolve_source(None), "vehicles")
        self.assertEqual(sync.resolve_source("api"), "api")
