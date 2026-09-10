"""
Checks over the punch-quality probe and the Paycor push.

Plain asserts and no pytest, matching test_region_monthly.py: these have to be
runnable on the reporting machine, which has neither.

No live calls. The Paycor client is exercised against its own guards and a
stubbed transport; the mapping, refusal and reconciliation logic is exercised
against fixtures shaped like the real feeds.

    python test_paycor_punches.py
"""

import os
import sys
from datetime import datetime, timedelta

# The Paycor client reads its environment gate at import time, so the sandbox
# default has to be established before the module is loaded.
os.environ.pop("PAYCOR_ENVIRONMENT", None)
os.environ.setdefault("PAYCOR_SUBSCRIPTION_KEY", "test-subscription-key")
os.environ.setdefault("PAYCOR_ACCESS_TOKEN", "test-access-token")
os.environ.setdefault("PAYCOR_LEGAL_ENTITY_ID", "LE-1")

import paycor_api
from paycor_api import PaycorClient, PaycorReadOnlyError, describe_write_contract

import probe_punch_quality as PQ
import push_paycor_timecards as PUSH

PASSED = 0
FAILED = []


def check(label, condition, detail=""):
    global PASSED
    if condition:
        PASSED += 1
    else:
        FAILED.append(f"{label}{(' -- ' + detail) if detail else ''}")


def section(title):
    print(f"\n--- {title} ---")


# =============================
# FIXTURES
# =============================
D = datetime(2026, 9, 9)


def punch(start, end=None, deleted=False, pid=1):
    """A ShiftPunch as the feed returns it: naive UTC strings."""
    return {
        "id": pid,
        "start_time": start.strftime("%Y-%m-%d %H:%M:%S"),
        "end_time": end.strftime("%Y-%m-%d %H:%M:%S") if end else None,
        "deleted": deleted,
    }


def shift(name, user_id, start, end, punches=(), deleted=False, vehicle="A-101"):
    return {
        "id": user_id * 100,
        "user_id": user_id,
        "shift_name": name,
        "vehicle_name": vehicle,
        "license_level": "Paramedic",
        "start_time": start.strftime("%Y-%m-%d %H:%M:%S"),
        "end_time": end.strftime("%Y-%m-%d %H:%M:%S"),
        "deleted": deleted,
        "punches": list(punches),
    }


def employee(user_id, num, first="Pat", last="Crew"):
    return {
        "user_id": user_id,
        "employee_num": num,
        "first_name": first,
        "last_name": last,
        "cost_center_name": "Columbus",
    }


ZERO = timedelta(0)


# =============================
# 1. PUNCH CLASSIFICATION
# =============================
section("punch classification")

# One shift that has ended, one still running, at a "now" between them.
now = D.replace(hour=20)
shifts = [
    # finished shift, punch closed properly
    shift("OH-A-CIN-06-18", 1, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D.replace(hour=18), pid=11)]),
    # finished shift, punch left open -> missed punch-out
    shift("OH-A-CIN-06-18", 2, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), None, pid=12)]),
    # shift still running, punch open -> not a problem
    shift("OH-A-CIN-19-07", 3, D.replace(hour=19), D.replace(hour=23),
          [punch(D.replace(hour=19), None, pid=13)]),
    # deleted punch is ignored entirely
    shift("OH-A-CIN-06-18", 4, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D.replace(hour=18), deleted=True, pid=14)]),
]

rows = PQ.collect_punches(shifts, ZERO, now)
check("deleted punches are dropped", len(rows) == 3, f"got {len(rows)}")
by_id = {r["start"]: r for r in rows}
missed = [r for r in rows if r["missed_punch_out"]]
running = [r for r in rows if r["open"] and not r["missed_punch_out"]]
closed = [r for r in rows if not r["open"]]
check("one closed punch", len(closed) == 1, f"got {len(closed)}")
check("one missed punch-out", len(missed) == 1, f"got {len(missed)}")
check("one still-running punch", len(running) == 1, f"got {len(running)}")
check("a running shift's open punch is not a missed punch-out",
      running[0]["profile"] == "OH-A-CIN-19-07")

summary = PQ.report_completion(rows, now)
check("completion scores only finished shifts",
      summary["completion_pct"] == 50.0,
      f"got {summary['completion_pct']} (1 closed of 2 finished)")
check("still-running punches excluded from completion",
      summary["still_running"] == 1)

# A deleted shift row contributes nothing.
rows_deleted = PQ.collect_punches(
    [shift("X", 9, D.replace(hour=6), D.replace(hour=18),
           [punch(D.replace(hour=6), D.replace(hour=18))], deleted=True)],
    ZERO, now,
)
check("deleted shift rows are dropped", rows_deleted == [])


# =============================
# 2. SUSPECT AND STALE PUNCHES
# =============================
section("suspect and stale punches")

short = PQ.collect_punches(
    [shift("Y", 5, D.replace(hour=6), D.replace(hour=18),
           [punch(D.replace(hour=6), D.replace(hour=6, minute=1))])],
    ZERO, now,
)
check("a one-minute punch is measured as such", abs(short[0]["minutes"] - 1.0) < 1e-6)
short_summary = PQ.report_completion(short, now)
check("a sub-threshold punch is flagged suspect", short_summary["suspect_short"] == 1)

# An open punch well past the shift end is stale; one just past it is not.
stale_rows = PQ.collect_punches(
    [shift("Z", 6, D.replace(hour=6), D.replace(hour=18), [punch(D.replace(hour=6))])],
    ZERO, D.replace(hour=23),
)
check("an open punch 5h past shift end is stale", stale_rows[0]["stale"] is True)
fresh_rows = PQ.collect_punches(
    [shift("Z", 6, D.replace(hour=6), D.replace(hour=18), [punch(D.replace(hour=6))])],
    ZERO, D.replace(hour=19),
)
check("an open punch 1h past shift end is not yet stale",
      fresh_rows[0]["stale"] is False and fresh_rows[0]["missed_punch_out"] is True)


# =============================
# 3. FABRICATED HOURS
# =============================
section("fabricated hours")

# Two crew on one unit for a 12h shift. Both punch in; one never punches out.
# The unit needs 2 crew, so worked hours are the time BOTH were on the clock.
fab_shifts = [
    shift("OH-A-CIN-06-18", 1, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D.replace(hour=18), pid=21)]),
    shift("OH-A-CIN-06-18", 2, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), None, pid=22)]),
]
after = D.replace(hour=23)
result = PQ.report_fabricated_hours(fab_shifts, ZERO, after, [D.date()])
check("as-shipped credits the full shift via the fallback",
      abs(result["worked_hours_as_shipped"] - 12.0) < 0.01,
      f"got {result['worked_hours_as_shipped']}")
check("punched-only credits nothing without a second closed punch",
      abs(result["worked_hours_measured"]) < 0.01,
      f"got {result['worked_hours_measured']}")
check("the whole denominator is fabricated in this case",
      abs(result["fabricated_pct"] - 100.0) < 0.01,
      f"got {result['fabricated_pct']}")

# Both punch out properly -> nothing fabricated.
clean_shifts = [
    shift("OH-A-CIN-06-18", 1, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D.replace(hour=18), pid=31)]),
    shift("OH-A-CIN-06-18", 2, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D.replace(hour=18), pid=32)]),
]
clean = PQ.report_fabricated_hours(clean_shifts, ZERO, after, [D.date()])
check("a fully punched day fabricates nothing",
      abs(clean["fabricated_hours"]) < 0.01, f"got {clean['fabricated_hours']}")
check("and its measured hours equal its shipped hours",
      abs(clean["worked_hours_measured"] - clean["worked_hours_as_shipped"]) < 0.01)


# =============================
# 4. PAYCOR READINESS
# =============================
section("paycor join readiness")

emps = [employee(1, "1001"), employee(2, "1002"), employee(3, "")]
ready = PQ.report_paycor_readiness(shifts, emps)
check("counts crew who actually punched", ready["punching_crew"] == 3,
      f"got {ready['punching_crew']}")
check("employee 3 has no employee_num", ready["missing_employee_num"] == 1)

dupes = PQ.report_paycor_readiness(
    shifts, [employee(1, "1001"), employee(2, "1001"), employee(3, "1003")]
)
check("a duplicate employee_num is caught", dupes["duplicate_employee_num"] == 1)


# =============================
# 5. EMPLOYEE MAPPING
# =============================
section("employee mapping")

paycor_roster = [
    {"employeeId": "px-1", "employeeNumber": "1001"},
    {"employeeId": "px-2", "employeeNumber": "1002"},
    {"employeeId": "px-3", "employeeNumber": "1002"},   # a collision
]
index = PUSH.paycor_employee_index(paycor_roster)

got, how = PUSH.resolve_employee(employee(1, "1001"), {}, index)
check("a unique employee_num resolves", got == "px-1" and how == "employee_num")

got, how = PUSH.resolve_employee(employee(2, "1002"), {}, index)
check("an ambiguous employee_num is refused", got is None, f"got {got}")
check("and says how many it matched", "2 Paycor employees" in how, how)

got, how = PUSH.resolve_employee(employee(9, "9999"), {}, index)
check("an employee_num matching nothing is refused", got is None)

got, how = PUSH.resolve_employee(employee(3, ""), {}, index)
check("no employee_num at all is refused", got is None)
check("and says why", "no employee_num" in how, how)

got, how = PUSH.resolve_employee(employee(2, "1002"), {"1002": "px-chosen"}, index)
check("an override beats an ambiguous match", got == "px-chosen" and how == "override")

got, how = PUSH.resolve_employee(employee(3, ""), {"user_id:3": "px-byuser"}, index)
check("an override can key on user_id", got == "px-byuser")

got, how = PUSH.resolve_employee(employee(1, "1001"), {}, {})
check("with no roster loaded the number is carried through", got == "1001")
check("and the plan says it is unchecked", "unchecked" in how, how)


# =============================
# 6. PUSH REFUSALS
# =============================
section("push refusals")

push_shifts = [
    shift("A", 1, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D.replace(hour=18), pid=41)]),      # fine
    shift("A", 2, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), pid=42)]),                          # open
    shift("A", 3, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=12), D.replace(hour=6), pid=43)]),      # backwards
    shift("A", 4, D.replace(hour=6), D.replace(hour=18),
          [punch(D.replace(hour=6), D + timedelta(days=3), pid=44)]),   # runaway
]
emps4 = [employee(i, f"100{i}") for i in (1, 2, 3, 4)]
idx4 = PUSH.paycor_employee_index(
    [{"employeeId": f"px-{i}", "employeeNumber": f"100{i}"} for i in (1, 2, 3, 4)]
)
sendable, refused = PUSH.build_plan(push_shifts, emps4, ZERO, after, {}, idx4)

check("only the sound punch is sendable", len(sendable) == 1, f"got {len(sendable)}")
check("the other three are refused", len(refused) == 3, f"got {len(refused)}")
reasons = " | ".join(r["refused"] for r in refused)
check("an open punch is refused", "open" in reasons, reasons)
check("a backwards punch is refused", "not after" in reasons, reasons)
check("a runaway punch is refused", "ceiling" in reasons, reasons)
check("refusals keep the person's name for the report",
      all("employee_name" in r for r in refused))

# Nothing sendable is ever open -- the rail that matters most.
check("no sendable punch lacks a clock-out",
      all(r["punch_out"] is not None for r in sendable))

# --date scopes the plan.
other_day = shift("A", 1, D + timedelta(days=1), D + timedelta(days=1, hours=12),
                  [punch(D + timedelta(days=1), D + timedelta(days=1, hours=12), pid=51)])
scoped, _ = PUSH.build_plan(
    push_shifts + [other_day], emps4, ZERO, after, {}, idx4, target_date=D.date()
)
check("--date keeps only that day's punches",
      all(r["punch_in"].date() == D.date() for r in scoped))


# =============================
# 7. RECONCILIATION
# =============================
section("reconciliation")

check("a Z-suffixed timestamp parses",
      PUSH.parse_paycor_time("2026-09-09T06:00:00Z") == D.replace(hour=6))
check("a fractional timestamp parses",
      PUSH.parse_paycor_time("2026-09-09T06:00:00.123") == D.replace(hour=6))
check("an offset timestamp parses to naive local",
      PUSH.parse_paycor_time("2026-09-09T06:00:00-04:00") == D.replace(hour=6))
check("an empty timestamp is None", PUSH.parse_paycor_time("") is None)

paycor_punches = [
    {"employeeId": "px-1", "punchInTime": "2026-09-09T06:00:00",
     "punchOutTime": "2026-09-09T18:00:00"},
]
pindex = PUSH.index_paycor_punches(paycor_punches)
row = dict(sendable[0])
row["paycor_employee_id"] = "px-1"
present, existing = PUSH.already_present(row, pindex)
check("an identical punch is recognised as present", present is True)

# Inside the tolerance still counts as the same punch.
near = PUSH.index_paycor_punches([
    {"employeeId": "px-1", "punchInTime": "2026-09-09T06:01:00",
     "punchOutTime": "2026-09-09T18:00:00"}
])
check("a punch one minute out is the same punch",
      PUSH.already_present(row, near)[0] is True)

far = PUSH.index_paycor_punches([
    {"employeeId": "px-1", "punchInTime": "2026-09-09T09:00:00",
     "punchOutTime": "2026-09-09T18:00:00"}
])
check("a punch three hours out is a different punch",
      PUSH.already_present(row, far)[0] is False)

recon = PUSH.report_reconciliation([row], pindex)
check("an agreeing punch reconciles as matched", recon["matched"] == 1)
check("and is not reported missing", recon["missing"] == 0)

drift = PUSH.index_paycor_punches([
    {"employeeId": "px-1", "punchInTime": "2026-09-09T06:00:00",
     "punchOutTime": "2026-09-09T17:00:00"}
])
recon_drift = PUSH.report_reconciliation([dict(row)], drift)
check("a clock-out an hour out is reported as drifted", recon_drift["drifted"] == 1)
check("and not silently counted as agreement", recon_drift["matched"] == 0)

recon_missing = PUSH.report_reconciliation([dict(row)], {})
check("an empty Paycor window reports nothing matched",
      recon_missing["matched"] == 0)


# =============================
# 8. CLIENT GUARDS
# =============================
section("paycor client guards")

check("the client defaults to the sandbox",
      paycor_api.SANDBOX_BASE_URL in PaycorClient().base_url,
      PaycorClient().base_url)
check("and is not in production mode by default",
      paycor_api.IS_PRODUCTION is False)

client = PaycorClient()
check("a default client is read-only", client.read_only is True)

raised = None
try:
    client.create_punch("px-1", D.replace(hour=6), D.replace(hour=18))
except PaycorReadOnlyError as exc:
    raised = exc
check("a read-only client refuses to write a punch", raised is not None)
check("and names the constructor argument that would allow it",
      raised is not None and "read_only=False" in str(raised))

# The guard is on `write`, not on the HTTP verb, so a POST-shaped read is not
# silently blocked and a GET-shaped write is not silently allowed.
raised_get = None
try:
    client.request("GET", "v1/anything", write=True)
except PaycorReadOnlyError as exc:
    raised_get = exc
check("write=True is refused even on a GET", raised_get is not None)

contract = describe_write_contract()
check("the write contract reports the sandbox", contract["environment"] == "sandbox")
check("and admits it is unverified", contract["verified"] is False)
check("and names every field it would send",
      set(contract["body_fields"]) == {"employee", "punch_in", "punch_out"})

body = client.build_punch_body("px-9", D.replace(hour=6), D.replace(hour=18))
check("the body names the employee", body["employeeId"] == "px-9")
check("the body carries a formatted clock-in",
      body["punchInTime"] == "2026-09-09T06:00:00", body.get("punchInTime"))
check("the body carries a formatted clock-out",
      body["punchOutTime"] == "2026-09-09T18:00:00")

open_body = client.build_punch_body("px-9", D.replace(hour=6), None)
check("an open punch body omits the clock-out entirely",
      "punchOutTime" not in open_body)

# A writable client is possible, and pagination handles both key spellings.
writable = PaycorClient(read_only=False)
check("a writable client can be constructed", writable.read_only is False)

pages = [
    {"records": [{"a": 1}], "continuationToken": "t1", "hasMoreResults": True},
    {"results": [{"a": 2}], "hasMoreResults": False},
]
calls = []


def fake_request(method, path, params=None, json_body=None, write=False):
    calls.append(dict(params or {}))
    return pages[len(calls) - 1]


writable.request = fake_request
walked = list(writable.paginate("v1/whatever"))
check("pagination walks both 'records' and 'results'", len(walked) == 2, str(walked))
check("and passes the continuation token on the second call",
      calls[1].get("continuationToken") == "t1", str(calls))


# =============================
# 9. OVERRIDES FILE
# =============================
section("overrides file")

example = os.path.join("state", "paycor_employee_overrides.example.json")
check("the example overrides file exists", os.path.exists(example))
if os.path.exists(example):
    import json
    with open(example, encoding="utf-8") as handle:
        data = json.load(handle)
    check("it is valid JSON with an overrides block", "overrides" in data)
    check("and explains itself", "note" in data)


# =============================
print("\n" + "=" * 60)
if FAILED:
    print(f"FAILED {len(FAILED)} of {PASSED + len(FAILED)} checks:")
    for failure in FAILED:
        print(f"  - {failure}")
    sys.exit(1)
print(f"All {PASSED} checks passed.")
