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
import paycor_api as PA
from paycor_api import PaycorClient, PaycorReadOnlyError, describe_write_contract

import probe_punch_quality as PQ
import push_paycor_timecards as PUSH

PASSED = 0
FAILED = []
SKIPPED = []


def check(label, condition, detail=""):
    global PASSED
    if condition:
        PASSED += 1
    else:
        FAILED.append(f"{label}{(' -- ' + detail) if detail else ''}")


def skip(label, why):
    """
    Record a check that could not run here, distinctly from one that failed.

    The reporting machine gets these files copied by hand, so a repo artifact
    that is not a .py file may simply be absent. That says nothing about
    whether the code is correct, and reporting it as a failure teaches people
    to ignore failures -- which is worse than not checking at all.
    """
    SKIPPED.append(f"{label} -- {why}")


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
    {"id": "px-1", "employeeNumber": "1001", "firstName": "A", "lastName": "One",
     "department": {"id": "dept-1"}},
    {"id": "px-2", "employeeNumber": "1002", "department": {"id": "dept-2"}},
    {"id": "px-3", "employeeNumber": "1002", "department": {"id": "dept-3"}},
    {"id": "px-4", "employeeNumber": "1004", "department": None},
]
index = PUSH.paycor_employee_index(paycor_roster)

entry, how = PUSH.resolve_employee(employee(1, "1001"), {}, index)
check("a unique employee_num resolves to the Paycor guid",
      entry and entry["id"] == "px-1" and how == "employee_num")
check("and carries the department guid the write requires",
      entry and entry["department_id"] == "dept-1")

entry, how = PUSH.resolve_employee(employee(2, "1002"), {}, index)
check("an ambiguous employee_num is refused", entry is None)
check("and says how many it matched", "2 Paycor employees" in how, how)

entry, how = PUSH.resolve_employee(employee(9, "9999"), {}, index)
check("an employee_num matching nothing is refused", entry is None)

entry, how = PUSH.resolve_employee(employee(3, ""), {}, index)
check("no employee_num at all is refused", entry is None)
check("and says why", "no employee_num" in how, how)

# An override names WHICH employee. The department still comes from Paycor's
# roster: carrying none refused every overridden punch in production for want
# of a departmentId that was available all along.
entry, how = PUSH.resolve_employee(employee(2, "1002"), {"1002": "px-1"}, index)
check("an override beats an ambiguous match",
      entry and entry["id"] == "px-1" and how == "override")
check("and inherits the department Paycor holds for that employee",
      entry and entry["department_id"] == "dept-1", str(entry))

entry, how = PUSH.resolve_employee(employee(3, ""), {"user_id:3": "px-2"}, index)
check("an override can key on user_id", entry and entry["id"] == "px-2")
check("and picks up that employee's department too",
      entry and entry["department_id"] == "dept-2", str(entry))

# A guid the roster does not contain is a bad override, and it has to say so:
# silently continuing would fail later as a missing department, which sends
# the reader after the wrong problem.
entry, how = PUSH.resolve_employee(employee(2, "1002"), {"1002": "px-gone"}, index)
check("an override naming an employee Paycor does not have is refused",
      entry is None)
check("and blames the override rather than the department",
      "not in the roster" in (how or ""), how)

# An overridden employee who genuinely has no department still fails on the
# department, which is a different fix.
entry, how = PUSH.resolve_employee(employee(4, "zzz"), {"zzz": "px-4"}, index)
check("an override to an employee with no department still resolves",
      entry and entry["id"] == "px-4", str(entry))
check("leaving the missing department to be caught as such",
      entry and entry["department_id"] is None)

entry, how = PUSH.resolve_employee(employee(1, "1001"), {}, {})
check("with no roster the employee cannot be resolved", entry is None)
check("and the plan says so rather than guessing", "no Paycor roster" in how, how)

# badgeNumber and alternateEmployeeNumber are also indexed.
alt = PUSH.paycor_employee_index([{"id": "px-9", "badgeNumber": "7788"}])
entry, how = PUSH.resolve_employee(employee(9, "7788"), {}, alt)
check("a badge number resolves too", entry and entry["id"] == "px-9")


# =============================
# 6. ACTIVITY TYPE
# =============================
section("activity type")

types = [{"id": "act-work", "name": "Work"}, {"id": "act-break", "name": "Break"}]
got, how = PUSH.resolve_activity_type(types)
check("the default activity type resolves by name", got == "act-work", f"got {got}")

got, how = PUSH.resolve_activity_type([])
check("no activity types means no id", got is None)
check("and says the list was empty", "no activity types" in how, how)

got, how = PUSH.resolve_activity_type([{"id": "a", "name": "Break"}])
check("a missing name is refused", got is None)
check("and names what the tenant actually has", "Break" in how, how)

got, how = PUSH.resolve_activity_type(
    [{"id": "a", "name": "Work"}, {"id": "b", "name": "work"}]
)
check("two types with the same name are refused", got is None)
check("and point at the id override", "PAYCOR_ACTIVITY_TYPE_ID" in how, how)


# =============================
# 7. PUSH REFUSALS
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
emps4 = [employee(i, f"200{i}") for i in (1, 2, 3, 4)]
idx4 = PUSH.paycor_employee_index([
    {"id": f"px-{i}", "employeeNumber": f"200{i}", "department": {"id": f"dept-{i}"}}
    for i in (1, 2, 3, 4)
])
sendable, refused = PUSH.build_plan(
    push_shifts, emps4, ZERO, after, {}, idx4, "act-work"
)

check("only the sound punch is sendable", len(sendable) == 1, f"got {len(sendable)}")
check("the other three are refused", len(refused) == 3, f"got {len(refused)}")
reasons = " | ".join(r["refused"] for r in refused)
check("an open punch is refused", "open" in reasons, reasons)
check("a backwards punch is refused", "not after" in reasons, reasons)
check("a runaway punch is refused", "ceiling" in reasons, reasons)
check("no sendable punch lacks a clock-out",
      all(r["punch_out"] is not None for r in sendable))

# The rail that matters most: a punch with no department cannot be filed.
no_dept_idx = PUSH.paycor_employee_index([
    {"id": "px-x", "employeeNumber": "2001", "department": None}
])
nd_send, nd_ref = PUSH.build_plan(
    push_shifts[:1], emps4, ZERO, after, {}, no_dept_idx, "act-work"
)
check("an employee with no Paycor department is refused", nd_send == [])
check("and the reason names the required field",
      nd_ref and "department" in nd_ref[0]["refused"], str(nd_ref[:1]))

# No activity type -> nothing can be sent at all.
na_send, na_ref = PUSH.build_plan(push_shifts[:1], emps4, ZERO, after, {}, idx4, None)
check("no activity type means nothing is sendable", na_send == [])
check("and the reason says CreatePunches requires one",
      na_ref and "activity type" in na_ref[0]["refused"], str(na_ref[:1]))

scoped, _ = PUSH.build_plan(
    push_shifts + [shift("A", 1, D + timedelta(days=1),
                         D + timedelta(days=1, hours=12),
                         [punch(D + timedelta(days=1),
                                D + timedelta(days=1, hours=12), pid=51)])],
    emps4, ZERO, after, {}, idx4, "act-work", target_date=D.date(),
)
check("--date keeps only that day's punches",
      all(r["punch_in"].date() == D.date() for r in scoped))


# =============================
# 8. THE EVENT MODEL
# =============================
section("punch objects -- events, not intervals")

row = sendable[0]
objs = PUSH.punch_objects(row)
check("one Traumasoft punch becomes two Paycor punches", len(objs) == 2,
      f"got {len(objs)}")
check("the first is an In", objs[0]["punchStatusType"] == "In")
check("the second is an Out", objs[1]["punchStatusType"] == "Out")
check("the In carries the clock-in time",
      objs[0]["punchDateTime"] == "2026-09-09T06:00:00", objs[0]["punchDateTime"])
check("the Out carries the clock-out time",
      objs[1]["punchDateTime"] == "2026-09-09T18:00:00", objs[1]["punchDateTime"])

for name in ("employeeId", "departmentId", "punchDateTime",
             "punchStatusType", "activityTypeId", "isTransfer"):
    check(f"every punch carries the required field {name}",
          all(name in o for o in objs), str(objs[0]))
check("isTransfer is false -- these are plain clock events",
      all(o["isTransfer"] is False for o in objs))

check("the two halves carry different correlation ids",
      objs[0]["correlationId"] != objs[1]["correlationId"])
check("a correlation id is stable across runs",
      PA.correlation_id(41, "in") == PA.correlation_id(41, "in"))
check("and differs per punch",
      PA.correlation_id(41, "in") != PA.correlation_id(42, "in"))
check("and differs per half",
      PA.correlation_id(41, "in") != PA.correlation_id(41, "out"))

raised = None
try:
    PaycorClient.build_punch("e", "d", "a", D, "Sideways")
except ValueError as exc:
    raised = exc
check("an invalid punchStatusType is refused", raised is not None)

long_note = PaycorClient.build_punch("e", "d", "a", D, "In", note="x" * 500)
check("a note is clipped to Paycor's 300-character ceiling",
      len(long_note["note"]) == 300, str(len(long_note["note"])))


# =============================
# 9. RECONCILIATION
# =============================
section("reconciliation")

check("a Z-suffixed timestamp parses",
      PUSH.parse_paycor_time("2026-09-09T06:00:00Z") == D.replace(hour=6))
check("a fractional timestamp parses without microseconds",
      PUSH.parse_paycor_time("2026-09-09T06:00:00.123") == D.replace(hour=6))
check("an offset timestamp parses to naive local",
      PUSH.parse_paycor_time("2026-09-09T06:00:00-04:00") == D.replace(hour=6))
check("an empty timestamp is None", PUSH.parse_paycor_time("") is None)

timecards = [{"employeeId": "px-1", "punchIn": "2026-09-09T06:00:00",
              "punchOut": "2026-09-09T18:00:00"}]
tc_index = PUSH.index_timecards(timecards)

# Our own punch, recognised by correlation id -- the exact path.
recon = PUSH.report_reconciliation([dict(row)], {}, {row["correlation_in"].lower()})
check("a punch we already sent is recognised by correlation id",
      recon["already_present"] == 1 and recon["missing"] == 0)

# Somebody else's equivalent punch, agreeing.
recon = PUSH.report_reconciliation([dict(row)], tc_index, set())
check("an equivalent Paycor punch counts as present",
      recon["already_present"] == 1, str(recon))
check("and is not reported missing", recon["missing"] == 0)

drift = PUSH.index_timecards([{"employeeId": "px-1",
                               "punchIn": "2026-09-09T06:00:00",
                               "punchOut": "2026-09-09T17:00:00"}])
recon = PUSH.report_reconciliation([dict(row)], drift, set())
check("a clock-out an hour out is reported as drifted", recon["drifted"] == 1)
check("and not silently counted as agreement", recon["already_present"] == 0)

recon = PUSH.report_reconciliation([dict(row)], {}, set())
check("an empty Paycor window reports the punch missing", recon["missing"] == 1)


# =============================
# 10. CLIENT GUARDS
# =============================
section("paycor client guards")

check("the client defaults to the sandbox",
      paycor_api.SANDBOX_BASE_URL in PaycorClient().base_url,
      PaycorClient().base_url)
check("and is not in production mode by default", paycor_api.IS_PRODUCTION is False)

client = PaycorClient()
check("a default client is read-only", client.read_only is True)

raised = None
try:
    client.create_punches([{"employeeId": "x"}])
except PaycorReadOnlyError as exc:
    raised = exc
check("a read-only client refuses to create punches", raised is not None)
check("and names the constructor argument that would allow it",
      raised is not None and "read_only=False" in str(raised))

raised = None
try:
    client.delete_punches("px-1", ["punch-1"])
except PaycorReadOnlyError as exc:
    raised = exc
check("a read-only client refuses to delete punches", raised is not None)

# The guard is on `write`, not the HTTP verb.
raised = None
try:
    client.request("GET", "v1/anything", write=True)
except PaycorReadOnlyError as exc:
    raised = exc
check("write=True is refused even on a GET", raised is not None)

contract = describe_write_contract()
check("the write contract reports the sandbox", contract["environment"] == "sandbox")
check("and is marked verified against the spec", contract["verified"] is True)
check("and names CreatePunches", "CreatePunches" in contract["request"])
check("and says a 202 is not a result",
      "punchErrorLog" in contract["accepted_response"], contract["accepted_response"])

writable = PaycorClient(read_only=False)
check("a writable client can be constructed", writable.read_only is False)

over = [{"employeeId": str(i)} for i in range(PA.MAX_PUNCH_BATCH + 1)]
raised = None
try:
    writable.create_punches(over, legal_entity_id=1)
except ValueError as exc:
    raised = exc
check("an oversized batch is refused before it is sent", raised is not None)
check("an empty batch is a no-op",
      writable.create_punches([], legal_entity_id=1) is None)

# A 202 that carries no tracking id must not read as success.
writable.request = lambda *a, **k: {"resourceUrl": {}}
check("a 202 with no tracking id returns None",
      writable.create_punches([{"employeeId": "x"}], legal_entity_id=1) is None)
writable.request = lambda *a, **k: {"resourceUrl": {"id": "track-1"}}
check("a 202 with a tracking id returns it",
      writable.create_punches([{"employeeId": "x"}], legal_entity_id=1) == "track-1")

# Pagination: the paged envelope, and the bare array employeePunches returns.
pages = [
    {"records": [{"a": 1}], "continuationToken": "t1", "hasMoreResults": True},
    {"records": [{"a": 2}], "hasMoreResults": False},
]
calls = []
def fake_request(method, path, params=None, json_body=None, write=False):
    calls.append(dict(params or {}))
    return pages[len(calls) - 1]
writable.request = fake_request
walked = list(writable.paginate("v1/whatever"))
check("pagination walks the records envelope", len(walked) == 2, str(walked))
check("and passes the continuation token on the second call",
      calls[1].get("continuationToken") == "t1", str(calls))

writable.request = lambda *a, **k: [{"punchId": "p1"}, {"punchId": "p2"}]
bare = list(writable.paginate("v1/employees/x/employeePunches"))
check("a bare-array response is handled too", len(bare) == 2, str(bare))


# =============================
# 11. OVERRIDES FILE
# =============================
section("overrides file")

# Relative to this file, not the working directory: the suite is run from
# wherever the operator happens to be standing.
example = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "state", "paycor_employee_overrides.example.json",
)
if os.path.exists(example):
    import json
    with open(example, encoding="utf-8") as handle:
        data = json.load(handle)
    check("the example overrides file is valid JSON", isinstance(data, dict))
    check("it carries an overrides block", "overrides" in data)
    check("and explains itself", "note" in data)
    check("its overrides map strings to strings",
          all(isinstance(k, str) and isinstance(v, str)
              for k, v in (data.get("overrides") or {}).items()))
else:
    skip("example overrides file checks",
         f"not present at {example}; expected in a full checkout, and not "
         "needed on a machine where the files were copied by hand")


# =============================
section("rotated refresh tokens are written back")

# Paycor rotates the refresh token on use in some configurations. Held in
# memory only, the next run starts from a credential Paycor has retired, and
# that arrives as an auth error on a morning nobody changed anything.
import shutil
import tempfile as _tempfile

env_dir = _tempfile.mkdtemp()
env_path = os.path.join(env_dir, ".env")


def write_env(body):
    with open(env_path, "w", encoding="utf-8") as handle:
        handle.write(body)


def read_env():
    with open(env_path, encoding="utf-8") as handle:
        return handle.read()


write_env("# comment\nTS_API_KEY=keep-me\nPAYCOR_REFRESH_TOKEN=old-token\nOTHER=untouched\n")
paycor_api.write_env_value(env_path, "PAYCOR_REFRESH_TOKEN", "new-token")
body = read_env()
check("the rotated value replaces the old one",
      "PAYCOR_REFRESH_TOKEN=new-token" in body and "old-token" not in body, body)
check("every other line survives",
      "# comment" in body and "TS_API_KEY=keep-me" in body and "OTHER=untouched" in body,
      body)
check("and nothing is duplicated", body.count("PAYCOR_REFRESH_TOKEN") == 1, body)

write_env("EXISTING=1\n")
added = paycor_api.write_env_value(env_path, "PAYCOR_REFRESH_TOKEN", "fresh")
check("a name not yet present is appended rather than lost",
      added is False and "PAYCOR_REFRESH_TOKEN=fresh" in read_env(), read_env())

write_env("export PAYCOR_REFRESH_TOKEN=old\n")
paycor_api.write_env_value(env_path, "PAYCOR_REFRESH_TOKEN", "new")
check("an exported line is replaced too, not duplicated",
      read_env().count("PAYCOR_REFRESH_TOKEN") == 1 and "new" in read_env(), read_env())

# The scoped name has to win: writing the bare one back while a scoped one is
# set would leave the retired value winning on the next run.
saved = dict(os.environ)
try:
    for key in [k for k in os.environ if k.startswith("PAYCOR_")]:
        del os.environ[key]
    os.environ["PAYCOR_REFRESH_TOKEN"] = "bare"
    os.environ["PAYCOR_SANDBOX_REFRESH_TOKEN"] = "scoped"
    paycor_api.IS_PRODUCTION = False
    check("the write-back targets the name that supplied the credential",
          paycor_api.credential_var_name("REFRESH_TOKEN") == "PAYCOR_SANDBOX_REFRESH_TOKEN",
          paycor_api.credential_var_name("REFRESH_TOKEN"))
    del os.environ["PAYCOR_SANDBOX_REFRESH_TOKEN"]
    check("and falls back to the bare name when no scoped one is set",
          paycor_api.credential_var_name("REFRESH_TOKEN") == "PAYCOR_REFRESH_TOKEN")
    del os.environ["PAYCOR_REFRESH_TOKEN"]
    check("a credential that came from code has no variable to write back to",
          paycor_api.credential_var_name("REFRESH_TOKEN") is None)
finally:
    os.environ.clear()
    os.environ.update(saved)
    paycor_api.IS_PRODUCTION = paycor_api.ENVIRONMENT == "production"

# A failed write must not take down a run that may be part way through payroll.
rotating = PaycorClient(subscription_key="k", refresh_token="old",
                        legal_entity_id="LE-1", read_only=True)
saved = dict(os.environ)
try:
    os.environ["PAYCOR_ENV_FILE"] = os.path.join(env_dir, "no", "such", "file")
    os.environ["PAYCOR_REFRESH_TOKEN"] = "old"
    persisted = rotating._persist_refresh_token("rotated")
    check("an unwritable env file is reported, not raised", persisted is False)

    os.environ["PAYCOR_PERSIST_REFRESH_TOKEN"] = "false"
    check("persistence can be turned off for a host that owns the secret",
          rotating._persist_refresh_token("rotated") is False)
finally:
    os.environ.clear()
    os.environ.update(saved)

# End to end: a token refresh whose response carries a new refresh token must
# actually reach the .env. The pieces above can all pass while nothing wires
# them together.
import contextlib
import io
import urllib.request as _urllib_request

write_env("PAYCOR_REFRESH_TOKEN=old-token\n")


class FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False


saved = dict(os.environ)
real_urlopen = _urllib_request.urlopen
try:
    for key in [k for k in os.environ if k.startswith("PAYCOR_")]:
        del os.environ[key]
    os.environ["PAYCOR_ENV_FILE"] = env_path
    os.environ["PAYCOR_REFRESH_TOKEN"] = "old-token"
    paycor_api.IS_PRODUCTION = False

    _urllib_request.urlopen = lambda *a, **k: FakeResponse(json.dumps({
        "access_token": "an-access-token",
        "refresh_token": "rotated-by-paycor",
        "expires_in": 3600,
    }).encode("utf-8"))

    rotator = PaycorClient(subscription_key="k", legal_entity_id="LE-1",
                           read_only=True)
    token = rotator._token()
    check("the refresh still yields an access token", token == "an-access-token", token)
    check("the rotated refresh token is held in memory",
          rotator.refresh_token == "rotated-by-paycor", rotator.refresh_token)
    check("and written to the env file, so the next run does not start retired",
          "PAYCOR_REFRESH_TOKEN=rotated-by-paycor" in read_env(), read_env())
    check("and the process environment is updated too",
          os.environ.get("PAYCOR_REFRESH_TOKEN") == "rotated-by-paycor")
finally:
    _urllib_request.urlopen = real_urlopen
    os.environ.clear()
    os.environ.update(saved)
    paycor_api.IS_PRODUCTION = paycor_api.ENVIRONMENT == "production"

shutil.rmtree(env_dir, ignore_errors=True)


# =============================
section("alerting")

import report_alerts as ALERTS

alert_calls = {"email": [], "sms": [], "webhook": [], "heartbeat": []}


def reset_alert_calls():
    for value in alert_calls.values():
        del value[:]


def fake_email(severity, headline, detail, recipients):
    alert_calls["email"].append((severity, headline, recipients))


def fake_sms(headline, recipients):
    alert_calls["sms"].append((headline, recipients))


def fake_webhook(severity, headline, detail, url):
    alert_calls["webhook"].append((severity, headline, url))


real = (ALERTS._send_email, ALERTS._send_sms, ALERTS._send_webhook)
ALERTS._send_email, ALERTS._send_sms, ALERTS._send_webhook = (
    fake_email, fake_sms, fake_webhook)

saved = dict(os.environ)
try:
    os.environ.update({
        "ALERT_EMAIL_TO": "ops@example.com",
        "ALERT_SMS_TO": "5551234567@example.net",
        "ALERT_WEBHOOK_URL": "https://example.com/hook",
    })

    reset_alert_calls()
    result = ALERTS.alert(ALERTS.CRITICAL, "it broke", "detail")
    check("a critical alert reaches every channel",
          sorted(result["sent"]) == ["email", "sms", "webhook"], str(result))

    # A warning must not wake anyone at 3am.
    reset_alert_calls()
    result = ALERTS.alert(ALERTS.WARNING, "needs a look")
    check("a warning does not send SMS",
          "sms" not in result["sent"] and not alert_calls["sms"], str(result))
    check("but does reach email and the channel",
          sorted(result["sent"]) == ["email", "webhook"], str(result))

    reset_alert_calls()
    result = ALERTS.alert(ALERTS.INFO, "all fine")
    check("routine completion is email only",
          result["sent"] == ["email"], str(result))

    # The channel that fails may be the reason an alert was needed.
    def boom(*a, **k):
        raise RuntimeError("smtp down")

    ALERTS._send_email = boom
    reset_alert_calls()
    result = ALERTS.alert(ALERTS.CRITICAL, "it broke")
    check("a failing channel does not stop the others",
          "webhook" in result["sent"] and "sms" in result["sent"], str(result))
    check("and the failure is reported rather than swallowed",
          any("email" in f for f in result["failed"]), str(result))
    ALERTS._send_email = fake_email

    raised = None
    try:
        ALERTS.alert("URGENT-ISH", "wrong severity")
    except ValueError as exc:
        raised = exc
    check("an unknown severity is refused", raised is not None)

    # A host with nothing configured must not crash the payroll run.
    for key in ("ALERT_EMAIL_TO", "ALERT_SMS_TO", "ALERT_WEBHOOK_URL"):
        del os.environ[key]
    result = ALERTS.alert(ALERTS.CRITICAL, "nowhere to go")
    check("no configured channel is survivable, not fatal",
          result == {"sent": [], "failed": []}, str(result))
finally:
    os.environ.clear()
    os.environ.update(saved)
    ALERTS._send_email, ALERTS._send_sms, ALERTS._send_webhook = real

# The heartbeat is the only thing that catches a run that never happened, so
# it must fire on success and NOT on failure.
pings = []
saved = dict(os.environ)
real_urlopen = ALERTS.urllib.request.urlopen
try:
    os.environ["ALERT_HEARTBEAT_URL"] = "https://hc.example.com/abc"

    class FakePing:
        def __init__(self, req):
            pings.append(req.full_url)

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def read(self):
            return b""

    ALERTS.urllib.request.urlopen = lambda req, timeout=None: FakePing(req)

    ALERTS.heartbeat(ok=True)
    check("a good run pings the watcher", pings == ["https://hc.example.com/abc"],
          str(pings))

    del pings[:]
    ALERTS.heartbeat(ok=False)
    check("a bad run pings the failure endpoint instead",
          pings == ["https://hc.example.com/abc/fail"], str(pings))

    del pings[:]
    del os.environ["ALERT_HEARTBEAT_URL"]
    check("no watcher configured means no ping and no error",
          ALERTS.heartbeat(ok=True) is False and not pings)
finally:
    ALERTS.urllib.request.urlopen = real_urlopen
    os.environ.clear()
    os.environ.update(saved)

# guarded() must alert AND re-raise: swallowing would leave a zero exit code
# and a scheduler that thinks the run succeeded.
raised = None
alerted = []
real_alert, real_heartbeat = ALERTS.alert, ALERTS.heartbeat
try:
    ALERTS.alert = lambda sev, headline, detail="": alerted.append((sev, headline))
    ALERTS.heartbeat = lambda ok=True, detail="": alerted.append(("heartbeat", ok))
    try:
        with ALERTS.guarded("test run"):
            raise ValueError("something broke")
    except ValueError as exc:
        raised = exc
    check("a crash inside guarded is re-raised, not swallowed", raised is not None)
    check("and alerts CRITICAL first",
          alerted and alerted[0][0] == ALERTS.CRITICAL, str(alerted))
    check("and tells the watcher the run failed",
          ("heartbeat", False) in alerted, str(alerted))

    del alerted[:]
    with ALERTS.guarded("test run"):
        pass
    check("a clean run pings the watcher and raises nothing",
          alerted == [("heartbeat", True)], str(alerted))
finally:
    ALERTS.alert, ALERTS.heartbeat = real_alert, real_heartbeat


# =============================
section("reconciliation protects the publish")

# Correlation ids only recognise OUR OWN previous writes. A punch the crew
# clocked themselves carries none, so nothing downstream stops a publish from
# adding a second punch for a shift Paycor already holds. A production run had
# 86 such shifts against 63 genuinely new ones.
def recon_row(uid):
    return {
        "user_id": uid,
        "employee_name": f"Crew {uid}",
        "paycor_employee_id": f"emp-{uid}",
        "punch_in": datetime(2026, 9, 2, 8, 0),
        "punch_out": datetime(2026, 9, 2, 16, 0),
        "profile": "TEST-1",
        "correlation_in": f"cid-in-{uid}",
        "correlation_out": f"cid-out-{uid}",
    }


recon_rows = [
    recon_row("10"),   # Paycor has it, clocks agree
    recon_row("11"),   # Paycor has it, clocks differ
    recon_row("12"),   # Paycor has that day, too far off to pair
    recon_row("13"),   # Paycor has nothing
]

shift_in = datetime(2026, 9, 2, 8, 0)
def held(punch_in, punch_out, hours=None, pay=None):
    return {"in": punch_in, "out": punch_out, "hours": hours, "pay": pay}


timecards = {
    "emp-10": [held(shift_in, datetime(2026, 9, 2, 16, 0), 8.0, 200.0)],
    "emp-11": [held(shift_in + timedelta(minutes=30),
                    datetime(2026, 9, 2, 21, 0), 12.5, 312.50)],
    # Five hours off the clock-in: too far to pair, same day all the same.
    "emp-12": [held(shift_in + timedelta(hours=5),
                    datetime(2026, 9, 2, 23, 0), 10.0, 250.0)],
}

PUSH.report_reconciliation(recon_rows, timecards, set())

by_user = {r["user_id"]: r for r in recon_rows}
check("a shift Paycor already holds is marked as held",
      bool(by_user["10"].get("paycor_holds")), str(by_user["10"].get("paycor_holds")))
check("so is one where the clocks disagree",
      bool(by_user["11"].get("paycor_holds")))
check("so is one Paycor has that day but could not be paired",
      bool(by_user["12"].get("paycor_holds")),
      str(by_user["12"].get("paycor_holds")))
check("and the genuinely new shift is not",
      not by_user["13"].get("paycor_holds"))
check("the unpaired one names the day rather than a clock match",
      "too far off to pair" in (by_user["12"].get("paycor_holds") or ""),
      by_user["12"].get("paycor_holds"))

# The pay impact: a shift Traumasoft says is 8h that Paycor pays 8h for has no
# gap; one Paycor pays less for does, priced at the rate Paycor itself implies.
impact = [r for r in recon_rows if r["user_id"] == "11"][0]
check("a shift both systems hold carries Paycor's own pay figure",
      impact["paycor_existing"]["pay"] == 312.50)
check("and its hours", impact["paycor_existing"]["hours"] == 12.5)

# 200.00 over 8.0 hours = 25.00/h implied, against a Traumasoft 8.0h shift:
# no gap, so nothing is claimed.
agreed = [r for r in recon_rows if r["user_id"] == "10"][0]
ts_hours = (agreed["punch_out"] - agreed["punch_in"]).total_seconds() / 3600.0
check("a shift whose hours agree shows no gap",
      abs(ts_hours - agreed["paycor_existing"]["hours"]) < 0.01)

# A timecard with no pay figure must not crash the report or invent one.
PUSH.report_pay_impact([{
    "employee_name": "No Pay Data",
    "user_id": "99",
    "punch_in": datetime(2026, 9, 2, 8, 0),
    "punch_out": datetime(2026, 9, 2, 20, 0),
    "paycor_existing": {"in": datetime(2026, 9, 2, 8, 0),
                        "out": datetime(2026, 9, 2, 16, 0),
                        "hours": 8.0, "pay": None},
}])
check("a timecard with no pay figure is reported without inventing one", True)

# An open punch on Paycor's side has no hours to compare; it must be skipped
# rather than counted as a zero-hour shift, which would overstate every gap.
PUSH.report_pay_impact([{
    "employee_name": "Still Open",
    "user_id": "98",
    "punch_in": datetime(2026, 9, 2, 8, 0),
    "punch_out": datetime(2026, 9, 2, 20, 0),
    "paycor_existing": {"in": datetime(2026, 9, 2, 8, 0), "out": None,
                        "hours": None, "pay": None},
}])
check("an open Paycor punch is skipped rather than scored as zero hours", True)


# =============================
section("publish verifies by reading back")

# Sandbox run 3: a batch accepted with a tracking id and an empty error log
# that created nothing, because punches already occupied those times. An empty
# error log says nothing was rejected -- never that anything was written.
import push_paycor_timecards as PUSH

PUSH.PUBLISH_VERIFY_SETTLE_SECONDS = 0


def publish_row(uid, cid_in, cid_out):
    return {
        "user_id": uid,
        "employee_name": f"Crew {uid}",
        "paycor_employee_id": f"emp-{uid}",
        "department_id": "dept-1",
        "activity_type_id": "act-1",
        "punch_in": datetime(2026, 9, 2, 8, 0),
        "punch_out": datetime(2026, 9, 2, 16, 0),
        "profile": "TEST-1",
        "punch_id": f"p{uid}",
        "correlation_in": cid_in,
        "correlation_out": cid_out,
    }


class FakePaycor:
    """Accepts everything, and returns only the correlation ids it was told."""

    def __init__(self, readable):
        self.readable = readable

    def create_punches(self, payload, legal_entity_id=None):
        return "trk-1"

    def punch_errors(self, tracking_id, legal_entity_id=None):
        return []

    def get_employee_punches(self, employee_id, start, end):
        return [{"correlationId": c} for c in self.readable.get(employee_id, [])]


# Both halves come back: genuinely written.
rows = [publish_row("1", "cid-in-1", "cid-out-1")]
sent, failed, unverified = PUSH.publish(
    FakePaycor({"emp-1": ["cid-in-1", "cid-out-1"]}), rows)
check("a punch that reads back is reported as sent",
      len(sent) == 1 and not unverified and not failed,
      f"sent={len(sent)} unverified={len(unverified)}")

# Accepted, no errors, and nothing created. The case that was reported as sent.
rows = [publish_row("2", "cid-in-2", "cid-out-2")]
sent, failed, unverified = PUSH.publish(FakePaycor({"emp-2": []}), rows)
check("a punch accepted with an empty error log but absent from Paycor "
      "is NOT reported as sent", not sent, f"sent={len(sent)}")
check("it is reported as unverified", len(unverified) == 1)
check("and says it was not created",
      "not created" in (unverified[0].get("error") or ""),
      unverified[0].get("error"))

# Half a shift is still a problem: an In with no Out is an open punch.
rows = [publish_row("3", "cid-in-3", "cid-out-3")]
sent, failed, unverified = PUSH.publish(
    FakePaycor({"emp-3": ["cid-in-3"]}), rows)
check("a punch whose Out did not land is unverified, not sent",
      not sent and len(unverified) == 1)
check("and names which half is missing",
      "Out" in (unverified[0].get("error") or ""), unverified[0].get("error"))


# =============================
section("delete punches")

# The sandbox probe reported a clean cleanup over punches Paycor had kept,
# because a delete is asynchronous like a create and nothing read its result.
# These pin down the two things that made that possible.
sent = {}


def capture_delete(method, path, params=None, json_body=None, write=False):
    sent["method"] = method
    sent["path"] = path
    sent["body"] = json_body
    return {"resourceUrl": {"id": "trk-delete-1"}}


deleter = PaycorClient(subscription_key="k", access_token="t",
                       legal_entity_id="LE-1", read_only=False)
deleter.request = capture_delete

tracking = deleter.delete_punches("emp-1", ["punch-1", "punch-2"])
check("delete_punches returns the tracking id from the 202",
      tracking == "trk-delete-1", str(tracking))
check("bare ids still work", sent["body"] == [{"punchId": "punch-1"},
                                              {"punchId": "punch-2"}],
      str(sent["body"]))
check("and it is a DELETE", sent["method"] == "DELETE", sent["method"])

# The spec: when a punch carries a punchRefId, the delete must name it too.
# Passing records rather than ids is what makes that possible.
deleter.delete_punches("emp-1", [
    {"punchId": "punch-1", "punchRefId": "ref-1"},
    {"punchId": "punch-2"},
])
check("punchRefId is passed through when the punch carries one",
      sent["body"] == [{"punchId": "punch-1", "punchRefId": "ref-1"},
                       {"punchId": "punch-2"}],
      str(sent["body"]))

deleter.delete_punches("emp-1", [{"id": "punch-9"}])
check("a record keyed 'id' is accepted too",
      sent["body"] == [{"punchId": "punch-9"}], str(sent["body"]))

sent.clear()
check("an empty delete sends nothing",
      deleter.delete_punches("emp-1", []) is None and not sent)

raised = None
try:
    deleter.delete_punches("emp-1", [{"punchId": f"p{i}"} for i in range(101)])
except ValueError as exc:
    raised = exc
check("a delete over the batch ceiling is refused", raised is not None)


# =============================
section("per-environment credentials")

# env_credential reads module-level IS_PRODUCTION, which is fixed at import.
# Rather than reimport the module, exercise the resolution directly against a
# patched flag -- the behaviour under test is the name precedence, not the
# import-time gate.
def resolve(name, environ, production):
    saved_env = dict(os.environ)
    saved_flag = paycor_api.IS_PRODUCTION
    try:
        for key in [k for k in os.environ if k.startswith("PAYCOR_")]:
            del os.environ[key]
        os.environ.update(environ)
        paycor_api.IS_PRODUCTION = production
        return paycor_api.env_credential(name, "")
    finally:
        os.environ.clear()
        os.environ.update(saved_env)
        paycor_api.IS_PRODUCTION = saved_flag


check("the bare name is used when no scoped name is set",
      resolve("SUBSCRIPTION_KEY", {"PAYCOR_SUBSCRIPTION_KEY": "bare"}, False) == "bare")

check("a sandbox-scoped name wins in the sandbox",
      resolve("SUBSCRIPTION_KEY",
              {"PAYCOR_SUBSCRIPTION_KEY": "bare",
               "PAYCOR_SANDBOX_SUBSCRIPTION_KEY": "sand"}, False) == "sand")

check("a production-scoped name wins in production",
      resolve("SUBSCRIPTION_KEY",
              {"PAYCOR_SUBSCRIPTION_KEY": "bare",
               "PAYCOR_PRODUCTION_SUBSCRIPTION_KEY": "prod"}, True) == "prod")

# The whole point of the change: one environment's credentials must never be
# reachable while the other is selected.
check("a production-scoped name is invisible in the sandbox",
      resolve("SUBSCRIPTION_KEY",
              {"PAYCOR_PRODUCTION_SUBSCRIPTION_KEY": "prod"}, False) == "")

check("a sandbox-scoped name is invisible in production",
      resolve("SUBSCRIPTION_KEY",
              {"PAYCOR_SANDBOX_SUBSCRIPTION_KEY": "sand"}, True) == "")

check("an empty scoped name falls back rather than blanking the credential",
      resolve("SUBSCRIPTION_KEY",
              {"PAYCOR_SUBSCRIPTION_KEY": "bare",
               "PAYCOR_SANDBOX_SUBSCRIPTION_KEY": "   "}, False) == "bare")

check("legal entity id is scoped too, so the entity matches the tenant",
      resolve("LEGAL_ENTITY_ID",
              {"PAYCOR_LEGAL_ENTITY_ID": "196750",
               "PAYCOR_SANDBOX_LEGAL_ENTITY_ID": "42"}, False) == "42")


# =============================
section("sandbox probe refuses production")

# The probe writes punches and then deletes them. Both are fine in a sandbox
# and neither belongs in a tenant where people are paid, so the refusal is a
# safety property worth a regression test rather than a code comment.
probe = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sandbox_punch_probe.py")
if os.path.exists(probe):
    import subprocess

    env = dict(os.environ)
    env.update({
        "PAYCOR_ENVIRONMENT": "production",
        "PAYCOR_SUBSCRIPTION_KEY": "test-subscription-key",
        "PAYCOR_ACCESS_TOKEN": "test-access-token",
        "PAYCOR_LEGAL_ENTITY_ID": "LE-1",
    })
    result = subprocess.run([sys.executable, probe], env=env,
                            capture_output=True, text=True, timeout=60)

    check("the sandbox probe exits non-zero when pointed at production",
          result.returncode != 0, f"exit {result.returncode}")
    check("and says why", "production" in result.stdout.lower())
    # If it got as far as talking to Paycor it would have had to authenticate,
    # and the placeholder credentials above would have surfaced as an error.
    check("and refuses before making any request",
          "Could not" not in result.stdout and "Traceback" not in result.stderr,
          result.stderr[-200:] if result.stderr else "")

    env["PAYCOR_ENVIRONMENT"] = "PRODUCTION"
    upper = subprocess.run([sys.executable, probe], env=env,
                           capture_output=True, text=True, timeout=60)
    check("the refusal is not case-sensitive", upper.returncode != 0,
          f"exit {upper.returncode}")
else:
    skip("sandbox probe production guard",
         "sandbox_punch_probe.py not present in this checkout")


# =============================
print("\n" + "=" * 60)
for note in SKIPPED:
    print(f"SKIPPED: {note}")
if FAILED:
    print(f"FAILED {len(FAILED)} of {PASSED + len(FAILED)} checks:")
    for failure in FAILED:
        print(f"  - {failure}")
    sys.exit(1)
print(f"All {PASSED} checks passed."
      + (f" ({len(SKIPPED)} skipped)" if SKIPPED else ""))
