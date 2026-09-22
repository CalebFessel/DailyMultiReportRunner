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

entry, how = PUSH.resolve_employee(employee(2, "1002"), {"1002": "px-chosen"}, index)
check("an override beats an ambiguous match",
      entry and entry["id"] == "px-chosen" and how == "override")

entry, how = PUSH.resolve_employee(employee(3, ""), {"user_id:3": "px-byuser"}, index)
check("an override can key on user_id", entry and entry["id"] == "px-byuser")

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
