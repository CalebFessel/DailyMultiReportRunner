"""
Push Traumasoft punches into Paycor as timecard punches.

The goal this serves: crews clock in twice, only Paycor pays them, so nobody
closes a Traumasoft punch and every unit-hour figure in the report bundle rests
on a record with no incentive behind it. Making Traumasoft the clock that pays
fixes the data at its source. This is the mechanism.

Built against Paycor's own OpenAPI spec (docs/paycor-public-api-v1.json), which
settles three things a guess got wrong:

  * **Punches are events, not intervals.** One Traumasoft punch row, with a
    start and an end, becomes TWO Paycor punches -- an In and an Out.
  * **Three guids are required per punch** and Paycor defaults none of them:
    employeeId, departmentId and activityTypeId. The first two come from the
    employee roster, the third is a per-tenant choice (`PAYCOR_ACTIVITY_TYPE`,
    default "Work").
  * **A 202 is not success.** CreatePunches validates asynchronously and
    returns a tracking id; whether the punches landed is only visible in the
    punch error log for that id. Every publish here reads that log back.

Idempotency is by correlation id, not by timestamp. Each punch carries a
correlationId derived deterministically from the Traumasoft punch id, and
Paycor returns it on read -- so "have I sent this already?" is an exact lookup
that survives a re-run, a restart, and clocks that disagree by a minute.

Three modes, meant to be used in order:

  1. DRY RUN (default) -- build the plan and print it. Reads only. Shows the
     exact JSON that would be sent and everything refused, with reasons.
  2. --reconcile -- read what Paycor already holds for the same window and
     compare. This is the test that matters before any write: it proves the
     employee mapping resolves, the clocks agree, and both systems describe the
     same shifts. Still read-only on both sides.
  3. --publish -- send them, then read the error log back.

Safety rails, none of which are configurable away:

  * An open punch is never sent. A punch with no clock-out is not a payable
    record, and the shift-end fallback that makes it usable for reporting would
    here mean inventing the end of somebody's paid day.
  * An employee who cannot be mapped unambiguously is refused, not guessed.
  * A punch Paycor already holds, by correlation id, is skipped.
  * The first failed batch stops the run.
  * Production needs PAYCOR_ENVIRONMENT=production *and*
    --i-understand-this-is-payroll.

The shift feed only returns today-1..today+2, so this can only push a narrow
window and must run daily to cover a pay period. A day missed is a day whose
punches the API will not hand back.

Usage:
    python push_paycor_timecards.py                          # dry run
    python push_paycor_timecards.py --reconcile              # compare only
    python push_paycor_timecards.py --date 2026-09-10 --json plan.json
    python push_paycor_timecards.py --reconcile --publish --limit 2
"""

import os
import sys
import json
import time
import logging
import argparse
from collections import defaultdict
from datetime import datetime, date, timedelta

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError
import traumasoft_reports as R
import paycor_api
from paycor_api import (
    PaycorClient,
    PaycorAPIError,
    PaycorAuthError,
    PaycorReadOnlyError,
    correlation_id,
    describe_write_contract,
    PUNCH_STATUS_IN,
    PUNCH_STATUS_OUT,
    MAX_PUNCH_BATCH,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("push-paycor")

STATE_DIR = os.getenv("TS_STATE_DIR", "state")
OVERRIDES_FILE = os.path.join(STATE_DIR, "paycor_employee_overrides.json")

# Paycor accepts a punch before it serves it back, so verification waits.
PUBLISH_VERIFY_SETTLE_SECONDS = int(os.getenv("PAYCOR_VERIFY_SETTLE", "5"))

# Which activity type a pushed punch is filed under. Required by Paycor, with
# no default of its own; "Work" is the ordinary productive type on most
# tenants, and the dry run prints what this tenant actually has.
ACTIVITY_TYPE_NAME = os.getenv("PAYCOR_ACTIVITY_TYPE", "Work").strip()
ACTIVITY_TYPE_ID = os.getenv("PAYCOR_ACTIVITY_TYPE_ID", "").strip()

# A department for people whose Paycor record names none. Left empty by
# default so those punches are refused rather than filed somewhere arbitrary.
DEFAULT_DEPARTMENT_ID = os.getenv("PAYCOR_DEFAULT_DEPARTMENT_ID", "").strip()

# A punch longer than this is refused as implausible. A runaway punch reaching
# payroll would pay a multi-day shift.
MAX_PUNCH_HOURS = float(os.getenv("PAYCOR_MAX_PUNCH_HOURS", "24"))

# Used only for the human-readable drift report; matching is by correlation id.
DRIFT_TOLERANCE_MINUTES = float(os.getenv("PAYCOR_MATCH_TOLERANCE_MINUTES", "2"))


# =============================
# MAPPING
# =============================
def load_overrides():
    """Traumasoft employee_num (or user_id:N) -> Paycor employee guid."""
    data = R.load_state_file(OVERRIDES_FILE, "Paycor employee overrides") or {}
    raw = data.get("overrides", data) if isinstance(data, dict) else {}
    return {str(k).strip().lower(): str(v).strip() for k, v in raw.items() if v}


def paycor_employee_index(paycor_employees):
    """
    Index the Paycor roster by every identifier a Traumasoft record might name.

    Returns number -> [ {id, department_id, label}, ... ]. A list rather than a
    single entry because a number shared by two people has to be refused, and
    that can only be seen by keeping both.
    """
    index = defaultdict(list)
    for emp in paycor_employees:
        emp_id = emp.get("id")
        if not emp_id:
            continue
        department = emp.get("department") or {}
        entry = {
            "id": str(emp_id),
            "department_id": department.get("id"),
            "label": " ".join(
                p for p in (emp.get("firstName"), emp.get("lastName")) if p
            ).strip(),
        }
        for key in ("employeeNumber", "alternateEmployeeNumber", "badgeNumber"):
            value = str(emp.get(key) or "").strip()
            if value:
                index[value.lower()].append(entry)
    return index


def resolve_activity_type(activity_types):
    """
    (activity_type_id, how) or (None, why not).

    An explicit id wins. Otherwise match on name, since the guid differs per
    tenant and a name is the only thing a human can reasonably put in .env.
    """
    if ACTIVITY_TYPE_ID:
        return ACTIVITY_TYPE_ID, "PAYCOR_ACTIVITY_TYPE_ID"
    if not activity_types:
        return None, "no activity types could be read from Paycor"
    wanted = ACTIVITY_TYPE_NAME.lower()
    matches = [a for a in activity_types if str(a.get("name", "")).strip().lower() == wanted]
    if len(matches) == 1:
        return str(matches[0]["id"]), f"name {ACTIVITY_TYPE_NAME!r}"
    available = ", ".join(sorted(str(a.get("name")) for a in activity_types))
    if not matches:
        return None, (f"no activity type named {ACTIVITY_TYPE_NAME!r}; "
                      f"this tenant has: {available}")
    return None, (f"{len(matches)} activity types named {ACTIVITY_TYPE_NAME!r}; "
                  "set PAYCOR_ACTIVITY_TYPE_ID to choose one")


def resolve_employee(ts_employee, overrides, paycor_index):
    """
    (entry, how) or (None, why not), where entry carries the Paycor guid.

    An override always wins, being a decision rather than an observation. Then
    employee_num against the Paycor roster. An ambiguous number is refused --
    two people sharing a payroll number is exactly where a guess pays the wrong
    person.
    """
    user_id = str(ts_employee.get("user_id") or "").strip()
    number = str(ts_employee.get("employee_num") or "").strip()

    for key in (number.lower(), f"user_id:{user_id}".lower()):
        if key and key in overrides:
            return {"id": overrides[key], "department_id": None, "label": None}, "override"

    if not number:
        return None, "Traumasoft employee carries no employee_num"
    if not paycor_index:
        return None, "no Paycor roster loaded, so the employee cannot be resolved"

    matches = paycor_index.get(number.lower()) or []
    if len(matches) == 1:
        return matches[0], "employee_num"
    if not matches:
        return None, f"employee_num {number!r} matches no Paycor employee"
    return None, f"employee_num {number!r} matches {len(matches)} Paycor employees"


# =============================
# PLAN
# =============================
def candidate_punches(shifts, offset, now, target_date=None):
    """
    Every Traumasoft punch that could be pushed, with the reason if it cannot.

    Nothing is filtered out silently: a refused punch stays in the plan with a
    `refused` reason, because the whole value of the dry run is seeing what
    would not go and why.
    """
    out = []
    for shift in shifts:
        if shift.get("deleted"):
            continue
        profile = R.profile_name(shift)
        shift_end = R.parse_shift_ts(shift.get("end_time"), offset)

        for punch in shift.get("punches") or []:
            if punch.get("deleted"):
                continue
            start = R.parse_shift_ts(punch.get("start_time"), offset)
            end = R.parse_shift_ts(punch.get("end_time"), offset)
            if not start:
                continue
            if target_date is not None and start.date() != target_date:
                continue

            refused = None
            if end is None:
                shift_over = bool(shift_end and now > shift_end)
                refused = (
                    "punch is open (missed punch-out)" if shift_over
                    else "punch is open (shift still running)"
                )
            elif end <= start:
                refused = "clock-out is not after clock-in"
            elif (end - start).total_seconds() / 3600.0 > MAX_PUNCH_HOURS:
                refused = (
                    f"punch spans {(end - start).total_seconds() / 3600.0:.1f}h, "
                    f"over the {MAX_PUNCH_HOURS:g}h ceiling"
                )

            out.append({
                "punch_id": punch.get("id"),
                "user_id": shift.get("user_id"),
                "profile": profile,
                "vehicle_name": shift.get("vehicle_name"),
                "punch_in": start,
                "punch_out": end,
                "refused": refused,
            })
    return out


def build_plan(shifts, ts_employees, offset, now, overrides, paycor_index,
               activity_type_id, target_date=None):
    by_user = {}
    for emp in ts_employees:
        uid = emp.get("user_id")
        if uid is not None:
            by_user[str(uid)] = emp

    sendable, refused = [], []
    for row in candidate_punches(shifts, offset, now, target_date):
        ts_emp = by_user.get(str(row["user_id"])) or {}
        row["employee_num"] = ts_emp.get("employee_num")
        row["employee_name"] = " ".join(
            p for p in (ts_emp.get("first_name"), ts_emp.get("last_name")) if p
        ).strip() or None
        row["cost_center"] = ts_emp.get("cost_center_name")

        if row["refused"]:
            refused.append(row)
            continue

        entry, how = resolve_employee(ts_emp, overrides, paycor_index)
        row["mapped_by"] = how
        if entry is None:
            row["refused"] = how
            refused.append(row)
            continue

        row["paycor_employee_id"] = entry["id"]
        department_id = entry.get("department_id") or DEFAULT_DEPARTMENT_ID
        if not department_id:
            row["refused"] = (
                "Paycor employee names no department, and CreatePunches "
                "requires one (set PAYCOR_DEFAULT_DEPARTMENT_ID to supply a fallback)"
            )
            refused.append(row)
            continue
        row["department_id"] = department_id

        if not activity_type_id:
            row["refused"] = "no activity type resolved; CreatePunches requires one"
            refused.append(row)
            continue
        row["activity_type_id"] = activity_type_id

        # The two halves, each with its own stable correlation id.
        row["correlation_in"] = correlation_id(row["punch_id"], "in")
        row["correlation_out"] = correlation_id(row["punch_id"], "out")
        sendable.append(row)

    return sendable, refused


def punch_objects(row):
    """The two Paycor punch objects one Traumasoft punch becomes."""
    note = f"TS {row['profile']}" if row.get("profile") else None
    return [
        PaycorClient.build_punch(
            row["paycor_employee_id"], row["department_id"], row["activity_type_id"],
            row["punch_in"], PUNCH_STATUS_IN,
            note=note, correlation=row["correlation_in"],
        ),
        PaycorClient.build_punch(
            row["paycor_employee_id"], row["department_id"], row["activity_type_id"],
            row["punch_out"], PUNCH_STATUS_OUT,
            note=note, correlation=row["correlation_out"],
        ),
    ]


def chunked(items, size):
    for i in range(0, len(items), size):
        yield items[i:i + size]


# =============================
# RECONCILE
# =============================
def parse_paycor_time(value):
    """
    Paycor timestamps arrive in a few shapes; take the first that parses.

    Sub-second precision is dropped rather than carried. A punch is a clock
    event to the minute; keeping microseconds would put a precision on payroll
    times that nothing behind them supports.
    """
    if not value:
        return None
    text = str(value).strip().replace("Z", "+00:00")
    for candidate in (text, text.split(".")[0], text[:19]):
        try:
            parsed = datetime.fromisoformat(candidate)
        except ValueError:
            continue
        return parsed.replace(tzinfo=None, microsecond=0)
    return None


def index_timecards(timecards):
    """paycor employee guid -> [(punch_in, punch_out), ...] from the paired read."""
    index = defaultdict(list)
    for row in timecards:
        emp_id = str(row.get("employeeId") or "").strip()
        if not emp_id:
            continue
        punch_in = parse_paycor_time(row.get("punchIn"))
        punch_out = parse_paycor_time(row.get("punchOut"))
        if punch_in:
            index[emp_id].append((punch_in, punch_out))
    return index


def collect_correlations(paycor, sendable, window_start, window_end):
    """
    Every correlation id Paycor already holds for the crew we are about to push.

    Read per employee because only the employeePunches endpoint returns
    correlationId; the legal-entity read gives pairs without it. Employees are
    visited once each, not once per punch.
    """
    seen = set()
    employees = sorted({row["paycor_employee_id"] for row in sendable})
    for employee_id in employees:
        try:
            events = paycor.get_employee_punches(
                employee_id, window_start.isoformat(), window_end.isoformat()
            )
        except PaycorAPIError as exc:
            # A 404 here means this employee has no punches in the window,
            # which is information rather than a failure.
            if exc.status_code == 404:
                continue
            raise
        for event in events:
            marker = event.get("correlationId")
            if marker:
                seen.add(str(marker).lower())
    return seen


def report_reconciliation(sendable, timecard_index, known_correlations):
    print("\n" + "=" * 78)
    print("RECONCILIATION -- Traumasoft against what Paycor already holds")
    print("=" * 78)

    already, missing, drifted = [], [], []
    for row in sendable:
        if str(row["correlation_in"]).lower() in known_correlations:
            already.append(row)
            continue
        # Not ours -- but Paycor may still hold an equivalent punch entered by
        # hand or by the crew themselves. That is the interesting case: the
        # same work recorded twice, by two clocks that may disagree.
        match = None
        for existing_in, existing_out in timecard_index.get(row["paycor_employee_id"], []):
            if abs((existing_in - row["punch_in"]).total_seconds()) <= 3600:
                match = (existing_in, existing_out)
                break
        if match is None:
            missing.append(row)
            continue
        row["paycor_existing"] = match
        gap_in = abs((match[0] - row["punch_in"]).total_seconds()) / 60.0
        gap_out = (
            abs((match[1] - row["punch_out"]).total_seconds()) / 60.0
            if match[1] and row["punch_out"] else None
        )
        row["drift_in_minutes"] = round(gap_in, 1)
        row["drift_out_minutes"] = round(gap_out, 1) if gap_out is not None else None
        worst = max([g for g in (gap_in, gap_out) if g is not None] or [0])
        if worst > DRIFT_TOLERANCE_MINUTES:
            drifted.append(row)
        else:
            already.append(row)

    total = len(sendable)
    print(f"  Traumasoft punches in window   : {total}")
    print(f"  Already in Paycor (ours)       : "
          f"{sum(1 for r in already if str(r['correlation_in']).lower() in known_correlations)}")
    print(f"  Matched an existing Paycor punch: "
          f"{sum(1 for r in already if 'paycor_existing' in r)}")
    print(f"  Same shift, clocks disagree    : {len(drifted)}")
    print(f"  Not in Paycor at all           : {len(missing)}")

    if drifted:
        print("\n  Both systems hold these, with times that differ. Both clocks are")
        print("  shown in full, with dates: a shift crossing midnight reads as a")
        print("  13-hour day or a 37-hour one depending on which date the out")
        print("  carries, and an out printed as a bare wall time hides that.")
        print()
        print(f"    {'crew':<22} {'Traumasoft in':<17} {'Traumasoft out':<17} "
              f"{'Paycor in':<17} {'Paycor out':<17}  drift")
        for row in sorted(drifted, key=lambda r: -max(
                r.get("drift_in_minutes") or 0, r.get("drift_out_minutes") or 0))[:15]:
            existing_in, existing_out = row.get("paycor_existing", (None, None))
            out_gap = row["drift_out_minutes"]
            drift = (f"in {row['drift_in_minutes']:.0f}m"
                     + (f" out {out_gap:.0f}m" if out_gap is not None else ""))
            print(f"    {(row['employee_name'] or row['user_id'])!s:<22} "
                  f"{row['punch_in']:%m-%d %H:%M}     "
                  f"{row['punch_out']:%m-%d %H:%M}     "
                  f"{existing_in:%m-%d %H:%M}     "
                  f"{(f'{existing_out:%m-%d %H:%M}' if existing_out else 'open'):<17}"
                  f"  {drift}")

        # A whole day of drift is a different problem from a few minutes, and
        # only one of them is a clock.
        day_off = [r for r in drifted
                   if r.get("drift_out_minutes") is not None
                   and abs(r["drift_out_minutes"] - 1440) <= 60]
        if day_off:
            print(f"\n  {len(day_off)} of these differ by close to exactly 24 hours on the")
            print("  clock-out, and every one crosses midnight. That is a date being")
            print("  assigned wrong rather than two clocks drifting apart. Check the")
            print("  Paycor dates above against what the crew actually worked before")
            print("  deciding which side is right -- and note that a publish would")
            print("  not correct these: it adds punches, it does not amend one.")

    if missing:
        print(f"\n  Punches Traumasoft has and Paycor does not ({len(missing)}).")
        print("  These are what a publish would add:")
        for row in missing[:15]:
            hours = (row["punch_out"] - row["punch_in"]).total_seconds() / 3600.0
            print(f"    {(row['employee_name'] or row['user_id'])!s:<24} "
                  f"{row['punch_in']:%m-%d %H:%M} - {row['punch_out']:%m-%d %H:%M} "
                  f"({hours:4.1f}h)  {row['profile']}")
        if len(missing) > 15:
            print(f"    ... and {len(missing) - 15} more")

    if total and not missing and not drifted:
        print("\n  Every Traumasoft punch is already in Paycor and the clocks agree.")
        print("  That is the result you want before cutting over: the two systems")
        print("  describe the same work, so making Traumasoft authoritative")
        print("  changes who is trusted, not what anyone is paid.")

    return {
        "already_present": len(already),
        "missing": len(missing),
        "drifted": len(drifted),
        "missing_rows": missing,
    }


# =============================
# OUTPUT
# =============================
def report_plan(sendable, refused, target_date, activity_note):
    print("\n" + "=" * 78)
    print(f"PLAN{f' for {target_date}' if target_date else ''}")
    print("=" * 78)
    print(f"  Traumasoft punches ready : {len(sendable)}")
    print(f"  Paycor punch objects     : {len(sendable) * 2}  (an In and an Out each)")
    print(f"  Refused                  : {len(refused)}")
    print(f"  Activity type            : {activity_note}")

    if sendable:
        payable = sum(
            (r["punch_out"] - r["punch_in"]).total_seconds() / 3600.0 for r in sendable
        )
        crew = len({r["user_id"] for r in sendable})
        print(f"  Payable hours represented: {payable:.2f} across {crew} crew")

    if refused:
        buckets = defaultdict(int)
        for row in refused:
            buckets[row["refused"]] += 1
        print("\n  Why punches were refused:")
        for reason, count in sorted(buckets.items(), key=lambda b: -b[1]):
            print(f"    {count:>5}  {reason}")
        missed = buckets.get("punch is open (missed punch-out)", 0)
        if missed:
            print(f"\n  {missed} of those are missed punch-outs -- work that happened")
            print("  and cannot be paid from this feed. That is the problem the")
            print("  cutover is meant to solve, quantified.")

    if sendable:
        print("\n  First punch, exactly as it would be sent:")
        for obj in punch_objects(sendable[0]):
            print("    " + json.dumps(obj))


def report_contract(activity_note):
    contract = describe_write_contract()
    print("\n" + "=" * 78)
    print("WRITE CONTRACT")
    print("=" * 78)
    print(f"  environment : {contract['environment']}")
    print(f"  base url    : {contract['base_url']}")
    print(f"  request     : {contract['request']}")
    print(f"  body        : {contract['body']}")
    print(f"  required    : {', '.join(contract['required_fields'])}")
    print(f"  punch model : {contract['punch_model']}")
    print(f"  on accept   : {contract['accepted_response']}")
    print(f"  activity    : {activity_note}")
    print(f"  source      : {contract['source']}")


def jsonable(row):
    out = {}
    for key, value in row.items():
        if isinstance(value, datetime):
            out[key] = value.isoformat(timespec="minutes")
        elif isinstance(value, tuple):
            out[key] = [v.isoformat(timespec="minutes") if isinstance(v, datetime) else v
                        for v in value]
        else:
            out[key] = value
    return out


# =============================
# PUBLISH
# =============================
def publish(paycor, queue, legal_entity_id=None):
    """
    Send the queue in batches, read each batch's error log, then read the
    punches back and confirm they exist.

    A 202 only says Paycor accepted the batch for processing, and an empty
    error log only says nothing was *rejected*. Neither says anything was
    *created*. Sandbox testing found exactly that case: a batch accepted with
    an empty error log that created nothing, because punches already occupied
    those times and Paycor silently kept them.

    So a row is reported as sent only once both of its correlation ids come
    back from Paycor. Anything else is unverified, which is a different thing
    from failed and has to be read by a human rather than retried blindly.
    """
    accepted, failed, unverified = [], [], []
    objects = []
    for row in queue:
        for obj in punch_objects(row):
            objects.append((row, obj))

    for batch in chunked(objects, MAX_PUNCH_BATCH):
        payload = [obj for _row, obj in batch]
        rows = []
        for row, _obj in batch:
            if row not in rows:
                rows.append(row)
        try:
            tracking_id = paycor.create_punches(payload, legal_entity_id)
        except PaycorReadOnlyError:
            raise
        except (PaycorAPIError, PaycorAuthError) as exc:
            for row in rows:
                row["error"] = str(exc)
            failed.extend(rows)
            log.error("Batch of %s punch objects rejected outright: %s", len(payload), exc)
            log.error("Stopping. Nothing further sent.")
            break

        log.info("Batch of %s punch objects accepted, tracking id %s",
                 len(payload), tracking_id)

        if not tracking_id:
            for row in rows:
                row["error"] = "accepted but no tracking id returned; cannot verify"
            unverified.extend(rows)
            continue

        try:
            errors = paycor.punch_errors(tracking_id, legal_entity_id)
        except (PaycorAPIError, PaycorAuthError) as exc:
            for row in rows:
                row["error"] = f"accepted, but the error log could not be read: {exc}"
            unverified.extend(rows)
            log.error("Could not read the punch error log for %s: %s", tracking_id, exc)
            continue

        if errors:
            for row in rows:
                row["error"] = "; ".join(
                    str(e.get("errorDetail") or e.get("punchDetail") or e)[:160]
                    for e in errors[:3]
                )
            failed.extend(rows)
            log.error("Paycor rejected punches in tracking id %s:", tracking_id)
            for entry in errors[:10]:
                log.error("  %s | %s", entry.get("errorDetail"), entry.get("punchDetail"))
            log.error("Stopping. Nothing further sent.")
            break

        for row in rows:
            row["tracking_id"] = tracking_id
        accepted.extend(rows)
        log.info("Tracking id %s reports no errors on %s punch(es); "
                 "verifying they exist.", tracking_id, len(rows))

    # ---- the part an empty error log does not tell you ----
    if not accepted:
        return [], failed, unverified

    # Paycor does not necessarily serve a punch back the instant it accepts it.
    time.sleep(PUBLISH_VERIFY_SETTLE_SECONDS)

    window_start = min(r["punch_in"] for r in accepted).date()
    window_end = max(r["punch_in"] for r in accepted).date() + timedelta(days=1)
    try:
        present = collect_correlations(paycor, accepted, window_start, window_end)
    except (PaycorAPIError, PaycorAuthError) as exc:
        log.error("Could not read the punches back to verify them: %s", exc)
        for row in accepted:
            row["error"] = f"accepted, but could not be read back to verify: {exc}"
        return [], failed, unverified + accepted

    sent = []
    for row in accepted:
        missing = [half for half, cid in (("In", row.get("correlation_in")),
                                          ("Out", row.get("correlation_out")))
                   if not cid or str(cid).lower() not in present]
        if missing:
            row["error"] = (
                f"accepted with no errors, but the {' and '.join(missing)} "
                "punch does not read back from Paycor -- it was not created"
            )
            unverified.append(row)
        else:
            sent.append(row)

    if len(sent) != len(accepted):
        log.error("%s of %s punch(es) were accepted without error but do NOT "
                  "exist in Paycor. An empty error log is not proof of a write.",
                  len(accepted) - len(sent), len(accepted))
    else:
        log.info("Verified: all %s punch(es) read back from Paycor.", len(sent))

    return sent, failed, unverified


# =============================
# MAIN
# =============================
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--date", help="only punches starting on this date (YYYY-MM-DD)")
    parser.add_argument("--reconcile", action="store_true",
                        help="read Paycor's punches and compare; writes nothing")
    parser.add_argument("--publish", action="store_true", help="actually send the punches")
    parser.add_argument("--i-understand-this-is-payroll", action="store_true",
                        dest="payroll_ack",
                        help="required alongside --publish against production")
    parser.add_argument("--json", help="write the plan here")
    parser.add_argument("--limit", type=int,
                        help="send at most this many Traumasoft punches "
                             "(for a first live test)")
    args = parser.parse_args()

    target_date = None
    if args.date:
        try:
            target_date = date.fromisoformat(args.date)
        except ValueError:
            log.error("--date must be YYYY-MM-DD, got %r", args.date)
            return 2

    mode = "PUBLISH" if args.publish else ("RECONCILE" if args.reconcile else "DRY RUN")
    print("=" * 78)
    print("TRAUMASOFT -> PAYCOR TIMECARD PUSH")
    print(f"mode: {mode}    run {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 78)

    if args.publish and paycor_api.IS_PRODUCTION and not args.payroll_ack:
        log.error(
            "Refusing to publish to PRODUCTION without "
            "--i-understand-this-is-payroll. This writes what people are paid."
        )
        return 2

    # ---- Traumasoft side (always needed, always read-only) ----
    ts = TraumasoftAPI()
    try:
        ts.detect_auth_mode()
    except TraumasoftAPIError as exc:
        log.error("Could not authenticate to Traumasoft: %s", exc)
        return 1

    log.info("Reading Traumasoft shifts, employees and a day of trips...")
    shifts = ts.list_shifts()
    ts_employees = ts.list_employees()
    legs = ts.get_trips(datetime.now().date(), range_days=1)
    offset = R.resolve_shift_offset(legs)
    now = R.tenant_now(offset)

    overrides = load_overrides()
    if overrides:
        log.info("Loaded %s Paycor employee override(s).", len(overrides))

    # ---- Paycor side ----
    paycor = None
    paycor_roster, activity_types = [], []
    need_paycor = args.reconcile or args.publish
    try:
        paycor = PaycorClient(read_only=not args.publish)
    except PaycorAuthError as exc:
        if need_paycor:
            log.error("Paycor credentials are required for this mode: %s", exc)
            return 1
        log.warning("No Paycor credentials, so the plan cannot be resolved: %s", exc)

    if paycor is not None:
        try:
            paycor_roster = paycor.list_employees()
            activity_types = paycor.list_activity_types()
            log.info("Paycor roster: %s employees, %s activity types.",
                     len(paycor_roster), len(activity_types))
        except (PaycorAPIError, PaycorAuthError) as exc:
            log.error("Could not read from Paycor: %s", exc)
            if need_paycor:
                return 1

    activity_type_id, activity_note = resolve_activity_type(activity_types)
    if activity_type_id:
        activity_note = f"{activity_type_id} (by {activity_note})"

    paycor_index = paycor_employee_index(paycor_roster)
    sendable, refused = build_plan(
        shifts, ts_employees, offset, now, overrides, paycor_index,
        activity_type_id, target_date,
    )
    report_plan(sendable, refused, target_date, activity_note)

    recon = None
    known_correlations = set()
    if paycor is not None and need_paycor and sendable:
        window_start = min(r["punch_in"] for r in sendable).date()
        window_end = max(r["punch_in"] for r in sendable).date() + timedelta(days=1)
        try:
            timecards = paycor.get_timecard_punches(
                window_start.isoformat(), window_end.isoformat()
            )
            known_correlations = collect_correlations(
                paycor, sendable, window_start, window_end
            )
            log.info("Paycor holds %s timecard record(s) and %s correlated punch(es) "
                     "in %s..%s.", len(timecards), len(known_correlations),
                     window_start, window_end)
        except (PaycorAPIError, PaycorAuthError) as exc:
            log.error("Could not read Paycor punches: %s", exc)
            return 1
        recon = report_reconciliation(
            sendable, index_timecards(timecards), known_correlations
        )

    report_contract(activity_note)

    # ---- Publish ----
    sent, failed, unverified, skipped = [], [], [], []
    if args.publish:
        if sendable and recon is None:
            log.error(
                "Refusing to publish without a reconciliation read. Without "
                "knowing what Paycor already holds, a re-run would duplicate "
                "punches and double-pay."
            )
            return 2

        queue = []
        for row in sendable:
            if str(row["correlation_in"]).lower() in known_correlations:
                skipped.append(row)
            else:
                queue.append(row)

        if args.limit:
            held = queue[args.limit:]
            queue = queue[:args.limit]
            if held:
                log.info("--limit %s: sending %s, holding %s back.",
                         args.limit, len(queue), len(held))

        print("\n" + "=" * 78)
        print(f"PUBLISHING {len(queue)} punch(es) -> {len(queue) * 2} Paycor objects "
              f"to {'PRODUCTION' if paycor_api.IS_PRODUCTION else 'the sandbox'}")
        print("=" * 78)
        sent, failed, unverified = publish(paycor, queue)
        print(f"\n  landed {len(sent)}   rejected {len(failed)}   "
              f"unverified {len(unverified)}   already present {len(skipped)}")
        if unverified:
            print("\n  'unverified' means Paycor accepted the punch and raised no")
            print("  error, but it could not be confirmed to exist afterwards --")
            print("  either the read-back failed, or Paycor created nothing. An")
            print("  accepted batch with an empty error log is not proof of a")
            print("  write; only reading the punch back is.")
            print("\n  Check each of these in Paycor before re-running. Paycor")
            print("  refuses an identical re-send, so a re-run cannot double")
            print("  them, but it will not fix them either:")
            for row in unverified[:20]:
                print(f"    {row.get('employee_name') or row.get('user_id')}"
                      f"  {row['punch_in']:%m-%d %H:%M} - {row['punch_out']:%H:%M}"
                      f"  {row.get('error', '')[:90]}")
            if len(unverified) > 20:
                print(f"    ... and {len(unverified) - 20} more (see --json)")
    elif sendable:
        print(f"\n  Dry run: nothing was sent. {len(sendable)} punch(es) are ready.")
        if not args.reconcile:
            print("  Next: --reconcile to compare against Paycor before writing.")

    if args.json:
        payload = {
            "mode": mode,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "target_date": target_date.isoformat() if target_date else None,
            "write_contract": describe_write_contract(),
            "activity_type": activity_note,
            "reconciliation": {k: v for k, v in (recon or {}).items()
                               if k != "missing_rows"},
            "sendable": [jsonable(r) for r in sendable],
            "refused": [jsonable(r) for r in refused],
            "sent": [jsonable(r) for r in sent],
            "failed": [jsonable(r) for r in failed],
            "unverified": [jsonable(r) for r in unverified],
        }
        with open(args.json, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, default=str)
        print(f"\nPlan written to {args.json}")

    if failed or unverified:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
