"""
Push Traumasoft punches into Paycor as timecard punches.

The goal this serves: crews clock in twice, only Paycor pays them, so nobody
closes a Traumasoft punch and every unit-hour figure in the report bundle rests
on a record with no incentive behind it. Making Traumasoft the clock that pays
fixes the data at its source. This is the mechanism.

There are three modes and they are meant to be used in order.

  1. DRY RUN (default) -- build the plan and print it. Nothing leaves the
     machine except reads. Shows exactly which punch would be sent for which
     person, and everything that was refused and why.

  2. --reconcile -- read what Paycor already holds for the same window and
     compare it punch by punch against Traumasoft. This is the test that
     matters before any write: it proves the employee mapping resolves, the
     clocks agree, and the two systems are describing the same shifts. Still
     read-only on both sides.

  3. --publish -- send them. Requires the Paycor client to be constructed
     writable, and refuses production unless you have said so twice.

Safety rails, none of which are configurable away:

  * An open punch is never sent. A punch with no clock-out is not a payable
    record, and the shift-end fallback that makes it usable for reporting would
    here mean inventing the end of somebody's paid day.
  * An employee who cannot be mapped unambiguously is refused, not guessed.
  * A punch Paycor already holds is skipped, so a re-run does not double-pay.
    That check needs --reconcile data; without it a publish refuses outright
    rather than risking duplicates.
  * Production needs PAYCOR_ENVIRONMENT=production *and* --i-understand-this-is-payroll.

The shift feed only ever returns today-1..today+2, so this can only push a
narrow window and must run daily to cover a pay period. A day missed is a day
whose punches the API will not hand back.

Usage:
    python push_paycor_timecards.py                          # dry run
    python push_paycor_timecards.py --reconcile              # compare, write nothing
    python push_paycor_timecards.py --date 2026-09-09 --json plan.json
    python push_paycor_timecards.py --reconcile --publish    # for real
"""

import os
import sys
import json
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
    describe_write_contract,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("push-paycor")

OVERRIDES_FILE = os.path.join(
    os.getenv("TS_STATE_DIR", "state"), "paycor_employee_overrides.json"
)

# Two punches this far apart or closer are treated as the same punch when
# reconciling. Clocks drift and the two systems round differently; without a
# tolerance every punch looks like a mismatch.
MATCH_TOLERANCE_MINUTES = float(os.getenv("PAYCOR_MATCH_TOLERANCE_MINUTES", "2"))

# A punch longer than this is refused as implausible rather than sent. A
# runaway punch that reached payroll would pay a multi-day shift.
MAX_PUNCH_HOURS = float(os.getenv("PAYCOR_MAX_PUNCH_HOURS", "24"))


# =============================
# MAPPING
# =============================
def load_overrides():
    """Traumasoft employee_num (or user_id) -> Paycor employee id."""
    data = R.load_state_file(OVERRIDES_FILE, "Paycor employee overrides") or {}
    raw = data.get("overrides", data) if isinstance(data, dict) else {}
    return {str(k).strip().lower(): str(v).strip() for k, v in raw.items() if v}


def paycor_employee_index(paycor_employees):
    """
    Index the Paycor roster by every identifier a Traumasoft record might name.

    Paycor's employee payload spells its own fields several ways depending on
    the endpoint, so candidates are gathered rather than assumed.
    """
    by_number = defaultdict(list)
    for emp in paycor_employees:
        emp_id = (
            emp.get("employeeId") or emp.get("id") or emp.get("employeeUuid")
        )
        if emp_id is None:
            continue
        for key in ("employeeNumber", "employee_number", "employeeNo", "badgeNumber"):
            value = str(emp.get(key) or "").strip()
            if value:
                by_number[value.lower()].append(str(emp_id))
    return by_number


def resolve_employee(ts_employee, overrides, paycor_index):
    """
    (paycor_employee_id, how) or (None, why not).

    An override always wins, being a decision rather than an observation. Then
    employee_num against the Paycor roster. An ambiguous number is refused --
    two people sharing a payroll number is exactly the case where a guess
    pays the wrong person.
    """
    user_id = str(ts_employee.get("user_id") or "").strip()
    number = str(ts_employee.get("employee_num") or "").strip()

    for key in (number.lower(), f"user_id:{user_id}".lower(), user_id.lower()):
        if key and key in overrides:
            return overrides[key], "override"

    if not number:
        return None, "Traumasoft employee carries no employee_num"

    if not paycor_index:
        # No roster to check against -- a dry run without Paycor credentials.
        # Carry the number through so the plan is still legible, but say so.
        return number, "employee_num (unchecked, no Paycor roster loaded)"

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
    Every punch that could be pushed, with the reason if it cannot be.

    Nothing is filtered out silently: a refused punch stays in the plan with a
    `refused` reason, because the whole value of the dry run is seeing what
    would not go and why.
    """
    out = []
    for shift in shifts:
        if shift.get("deleted"):
            continue
        profile = R.profile_name(shift)
        shift_start = R.parse_shift_ts(shift.get("start_time"), offset)
        shift_end = R.parse_shift_ts(shift.get("end_time"), offset)
        user_id = shift.get("user_id")

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
                "user_id": user_id,
                "profile": profile,
                "vehicle_name": shift.get("vehicle_name"),
                "shift_start": shift_start,
                "shift_end": shift_end,
                "punch_in": start,
                "punch_out": end,
                "refused": refused,
            })
    return out


def build_plan(shifts, ts_employees, offset, now, overrides, paycor_index,
               target_date=None):
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

        paycor_id, how = resolve_employee(ts_emp, overrides, paycor_index)
        row["paycor_employee_id"] = paycor_id
        row["mapped_by"] = how
        if paycor_id is None:
            row["refused"] = how
            refused.append(row)
        else:
            sendable.append(row)

    return sendable, refused


# =============================
# RECONCILE
# =============================
def parse_paycor_time(value):
    """
    Paycor timestamps arrive in a few shapes; take the first that parses.

    Sub-second precision is dropped rather than carried. A punch is a clock
    event to the minute; keeping microseconds would put a precision on payroll
    times that nothing behind them supports, and makes two records of the same
    punch compare unequal for a reason no one could act on.
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


def index_paycor_punches(paycor_punches):
    """paycor_employee_id -> [(in, out), ...] for the reconciliation window."""
    index = defaultdict(list)
    for row in paycor_punches:
        emp_id = str(
            row.get("employeeId") or row.get("employee_id") or row.get("id") or ""
        ).strip()
        if not emp_id:
            continue
        punch_in = parse_paycor_time(
            row.get("punchInTime") or row.get("inTime") or row.get("startTime")
        )
        punch_out = parse_paycor_time(
            row.get("punchOutTime") or row.get("outTime") or row.get("endTime")
        )
        if punch_in:
            index[emp_id].append((punch_in, punch_out))
    return index


def already_present(row, paycor_index):
    """Is this punch already in Paycor, within the clock tolerance?"""
    tolerance = timedelta(minutes=MATCH_TOLERANCE_MINUTES)
    for existing_in, existing_out in paycor_index.get(str(row["paycor_employee_id"]), []):
        if abs(existing_in - row["punch_in"]) <= tolerance:
            return True, (existing_in, existing_out)
    return False, None


def report_reconciliation(sendable, paycor_index):
    print("\n" + "=" * 78)
    print("RECONCILIATION -- Traumasoft against what Paycor already holds")
    print("=" * 78)

    if not paycor_index:
        print("  Paycor returned no punches for this window.")
        print("  Either the window is genuinely empty, the legal entity id is")
        print("  wrong, or this key cannot read timecards. Resolve that before")
        print("  reading anything below as agreement.")
        return {"matched": 0, "missing": 0, "drifted": 0}

    matched, missing, drifted = [], [], []
    for row in sendable:
        present, existing = already_present(row, paycor_index)
        if not present:
            missing.append(row)
            continue
        row["paycor_existing"] = existing
        # Same clock-in, but does the clock-out agree?
        if row["punch_out"] and existing[1]:
            gap = abs((row["punch_out"] - existing[1]).total_seconds()) / 60.0
            if gap > MATCH_TOLERANCE_MINUTES:
                row["drift_minutes"] = round(gap, 1)
                drifted.append(row)
                continue
        matched.append(row)

    total = len(sendable)
    print(f"  Traumasoft punches in window : {total}")
    print(f"  Already in Paycor, agreeing   : {len(matched)}")
    print(f"  Already in Paycor, differing  : {len(drifted)}")
    print(f"  Not in Paycor                 : {len(missing)}")

    if drifted:
        print("\n  Punches both systems hold with clock-outs that disagree. These")
        print("  are the ones to understand before trusting either clock:")
        for row in sorted(drifted, key=lambda r: -r["drift_minutes"])[:15]:
            print(f"    {row['drift_minutes']:>7.1f} min  "
                  f"{(row['employee_name'] or row['user_id']):<26} "
                  f"TS {row['punch_in']:%m-%d %H:%M}-{row['punch_out']:%H:%M}  "
                  f"Paycor -{row['paycor_existing'][1]:%H:%M}")

    if missing:
        print(f"\n  Punches Traumasoft has and Paycor does not ({len(missing)}).")
        print("  On a shadow run these are what a publish would add:")
        for row in missing[:15]:
            print(f"    {(row['employee_name'] or row['user_id']):<26} "
                  f"{row['punch_in']:%m-%d %H:%M} - "
                  f"{row['punch_out']:%H:%M}  {row['profile']}")
        if len(missing) > 15:
            print(f"    ... and {len(missing) - 15} more")

    if total and len(matched) == total:
        print("\n  Every Traumasoft punch is already in Paycor and the clocks agree.")
        print("  That is the result you want before cutting over: the two systems")
        print("  describe the same work, so making Traumasoft authoritative")
        print("  changes who is trusted, not what anyone is paid.")

    return {"matched": len(matched), "missing": len(missing), "drifted": len(drifted)}


# =============================
# OUTPUT
# =============================
def report_plan(sendable, refused, target_date):
    print("\n" + "=" * 78
          + f"\nPLAN{f' for {target_date}' if target_date else ''}\n" + "=" * 78)
    print(f"  Punches that would be sent : {len(sendable)}")
    print(f"  Refused                    : {len(refused)}")

    if sendable:
        payable = sum(
            (r["punch_out"] - r["punch_in"]).total_seconds() / 3600.0 for r in sendable
        )
        crew = len({r["user_id"] for r in sendable})
        print(f"  Payable hours represented  : {payable:.2f} across {crew} crew")

    if refused:
        buckets = defaultdict(int)
        for row in refused:
            buckets[row["refused"]] += 1
        print("\n  Why punches were refused:")
        for reason, count in sorted(buckets.items(), key=lambda b: -b[1]):
            print(f"    {count:>5}  {reason}")
        open_missed = buckets.get("punch is open (missed punch-out)", 0)
        if open_missed:
            print(f"\n  {open_missed} of those are missed punch-outs -- work that")
            print("  happened and cannot be paid from this feed. That is the")
            print("  problem the cutover is meant to solve, quantified.")

    if sendable:
        print("\n  First few, as they would be sent:")
        for row in sendable[:8]:
            print(f"    {(row['employee_name'] or row['user_id']):<24} "
                  f"paycor={row['paycor_employee_id']:<12} "
                  f"{row['punch_in']:%m-%d %H:%M} - {row['punch_out']:%H:%M}"
                  f"  ({row['mapped_by']})")


def report_contract():
    contract = describe_write_contract()
    print("\n" + "=" * 78)
    print("WRITE CONTRACT -- unverified")
    print("=" * 78)
    print(f"  environment : {contract['environment']}")
    print(f"  base url    : {contract['base_url']}")
    print(f"  request     : {contract['method']} {contract['path']}")
    print(f"  body fields : {contract['body_fields']}")
    print(f"  time format : {contract['time_format']}")
    print("\n  developers.paycor.com was unreachable when this was written, so the")
    print("  above is a best reading of secondary sources. Check it against")
    print("  Paycor's own reference before publishing; every part of it is an")
    print("  environment variable, so a correction costs an .env edit.")


def jsonable(row):
    out = {}
    for key, value in row.items():
        if isinstance(value, datetime):
            out[key] = value.isoformat(timespec="minutes")
        elif isinstance(value, tuple):
            out[key] = [v.isoformat(timespec="minutes") if isinstance(v, datetime)
                        else v for v in value]
        else:
            out[key] = value
    return out


# =============================
# MAIN
# =============================
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--date", help="only punches starting on this date (YYYY-MM-DD)")
    parser.add_argument("--reconcile", action="store_true",
                        help="read Paycor's punches and compare; still writes nothing")
    parser.add_argument("--publish", action="store_true",
                        help="actually send the punches")
    parser.add_argument("--i-understand-this-is-payroll", action="store_true",
                        dest="payroll_ack",
                        help="required alongside --publish against production")
    parser.add_argument("--json", help="write the plan here")
    parser.add_argument("--limit", type=int,
                        help="send at most this many punches (for a first live test)")
    args = parser.parse_args()

    target_date = None
    if args.date:
        try:
            target_date = date.fromisoformat(args.date)
        except ValueError:
            log.error("--date must be YYYY-MM-DD, got %r", args.date)
            return 2

    print("=" * 78)
    print("TRAUMASOFT -> PAYCOR TIMECARD PUSH")
    mode = "PUBLISH" if args.publish else ("RECONCILE" if args.reconcile else "DRY RUN")
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

    # ---- Paycor side (optional in a dry run) ----
    paycor = None
    paycor_roster = []
    paycor_punch_index = {}
    need_paycor = args.reconcile or args.publish
    try:
        paycor = PaycorClient(read_only=not args.publish)
    except PaycorAuthError as exc:
        if need_paycor:
            log.error("Paycor credentials are required for this mode: %s", exc)
            return 1
        log.warning("No Paycor credentials, so the plan is unchecked: %s", exc)

    if paycor is not None:
        try:
            paycor_roster = paycor.list_employees()
            log.info("Paycor roster: %s employees.", len(paycor_roster))
        except (PaycorAPIError, PaycorAuthError) as exc:
            log.error("Could not read the Paycor roster: %s", exc)
            if need_paycor:
                return 1

    paycor_index = paycor_employee_index(paycor_roster)
    sendable, refused = build_plan(
        shifts, ts_employees, offset, now, overrides, paycor_index, target_date
    )
    report_plan(sendable, refused, target_date)

    recon = None
    if paycor is not None and need_paycor and sendable:
        window_start = min(r["punch_in"] for r in sendable).date()
        window_end = max(r["punch_in"] for r in sendable).date() + timedelta(days=1)
        try:
            existing = paycor.get_timecard_punches(
                window_start.isoformat(), window_end.isoformat()
            )
            paycor_punch_index = index_paycor_punches(existing)
            log.info("Paycor holds %s punch record(s) in %s..%s.",
                     len(existing), window_start, window_end)
        except (PaycorAPIError, PaycorAuthError) as exc:
            log.error("Could not read Paycor punches: %s", exc)
            return 1
        recon = report_reconciliation(sendable, paycor_punch_index)

    report_contract()

    # ---- Publish ----
    sent, failed, skipped = [], [], []
    if args.publish:
        # The guard is that the reconciliation read happened, not that it found
        # anything. A window Paycor genuinely holds nothing for is the normal
        # first push; a window we never asked about is the one that duplicates.
        if sendable and recon is None:
            log.error(
                "Refusing to publish without a reconciliation read. Without "
                "knowing what Paycor already holds, a re-run would duplicate "
                "punches and double-pay."
            )
            return 2

        queue, skipped = [], []
        for row in sendable:
            (skipped if already_present(row, paycor_punch_index)[0] else queue).append(row)
        if args.limit:
            held = queue[args.limit:]
            queue = queue[:args.limit]
            if held:
                log.info("--limit %s: sending %s, holding %s back.",
                         args.limit, len(queue), len(held))

        print("\n" + "=" * 78)
        print(f"PUBLISHING {len(queue)} punch(es) to "
              f"{'PRODUCTION' if paycor_api.IS_PRODUCTION else 'the sandbox'}")
        print("=" * 78)
        for row in queue:
            try:
                paycor.create_punch(
                    row["paycor_employee_id"], row["punch_in"], row["punch_out"]
                )
                sent.append(row)
                log.info("sent %s %s-%s",
                         row["employee_name"] or row["user_id"],
                         f"{row['punch_in']:%m-%d %H:%M}",
                         f"{row['punch_out']:%H:%M}")
            except PaycorReadOnlyError as exc:
                log.error("%s", exc)
                return 1
            except (PaycorAPIError, PaycorAuthError) as exc:
                row["error"] = str(exc)
                failed.append(row)
                log.error("FAILED %s: %s", row["punch_id"], exc)
                # Stop on the first failure rather than hammering payroll with
                # a systematically wrong body. One bad punch is a fix; two
                # hundred is an incident.
                log.error("Stopping after the first failure. Nothing further sent.")
                break
        print(f"\n  sent {len(sent)}   failed {len(failed)}   "
              f"already present {len(skipped)}")
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
            "reconciliation": recon,
            "sendable": [jsonable(r) for r in sendable],
            "refused": [jsonable(r) for r in refused],
            "sent": [jsonable(r) for r in sent],
            "failed": [jsonable(r) for r in failed],
        }
        with open(args.json, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, default=str)
        print(f"\nPlan written to {args.json}")

    if failed:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
