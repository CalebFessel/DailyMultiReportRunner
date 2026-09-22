"""
Exercise the Paycor write path end to end, in the sandbox, reversibly.

Everything validated so far is read-side. CreatePunches has never run, which
means the first POST would otherwise be against real payroll -- and its failure
modes are quiet ones. A 202 is an acknowledgement, not a result: Paycor
validates asynchronously and a batch it throws away looks identical to a batch
it keeps until you read the error log.

This probe answers the questions that matter before that first production
write, using sandbox employees rather than the Traumasoft mapping (a different
tenant has different guids, so the mapping cannot be tested here and is not
the point):

  1. Does Paycor accept the punch objects this client builds?
  2. Does the tracking id resolve, and is the error log empty?
  3. Do the punches read back, carrying the correlationId that was sent?
  4. Does sending the SAME correlation ids twice duplicate the punches?

Question 4 is the one worth the whole exercise. The push treats correlationId
as an idempotency key -- that is what lets a re-run recognise its own writes
instead of paying someone twice. That assumption has never been tested against
Paycor, and the safe place to find out it is wrong is a tenant where nobody
gets paid.

Refuses to run against production, unconditionally. There is no override flag,
because there is no version of this that should touch payroll.

    python sandbox_punch_probe.py                 # create, verify, clean up
    python sandbox_punch_probe.py --twice         # also test idempotency
    python sandbox_punch_probe.py --keep          # leave the punches behind
"""

import sys
import json
import argparse
from datetime import datetime, date, timedelta

# traumasoft_api loads .env as an import side effect; paycor_api reads
# PAYCOR_ENVIRONMENT at import time to choose its base url. This order, or a
# .env saying sandbox still talks to production.
import traumasoft_api  # noqa: F401
import paycor_api
from paycor_api import (
    PaycorClient,
    PaycorAPIError,
    PaycorAuthError,
    correlation_id,
)

PUNCH_IN_HOUR = 8
PUNCH_OUT_HOUR = 16


def pick_employee(roster, wanted_id=None):
    """
    An employee to punch against, and why that one.

    Needs a department: CreatePunches requires departmentId and Paycor supplies
    no default, so an employee without one cannot be punched at all.
    """
    if wanted_id:
        for emp in roster:
            if str(emp.get("id")) == str(wanted_id):
                dept = (emp.get("department") or {}).get("id")
                if not dept:
                    return None, f"employee {wanted_id} names no department"
                return emp, "chosen with --employee-id"
        return None, f"employee {wanted_id} is not in the sandbox roster"

    for emp in roster:
        if emp.get("id") and (emp.get("department") or {}).get("id"):
            return emp, "first sandbox employee carrying a department"
    return None, "no sandbox employee carries a department id"


def describe(emp):
    name = " ".join(p for p in (emp.get("firstName"), emp.get("lastName")) if p)
    number = emp.get("employeeNumber") or emp.get("badgeNumber") or "?"
    return f"{name or '(unnamed)'} (#{number})"


def read_back(paycor, employee_id, day, correlations):
    """The probe's punches as Paycor now holds them, keyed by correlation id."""
    window_start = (day - timedelta(days=1)).isoformat()
    window_end = (day + timedelta(days=1)).isoformat()
    try:
        punches = paycor.get_employee_punches(employee_id, window_start, window_end)
    except PaycorAPIError as exc:
        print(f"  Could not read punches back: {exc}")
        return {}

    found = {}
    for punch in punches:
        cid = str(punch.get("correlationId") or "").lower()
        if cid in correlations:
            found.setdefault(cid, []).append(punch)
    return found


def send(paycor, punches, label):
    """POST one batch and resolve its tracking id into an actual result."""
    print(f"\n  {label}: sending {len(punches)} punch object(s)...")
    try:
        tracking_id = paycor.create_punches(punches)
    except (PaycorAPIError, PaycorAuthError) as exc:
        print(f"  REJECTED outright: {exc}")
        return None, False

    if not tracking_id:
        print("  Accepted, but no tracking id came back -- errors for this")
        print("  batch cannot be checked, which is itself a finding.")
        return None, False

    print(f"  Accepted (202), tracking id {tracking_id}")

    try:
        errors = paycor.punch_errors(tracking_id)
    except PaycorAPIError as exc:
        print(f"  Could not read the error log: {exc}")
        return tracking_id, False

    if errors:
        print(f"  Paycor rejected {len(errors)} punch(es) from this batch:")
        for entry in errors:
            print(f"    {entry.get('errorDetail')} | {entry.get('punchDetail')}")
        return tracking_id, False

    print("  Error log is empty -- the batch landed.")
    return tracking_id, True


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--employee-id", help="sandbox employee guid to punch against")
    ap.add_argument("--date", help="punch date, YYYY-MM-DD (default: 7 days ago)")
    ap.add_argument("--tag", help="correlation tag; the same tag rebuilds the same "
                                  "correlation ids (default: derived from the date)")
    ap.add_argument("--twice", action="store_true",
                    help="send the identical batch a second time to test idempotency")
    ap.add_argument("--keep", action="store_true",
                    help="leave the punches in the sandbox instead of deleting them")
    args = ap.parse_args()

    print("=" * 78)
    print("PAYCOR SANDBOX WRITE PROBE")
    print(f"environment: {paycor_api.ENVIRONMENT}")
    print("=" * 78)

    # The one guard that matters. No flag overrides this.
    if paycor_api.IS_PRODUCTION:
        print("\nPAYCOR_ENVIRONMENT is production. This probe writes punches and")
        print("then deletes them; neither belongs in a tenant where people are")
        print("paid. Set PAYCOR_ENVIRONMENT=sandbox and supply the sandbox")
        print("credentials (PAYCOR_SANDBOX_*).")
        return 2

    if args.date:
        try:
            day = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print(f"--date must be YYYY-MM-DD, got {args.date!r}")
            return 2
    else:
        # A week back, so the probe cannot collide with anything a person is
        # looking at today.
        day = date.today() - timedelta(days=7)

    tag = args.tag or f"sandbox-probe-{day.isoformat()}"

    try:
        paycor = PaycorClient(read_only=False)
    except PaycorAuthError as exc:
        print(f"\nCould not build a sandbox client: {exc}")
        return 1

    print(f"base url:    {paycor.base_url}")
    print(f"entity:      {paycor.legal_entity_id or '(not set)'}")
    print(f"punch date:  {day.isoformat()}")
    print(f"tag:         {tag}")

    try:
        roster = paycor.list_employees()
        activity_types = paycor.list_activity_types()
    except (PaycorAPIError, PaycorAuthError) as exc:
        print(f"\nCould not read the sandbox tenant: {exc}")
        return 1

    print(f"\nsandbox roster: {len(roster)} employee(s), "
          f"{len(activity_types)} activity type(s)")

    if not roster:
        print("\nThe sandbox roster is empty, so there is no employee to punch")
        print("against. Ask Paycor to seed the sandbox tenant, or supply a")
        print("guid with --employee-id if you know one.")
        return 1
    if not activity_types:
        print("\nThe sandbox tenant defines no activity types. CreatePunches")
        print("requires an activityTypeId and Paycor supplies no default.")
        return 1

    employee, why = pick_employee(roster, args.employee_id)
    if employee is None:
        print(f"\nNo usable employee: {why}")
        return 1

    employee_id = str(employee["id"])
    department_id = (employee.get("department") or {}).get("id")
    activity = activity_types[0]
    activity_id = str(activity["id"])

    print(f"employee:    {describe(employee)}  [{why}]")
    print(f"department:  {department_id}")
    print(f"activity:    {activity.get('name')} ({activity_id})")

    # Same derivation the real push uses, so this exercises the actual
    # idempotency mechanism rather than a lookalike.
    cid_in = correlation_id(tag, "in")
    cid_out = correlation_id(tag, "out")
    correlations = {cid_in.lower(), cid_out.lower()}

    punches = [
        paycor_api.PaycorClient.build_punch(
            employee_id, department_id, activity_id,
            datetime.combine(day, datetime.min.time()).replace(hour=PUNCH_IN_HOUR),
            "In", note=f"probe {tag}", correlation=cid_in,
        ),
        paycor_api.PaycorClient.build_punch(
            employee_id, department_id, activity_id,
            datetime.combine(day, datetime.min.time()).replace(hour=PUNCH_OUT_HOUR),
            "Out", note=f"probe {tag}", correlation=cid_out,
        ),
    ]

    print("\n" + "-" * 78)
    print("WHAT WILL BE SENT")
    print("-" * 78)
    for obj in punches:
        print("  " + json.dumps(obj))

    # ---- 1. first write ----
    print("\n" + "-" * 78)
    print("FIRST WRITE")
    print("-" * 78)
    _tracking, landed = send(paycor, punches, "first batch")
    if not landed:
        print("\nThe first batch did not land cleanly. Stopping before the")
        print("idempotency test, which would only add noise.")
        return 1

    first = read_back(paycor, employee_id, day, correlations)
    print(f"\n  Read back {sum(len(v) for v in first.values())} punch(es) "
          f"carrying the probe's correlation ids.")
    for cid, entries in sorted(first.items()):
        for entry in entries:
            print(f"    {entry.get('punchStatusType'):<4} "
                  f"{entry.get('punchDateTime')}  punchId={entry.get('punchId')}")
    if not first:
        print("    None. Paycor accepted the batch and reported no errors, but")
        print("    the punches do not read back -- correlationId may not be")
        print("    returned on this endpoint, which would break the push's")
        print("    ability to recognise its own writes.")

    # ---- 2. idempotency ----
    second = {}
    if args.twice:
        print("\n" + "-" * 78)
        print("SECOND WRITE (identical correlation ids)")
        print("-" * 78)
        send(paycor, punches, "same batch again")
        second = read_back(paycor, employee_id, day, correlations)

        before = sum(len(v) for v in first.values())
        after = sum(len(v) for v in second.values())
        print(f"\n  Punches carrying these correlation ids: {before} -> {after}")
        if after == before:
            print("  Paycor treated the correlation id as an idempotency key.")
            print("  A re-run of the same day is safe.")
        else:
            print("  DUPLICATED. Paycor does NOT dedupe on correlationId, so a")
            print("  re-run would punch people twice. The push must read")
            print("  existing punches and skip its own before sending -- which")
            print("  is what --reconcile does, and it is now load-bearing")
            print("  rather than a convenience.")

    # ---- 3. cleanup ----
    print("\n" + "-" * 78)
    print("CLEANUP")
    print("-" * 78)
    if args.keep:
        print("  --keep given; the probe's punches are still in the sandbox.")
    else:
        # Only ever the punches this probe created, identified by correlation
        # id. Never a blanket delete over a time window.
        latest = second or first
        punch_ids = [e.get("punchId") for entries in latest.values()
                     for e in entries if e.get("punchId")]
        if not punch_ids:
            print("  Nothing to delete -- no punch ids read back.")
        else:
            try:
                paycor.delete_punches(employee_id, punch_ids)
                print(f"  Deleted {len(punch_ids)} punch(es) this probe created.")
            except (PaycorAPIError, PaycorAuthError) as exc:
                print(f"  Could not delete: {exc}")
                print(f"  Left behind: {', '.join(str(p) for p in punch_ids)}")
                return 1

            remaining = read_back(paycor, employee_id, day, correlations)
            left = sum(len(v) for v in remaining.values())
            print(f"  Verified: {left} punch(es) with these correlation ids remain.")

    print("\n" + "=" * 78)
    print("The write path works end to end in the sandbox." if landed
          else "The write path did not complete.")
    print("=" * 78)
    return 0


if __name__ == "__main__":
    sys.exit(main())
