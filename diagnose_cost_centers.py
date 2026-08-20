"""
Explain the "No Cost Center Assigned" bucket.

Trips carry no cost center -- the API's only trip enrichment is shift_name,
trip_status, trip_timestamp and post_id -- so a leg is attributed through
shift_name -> crew -> employee.cost_center_name, accumulated in
state/shift_cost_center_map.json. A leg lands in the bucket when that chain
breaks, and it can break in several places that want telling apart:

  * the leg names no shift at all, which is normal for a call cancelled
    before it was ever assigned to a unit;
  * it names a shift the map has never seen;
  * it names a shift the map has seen, but whose crew carry no cost center.

This counts each case, and measures how much a vehicle-based fallback would
recover before one is built.

Read-only.

Usage:
    python diagnose_cost_centers.py [YYYY-MM-DD]
"""

import sys
import logging
import argparse
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

from traumasoft_api import TraumasoftAPI
import traumasoft_reports as R

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("diagnose-cc")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("metrics_date", nargs="?", default=None)
    args = parser.parse_args()
    target = (datetime.strptime(args.metrics_date, "%Y-%m-%d").date()
              if args.metrics_date else date.today() - timedelta(days=1))

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("%s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Credentials rejected under every signing scheme.")
        return 1

    log.info("Fetching trips for %s ...", target)
    legs = api.get_trips(target, range_days=1)
    log.info("Fetching shifts ...")
    shifts = api.list_shifts()
    log.info("Fetching employees ...")
    employees = api.list_employees()
    log.info("%s legs, %s shift rows, %s employees", len(legs), len(shifts), len(employees))

    cc_map = R.CostCenterMap()
    cc_map.update(shifts, employees)

    feed_names = {s.get("shift_name") for s in shifts if s.get("shift_name")}
    emp_cc = {
        str(e.get("user_id")): e.get("cost_center_name")
        for e in employees if e.get("user_id")
    }
    crew_by_shift = defaultdict(list)
    for s in shifts:
        if s.get("shift_name"):
            crew_by_shift[s["shift_name"]].append(str(s.get("user_id")))

    resolved = 0
    no_shift_name = Counter()
    unknown_shift = Counter()
    known_but_blank = Counter()
    by_status_unresolved = Counter()
    vehicle_of_unresolved = []
    vehicle_cc = defaultdict(Counter)

    for leg in legs:
        name = (leg.get("shift_name") or "").strip()
        centre = cc_map.resolve(leg.get("shift_name"))
        status = (leg.get("trip_status") or "(none)").strip()
        if centre:
            resolved += 1
            if leg.get("vehicle_id"):
                vehicle_cc[str(leg["vehicle_id"])][centre] += 1
            continue

        by_status_unresolved[status] += 1
        if leg.get("vehicle_id"):
            vehicle_of_unresolved.append(str(leg["vehicle_id"]))
        if not name:
            no_shift_name[status] += 1
        elif name not in cc_map.counts:
            unknown_shift[name] += 1
        else:
            known_but_blank[name] += 1

    total = len(legs)
    unresolved = total - resolved
    print()
    print("=" * 72)
    print(f"COST CENTER ATTRIBUTION  {target}")
    print("=" * 72)
    print(f"\n  {total} legs: {resolved} resolved, {unresolved} unresolved "
          f"({unresolved / total:.0%})" if total else "\n  no legs")

    print("\n  Why the unresolved ones failed")
    print(f"    no shift_name on the leg          {sum(no_shift_name.values()):>5}")
    print(f"    shift_name the map never saw      {sum(unknown_shift.values()):>5}")
    print(f"    shift known, but crew have no CC  {sum(known_but_blank.values()):>5}")

    if no_shift_name:
        print("\n  Legs with no shift_name, by status:")
        for status, n in no_shift_name.most_common(10):
            print(f"    {status:<28} {n:>5}")

    if unknown_shift:
        print("\n  Shift names the map has never seen:")
        print(f"    {'shift_name':<40} {'legs':>5}  in today's shift feed?")
        for name, n in unknown_shift.most_common(20):
            crew = crew_by_shift.get(name, [])
            with_cc = sum(1 for u in crew if emp_cc.get(u))
            present = (f"yes, {len(crew)} crew, {with_cc} with a cost center"
                       if name in feed_names else "NO -- outside the rolling window")
            print(f"    {name:<40} {n:>5}  {present}")

    if known_but_blank:
        print("\n  Shifts the map knows but that resolve to nothing:")
        for name, n in known_but_blank.most_common(20):
            print(f"    {name:<40} {n:>5}")

    print("\n  Unresolved by trip status")
    for status, n in by_status_unresolved.most_common(12):
        # Ask the report's own predicate rather than restating its rule, so
        # this cannot drift from what the workbooks actually count.
        probe = {"trip_status": "" if status == "(none)" else status}
        if not R.has_status(probe):
            note = ("no status -- counted" if R.COUNT_STATUSLESS_LEGS
                    else "no status -- not counted as a run")
        elif R.is_run(probe):
            note = "counts as a run"
        else:
            note = "cancelled/disregarded"
        print(f"    {status:<28} {n:>5}   ({note})")

    # How much would falling back to the vehicle's own cost center recover?
    recoverable = sum(1 for vid in vehicle_of_unresolved if vehicle_cc.get(vid))
    print("\n  Vehicle fallback")
    print(f"    unresolved legs naming a vehicle              {len(vehicle_of_unresolved):>5}")
    print(f"    ... whose vehicle resolves elsewhere today    {recoverable:>5}")
    if unresolved:
        print(f"    would recover {recoverable / unresolved:.0%} of the bucket")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
