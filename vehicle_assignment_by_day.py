"""
Which of one cost center's vehicles were assigned a call each day, and which were not.

Asked for Boardman, July, per day -- with the not-assigned list counting
Boardman's vehicles only.

    python vehicle_assignment_by_day.py --cost-center Boardman --month 2026-07

RUN THIS SOON. Trips backfill roughly 90 days through GetTrips, and that is an
observation rather than a documented guarantee. On 29 September 2026, 1 July
is exactly 90 days back -- the front of the month is at the edge of the window
and another day falls off it every day. Section 1 reports the coverage it
actually got rather than assuming it got the lot.

A DAY WITH NO LEGS AT ALL IS NOT A QUIET DAY. It is a day outside the backfill
window, and reading it as "nothing was assigned" would invent an idle fleet. A
day where the API returned legs for other cost centers but none for this one is
a real zero. The two are told apart and reported separately.

THREE STATES PER VEHICLE-DAY, NOT TWO:

    assigned            a leg for this cost center
    assigned elsewhere  a leg, but for somewhere else
    not assigned        no leg at all

The middle one matters. A Boardman truck loaned to Cincinnati on the 12th was
not idle, and a two-state report would file it as not assigned and invite
somebody to ask why Boardman had a truck doing nothing.

WHOSE VEHICLES ARE BOARDMAN'S IS THE HARD PART. Cost center is on neither
vehicles nor trips -- only on employees -- so the fleet has to be assembled,
and the not-assigned list is only ever as complete as that assembly. A truck
that sat idle for the whole of July ran no legs in July, so defining the fleet
from July's assignments alone would make exactly the vehicle being asked about
invisible. Sources, best first, each labelled per row:

    1. ran a leg for this cost center during the period
    2. ran one in a lookback window before it (--lookback-days)
    3. its live shift_name maps to this cost center
    4. named in a roster file (--roster), which is the only source that
       cannot miss a truck that has not moved in months

STRICTLY READ-ONLY. GETs only.

PHI: no patient-identifying value is read or printed. Vehicle names, cost
centers and call counts are operational.
"""

import argparse
import json
import logging
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta

import traumasoft_reports as R
import vehicle_service_days as VS

log = logging.getLogger(__name__)

ASSIGNED = "assigned"
ELSEWHERE = "assigned elsewhere"
IDLE = "not assigned"
NO_DATA = "no data"

# One character per state, for the day-by-day matrix a reader scans across.
MATRIX_MARKS = {ASSIGNED: "A", ELSEWHERE: "E", IDLE: ".", NO_DATA: "?"}


# =============================
# THE COST CENTRE
# =============================
def normalize_centre(value):
    """
    A cost center name reduced to what is comparable.

    Employee records carry the legal entity wrapper -- 'Lynx EMS LLC dba Lynx
    Boardman' and 'Boardman' are one station -- so the wrapper is dropped along
    with spacing and punctuation. Same rule as probe_samsara_tags.normalize, so
    the two cannot disagree about what a station is called.
    """
    text = str(value or "").lower()
    text = re.sub(r"\b(lynx|ems|llc|dba|inc|the)\b", " ", text)
    return re.sub(r"[^a-z0-9]+", "", text)


def matching_centres(wanted, known):
    """
    Every cost center name that means the station being asked for.

    Returns (matched, how). Exact first, then normalized, then substring --
    and the names that matched are printed by the caller, because a substring
    match is how 'Columbus' quietly claims Columbus Indiana as well as
    Columbus Ohio. Seeing which names were folded together is the only way to
    catch that before the number ships.
    """
    wanted_raw = str(wanted or "").strip().lower()
    wanted_norm = normalize_centre(wanted)
    if not wanted_norm:
        return [], "nothing asked for"

    exact = sorted({n for n in known if str(n).strip().lower() == wanted_raw})
    if exact:
        return exact, "exact"
    normalized = sorted({n for n in known if normalize_centre(n) == wanted_norm})
    if normalized:
        return normalized, "name match ignoring the legal entity wrapper"
    loose = sorted({n for n in known if wanted_norm in normalize_centre(n)})
    if loose:
        return loose, "SUBSTRING -- check these are all one station"
    return [], "no cost center matched"


# =============================
# THE FLEET
# =============================
def load_roster(path):
    """
    Vehicle names that belong to the cost center, named by hand.

    The only source that cannot miss a truck which has not moved in months.
    Accepts a bare list or {"vehicles": [...]}.
    """
    if not path:
        return set()
    if not os.path.exists(path):
        log.warning("No roster at %s; the fleet comes from the data alone.", path)
        return set()
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, ValueError) as exc:
        log.warning("Could not read %s (%s).", path, exc)
        return set()
    names = payload.get("vehicles") if isinstance(payload, dict) else payload
    return {VS.normalize_name(n) for n in (names or []) if str(n).strip()}


def legs_by_vehicle_day(legs, centres, cost_center_map):
    """
    (normalized vehicle -> {day -> Counter of cost centers}), and the days seen.

    `centres` is the set of cost center names that mean the station asked for.
    Every leg is kept, not just that station's: a leg for somewhere else is
    what tells a loaned truck apart from an idle one.
    """
    per_vehicle = defaultdict(lambda: defaultdict(Counter))
    days_with_legs = set()
    unattributed = 0
    for leg in legs or []:
        day = VS.leg_date(leg)
        if day is None:
            continue
        days_with_legs.add(day)
        name = VS.normalize_name(leg.get("vehicle_name"))
        if not name:
            continue
        centre = cost_center_map.resolve(R.profile_name(leg))
        if centre is None:
            unattributed += 1
            centre = ""
        per_vehicle[name][day][centre] += 1
    return dict(per_vehicle), days_with_legs, unattributed


def build_fleet(period_activity, lookback_activity, roster_names, roster_vehicles,
                centres, cost_center_map):
    """
    normalized vehicle -> how it was placed in this cost center's fleet.

    Ordered best source first. The order is the point: a vehicle placed by a
    leg it ran this period is a stronger claim than one placed by the shift it
    happens to be on today, and the report says which it had so a reader can
    discount the weak ones rather than discovering them later.
    """
    placed = {}
    for name, days in period_activity.items():
        if any(centre in centres for day in days.values() for centre in day):
            placed[name] = "ran for it this period"

    for name, days in (lookback_activity or {}).items():
        if name in placed:
            continue
        if any(centre in centres for day in days.values() for centre in day):
            placed[name] = "ran for it before the period"

    for vehicle in roster_vehicles or []:
        name = VS.normalize_name(vehicle.get("name"))
        if not name or name in placed:
            continue
        centre = cost_center_map.resolve(str(vehicle.get("shift_name") or "").strip())
        if centre in centres:
            placed[name] = "its live shift says so"

    for name in roster_names or ():
        placed.setdefault(name, "named in the roster file")
    return placed


# =============================
# THE GRID
# =============================
def build_grid(fleet, period_activity, days, covered_days, centres):
    """
    {normalized vehicle -> {day -> state}}.

    A day the backfill did not reach is NO_DATA, never IDLE. That distinction
    is the whole difference between reporting a quiet fleet and inventing one.
    """
    grid = {}
    for name in fleet:
        row = {}
        activity = period_activity.get(name, {})
        for day in days:
            if day not in covered_days:
                row[day] = NO_DATA
                continue
            counts = activity.get(day)
            if not counts:
                row[day] = IDLE
            elif any(centre in centres for centre in counts):
                row[day] = ASSIGNED
            else:
                row[day] = ELSEWHERE
        grid[name] = row
    return grid


def display_names(fleet, period_activity, lookback_activity, roster_vehicles, legs):
    """normalized vehicle -> the name to print, preferring the live roster's."""
    names = {}
    for vehicle in roster_vehicles or []:
        key = VS.normalize_name(vehicle.get("name"))
        if key:
            names.setdefault(key, str(vehicle.get("name")))
    for leg in legs or []:
        key = VS.normalize_name(leg.get("vehicle_name"))
        if key:
            names.setdefault(key, str(leg.get("vehicle_name")))
    for key in fleet:
        names.setdefault(key, key)
    return names


def per_day_rows(grid, days, names):
    rows = []
    for day in days:
        states = Counter(grid[name][day] for name in grid)
        idle = sorted(names[n] for n in grid if grid[n][day] == IDLE)
        rows.append({
            "day": day.isoformat(),
            "fleet": len(grid),
            "assigned": states[ASSIGNED],
            "assigned_elsewhere": states[ELSEWHERE],
            "not_assigned": states[IDLE],
            "no_data": states[NO_DATA],
            "not_assigned_vehicles": "; ".join(idle),
        })
    return rows


def per_vehicle_rows(grid, days, names, fleet, period_activity, centres):
    rows = []
    for name, row in grid.items():
        states = Counter(row.values())
        activity = period_activity.get(name, {})
        ours = sum(count for day in activity.values()
                   for centre, count in day.items() if centre in centres)
        # Kept apart from `theirs` deliberately. A leg whose shift profile is
        # not in the cost-center map is a leg nobody can attribute -- it may
        # well have been this station's. Counting it as another station's work
        # would quietly undercount the days this station had a truck out, and
        # the direction of that error is the one that makes a station look
        # idler than it was.
        unknown = sum(count for day in activity.values()
                      for centre, count in day.items() if not centre)
        theirs = sum(count for day in activity.values()
                     for centre, count in day.items()
                     if centre and centre not in centres)
        rows.append({
            "vehicle_name": names[name],
            "in_fleet_because": fleet[name],
            "days_assigned": states[ASSIGNED],
            "days_assigned_elsewhere": states[ELSEWHERE],
            "days_not_assigned": states[IDLE],
            "days_no_data": states[NO_DATA],
            "days_observed": len(days) - states[NO_DATA],
            "legs_for_this_cost_center": ours,
            "legs_for_others": theirs,
            "legs_unattributed": unknown,
        })
    return sorted(rows, key=lambda r: (-r["days_not_assigned"], r["vehicle_name"]))


def matrix_rows(grid, days, names):
    """One row per vehicle, one column per day, for the sheet people scan."""
    rows = []
    for name, row in sorted(grid.items(), key=lambda kv: names[kv[0]]):
        entry = {"vehicle_name": names[name]}
        for day in days:
            entry[day.isoformat()] = MATRIX_MARKS[row[day]]
        rows.append(entry)
    return rows


# =============================
# OUTPUT
# =============================
def print_report(centre_label, matched, how, days, covered_days, grid, names,
                 fleet, period_activity, centres, unattributed, legs_total):
    missing = [day for day in days if day not in covered_days]

    print(f"\n1. WHAT THE BACKFILL ACTUALLY REACHED")
    print("   " + "-" * 66)
    print(f"   Days in the period:          {len(days)}")
    print(f"   Days the API returned legs:  {len(covered_days)}")
    print(f"   Days it returned nothing:    {len(missing)}")
    print(f"   Legs fetched:                {legs_total}")
    if missing:
        print("\n   A day with no legs for ANY cost center is a day outside the")
        print("   backfill window, not a quiet day. Trips reach roughly 90 days")
        print("   and that is observed, not promised. Those days are reported as")
        print("   'no data' and never counted as a vehicle sitting idle.\n")
        for run in VS.compress_days(missing)[:12]:
            print(f"      {run}")
        if min(days) in missing:
            print("\n   THE FRONT OF THE PERIOD IS GONE. Every day that passes loses")
            print("   another. If this month matters, pull it now and keep the CSV.")

    print(f"\n2. WHOSE VEHICLES COUNT AS {centre_label.upper()}'S")
    print("   " + "-" * 66)
    print(f"   Cost center match: {how}")
    for name in matched:
        print(f"      {name}")
    if not matched:
        print("      Nothing matched. Every vehicle below would be empty --")
        print("      check the spelling against --list-cost-centers.")
    print(f"\n   Cost center is on neither vehicles nor trips, so the fleet is")
    print("   assembled. The not-assigned list is only as complete as this:\n")
    print(f"   {'placed by':<34}{'vehicles':>9}")
    print("   " + "-" * 43)
    for reason, count in Counter(fleet.values()).most_common():
        print(f"   {reason:<34}{count:>9}")
    print(f"   {'':<34}{'-' * 9:>9}")
    print(f"   {'total':<34}{len(fleet):>9}")
    if not any(r == "named in the roster file" for r in fleet.values()):
        print("\n   No roster file. A truck that sat idle for the whole period ran")
        print("   no legs in it, so it can only be here via the lookback or its")
        print("   live shift -- and if it has neither, it is missing from the")
        print("   not-assigned list entirely. --roster is the fix.")

    print(f"\n3. BY DAY")
    print("   " + "-" * 66)
    print(f"   {'day':<12}{'fleet':>7}{'assigned':>10}{'elsewhere':>11}"
          f"{'not asgn':>10}{'no data':>9}")
    print("   " + "-" * 61)
    for row in per_day_rows(grid, days, names):
        print(f"   {row['day']:<12}{row['fleet']:>7}{row['assigned']:>10}"
              f"{row['assigned_elsewhere']:>11}{row['not_assigned']:>10}"
              f"{row['no_data']:>9}")

    print(f"\n4. BY VEHICLE")
    print("   " + "-" * 66)
    observed = len(covered_days)
    print(f"   Out of {observed} observed day(s).\n")
    print(f"   {'vehicle':<16}{'asgn':>6}{'elsew':>7}{'idle':>6}{'n/d':>5}"
          f"   {'placed by'}")
    print("   " + "-" * 68)
    for row in per_vehicle_rows(grid, days, names, fleet, period_activity, centres):
        print(f"   {str(row['vehicle_name'])[:15]:<16}{row['days_assigned']:>6}"
              f"{row['days_assigned_elsewhere']:>7}{row['days_not_assigned']:>6}"
              f"{row['days_no_data']:>5}   {row['in_fleet_because']}")

    print(f"\n5. WHAT TO CHECK")
    print("   " + "-" * 66)
    if unattributed:
        print(f"   {unattributed} leg(s) in the period belong to no cost center --")
        print("   their shift profile is not in the map, so nobody can say whose")
        print("   work they were. They may well have been this station's. Add the")
        print("   profiles to state/shift_cost_center_overrides.json.")
        # A vehicle whose ONLY activity is unattributable is a vehicle whose
        # verdict could flip once the map is fixed, so it is named rather than
        # left inside a total.
        rows = per_vehicle_rows(grid, days, names, fleet, period_activity, centres)
        at_risk = [r for r in rows if r["legs_unattributed"]
                   and not r["legs_for_this_cost_center"]
                   and not r["legs_for_others"]]
        if at_risk:
            print(f"\n   {len(at_risk)} of them ran nothing else all period, so their")
            print("   verdict here could change once those profiles are mapped:")
            print("      " + ", ".join(r["vehicle_name"] for r in at_risk)[:200])
    loaned = [n for n in grid if any(s == ELSEWHERE for s in grid[n].values())]
    if loaned:
        print(f"\n   {len(loaned)} vehicle(s) ran for another cost center on at least")
        print("   one day. Those days are 'elsewhere', not 'not assigned' -- the")
        print("   truck was working, just not here.")
        print("      " + ", ".join(names[n] for n in loaned[:15])[:200])
    never = sorted(
        (n for n in grid if not any(s == ASSIGNED for s in grid[n].values())),
        key=lambda n: names[n],
    )
    if never:
        print(f"\n   {len(never)} vehicle(s) were never assigned a "
              f"{centre_label} call in\n   the whole period:")
        print("      " + ", ".join(names[n] for n in never[:20])[:220])
    if "SUBSTRING" in how:
        print("\n   The cost center matched on a substring. Check the names listed")
        print("   in section 2 really are one station before sending this on.")


def write_outputs(day_rows, vehicle_rows, matrix, csv_path=None, xlsx_path=None):
    import pandas as pd

    if csv_path:
        pd.DataFrame(day_rows).to_csv(csv_path, index=False)
        print(f"  Written to {csv_path}")
    if xlsx_path:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            pd.DataFrame(day_rows).to_excel(writer, sheet_name="By Day", index=False)
            pd.DataFrame(vehicle_rows).to_excel(
                writer, sheet_name="By Vehicle", index=False)
            pd.DataFrame(matrix).to_excel(
                writer, sheet_name="Day Grid", index=False)
        print(f"  Written to {xlsx_path}")
        print("  Day Grid legend: A assigned, E assigned elsewhere, "
              ". not assigned, ? no data")


# =============================
# CLI
# =============================
def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--cost-center", required=True,
                        help='The station, e.g. "Boardman".')
    parser.add_argument("--month", help="YYYY-MM.")
    parser.add_argument("--start", help="YYYY-MM-DD.")
    parser.add_argument("--end", help="YYYY-MM-DD.")
    parser.add_argument("--lookback-days", type=int, default=60,
                        help="Days before the period to search for vehicles that "
                             "ran nothing inside it (default 60). Used only to "
                             "assemble the fleet, never counted as activity.")
    parser.add_argument("--roster",
                        help='JSON list of this station\'s vehicle names, or '
                             '{"vehicles": [...]}. The only source that catches '
                             "a truck which has not moved in months.")
    parser.add_argument("--cost-center-map", default="state/shift_cost_center_map.json")
    parser.add_argument("--list-cost-centers", action="store_true",
                        help="Print every cost center the period's legs resolve "
                             "to, and stop.")
    parser.add_argument("--csv", dest="csv_path")
    parser.add_argument("--xlsx", dest="xlsx_path")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    if args.month:
        start, end = VS.month_bounds(args.month)
    elif args.start and args.end:
        start, end = date.fromisoformat(args.start), date.fromisoformat(args.end)
    else:
        parser.error("Give --month YYYY-MM, or both --start and --end.")
    if end < start:
        parser.error("--end is before --start.")

    print("=" * 72)
    print(f"  Vehicle assignment by day -- {args.cost_center}")
    print(f"  {start} to {end}")
    print("  Read-only. GET only, no writes. No patient data.")
    print("=" * 72)

    try:
        from traumasoft_api import TraumasoftAPI
        api = TraumasoftAPI()
        roster_vehicles = api.list_vehicles()
        legs = VS.fetch_legs(api, start, end)
        lookback_legs = []
        if args.lookback_days > 0:
            lookback_legs = VS.fetch_legs(
                api, start - timedelta(days=args.lookback_days),
                start - timedelta(days=1),
            )
    except Exception as exc:  # noqa: BLE001 -- reporting the failure IS the result
        log.error("Traumasoft unavailable (%s: %s). Nothing can be reported.",
                  type(exc).__name__, exc)
        return 2

    cost_center_map = R.CostCenterMap(path=args.cost_center_map)
    known = {cost_center_map.resolve(R.profile_name(leg)) for leg in legs}
    known = {name for name in known if name}

    if args.list_cost_centers:
        print(f"\n   Cost centers the period's legs resolve to ({len(known)}):")
        for name in sorted(known):
            print(f"      {name}")
        return 0

    matched, how = matching_centres(args.cost_center, known)
    centres = set(matched)

    period_activity, covered_days, unattributed = legs_by_vehicle_day(
        legs, centres, cost_center_map)
    lookback_activity, _, _ = legs_by_vehicle_day(
        lookback_legs, centres, cost_center_map)

    fleet = build_fleet(period_activity, lookback_activity, load_roster(args.roster),
                        roster_vehicles, centres, cost_center_map)
    days = VS.days_in(start, end)
    grid = build_grid(fleet, period_activity, days, covered_days, centres)
    names = display_names(fleet, period_activity, lookback_activity,
                          roster_vehicles, legs + lookback_legs)

    print_report(args.cost_center, matched, how, days, covered_days, grid, names,
                 fleet, period_activity, centres, unattributed, len(legs))

    print("\n" + "=" * 72)
    print("  Safe to paste. Operational values only, no patient data.")
    print("=" * 72 + "\n")

    write_outputs(
        per_day_rows(grid, days, names),
        per_vehicle_rows(grid, days, names, fleet, period_activity, centres),
        matrix_rows(grid, days, names),
        args.csv_path, args.xlsx_path,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
