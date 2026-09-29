"""
For a period, how many days each vehicle was in service and how many it was not.

Management asked for July, by cost center, split AMB / MH / WC, and said the
thing that matters: *"A total count is misleading because they were not OOS the
whole month. For each vehicle, I seek something close to out of 31 days it was
running X days."*

That is a per-day question, and the Traumasoft ThirdParty API cannot answer it.
There is no vehicle status history and no work-order endpoint -- `vehicle_status`
is a single current value, so asking the API today tells you about today and
nothing about July. See docs/API_MIGRATION.md.

WHAT CAN ANSWER IT is the daily report's own archive.
`Daily_Vehicle_Overview_APPEND.xlsx` has been snapshotting the Out Of Service
sheet every day, one row per out-of-service vehicle per day, with 730-day
retention. A vehicle in the 5 July snapshot was out of service on 5 July. Count
the snapshot days it appears on and the answer is exact, not inferred.

THE DENOMINATOR IS NOT 31. It is the number of days in the period the daily
report actually ran. A day nobody snapshotted is a day nobody observed, and
counting it as in service would quietly inflate every vehicle's uptime -- which
is exactly the direction that flatters the fleet and misleads the reader. So
every vehicle gets three counts that sum to the calendar period:

    days_out_of_service    seen out of service in a snapshot
    days_in_service        snapshotted, and not out of service
    days_not_observed      no snapshot exists for that day

TWO DIFFERENT NUMBERS, BOTH REPORTED. "In service" means not carrying an
out-of-service label. It does not mean the truck worked -- a unit can be in
service, staffed and simply never dispatched. `days_ran_a_call` is counted
separately from the period's trip legs, which backfill reliably, so the gap
between the two is visible instead of assumed away.

COST CENTER AND VEHICLE CLASS ARE NOT ON A VEHICLE. The API's vehicle field
allowlist is closed -- id, name, vehicle_status, vin, odometer, disabled,
deleted, plus live-shift enrichment -- and `cost_center_id`, `cost_center_name`
and `status_reason` were confirmed absent against the live API. So:

    Cost center   the period's own trip legs: vehicle -> shift_name -> the
                  accumulated shift_cost_center_map. Period-accurate, unlike
                  the vehicle row's live shift_name, which is today's.
    Class         the unit name. A naming convention, not data. Every distinct
                  prefix found is printed so the mapping can be corrected in
                  one pass rather than trusted.

STRICTLY READ-ONLY. It reads an archive workbook and issues GETs. Nothing is
written back to the append file.

PHI: no patient-identifying value is read or printed.

Usage:
    python vehicle_service_days.py --month 2026-07
    python vehicle_service_days.py --start 2026-07-01 --end 2026-07-31 --xlsx July.xlsx
    python vehicle_service_days.py --month 2026-07 --append path/to/Append
"""

import argparse
import calendar
import json
import logging
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta

log = logging.getLogger(__name__)

APPEND_WORKBOOK = "Daily_Vehicle_Overview_APPEND.xlsx"
OOS_SHEET = "Out Of Service"
SUMMARY_SHEET = "Summary"

# Vehicle class read off the unit name, longest prefix first so "wc-" is not
# shadowed by a shorter rule.
#
# CONFIRM THIS AGAINST THE ROSTER BEFORE SENDING THE NUMBERS ANYWHERE. This
# repository documents `WC-` as wheelchair and `M-` as secure car
# (probe_uhu_sources.SINGLE_CREW_VEHICLE_PREFIXES); management asked for
# AMB / MH / WC. Whether their MH is this tenant's M- is not something the
# data says, so the report prints every prefix it found and every name it
# could not place, and VEHICLE_CLASS_PATTERNS overrides the lot.
DEFAULT_CLASS_PATTERNS = "AMB=a-;MH=m-;WC=wc-"

NON_FLEET_STATUSES = {"Retired", "New - Waiting for Delivery", "Waiting for Inspection"}
OUT_OF_SERVICE_STATUSES = {"Out of Service", "Out of Service - Collision"}


# =============================
# PERIOD
# =============================
def month_bounds(text):
    """(first, last) of a YYYY-MM."""
    year, month = (int(part) for part in text.split("-", 1))
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def days_in(start, end):
    return [start + timedelta(days=n) for n in range((end - start).days + 1)]


# =============================
# CLASSIFICATION
# =============================
def parse_class_patterns(text):
    """
    "AMB=a-;MH=m-;WC=wc-" -> [(prefix, label)], longest prefix first.

    Sorted by length so a longer rule always wins: with "m-" and "mh-" both
    defined, "MH-12" must not fall to the shorter one.
    """
    rules = []
    for chunk in str(text or "").split(";"):
        chunk = chunk.strip()
        if not chunk or "=" not in chunk:
            continue
        label, prefix = (part.strip() for part in chunk.split("=", 1))
        if label and prefix:
            rules.append((prefix.lower(), label))
    return sorted(rules, key=lambda rule: -len(rule[0]))


def vehicle_class(name, rules):
    """
    The class a unit name says it is, or None.

    Names like '(SC)WC-101' carry the prefix after a parenthetical, so the
    leading bracketed group is stripped before matching -- the same case
    probe_uhu_sources handles.
    """
    lowered = str(name or "").strip().lower()
    stripped = re.sub(r"^\([^)]*\)\s*", "", lowered)
    for prefix, label in rules:
        if stripped.startswith(prefix) or lowered.startswith(prefix):
            return label
    return None


def name_prefix(name):
    """The leading letters of a unit name, for reporting what is unplaced."""
    stripped = re.sub(r"^\([^)]*\)\s*", "", str(name or "").strip().lower())
    match = re.match(r"[a-z]+", stripped)
    return match.group(0) if match else ""


# =============================
# THE ARCHIVE
# =============================
def read_oos_snapshots(append_dir, start, end):
    """
    {snapshot date -> {normalized vehicle name -> row}} over the period.

    Returns (snapshots, error). An error is returned rather than raised so the
    caller can say precisely what is missing -- this file is the whole answer,
    and "no data" and "the file is not where I looked" are different problems
    with different fixes.
    """
    import pandas as pd

    path = os.path.join(append_dir, APPEND_WORKBOOK)
    if not os.path.exists(path):
        return None, (
            f"{path} does not exist. That workbook is the only per-day record "
            "of vehicle status -- the API has no status history -- so without "
            "it this period cannot be answered at all."
        )
    try:
        frame = pd.read_excel(path, sheet_name=OOS_SHEET, engine="openpyxl")
    except ValueError:
        return None, f"{path} has no '{OOS_SHEET}' sheet."
    except Exception as exc:  # noqa: BLE001
        return None, f"Could not read {path}: {type(exc).__name__}: {exc}"

    if "snapshot_date" not in frame.columns:
        return None, (
            f"'{OOS_SHEET}' in {path} has no snapshot_date column, so its rows "
            "cannot be placed on a day."
        )

    dates = pd.to_datetime(frame["snapshot_date"], errors="coerce").dt.date
    snapshots = defaultdict(dict)
    for (_, row), day in zip(frame.iterrows(), dates):
        if day is None or not (start <= day <= end):
            continue
        name = normalize_name(row.get("vehicle_name"))
        if name:
            snapshots[day][name] = row.to_dict()
    return dict(snapshots), None


def read_summary_counts(append_dir, start, end):
    """
    {snapshot date -> {metric -> value}}, for cross-checking the OOS sheet.

    Absent is not an error: it only costs the cross-check.
    """
    import pandas as pd

    path = os.path.join(append_dir, APPEND_WORKBOOK)
    if not os.path.exists(path):
        return {}
    try:
        frame = pd.read_excel(path, sheet_name=SUMMARY_SHEET, engine="openpyxl")
    except Exception:  # noqa: BLE001
        return {}
    if not {"snapshot_date", "metric", "value"} <= set(frame.columns):
        return {}

    dates = pd.to_datetime(frame["snapshot_date"], errors="coerce").dt.date
    counts = defaultdict(dict)
    for (_, row), day in zip(frame.iterrows(), dates):
        if day is not None and start <= day <= end:
            counts[day][str(row.get("metric"))] = row.get("value")
    return dict(counts)


def normalize_name(value):
    """A unit name reduced to what is comparable across systems and snapshots."""
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


# =============================
# COST CENTRE
# =============================
def cost_centers_from_legs(legs, cost_center_map):
    """
    normalized vehicle name -> the cost center it ran under in the period.

    Taken from the period's own trips rather than the vehicle row's live
    `shift_name`, which is today's shift and says nothing about July. A vehicle
    that ran under more than one profile is attributed to the one it ran most
    legs under, and the runners-up are returned so a split unit is visible
    rather than silently rounded.

    `cost_center_map` is traumasoft_reports.CostCenterMap, not a plain dict:
    it carries the hand-written overrides for profiles the crew route can never
    reach, and it already handles the casing and trailing-space drift between
    the trip and shift feeds. Reimplementing the lookup here would silently
    disagree with every other report in the repository.
    """
    if cost_center_map is None:
        return {}, {}
    tally = defaultdict(Counter)
    for leg in legs or []:
        name = normalize_name(leg.get("vehicle_name"))
        if not name:
            continue
        centre = cost_center_map.resolve(str(leg.get("shift_name") or "").strip())
        if centre:
            tally[name][centre] += 1

    resolved, split = {}, {}
    for name, counter in tally.items():
        # most_common breaks ties by insertion order; sort for determinism.
        resolved[name] = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
        if len(counter) > 1:
            split[name] = dict(counter)
    return resolved, split


# =============================
# ASSEMBLY
# =============================
def build_rows(snapshots, roster, legs, cost_center_map, start, end, class_rules,
               include_non_fleet=False, lookback_legs=None):
    """
    One row per vehicle, with the three day counts and how it was attributed.

    The vehicle set is the current roster UNION every vehicle seen out of
    service during the period. The union matters: a truck that was out of
    service through July and retired in August is gone from today's roster,
    and dropping it would understate exactly the thing being measured.
    """
    observed = sorted(snapshots)
    period = days_in(start, end)
    not_observed = [day for day in period if day not in snapshots]

    by_name = {}
    for vehicle in roster or []:
        status = str(vehicle.get("vehicle_status") or "").strip()
        if not include_non_fleet and status in NON_FLEET_STATUSES:
            continue
        name = normalize_name(vehicle.get("name"))
        if name:
            by_name[name] = vehicle

    # Vehicles the archive saw that the roster no longer carries.
    archived_only = {}
    for day_rows in snapshots.values():
        for name, row in day_rows.items():
            if name not in by_name:
                archived_only.setdefault(name, row)

    # Cost-center attribution, best source first. A vehicle that was out of
    # service for the whole period ran nothing in it -- and those are exactly
    # the vehicles this report is about, so leaving them unattributed would
    # strand the rows management most wants grouped. The fallbacks are labelled
    # per row rather than blended, because a cost center read from March is a
    # weaker claim than one read from July and the reader should be able to
    # tell which they have.
    centres, split = cost_centers_from_legs(legs, cost_center_map)
    sources = {name: "period" for name in centres}

    earlier, earlier_split = cost_centers_from_legs(lookback_legs, cost_center_map)
    for name, centre in earlier.items():
        if name not in centres:
            centres[name] = centre
            sources[name] = "earlier legs"
            if name in earlier_split:
                split[name] = earlier_split[name]
    ran_days = defaultdict(set)
    for leg in legs or []:
        name = normalize_name(leg.get("vehicle_name"))
        day = leg_date(leg)
        if name and day and start <= day <= end:
            ran_days[name].add(day)

    rows = []
    for name in sorted(set(by_name) | set(archived_only)):
        vehicle = by_name.get(name)
        archived = archived_only.get(name)
        display = (vehicle or {}).get("name") or (archived or {}).get("vehicle_name") or name

        centre = centres.get(name)
        source = sources.get(name, "")
        if centre is None and vehicle is not None:
            # Last resort: the vehicle row's live shift enrichment. It is
            # today's shift, not the period's, so it is labelled as such.
            centre = cost_center_map.resolve(
                str(vehicle.get("shift_name") or "").strip()
            ) if cost_center_map is not None else None
            if centre:
                source = "current shift"

        oos_days = [day for day in observed if name in snapshots[day]]
        in_service_days = [day for day in observed if name not in snapshots[day]]

        rows.append({
            "vehicle_name": display,
            "vehicle_class": vehicle_class(display, class_rules) or "UNCLASSIFIED",
            "cost_center": centre or "UNKNOWN",
            "cost_center_source": source,
            "days_in_service": len(in_service_days),
            "days_out_of_service": len(oos_days),
            "days_not_observed": len(not_observed),
            "days_in_period": len(period),
            "days_observed": len(observed),
            "days_ran_a_call": len(ran_days.get(name, ())),
            "in_current_fleet": vehicle is not None,
            "current_status": (vehicle or {}).get("vehicle_status") or "",
            "cost_center_split": json.dumps(split[name]) if name in split else "",
            "first_day_out": min(oos_days).isoformat() if oos_days else "",
            "last_day_out": max(oos_days).isoformat() if oos_days else "",
        })
    return rows, observed, not_observed


def fetch_legs(api, start, end):
    """
    Every trip leg in a window.

    GetTrips caps range_days at 31, so anything longer is walked rather than
    quietly truncated to the first month -- which would look like a fleet that
    stopped running.
    """
    legs, cursor = [], start
    while cursor <= end:
        span = min(31, (end - cursor).days + 1)
        legs.extend(api.get_trips(cursor.isoformat(), range_days=span))
        cursor += timedelta(days=span)
    return legs


def leg_date(leg):
    """The day a leg belongs to, from whichever stamp it carries."""
    for key in ("pickup_time", "requested_pickup_time", "appt_time", "created"):
        value = leg.get(key)
        if not value:
            continue
        text = str(value).strip()[:10]
        try:
            return date.fromisoformat(text)
        except ValueError:
            continue
    return None


def rollup(rows, by):
    """Totals grouped by one or more row keys."""
    groups = defaultdict(lambda: {
        "vehicles": 0, "days_in_service": 0, "days_out_of_service": 0,
        "days_not_observed": 0, "days_ran_a_call": 0, "vehicles_ever_out": 0,
    })
    for row in rows:
        key = tuple(row[field] for field in by)
        entry = groups[key]
        entry["vehicles"] += 1
        entry["days_in_service"] += row["days_in_service"]
        entry["days_out_of_service"] += row["days_out_of_service"]
        entry["days_not_observed"] += row["days_not_observed"]
        entry["days_ran_a_call"] += row["days_ran_a_call"]
        if row["days_out_of_service"]:
            entry["vehicles_ever_out"] += 1
    return dict(groups)


# =============================
# OUTPUT
# =============================
def print_report(rows, observed, not_observed, start, end, summary_counts,
                 snapshots, class_rules, args_append):
    period_days = (end - start).days + 1

    print(f"\n1. WHAT WAS ACTUALLY OBSERVED, {start} TO {end}")
    print("   " + "-" * 66)
    print(f"   Days in the period:        {period_days}")
    print(f"   Days a snapshot exists:    {len(observed)}")
    print(f"   Days never snapshotted:    {len(not_observed)}")
    if not_observed:
        print("\n   THE DENOMINATOR IS NOT "
              f"{period_days}, IT IS {len(observed)}. A day nobody")
        print("   snapshotted is a day nobody observed. Counting it as in")
        print("   service would inflate every vehicle's uptime, so those days")
        print("   are reported separately and never folded in.\n")
        runs = compress_days(not_observed)
        for run in runs[:12]:
            print(f"      {run}")
        if len(runs) > 12:
            print(f"      ... and {len(runs) - 12} more gaps")

    if summary_counts:
        print("\n   Cross-check against the Summary sheet:")
        mismatched = []
        for day in observed:
            stated = summary_counts.get(day, {}).get("Out Of Service")
            actual = len(snapshots[day])
            if stated is not None and int(stated) != actual:
                mismatched.append((day, int(stated), actual))
        if mismatched:
            print(f"      {len(mismatched)} day(s) where the Summary count and the")
            print("      Out Of Service row count disagree. One of them is wrong.")
            for day, stated, actual in mismatched[:10]:
                print(f"         {day}   summary {stated:>4}   rows {actual:>4}")
        else:
            print("      Every observed day agrees. The row counts are sound.")

    print(f"\n2. BY COST CENTER AND CLASS")
    print("   " + "-" * 66)
    print(f"   {'cost center':<22}{'class':<14}{'vehicles':>9}{'ever out':>9}"
          f"{'days in':>9}{'days out':>9}")
    print("   " + "-" * 72)
    grouped = rollup(rows, ("cost_center", "vehicle_class"))
    for (centre, klass), entry in sorted(grouped.items()):
        print(f"   {str(centre)[:21]:<22}{str(klass)[:13]:<14}"
              f"{entry['vehicles']:>9}{entry['vehicles_ever_out']:>9}"
              f"{entry['days_in_service']:>9}{entry['days_out_of_service']:>9}")

    print(f"\n3. HOW MANY WERE SUPPOSED TO BE ON THE ROAD")
    print("   " + "-" * 66)
    print("   Read as: of the vehicles carried in the fleet, how many were not")
    print("   labelled out of service. A vehicle counts once even if it was out")
    print("   for part of the period, so the 'ever out' column is the honest")
    print("   answer to 'how many were out' and the day counts are the rest.\n")
    per_class = rollup(rows, ("vehicle_class",))
    print(f"   {'class':<15}{'vehicles':>9}{'ever out':>10}{'never out':>11}"
          f"{'% days in':>11}")
    print("   " + "-" * 56)
    for (klass,), entry in sorted(per_class.items()):
        observed_days = entry["days_in_service"] + entry["days_out_of_service"]
        share = (100.0 * entry["days_in_service"] / observed_days) if observed_days else 0.0
        print(f"   {str(klass)[:14]:<15}{entry['vehicles']:>9}"
              f"{entry['vehicles_ever_out']:>10}"
              f"{entry['vehicles'] - entry['vehicles_ever_out']:>11}"
              f"{share:>10.1f}%")

    print(f"\n4. PER VEHICLE -- OUT OF {len(observed)} OBSERVED DAYS")
    print("   " + "-" * 66)
    out_at_all = sorted(
        (r for r in rows if r["days_out_of_service"]),
        key=lambda r: (-r["days_out_of_service"], r["vehicle_name"]),
    )
    if not out_at_all:
        print("   No vehicle was out of service on any observed day.")
    else:
        print(f"   {len(out_at_all)} vehicle(s) were out of service at some point.\n")
        print(f"   {'vehicle':<16}{'class':<14}{'cost center':<18}"
              f"{'in':>5}{'out':>5}{'ran':>5}")
        print("   " + "-" * 69)
        for row in out_at_all[:40]:
            print(f"   {str(row['vehicle_name'])[:15]:<16}"
                  f"{str(row['vehicle_class'])[:13]:<14}"
                  f"{str(row['cost_center'])[:17]:<18}"
                  f"{row['days_in_service']:>5}{row['days_out_of_service']:>5}"
                  f"{row['days_ran_a_call']:>5}")
        if len(out_at_all) > 40:
            print(f"   ... and {len(out_at_all) - 40} more, all in the export")

    print("\n5. WHAT TO CHECK BEFORE SENDING THIS ON")
    print("   " + "-" * 66)
    unplaced = [r for r in rows if r["vehicle_class"] == "UNCLASSIFIED"]
    prefixes = Counter(name_prefix(r["vehicle_name"]) for r in rows)
    mapped = {prefix.rstrip("-") for prefix, _ in class_rules}
    print(f"   Class comes from the unit name, not from data. Rules in use: "
          f"{', '.join(f'{l}={p}' for p, l in class_rules)}")
    print(f"\n   {'name prefix':<16}{'vehicles':>9}   mapped to")
    print("   " + "-" * 50)
    for prefix, count in prefixes.most_common(15):
        label = next((l for p, l in class_rules
                      if prefix and (prefix + "-").startswith(p)), "-- nothing --")
        print(f"   {(prefix or '(none)')[:15]:<16}{count:>9}   {label}")
    if unplaced:
        print(f"\n   {len(unplaced)} vehicle(s) match no class rule:")
        print("      " + ", ".join(str(r["vehicle_name"]) for r in unplaced[:20])[:220])
        print("      Set VEHICLE_CLASS_PATTERNS to place them, e.g.")
        print("      VEHICLE_CLASS_PATTERNS='AMB=a-;MH=mh-;WC=wc-'")
    if "MH" in {label for _, label in class_rules}:
        print("\n   CONFIRM MH. This repository documents `M-` as secure car, not")
        print("   Medicar -- see probe_uhu_sources.SINGLE_CREW_VEHICLE_PREFIXES.")
        print("   Whether management's MH is this tenant's M- is not something")
        print("   the data says. Check the prefix table above against the roster.")

    print("\n   Cost center is not on a vehicle in this API. Where each one")
    print("   came from:")
    for source, count in Counter(
        r["cost_center_source"] or "-- nothing --" for r in rows
    ).most_common():
        label = {
            "period": "the period's own trip legs (period-accurate)",
            "earlier legs": "trips before the period (weaker -- a unit can move)",
            "current shift": "today's shift enrichment (weakest -- not the period)",
        }.get(source, "unattributed")
        print(f"      {count:>4}  {label}")

    unknown_centre = [r for r in rows if r["cost_center"] == "UNKNOWN"]
    if unknown_centre:
        print(f"\n   {len(unknown_centre)} vehicle(s) could not be placed at all --")
        print("   they ran nothing in the period or the lookback and carry no")
        print("   live shift. Widen --attribution-lookback-days, or add them to")
        print("   state/shift_cost_center_overrides.json.")
        print("      " + ", ".join(str(r["vehicle_name"]) for r in unknown_centre[:20])[:220])

    split = [r for r in rows if r["cost_center_split"]]
    if split:
        print(f"\n   {len(split)} vehicle(s) ran under more than one cost center and")
        print("   are attributed to the one they ran most legs under. The full")
        print("   split is in the export.")

    archived = [r for r in rows if not r["in_current_fleet"]]
    if archived:
        print(f"\n   {len(archived)} vehicle(s) appear in the archive but not in today's")
        print("   fleet -- retired or removed since. They are counted, because")
        print("   dropping them would understate the period's out-of-service days.")
        print("      " + ", ".join(str(r["vehicle_name"]) for r in archived[:20])[:220])


def compress_days(days):
    """['2026-07-04', '2026-07-06 to 2026-07-08'] from a list of dates."""
    if not days:
        return []
    runs, start, previous = [], days[0], days[0]
    for day in days[1:]:
        if (day - previous).days == 1:
            previous = day
            continue
        runs.append((start, previous))
        start = previous = day
    runs.append((start, previous))
    return [
        first.isoformat() if first == last else f"{first} to {last}"
        for first, last in runs
    ]


def write_outputs(rows, start, end, csv_path=None, xlsx_path=None):
    import pandas as pd

    frame = pd.DataFrame(rows)
    if csv_path:
        frame.to_csv(csv_path, index=False)
        print(f"  Written to {csv_path}")
    if xlsx_path:
        by_centre = rollup(rows, ("cost_center", "vehicle_class"))
        by_class = rollup(rows, ("vehicle_class",))
        sheets = {
            "Per Vehicle": frame,
            "By Cost Center": pd.DataFrame([
                {"cost_center": centre, "vehicle_class": klass, **entry}
                for (centre, klass), entry in sorted(by_centre.items())
            ]),
            "By Class": pd.DataFrame([
                {"vehicle_class": klass, **entry}
                for (klass,), entry in sorted(by_class.items())
            ]),
        }
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            for name, sheet in sheets.items():
                sheet.to_excel(writer, sheet_name=name, index=False)
        print(f"  Written to {xlsx_path}")


# =============================
# CLI
# =============================
def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--month", help="YYYY-MM. Shorthand for the whole month.")
    parser.add_argument("--start", help="First day, YYYY-MM-DD.")
    parser.add_argument("--end", help="Last day, YYYY-MM-DD.")
    parser.add_argument("--append", default=None,
                        help="Directory holding Daily_Vehicle_Overview_APPEND.xlsx. "
                             "Defaults to the daily run's own append directory.")
    parser.add_argument("--cost-center-map", default="state/shift_cost_center_map.json",
                        help="shift_name -> cost center map accumulated by the daily run.")
    parser.add_argument("--class-patterns",
                        default=os.getenv("VEHICLE_CLASS_PATTERNS", DEFAULT_CLASS_PATTERNS),
                        help='e.g. "AMB=a-;MH=mh-;WC=wc-"')
    parser.add_argument("--attribution-lookback-days", type=int, default=90,
                        help="Days before the period to search for trips when a "
                             "vehicle ran nothing inside it, for cost-center "
                             "attribution only (default 90). Never counted as "
                             "a day in service or a day run.")
    parser.add_argument("--include-non-fleet", action="store_true",
                        help="Also count retired and undelivered vehicles.")
    parser.add_argument("--no-trips", action="store_true",
                        help="Skip the Traumasoft calls. Cost centers and "
                             "days-ran will be blank.")
    parser.add_argument("--csv", dest="csv_path")
    parser.add_argument("--xlsx", dest="xlsx_path")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    if args.month:
        start, end = month_bounds(args.month)
    elif args.start and args.end:
        start, end = date.fromisoformat(args.start), date.fromisoformat(args.end)
    else:
        parser.error("Give --month YYYY-MM, or both --start and --end.")
    if end < start:
        parser.error("--end is before --start.")

    append_dir = args.append
    if append_dir is None:
        import report_output as OUT
        append_dir = OUT.APPEND_DIR

    print("=" * 72)
    print("  Vehicle service days")
    print(f"  {start} to {end}   archive {os.path.join(append_dir, APPEND_WORKBOOK)}")
    print("  Read-only. The append workbook is never written back.")
    print("=" * 72)

    snapshots, error = read_oos_snapshots(append_dir, start, end)
    if error:
        log.error("%s", error)
        print("\n  The Traumasoft API cannot substitute for it: there is no")
        print("  vehicle status history endpoint, so today's status is all it")
        print("  offers. If the workbook exists on another machine, point")
        print("  --append at it. See docs/VEHICLE_SERVICE_DAYS.md.")
        return 2
    if not snapshots:
        log.error("The archive holds no snapshot between %s and %s.", start, end)
        print("\n  The file is there but covers a different period. Check its")
        print("  retention window -- APPEND_RETENTION_DAYS prunes old rows.")
        return 2

    legs, lookback_legs, roster = [], [], []
    if not args.no_trips:
        try:
            from traumasoft_api import TraumasoftAPI
            api = TraumasoftAPI()
            roster = api.list_vehicles()
            legs = fetch_legs(api, start, end)
            if args.attribution_lookback_days > 0:
                lookback_legs = fetch_legs(
                    api,
                    start - timedelta(days=args.attribution_lookback_days),
                    start - timedelta(days=1),
                )
        except Exception as exc:  # noqa: BLE001
            log.warning("Traumasoft unavailable (%s: %s).", type(exc).__name__, exc)
            log.warning("Cost centers and days-ran will be blank. The day counts "
                        "come from the archive and are unaffected.")

    import traumasoft_reports as R
    cost_center_map = R.CostCenterMap(path=args.cost_center_map)
    if not cost_center_map.counts and not cost_center_map.override_names \
            and not args.no_trips:
        log.warning("No shift -> cost center map at %s and no overrides, so "
                    "every vehicle will read as UNKNOWN cost center.",
                    args.cost_center_map)

    class_rules = parse_class_patterns(args.class_patterns)
    rows, observed, not_observed = build_rows(
        snapshots, roster, legs, cost_center_map, start, end, class_rules,
        include_non_fleet=args.include_non_fleet, lookback_legs=lookback_legs,
    )
    summary_counts = read_summary_counts(append_dir, start, end)

    print_report(rows, observed, not_observed, start, end, summary_counts,
                 snapshots, class_rules, append_dir)

    print("\n" + "=" * 72)
    print("  Safe to paste. Operational values only, no patient data.")
    print("=" * 72 + "\n")

    write_outputs(rows, start, end, args.csv_path, args.xlsx_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
