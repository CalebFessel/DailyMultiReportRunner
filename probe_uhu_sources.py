"""
Measure the two things that could put UHU on a sound footing.

The ePCR is closed -- Data/Epcr/Huly answers 501 to a bare GET and 404 to every
read-shaped rtype, so it is a write surface. What the access sweep turned up
instead is that shifts carry more than we were using:

  * `punches` -- an array of clock-in/clock-out intervals per shift row. That is
    time actually worked, which is a denominator that does not depend on the
    schedule being honoured. A shift rostered but never worked currently counts
    its full scheduled hours against utilization.

  * `vehicle_name` -- a direct shift-to-truck link, which nothing else in the
    API provides.

And Data/Cad/Timestamps lists the timestamps this tenant has configured, with a
`type`. Cross-referencing that catalogue against how often each name actually
appears on a leg answers the open UHU_SPAN question with measurement rather than
argument: a stamp pair is only usable if both halves are reliably recorded.

Read-only.

Usage:
    python probe_uhu_sources.py [YYYY-MM-DD]
"""

import os
import sys
import json
import logging
import argparse
from pathlib import Path
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError
import traumasoft_reports as R

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("probe-uhu-sources")


# Vehicle name prefixes whose units run single-crewed. On this tenant the fleet
# numbers say it plainly: WC- is wheelchair and M- is secure car, while A- is an
# ambulance needing two. Overridable, since a fleet can renumber.
SINGLE_CREW_VEHICLE_PREFIXES = tuple(
    p.strip().lower()
    for p in os.getenv("SINGLE_CREW_VEHICLE_PREFIXES", "wc-,m-").split(",")
    if p.strip()
)

# Markers in a profile name that say the same thing, for rows naming no vehicle.
SINGLE_CREW_NAME_MARKERS = ("w/c", "wheelchair", "secure car")


def vehicle_class(name):
    """'wheelchair'/'secure car' style classes, read off the fleet number."""
    lowered = str(name or "").strip().lower()
    # Names like '(SC)WC-101' carry the prefix after a parenthetical.
    for prefix in SINGLE_CREW_VEHICLE_PREFIXES:
        if lowered.startswith(prefix) or f")({prefix}" in lowered or f"){prefix}" in lowered:
            return prefix
    return None


def suggest_min_crew(profile, vehicles, default):
    """
    What this profile probably needs, and why.

    Returns (min_crew, reason). A profile whose trucks are all wheelchair or
    secure car runs single-crewed; one mixing classes is left at the default and
    flagged, since only the operation knows which way it should fall.
    """
    lowered = str(profile or "").lower()
    marker = next((m for m in SINGLE_CREW_NAME_MARKERS if m in lowered), None)
    classes = {vehicle_class(v) for v in vehicles} if vehicles else set()

    if vehicles and None not in classes and classes:
        return 1, f"all vehicles are {'/'.join(sorted(c.rstrip('-') for c in classes))}"
    if vehicles and classes - {None}:
        mixed = ", ".join(sorted(vehicles))
        return default, f"MIXED vehicle classes ({mixed}) -- confirm"
    if marker:
        return 1, f"name says '{marker}' but its vehicles are not -- confirm"
    return default, ""


def hours(start, end):
    if not start or not end or end <= start:
        return 0.0
    return (end - start).total_seconds() / 3600.0


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

    log.info("Fetching shifts ...")
    shifts = api.list_shifts()
    log.info("Fetching trips for %s ...", target)
    legs = api.get_trips(target, range_days=1)
    log.info("Fetching the timestamp catalogue ...")
    try:
        catalogue = list(api.paginate("ThirdParty/Data/Cad/Timestamps"))
    except TraumasoftAPIError as exc:
        log.warning("Timestamps refused: %s", exc)
        catalogue = []
    log.info("%s shift rows, %s legs, %s configured timestamps",
             len(shifts), len(legs), len(catalogue))

    offset = R.resolve_shift_offset(legs)
    dump = {"metrics_date": target.isoformat(), "shift_offset_hours":
            offset.total_seconds() / 3600.0}

    # ---------- punches ----------
    print()
    print("=" * 78)
    print("SHIFT PUNCHES -- time actually worked")
    print("=" * 78)

    with_punches = [s for s in shifts if s.get("punches")]
    print(f"\n  {len(with_punches)} of {len(shifts)} shift rows carry punches "
          f"({len(with_punches) / len(shifts):.0%})" if shifts else "\n  no shifts")

    in_window = sorted({
        s.date() for s in (R.parse_shift_ts(x.get("start_time"), offset) for x in shifts)
        if s
    })
    if target not in in_window:
        print()
        print("  !! No shift starts on this date. The shift feed only covers")
        print(f"     {in_window[0]} .. {in_window[-1]} and cannot be asked for an older one,"
              if in_window else "     nothing at all,")
        print("     so punches and worked hours are unavailable for it. Re-run with")
        if in_window:
            # The window opens at yesterday, which is the most recent complete day.
            print(f"     a date in that range -- {in_window[0]} is the most recent")
            print("     complete one.")

    if with_punches:
        sample = with_punches[0]
        print(f"\n  Structure of one punch array ({sample.get('shift_name')}):")
        print("   ", json.dumps(sample.get("punches")[:3], indent=2, default=str)
              .replace("\n", "\n    "))
        print(f"\n  Its shift row says start_time={sample.get('start_time')} "
              f"end_time={sample.get('end_time')}")
        print(f"  Detected offset to local: {offset}")

    # Scheduled versus punched, per unit-shift starting on the target date.
    rules = R.UnitStaffingRules()
    print(f"\n  Scheduled vs punched hours for unit-shifts starting {target}")
    print(f"    {'shift_profile':<34} {'need':>4} {'sched':>7} {'worked':>7} "
          f"{'crew':>5}  note")
    rows = []
    by_profile = defaultdict(list)
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = R.profile_name(shift)
        start = R.parse_shift_ts(shift.get("start_time"), offset)
        if not name or not start or start.date() != target:
            continue
        by_profile[name].append(shift)

    for name, crew_rows in sorted(by_profile.items()):
        spans = set()
        punched = 0.0
        punch_rows = 0
        open_punches = 0
        for shift in crew_rows:
            s = R.parse_shift_ts(shift.get("start_time"), offset)
            e = R.parse_shift_ts(shift.get("end_time"), offset)
            if s and e and e > s:
                spans.add((s, e))
            for punch in shift.get("punches") or []:
                if punch.get("deleted"):
                    continue
                ps = R.parse_shift_ts(punch.get("start_time"), offset)
                pe = R.parse_shift_ts(punch.get("end_time"), offset)
                if ps and not pe:
                    open_punches += 1
                    continue
                h = hours(ps, pe)
                if h:
                    punched += h
                    punch_rows += 1
        scheduled = sum(hours(s, e) for s, e in R._merge_spans(spans))
        needed = rules.min_crew(name)
        # The figure the report will actually bill: hours with enough crew on
        # the clock together, not the sum of what everyone punched.
        grouped, _ = R.unit_punches_by_instance(crew_rows, offset, target)
        worked = sum(
            R.staffed_hours(punches, needed)
            for punches in grouped.get(name, {}).values()
        )
        note = []
        if not punch_rows:
            note.append("no completed punches")
        if open_punches:
            note.append(f"{open_punches} still open")
        if scheduled and not worked:
            note.append(f"NEVER REACHED {needed} CREW -- contributes no unit hours")
        rows.append({"shift_profile": name, "min_crew": needed,
                     "scheduled_hours": round(scheduled, 2),
                     "worked_hours": round(worked, 2),
                     "punched_hours": round(punched, 2), "crew": len(crew_rows),
                     "punch_rows": punch_rows, "open_punches": open_punches})
        print(f"    {name:<34} {needed:>4} {scheduled:>7.2f} {worked:>7.2f} "
              f"{len(crew_rows):>5}  {' '.join(note)}")

    total_sched = sum(r["scheduled_hours"] for r in rows)
    total_worked = sum(r["worked_hours"] for r in rows)
    unstaffed = [r for r in rows if r["scheduled_hours"] and not r["worked_hours"]]
    print(f"\n    {'TOTAL':<34} {'':>4} {total_sched:>7.2f} {total_worked:>7.2f}")
    print(f"    {len(unstaffed)} profile(s) rostered but never crewed to minimum")
    print("\n    'need' is the crew this profile must have on the clock together.")
    print("    'worked' is what the report bills -- hours at or above that number,")
    print("    not the sum of individual punches. Any profile that should need one")
    print("    rather than two belongs in the rules file below.")
    dump["profiles"] = rows

    # ---------- a rules file to edit rather than write from scratch ----------
    vehicles_by_profile = defaultdict(set)
    for shift in shifts:
        name = R.profile_name(shift)
        if name and shift.get("vehicle_name"):
            vehicles_by_profile[name].add(str(shift["vehicle_name"]).strip())

    suggested = {}
    evidence = {}
    for r in rows:
        name = r["shift_profile"]
        vehicles = sorted(vehicles_by_profile.get(name, ()))
        crew, reason = suggest_min_crew(name, vehicles, rules.default)
        suggested[name] = crew
        if reason or vehicles:
            evidence[name] = {"vehicles": vehicles, "suggested": crew, "why": reason}

    single = [n for n, c in suggested.items() if c == 1]
    confirm = [n for n, e in evidence.items() if "confirm" in e["why"]]
    print(f"\n    Suggested {len(single)} profile(s) as single-crew from their fleet "
          f"numbers; {len(confirm)} need a look.")
    for name in sorted(confirm):
        print(f"      {name:<34} {evidence[name]['why']}")

    proposal = {
        "note": [
            "Crew each shift profile must have on the clock at once for its",
            "unit-shift to count. Generated from the profiles seen on "
            f"{target}; every one starts at the default.",
            "",
            "Values are suggested from the fleet numbers of the vehicles each",
            "profile actually ran -- wheelchair and secure car units run",
            "single-crewed, ambulances do not. Check them: _evidence records the",
            "vehicles behind each, and anything marked 'confirm' mixed classes or",
            "disagreed with its own name. Delete an entry to fall back to",
            "default_min_crew. _evidence is ignored when the file is read.",
            "",
            "Rename to unit_staffing_rules.json to put it in effect.",
        ],
        "default_min_crew": rules.default,
        "by_pattern": suggested,
        "_evidence": dict(sorted(evidence.items())),
    }
    proposed_path = Path(R.UNIT_STAFFING_RULES_FILE).with_suffix(".proposed.json")
    proposed_path.parent.mkdir(parents=True, exist_ok=True)
    proposed_path.write_text(json.dumps(proposal, indent=2), encoding="utf-8")
    print(f"\n    Wrote {proposed_path} with {len(rows)} profile(s) to review.")

    # ---------- which timestamps actually get recorded ----------
    print()
    print("=" * 78)
    print("TIMESTAMP CATALOGUE vs WHAT LEGS ACTUALLY CARRY")
    print("=" * 78)

    seen = Counter()
    for leg in legs:
        for name in R.timestamp_map(leg):
            seen[name] += 1
    total_legs = len(legs) or 1
    ran = [leg for leg in legs if R.is_run(leg)]
    seen_on_runs = Counter()
    for leg in ran:
        for name in R.timestamp_map(leg):
            seen_on_runs[name] += 1
    total_runs = len(ran) or 1

    # The catalogue's `name` is a display label ("At Scene"); the key a leg
    # actually carries is its `type` ("at_scene"). Joining on name matches
    # nothing and makes every stamp look unconfigured.
    configured = {}
    for t in catalogue:
        key = str(t.get("type") or t.get("name") or "").strip()
        if key:
            configured[key] = t
    print(f"\n  {len(configured)} configured, {len(seen)} seen on legs for {target}")
    print(f"\n    {'timestamp':<38} {'catalogue':<22} {'all legs':>10} {'runs only':>11}")
    for name in sorted(set(configured) | set(seen)):
        t = configured.get(name) or {}
        kind = str(t.get("name") or ("-" if name in configured else "not in catalogue"))
        a = seen.get(name, 0)
        r = seen_on_runs.get(name, 0)
        print(f"    {name:<38} {kind:<22} {a:>6} {a / total_legs:>5.0%} "
              f"{r:>6} {r / total_runs:>5.0%}")

    dump["timestamp_fill"] = {
        name: {"catalogue_name": (configured.get(name) or {}).get("name"),
               "on_all_legs": seen.get(name, 0), "on_runs": seen_on_runs.get(name, 0)}
        for name in sorted(set(configured) | set(seen))
    }
    dump["totals"] = {"legs": len(legs), "runs": len(ran)}

    # ---------- does the shift row name a vehicle? ----------
    named = [s for s in shifts if s.get("vehicle_name")]
    print()
    print("=" * 78)
    print("SHIFT -> VEHICLE LINK")
    print("=" * 78)
    print(f"\n  {len(named)} of {len(shifts)} shift rows name a vehicle "
          f"({len(named) / len(shifts):.0%})" if shifts else "")
    if named:
        pairs = {(R.profile_name(s), s.get("vehicle_name")) for s in named}
        print(f"  {len(pairs)} distinct (shift_profile, vehicle) pairs")
        for shift_name, vehicle in sorted(pairs)[:10]:
            print(f"    {str(shift_name):<34} -> {vehicle}")
    dump["shift_vehicle_pairs"] = sorted(
        {(str(R.profile_name(s)), str(s.get("vehicle_name"))) for s in named}
    )

    out = Path("uhu_sources.json")
    out.write_text(json.dumps(dump, indent=2, default=str), encoding="utf-8")
    print(f"\nWritten to {out.resolve()} -- send me this file.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
