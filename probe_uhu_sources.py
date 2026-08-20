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

    if with_punches:
        sample = with_punches[0]
        print(f"\n  Structure of one punch array ({sample.get('shift_name')}):")
        print("   ", json.dumps(sample.get("punches")[:3], indent=2, default=str)
              .replace("\n", "\n    "))
        print(f"\n  Its shift row says start_time={sample.get('start_time')} "
              f"end_time={sample.get('end_time')}")
        print(f"  Detected offset to local: {offset}")

    # Scheduled versus punched, per unit-shift starting on the target date.
    print(f"\n  Scheduled vs punched hours for unit-shifts starting {target}")
    print(f"    {'shift_profile':<34} {'sched':>7} {'punched':>8} {'crew':>5}  note")
    rows = []
    by_profile = defaultdict(list)
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = shift.get("shift_name")
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
        note = []
        if not punch_rows:
            note.append("no completed punches")
        if open_punches:
            note.append(f"{open_punches} still open")
        rows.append({"shift_profile": name, "scheduled_hours": round(scheduled, 2),
                     "punched_hours": round(punched, 2), "crew": len(crew_rows),
                     "punch_rows": punch_rows, "open_punches": open_punches})
        print(f"    {name:<34} {scheduled:>7.2f} {punched:>8.2f} {len(crew_rows):>5}  "
              f"{' '.join(note)}")

    total_sched = sum(r["scheduled_hours"] for r in rows)
    total_punched = sum(r["punched_hours"] for r in rows)
    unpunched = [r for r in rows if not r["punch_rows"]]
    print(f"\n    {'TOTAL':<34} {total_sched:>7.2f} {total_punched:>8.2f}")
    print(f"    {len(unpunched)} profile(s) scheduled with no completed punch")
    print("\n    Note: punched hours are per crew member, scheduled hours are per")
    print("    unit. Two medics on one truck punch twice for one unit-shift, so")
    print("    these columns are not directly comparable -- the useful signal is")
    print("    which profiles show zero.")
    dump["profiles"] = rows

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

    configured = {str(t.get("name")): t for t in catalogue}
    print(f"\n  {len(configured)} configured, {len(seen)} seen on legs for {target}")
    print(f"\n    {'timestamp':<38} {'type':<12} {'all legs':>10} {'runs only':>11}")
    for name in sorted(set(configured) | set(seen)):
        t = configured.get(name) or {}
        kind = str(t.get("type") or ("-" if name in configured else "NOT CONFIGURED"))
        a = seen.get(name, 0)
        r = seen_on_runs.get(name, 0)
        print(f"    {name:<38} {kind:<12} {a:>6} {a / total_legs:>5.0%} "
              f"{r:>6} {r / total_runs:>5.0%}")

    dump["timestamp_fill"] = {
        name: {"type": (configured.get(name) or {}).get("type"),
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
        pairs = {(s.get("shift_name"), s.get("vehicle_name")) for s in named}
        print(f"  {len(pairs)} distinct (shift_profile, vehicle) pairs")
        for shift_name, vehicle in sorted(pairs)[:10]:
            print(f"    {str(shift_name):<34} -> {vehicle}")
    dump["shift_vehicle_pairs"] = sorted(
        {(str(s.get("shift_name")), str(s.get("vehicle_name"))) for s in named}
    )

    out = Path("uhu_sources.json")
    out.write_text(json.dumps(dump, indent=2, default=str), encoding="utf-8")
    print(f"\nWritten to {out.resolve()} -- send me this file.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
