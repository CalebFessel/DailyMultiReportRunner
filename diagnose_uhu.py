"""
Dump the raw shift and trip records behind a UHU row.

The UHU report compares scheduled hours from /Schedule/Shifts against utilized
hours from trip timestamps. When a row looks wrong -- a twelve-hour overnight
profile showing one scheduled hour, a single run showing twenty-four utilized
hours -- the report itself cannot say which side is at fault. This prints the
records both numbers were built from, and writes them to JSON.

Read-only: it issues the same GETs the report does and writes nothing but the
dump file.

Usage:
    python diagnose_uhu.py [YYYY-MM-DD] [--profile NAME] [--top N]

    YYYY-MM-DD   date to inspect (default: yesterday, matching the report)
    --profile    inspect these profiles by name; repeatable. Without it the
                 worst offenders by ratio and by span are chosen automatically.
    --top        how many offenders to pick automatically (default 8)
"""

import sys
import json
import logging
import argparse
from pathlib import Path
from datetime import date, datetime, time, timedelta
from collections import defaultdict

from traumasoft_api import TraumasoftAPI
import traumasoft_reports as R

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("diagnose-uhu")

# A leg running longer than this almost certainly never got a clear stamp.
IMPLAUSIBLE_SPAN_HOURS = 8.0


def _fmt(value):
    return "-" if value in (None, "") else str(value)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("metrics_date", nargs="?", default=None)
    parser.add_argument("--profile", action="append", default=[])
    parser.add_argument("--top", type=int, default=8)
    args = parser.parse_args()

    if args.metrics_date:
        target = datetime.strptime(args.metrics_date, "%Y-%m-%d").date()
    else:
        target = date.today() - timedelta(days=1)

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("%s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Credentials rejected under every signing scheme.")
        return 1

    window_start = datetime.combine(target, time(0, 0))
    window_end = window_start + timedelta(days=1)

    log.info("Fetching trips for %s ...", target)
    legs = api.get_trips(target, range_days=1)
    log.info("Fetching trips for %s (overnight tail) ...", target + timedelta(days=1))
    legs += api.get_trips(target + timedelta(days=1), range_days=1)
    log.info("Fetching shifts (rolling window) ...")
    shifts = api.list_shifts()
    log.info("%s legs, %s shift rows", len(legs), len(shifts))

    # Shift times arrive as UTC while trips are local; the trips carry the
    # tenant's offset, so it is read from them rather than configured.
    shift_offset = R.resolve_shift_offset(legs)
    log.info("Tenant UTC offset from trip timestamps: %s", shift_offset)

    start_key, end_key = R.UHU_SPANS.get(R.UHU_SPAN, R.UHU_SPANS["task"])

    # --- group the raw records by profile ---
    shifts_by_profile = defaultdict(list)
    for shift in shifts:
        name = shift.get("shift_name")
        if name:
            shifts_by_profile[name].append(shift)

    legs_by_profile = defaultdict(list)
    for leg in legs:
        name = leg.get("shift_name")
        if name:
            legs_by_profile[name].append(leg)

    # --- pick what to inspect ---
    if args.profile:
        chosen = list(dict.fromkeys(args.profile))
    else:
        df = R.build_uhu(shifts, legs, R.CostCenterMap(), target, shift_offset=shift_offset)
        chosen = []
        if not df.empty:
            worst_ratio = df[df["scheduled_hours"] > 0].nlargest(args.top, "uhu_ratio")
            worst_run = df.nlargest(args.top, "hours_per_run")
            chosen = list(dict.fromkeys(
                list(worst_ratio["shift_profile_name"]) + list(worst_run["shift_profile_name"])
            ))
        log.info("No --profile given; inspecting the %s worst rows.", len(chosen))

    dump = {
        "metrics_date": target.isoformat(),
        "window": [window_start.isoformat(), window_end.isoformat()],
        "attribution": "shift-instance",
        "shift_offset_hours": shift_offset.total_seconds() / 3600.0,
        "uhu_span": R.UHU_SPAN,
        "span_keys": [start_key, end_key],
        "generated": datetime.now().isoformat(timespec="seconds"),
        "profiles": {},
    }

    for name in chosen:
        rows = shifts_by_profile.get(name, [])
        profile_legs = legs_by_profile.get(name, [])

        print()
        print("=" * 78)
        print(f"PROFILE: {name}")
        print("=" * 78)

        instances = R.shift_instances(shifts, shift_offset).get(name, [])
        todays = [s for s in instances if s[0].date() == target]
        print(f"\n  Shift records ({len(rows)}), shown as raw then local "
              f"(offset {shift_offset})")
        print(f"  Unit-shift instances starting {target}: {len(todays)}")
        for s, e in todays:
            print(f"    {s} -> {e}   ({(e - s).total_seconds() / 3600.0:.2f}h)")
        if not rows:
            print("    NONE. The rolling shift feed has nothing for this profile, so")
            print("    scheduled_hours is 0 and the ratio is meaningless.")
        seen = set()
        for row in rows:
            start = R.parse_shift_ts(row.get("start_time"), shift_offset)
            end = R.parse_shift_ts(row.get("end_time"), shift_offset)
            key = (name, start, end)
            duplicate = key in seen
            seen.add(key)
            flags = []
            if row.get("deleted"):
                flags.append("deleted")
            if duplicate:
                flags.append("dup-not-counted")
            if start and end and end <= start:
                flags.append("END <= START")
            if start and end and end.date() != start.date():
                flags.append("crosses-midnight")
            print(
                f"    raw={_fmt(row.get('start_time')):<18}->{_fmt(row.get('end_time')):<18} "
                f"local={start} -> {end}  {' '.join(flags)}"
            )
        # The instances above are what the report bills; overlapping crew rows
        # have already been merged into one covered period.
        scheduled_total = sum((e - s).total_seconds() for s, e in todays) / 3600.0
        print(f"    -> scheduled_hours = {scheduled_total:.2f}")

        print(f"\n  Legs ({len(profile_legs)})")
        utilized_total = 0.0
        for leg in profile_legs:
            stamps = R.timestamp_map(leg)
            span = R.span_minutes(leg, start_key, end_key)
            raw_hours = (span or 0.0) / 60.0
            instance = R.assign_leg(leg, todays, start_key, end_key)
            s_start_c = R.parse_ts(stamps.get(start_key))
            s_end_c = R.parse_ts(stamps.get(end_key))
            if instance and s_start_c and s_end_c:
                hours = R.overlap_minutes(s_start_c, s_end_c, instance[0], instance[1]) / 60.0
            else:
                hours = raw_hours
            status = (leg.get("trip_status") or "").strip()
            counted = status.lower() not in R.UHU_EXCLUDED_TRIP_STATUSES
            if counted:
                utilized_total += hours
            flags = []
            if not counted:
                flags.append(f"excluded({status})")
            if todays and instance is None:
                flags.append("NOT IN ANY INSTANCE -- adjacent night's unit")
            elif instance and abs(raw_hours - hours) > 0.01:
                flags.append(f"clipped from {raw_hours:.2f}h to the instance")
            if span is None:
                flags.append(f"NO SPAN ({start_key}/{end_key} missing)")
            elif raw_hours > IMPLAUSIBLE_SPAN_HOURS:
                flags.append(f"RAW SPAN {raw_hours:.1f}h -- likely never cleared")
            s_start = R.parse_ts(stamps.get(start_key))
            s_end = R.parse_ts(stamps.get(end_key))
            if s_start and s_end and s_end.date() != s_start.date():
                flags.append("crosses-midnight")
            print(
                f"    pickup={_fmt(leg.get('pickup_time')):<22} "
                f"{start_key}={_fmt(stamps.get(start_key)):<22} "
                f"{end_key}={_fmt(stamps.get(end_key)):<22} "
                f"span={hours:6.2f}h  {' '.join(flags)}"
            )
        print(f"    -> utilized_hours = {utilized_total:.2f}")

        dump["profiles"][name] = {
            "shift_records": rows,
            "legs": [
                {
                    "trip_status": leg.get("trip_status"),
                    "pickup_time": leg.get("pickup_time"),
                    "timestamps": R.timestamp_map(leg),
                }
                for leg in profile_legs
            ],
        }

    # ---- which stamp pair to trust ----
    #
    # enroute -> clear is the natural definition of a committed unit, but it is
    # only as good as the crew's habit of pressing Clear. When Clear is really
    # being set as the next call is assigned, consecutive legs butt up against
    # each other and the span tiles the whole shift. That is measurable, so
    # measure it rather than arguing about it.
    print()
    print("=" * 78)
    print("CANDIDATE SPANS")
    print("=" * 78)

    todays_all = {
        name: [s for s in spans if s[0].date() == target]
        for name, spans in R.shift_instances(shifts, shift_offset).items()
    }
    todays_all = {n: s for n, s in todays_all.items() if s}
    scheduled_total = sum(
        (e - s).total_seconds() for spans in todays_all.values() for s, e in spans
    ) / 3600.0

    print(f"\n  Scheduled unit-hours on {target}: {scheduled_total:.1f}")
    print(f"  {'span':<12} {'keys':<34} {'legs':>6} {'missing':>8} {'hours':>9} {'UHU':>7}")
    for label, (s_key, e_key) in R.UHU_SPANS.items():
        hours = 0.0
        counted = missing = 0
        for name, spans in todays_all.items():
            for leg in legs_by_profile.get(name, []):
                status = (leg.get("trip_status") or "").strip().lower()
                if status in R.UHU_EXCLUDED_TRIP_STATUSES:
                    continue
                instance = R.assign_leg(leg, spans, s_key, e_key)
                if instance is None:
                    continue
                stamps = R.timestamp_map(leg)
                a = R.parse_ts(stamps.get(s_key))
                b = R.parse_ts(stamps.get(e_key))
                if not a or not b or b <= a:
                    missing += 1
                    continue
                counted += 1
                hours += R.overlap_minutes(a, b, instance[0], instance[1]) / 60.0
        ratio = hours / scheduled_total if scheduled_total else 0.0
        print(f"  {label:<12} {s_key + ' -> ' + e_key:<34} {counted:>6} {missing:>8} "
              f"{hours:>9.1f} {ratio:>7.1%}")

    # ---- is `clear` really the end of the call? ----
    print()
    print("  Gap between one leg clearing and the same unit going enroute again:")
    gaps = []
    for name, spans in todays_all.items():
        ordered = []
        for leg in legs_by_profile.get(name, []):
            status = (leg.get("trip_status") or "").strip().lower()
            if status in R.UHU_EXCLUDED_TRIP_STATUSES:
                continue
            stamps = R.timestamp_map(leg)
            a = R.parse_ts(stamps.get("enroute"))
            b = R.parse_ts(stamps.get("clear"))
            if a and b:
                ordered.append((a, b))
        ordered.sort()
        for (_, prev_clear), (next_enroute, _) in zip(ordered, ordered[1:]):
            gaps.append((next_enroute - prev_clear).total_seconds())

    if gaps:
        gaps.sort()
        within = lambda s: sum(1 for g in gaps if g <= s)
        print(f"    {len(gaps)} consecutive pairs")
        print(f"    within 10s : {within(10):>4}  ({within(10) / len(gaps):.0%})")
        print(f"    within 60s : {within(60):>4}  ({within(60) / len(gaps):.0%})")
        print(f"    within 5m  : {within(300):>4}  ({within(300) / len(gaps):.0%})")
        print(f"    median     : {gaps[len(gaps) // 2] / 60.0:.1f} min")
        if within(60) > len(gaps) * 0.5:
            print()
            print("    More than half the time this unit goes enroute on the next call")
            print("    within a minute of clearing the last one. Clear is being set as")
            print("    the next call is assigned, not when the call ended, so")
            print("    enroute -> clear tiles the shift and overstates utilization.")
    else:
        print("    (no consecutive pairs to compare)")
    print()

    out = Path("uhu_diagnosis.json")
    out.write_text(json.dumps(dump, indent=2, default=str), encoding="utf-8")
    print()
    print(f"Written to {out.resolve()} -- send me this file.")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
