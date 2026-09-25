"""
Whether Samsara can be the source of truth for vehicle usage, and how far back.

The Daily Vehicle Overview currently says which trucks dispatch assigned calls
to, not which trucks moved -- docs/VEHICLE_USAGE_DATA_SOURCES.md sets out why.
The decision has been taken to move In Service Used / In Service Not Used onto
Samsara instead:

    Unused   a vehicle whose longest continuous engine-off block within the
             tenant-local calendar day reaches 8 hours
    Status   Samsara's `Asset Status` attribute, not Traumasoft's
    Roster   Samsara. A vehicle absent from Samsara is out of scope.
    Join     VIN

Three things have to be true before that can be built, and none of them are
knowable from the documentation:

  1. the stats history endpoint returns engine states in the shape assumed;
  2. history reaches far enough back to matter -- a separate worksheet wants
     May, June and July 2026, and whether those are answerable at all rests
     entirely on Samsara's retention, which is not published;
  3. an 8-hour threshold actually separates the fleet. A rule that marks
     nearly every vehicle Unused is not a rule, and overnight shutdown makes
     that a real possibility on a calendar-day window.

So this measures all three against the live API rather than asserting them.
It reports the endpoint's real response shape before computing anything from
it, because the cost of assuming wrongly here is a number that looks right.

STRICTLY READ-ONLY. The Samsara client is constructed read_only=True and
nothing here passes write=True.

Vehicle names and movement times are operational, not patient data. Safe to
paste.

Usage:
    python probe_samsara_movement.py
    python probe_samsara_movement.py --day 2026-09-22 --json movement.json
    python probe_samsara_movement.py --threshold 8
"""

import argparse
import json
import logging
import os
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone

log = logging.getLogger(__name__)

STATS_PATH = "fleet/vehicles/stats/history"

# Samsara reports engine state as a small vocabulary. Only "off" is treated as
# the vehicle not being used, per the decision that engine-on counts as moved:
# an EMS unit idles with the engine running for climate and equipment power,
# and that is a truck in use. The real vocabulary is reported rather than
# assumed -- anything unrecognised is listed rather than silently bucketed.
OFF_VALUES = {"off"}

# Depths, in days back from today, at which to ask for a day of history. The
# ladder is what establishes retention; the explicit dates are the months the
# vehicle status worksheet asks about.
RETENTION_DEPTHS = [1, 7, 30, 60, 90, 120, 150, 180, 270, 365]
WORKSHEET_DAYS = ["2026-05-15", "2026-06-15", "2026-07-15"]


def tenant_zone():
    """
    The tenant's zone, for calendar-day boundaries.

    A calendar day is the window the rule is measured over, so getting this
    wrong shifts every boundary by the offset and silently reassigns a night
    shift's stillness to the wrong date.
    """
    name = os.getenv("SAMSARA_TENANT_TIMEZONE", "").strip()
    if name:
        try:
            from zoneinfo import ZoneInfo
            return ZoneInfo(name), name
        except Exception as exc:  # noqa: BLE001 -- bad zone name is user input
            log.warning("SAMSARA_TENANT_TIMEZONE=%r unusable (%s)", name, exc)

    fixed = os.getenv("SAMSARA_TENANT_UTC_OFFSET", "").strip()
    if fixed:
        sign = -1 if fixed.startswith("-") else 1
        body = fixed.lstrip("+-")
        hours, _, minutes = body.partition(":")
        delta = timedelta(hours=int(hours or 0), minutes=int(minutes or 0)) * sign
        return timezone(delta), f"fixed {fixed}"

    raise RuntimeError(
        "Neither SAMSARA_TENANT_TIMEZONE nor SAMSARA_TENANT_UTC_OFFSET is set. "
        "A calendar-day rule cannot be measured without one -- see .env.example."
    )


def day_window(day, zone, lookback_hours=24):
    """
    (start, end) in RFC3339 UTC for one tenant-local calendar day.

    `lookback_hours` extends the start backwards only to establish what state
    a vehicle was already in at midnight. Without it a truck shut down at 20:00
    yesterday looks like it has no state until its first event today, and the
    stillness that actually spans the boundary is lost.
    """
    start_local = datetime(day.year, day.month, day.day, tzinfo=zone)
    end_local = start_local + timedelta(days=1)
    fetch_from = start_local - timedelta(hours=lookback_hours)
    fmt = lambda m: m.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")  # noqa: E731
    return fmt(fetch_from), fmt(end_local), start_local, end_local


def parse_ts(text):
    if not text:
        return None
    raw = str(text).strip().replace("Z", "+00:00")
    try:
        moment = datetime.fromisoformat(raw)
    except ValueError:
        return None
    return moment if moment.tzinfo else moment.replace(tzinfo=timezone.utc)


def fetch_engine_states(samsara, start, end):
    """One stats-history query for engineStates over a window, all vehicles."""
    return list(
        samsara.paginate(
            STATS_PATH,
            params={"startTime": start, "endTime": end, "types": "engineStates"},
        )
    )


# =============================
# SECTIONS
# =============================
def report_shape(samsara, zone, day, out):
    """
    The endpoint's real response, before anything is computed from it.

    Printed as observed keys and a redacted sample rather than checked against
    an expectation, because the point is to find out where the assumption is
    wrong, and an assertion would only say that it is.
    """
    print("\n1. STATS HISTORY ENDPOINT")
    print("   " + "-" * 66)
    start, end, _, _ = day_window(day, zone, lookback_hours=0)
    print(f"   GET /{STATS_PATH}")
    print(f"   types=engineStates  {start} .. {end}\n")

    try:
        rows = fetch_engine_states(samsara, start, end)
    except Exception as exc:  # noqa: BLE001 -- reporting the failure IS the result
        print(f"   FAILED: {type(exc).__name__}: {exc}")
        print("\n   If this is a 4xx the endpoint or its parameters differ from")
        print("   what was assumed, and nothing below can be trusted. If it is a")
        print("   permission error the API token lacks the telematics scope.")
        out["endpoint"] = {"ok": False, "error": str(exc)}
        return None

    print(f"   OK -- {len(rows)} vehicle row(s) returned.")
    if not rows:
        print("   No vehicles in the window. Try another --day.")
        out["endpoint"] = {"ok": True, "rows": 0}
        return rows

    sample = rows[0]
    print(f"\n   Keys on a vehicle row: {', '.join(sorted(sample.keys()))}")
    states = sample.get("engineStates") or []
    print(f"   engineStates entries on that row: {len(states)}")
    if states:
        print(f"   Keys on an engineState: {', '.join(sorted(states[0].keys()))}")
        print(f"   Sample: {json.dumps(states[0], default=str)[:120]}")

    vocabulary = Counter()
    for row in rows:
        for event in row.get("engineStates") or []:
            vocabulary[str(event.get("value"))] += 1
    print(f"\n   Engine state vocabulary: "
          f"{', '.join(f'{v} ({c})' for v, c in vocabulary.most_common())}")
    unknown = {v for v in vocabulary if v.lower() not in OFF_VALUES} - {"On", "Idle", "Running"}
    if unknown:
        print(f"   NOT RECOGNISED, treated as engine-on: {', '.join(sorted(unknown))}")

    out["endpoint"] = {
        "ok": True,
        "rows": len(rows),
        "vehicle_keys": sorted(sample.keys()),
        "state_vocabulary": dict(vocabulary),
    }
    return rows


def report_retention(samsara, zone, out):
    """
    How far back history actually reaches.

    This is the finding the whole backfill question turns on, and Samsara does
    not publish it, so it is asked rather than looked up: one day of history at
    increasing depth until it stops coming back.
    """
    print("\n2. HISTORY RETENTION")
    print("   " + "-" * 66)
    print("   One day of engineStates at increasing depth.\n")
    print(f"   {'date':<14}{'days back':>10}{'vehicles':>10}{'events':>9}   result")
    print("   " + "-" * 66)

    today = date.today()
    probes = [(today - timedelta(days=d), d) for d in RETENTION_DEPTHS]
    probes += [
        (date.fromisoformat(d), (today - date.fromisoformat(d)).days)
        for d in WORKSHEET_DAYS
    ]

    findings = []
    for day, depth in sorted(probes, key=lambda p: p[1]):
        start, end, _, _ = day_window(day, zone, lookback_hours=0)
        try:
            rows = fetch_engine_states(samsara, start, end)
            events = sum(len(r.get("engineStates") or []) for r in rows)
            with_data = sum(1 for r in rows if r.get("engineStates"))
            verdict = "data" if events else "EMPTY"
            print(f"   {day.isoformat():<14}{depth:>10}{with_data:>10}{events:>9}   {verdict}")
            findings.append({"date": day.isoformat(), "days_back": depth,
                             "vehicles": with_data, "events": events, "ok": True})
        except Exception as exc:  # noqa: BLE001
            print(f"   {day.isoformat():<14}{depth:>10}{'-':>10}{'-':>9}   "
                  f"{type(exc).__name__}: {str(exc)[:28]}")
            findings.append({"date": day.isoformat(), "days_back": depth,
                             "ok": False, "error": str(exc)})

    deepest = max((f for f in findings if f.get("events")), key=lambda f: f["days_back"],
                  default=None)
    if deepest:
        print(f"\n   Deepest day with data: {deepest['date']} "
              f"({deepest['days_back']} days back).")
        worksheet = [f for f in findings if f["date"] in WORKSHEET_DAYS]
        answered = [f for f in worksheet if f.get("events")]
        if len(answered) == len(worksheet) and worksheet:
            print("   All three worksheet months are reachable.")
        elif answered:
            print(f"   Worksheet months reachable: "
                  f"{', '.join(f['date'] for f in answered)} -- the rest are not.")
        else:
            print("   NONE of the worksheet months have data. Those columns")
            print("   cannot be backfilled from Samsara either.")
    else:
        print("\n   No day returned any data. Either the token lacks telematics")
        print("   scope or the endpoint is not what was assumed -- see section 1.")

    out["retention"] = findings


def longest_off_block(events, day_start, day_end):
    """
    The longest continuous engine-off stretch inside the day, in hours.

    Returns (hours, had_any_signal). The second value matters as much as the
    first: a vehicle with no events and no carried-in state is a gateway that
    said nothing, not a truck that sat still, and counting those as Unused
    would quietly inflate every total.
    """
    timeline = []
    for event in events:
        moment = parse_ts(event.get("time"))
        if moment:
            timeline.append((moment, str(event.get("value") or "").lower()))
    timeline.sort()
    if not timeline:
        return None, False

    # The state in force at midnight is the last one set before it.
    carried = None
    for moment, value in timeline:
        if moment <= day_start:
            carried = value
        else:
            break
    inside = [(m, v) for m, v in timeline if day_start < m < day_end]
    if carried is None and not inside:
        return None, False

    longest = timedelta(0)
    state = carried
    cursor = day_start
    for moment, value in inside + [(day_end, None)]:
        if state in OFF_VALUES:
            longest = max(longest, moment - cursor)
        cursor, state = moment, value
    return longest.total_seconds() / 3600.0, True


def report_idle_rule(rows, zone, day, threshold, out):
    """
    What the threshold actually selects, shown as a distribution first.

    A cutoff nobody has seen the shape of is a guess. The bands are printed
    before the verdict so the choice of eight hours can be judged against the
    fleet rather than defended after the fact.
    """
    print(f"\n3. THE {threshold}-HOUR RULE ON {day.isoformat()}")
    print("   " + "-" * 66)
    if not rows:
        print("   No data for this day.")
        out["rule"] = None
        return

    _, _, day_start, day_end = day_window(day, zone)
    scored, silent = {}, []
    for row in rows:
        name = row.get("name") or row.get("id") or "?"
        hours, had_signal = longest_off_block(
            row.get("engineStates") or [], day_start, day_end
        )
        if not had_signal:
            silent.append(name)
        else:
            scored[name] = hours

    bands = [(0, 1), (1, 4), (4, 8), (8, 12), (12, 20), (20, 24.01)]
    print(f"   {'longest engine-off block':<30}{'vehicles':>9}")
    print("   " + "-" * 42)
    for low, high in bands:
        count = sum(1 for h in scored.values() if low <= h < high)
        label = f"{low:g} to {high:g} h" if high <= 24 else f"{low:g} h and up"
        print(f"   {label:<30}{count:>9}")

    unused = sorted(n for n, h in scored.items() if h >= threshold)
    used = sorted(n for n, h in scored.items() if h < threshold)
    print(f"\n   Vehicles with usable signal: {len(scored)}")
    print(f"   Unused (>= {threshold}h still):    {len(unused)}")
    print(f"   Used   (<  {threshold}h still):    {len(used)}")
    if silent:
        print(f"   No signal at all:            {len(silent)}")
        print("   Those are reported separately, never as Unused -- a silent")
        print("   gateway is not an idle truck.")
        for name in silent[:15]:
            print(f"      {name}")

    if scored and len(unused) / len(scored) > 0.8:
        print(f"\n   WARNING: {threshold}h marks "
              f"{100.0 * len(unused) / len(scored):.0f}% of the fleet Unused.")
        print("   A rule that selects nearly everything is not separating")
        print("   anything. The bands above suggest where a cutoff would.")

    out["rule"] = {
        "day": day.isoformat(),
        "threshold_hours": threshold,
        "scored": len(scored),
        "unused": unused,
        "used": used,
        "no_signal": silent,
        "hours": {k: round(v, 2) for k, v in sorted(scored.items())},
    }


def report_vs_dispatch(rows, zone, day, threshold, out):
    """
    How the new definition differs from the one it replaces.

    The existing sheet calls a truck used when a leg was assigned to it. Both
    answers exist for the same day, so the disagreement can be counted instead
    of discovered after the change ships.
    """
    print("\n4. AGAINST THE CURRENT DISPATCH-BASED DEFINITION")
    print("   " + "-" * 66)
    try:
        from traumasoft_api import TraumasoftAPI
        import traumasoft_reports as R
        api = TraumasoftAPI()
        legs = api.get_trips(day.isoformat(), range_days=1)
        ts_vehicles = api.list_vehicles()
    except Exception as exc:  # noqa: BLE001
        print(f"   Traumasoft unavailable ({type(exc).__name__}: {exc}).")
        print("   Skipping the comparison.")
        out["comparison"] = None
        return

    used_ids = R.used_vehicle_ids(legs)
    ts_by_id = {str(v.get("id")): v for v in ts_vehicles}
    dispatch_used = {
        str(ts_by_id[i].get("name") or "").strip()
        for i in used_ids if i in ts_by_id
    }

    _, _, day_start, day_end = day_window(day, zone)
    movement_used = set()
    for row in rows or []:
        hours, had_signal = longest_off_block(
            row.get("engineStates") or [], day_start, day_end
        )
        if had_signal and hours < threshold:
            movement_used.add(str(row.get("name") or "").strip())

    both = dispatch_used & movement_used
    moved_only = sorted(movement_used - dispatch_used)
    dispatched_only = sorted(dispatch_used - movement_used)

    print(f"   Dispatch says used: {len(dispatch_used)}")
    print(f"   Movement says used: {len(movement_used)}")
    print(f"   Both agree:         {len(both)}\n")
    if moved_only:
        print(f"   Moved but never dispatched ({len(moved_only)}) -- the gap the")
        print(f"   current sheet cannot see:\n      {', '.join(moved_only)[:300]}")
    if dispatched_only:
        print(f"\n   Dispatched but reads still ({len(dispatched_only)}) -- worth a")
        print(f"   look, a leg was assigned yet the engine stayed off:")
        print(f"      {', '.join(dispatched_only)[:300]}")

    out["comparison"] = {
        "dispatch_used": sorted(dispatch_used),
        "movement_used": sorted(movement_used),
        "moved_not_dispatched": moved_only,
        "dispatched_not_moved": dispatched_only,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--day", help="Day to analyse, YYYY-MM-DD. Default yesterday.")
    parser.add_argument("--threshold", type=float, default=8.0,
                        help="Hours of continuous engine-off that mark a vehicle "
                             "Unused (default 8).")
    parser.add_argument("--skip-retention", action="store_true",
                        help="Skip the retention ladder, which is the slow part.")
    parser.add_argument("--json", dest="json_path", help="Also write findings as JSON.")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    day = date.fromisoformat(args.day) if args.day else date.today() - timedelta(days=1)

    try:
        zone, zone_label = tenant_zone()
    except RuntimeError as exc:
        log.error("%s", exc)
        return 2

    try:
        from samsara_api import SamsaraClient
        samsara = SamsaraClient(read_only=True)
    except Exception as exc:  # noqa: BLE001
        log.error("Samsara not reachable (%s: %s).", type(exc).__name__, exc)
        return 2

    print("=" * 72)
    print("  Samsara movement probe -- can Samsara define vehicle usage?")
    print(f"  Day {day}   zone {zone_label}   threshold {args.threshold}h")
    print("  Read-only. GET only, no writes.")
    print("=" * 72)

    out = {"day": day.isoformat(), "zone": zone_label, "threshold": args.threshold}
    rows = report_shape(samsara, zone, day, out)
    if not args.skip_retention:
        report_retention(samsara, zone, out)

    # Re-fetch with the midnight lookback so the rule sees state carried in
    # across the boundary; section 1 deliberately asked for the bare day.
    if rows is not None:
        start, end, _, _ = day_window(day, zone)
        try:
            rows = fetch_engine_states(samsara, start, end)
        except Exception as exc:  # noqa: BLE001
            log.warning("Could not re-fetch with lookback: %s", exc)
        report_idle_rule(rows, zone, day, args.threshold, out)
        report_vs_dispatch(rows, zone, day, args.threshold, out)

    print("\n" + "=" * 72)
    print("  Safe to paste. Operational values only, no patient data.")
    print("=" * 72 + "\n")

    if args.json_path:
        with open(args.json_path, "w", encoding="utf-8") as handle:
            json.dump(out, handle, indent=2, default=str)
        print(f"  Written to {args.json_path}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
