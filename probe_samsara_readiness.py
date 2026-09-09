"""
Answer the questions the Samsara route mapping currently guesses at.

Strictly read-only. GETs a window of trips from Traumasoft and, if a Samsara
token is present, the vehicle list from Samsara, then reports:

    1. Field coverage      -- how often the fields the mapping depends on are
                              actually populated
    2. Timezone handling   -- whether trip stamps carry a UTC offset, which
                              decides whether every dispatch time is right or
                              four hours early
    3. Call type / LOS     -- the real vocabulary, and which values the
                              exclusion list currently catches
    4. Eligibility funnel  -- how many legs would reach Samsara, and exactly
                              why the rest would not
    5. Drop-off provenance -- appointment time vs ETA vs 45-minute estimate
    6. Vehicle join        -- the real Traumasoft x Samsara match table

    python probe_samsara_readiness.py              # last 7 days
    python probe_samsara_readiness.py --days 30    # a fuller vocabulary
    python probe_samsara_readiness.py --json out.json

PHI: THIS SCRIPT NEVER PRINTS PATIENT-IDENTIFYING VALUES. Patient names, MRNs,
street addresses and phone numbers are counted, never shown -- coverage only.
Vehicle names, call types and levels of service are operational and are shown
in full, because those are the values the mapping has to be configured
against. The output is safe to paste into a chat or an issue.
"""

import argparse
import json
import logging
import sys
from collections import Counter, defaultdict
from datetime import date, timedelta

import samsara_routes as SR
import traumasoft_reports as R
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

log = logging.getLogger(__name__)

# Fields the route mapping reads. Shown as coverage percentages so a field
# that is populated on a tenth of legs is obvious before it silently halves
# the number of routable trips.
MAPPING_FIELDS = [
    "vehicle_name", "pickup_time", "appt_time", "dropoff_eta", "call_type", "los",
    "pu_facility_name", "pu_lat", "pu_lon", "pu_address1", "pu_city", "pu_state",
    "do_facility_name", "do_lat", "do_lon", "do_address1", "do_city", "do_state",
    "run_number", "trip_number", "response_priority", "transport_priority",
    # Would carry the tenant's zone, and would place a future day's trips
    # date-correctly without configuring anything. Empty on this tenant, which
    # is exactly why SAMSARA_TENANT_TIMEZONE has to be set instead.
    "timezone",
]

# Counted, never printed. Whether these are populated matters -- what they say
# does not, and must not leave the machine.
PHI_FIELDS = [
    "patient_first_name", "patient_last_name", "patient_mrn", "patient_id",
    "pu_phone", "do_phone",
]


def populated(value):
    if value is None:
        return False
    text = str(value).strip()
    if not text or text.lower() in ("none", "null"):
        return False
    # 0/0 coordinates are the Atlantic, not a location.
    if text in ("0", "0.0"):
        return False
    return True


def pct(count, total):
    return f"{(100.0 * count / total):5.1f}%" if total else "    -"


def bar(count, total, width=24):
    filled = int(round(width * count / total)) if total else 0
    return "█" * filled + "·" * (width - filled)


# =============================
# SECTIONS
# =============================
def report_coverage(legs, out):
    total = len(legs)
    print(f"\n1. FIELD COVERAGE  ({total} legs)")
    print("   " + "-" * 66)
    coverage = {}
    for field in MAPPING_FIELDS:
        hits = sum(1 for leg in legs if populated(leg.get(field)))
        coverage[field] = hits
        flag = ""
        if field in ("pickup_time", "vehicle_name") and hits < total:
            flag = "  <- required, gaps drop the leg"
        elif field in ("pu_lat", "do_lat") and total and hits / total < 0.9:
            flag = "  <- falls back to postal address"
        print(f"   {field:<22}{pct(hits, total)}  {bar(hits, total)}{flag}")

    print(f"\n   PHI fields (counted, not shown):")
    for field in PHI_FIELDS:
        hits = sum(1 for leg in legs if populated(leg.get(field)))
        coverage[field] = hits
        print(f"   {field:<22}{pct(hits, total)}  populated on {hits} leg(s)")
    out["coverage"] = coverage
    out["total_legs"] = total


def report_timezones(legs, out):
    print("\n2. TIMESTAMP OFFSETS")
    print("   " + "-" * 66)
    with_offset = without = 0
    for leg in legs:
        parsed = R.parse_ts_aware(leg.get("pickup_time"))
        if parsed is None:
            continue
        if parsed.tzinfo is not None:
            with_offset += 1
        else:
            without += 1
    # The same resolver the publisher uses, so the probe reports the answer
    # that will actually be applied rather than a second opinion.
    tenant, source = SR.resolve_tenant_offset(legs, R.parse_ts_aware)

    print(f"   pickup_time carries an offset : {with_offset}")
    print(f"   pickup_time is naive          : {without}")
    print(f"   tenant offset                 : {SR.format_offset(tenant)}  (from {source})")
    if without and tenant is None:
        print("   !! Naive stamps and no resolvable offset. Publishing is refused")
        print("      rather than dispatching hours early.")
        print("      Set SAMSARA_TENANT_TIMEZONE=America/New_York (follows DST).")
    elif without:
        print(f"   Naive stamps will be treated as {SR.format_offset(tenant)}.")

    # This window is history, so status timestamps are available to it. A day
    # of FUTURE work has none, and that is the day this job exists to push.
    zones = Counter(
        str(leg.get("timezone") or "").strip() for leg in legs
        if str(leg.get("timezone") or "").strip()
    )
    if source == "status timestamps":
        print()
        print("   NOTE: that offset came from status timestamps, which only exist on")
        print("   trips that have already run. Tomorrow's trips carry none, so the")
        if zones:
            print(f"   publisher will fall back to the timezone field "
                  f"({', '.join(sorted(zones))}).")
        else:
            print("   timezone field is empty on this tenant too, so nothing in a day")
            print("   of future work states its zone.")
            print()
            print("   ACTION: set SAMSARA_TENANT_TIMEZONE=America/New_York in .env.")
            print("   An IANA zone is resolved against each trip's own date, so it")
            print("   follows daylight saving; a fixed SAMSARA_TENANT_UTC_OFFSET works")
            print("   but is wrong for half the year unless changed twice a year.")
    out["timestamps"] = {
        "with_offset": with_offset,
        "naive": without,
        "tenant_offset": SR.format_offset(tenant),
        "offset_source": source,
    }
    return tenant


def report_vocabulary(legs, out):
    print("\n3. CALL TYPE AND LEVEL OF SERVICE")
    print("   " + "-" * 66)
    call_types = Counter(str(leg.get("call_type") or "(blank)").strip() for leg in legs)
    print(f"   call_type -- {len(call_types)} distinct value(s)")
    excluded_now = []
    for value, count in call_types.most_common():
        hit = SR.is_excluded_call_type(value)
        if hit:
            excluded_now.append(value)
        print(f"     {count:>6}  {'EXCLUDED' if hit else '        '}  {value}")
    print(f"\n   Currently excluded by SAMSARA_EXCLUDED_CALL_TYPES: "
          f"{', '.join(excluded_now) if excluded_now else '(none)'}")
    print("   Patterns in effect: " + ", ".join(SR.EXCLUDED_CALL_TYPE_PATTERNS))
    print("   -> Anything above that should NOT reach a driver's route "
          "(911, standby,\n      cancellations) needs adding to that list.")

    los = Counter(str(leg.get("los") or "(blank)").strip() for leg in legs)
    print(f"\n   los -- {len(los)} distinct value(s)")
    for value, count in los.most_common():
        print(f"     {count:>6}  {value}")
    out["call_types"] = dict(call_types)
    out["los"] = dict(los)
    out["excluded_call_types"] = excluded_now


def report_funnel(legs, out):
    print("\n4. ELIGIBILITY FUNNEL")
    print("   " + "-" * 66)
    kept, skipped = SR.eligible_legs(legs, R.parse_ts_aware)
    total = len(legs)
    print(f"   {total} leg(s) in  ->  {len(kept)} routable "
          f"({pct(len(kept), total).strip()})")
    reasons = Counter(reason for _, reason in skipped)
    for reason, count in reasons.most_common():
        print(f"     {count:>6}  {reason}")
    out["funnel"] = {"total": total, "eligible": len(kept), "skipped": dict(reasons)}
    return kept


def report_dropoff_sources(kept, out, offset):
    """
    Where each drop-off time would come from, and how often a guessed one
    would have collided with the unit's next pickup.

    Grouped by vehicle and day exactly as the publisher groups a route, so
    the collision count is the real one rather than a hypothetical.

    The offset only affects how a time is rendered, never which source it
    came from, so an unresolved offset still gives a truthful breakdown.
    """
    print("\n5. DROP-OFF TIME PROVENANCE")
    print("   " + "-" * 66)
    offset = offset if offset is not None else timedelta(0)

    by_unit_day = defaultdict(list)
    for leg in kept:
        pickup = R.parse_ts_aware(leg.get("pickup_time"))
        if pickup is None:
            continue
        by_unit_day[(str(leg.get("vehicle_name") or "").strip(), pickup.date())].append(leg)

    sources = Counter()
    for legs_for_unit in by_unit_day.values():
        legs_for_unit.sort(key=lambda l: R.parse_ts_aware(l.get("pickup_time")))
        pickups = [R.parse_ts_aware(l.get("pickup_time")) for l in legs_for_unit]
        for index, leg in enumerate(legs_for_unit):
            next_pickup = pickups[index + 1] if index + 1 < len(pickups) else None
            _, source = SR.leg_stops(
                leg, R.parse_ts_aware, default_offset=offset, next_pickup=next_pickup
            )
            sources[(source or "none")] += 1
    total = sum(sources.values())
    for source, count in sources.most_common():
        label = {
            "appt_time": "appointment time (best -- a real due time)",
            "dropoff_eta": "drop-off ETA",
            "estimated": f"estimated at pickup + {SR.DEFAULT_TRANSPORT_MINUTES} min",
            "estimated (capped at next pickup)":
                "estimated, then CAPPED -- the guess ran into the next pickup",
        }.get(source, source)
        print(f"     {count:>6}  {pct(count, total)}  {label}")

    capped = sources.get("estimated (capped at next pickup)", 0)
    guessed = capped + sources.get("estimated", 0)
    if guessed:
        print(f"\n   {pct(guessed, total).strip()} of drop-offs are guesses.")
    if capped:
        print(f"   {capped} of them would have sorted past the unit's next pickup and")
        print("   told the driver to collect the next patient before delivering the")
        print(f"   one aboard. Lowering SAMSARA_DEFAULT_TRANSPORT_MINUTES from "
              f"{SR.DEFAULT_TRANSPORT_MINUTES} would")
        print("   reduce that; capping already prevents the misordering either way.")
    out["dropoff_sources"] = dict(sources)


def report_vehicle_join(kept, out):
    print("\n6. VEHICLE JOIN")
    print("   " + "-" * 66)
    ts_names = sorted({str(leg.get("vehicle_name") or "").strip() for leg in kept} - {""})
    leg_counts = Counter(str(leg.get("vehicle_name") or "").strip() for leg in kept)
    print(f"   {len(ts_names)} Traumasoft unit(s) carried trips in this window.")

    try:
        from samsara_api import SamsaraAPIError, SamsaraClient
        samsara = SamsaraClient(read_only=True)
        vehicles = samsara.list_vehicles()
    except Exception as exc:  # noqa: BLE001 -- any failure means "no Samsara half"
        print(f"\n   Samsara not reachable ({type(exc).__name__}: {exc}).")
        print("   Traumasoft-side prefixes only -- these are what the join uses:\n")
        print(f"   {'Traumasoft vehicle_name':<34}{'prefix':<14}{'legs':>6}")
        print("   " + "-" * 54)
        for name in ts_names:
            print(f"   {name[:33]:<34}{str(SR.unit_prefix(name)):<14}{leg_counts[name]:>6}")
        out["vehicles"] = {
            "traumasoft": [
                {"name": n, "prefix": SR.unit_prefix(n), "legs": leg_counts[n]}
                for n in ts_names
            ],
            "samsara": None,
        }
        return

    overrides = SR.load_vehicle_overrides()
    matched, unmatched, ambiguous = SR.match_vehicles(ts_names, vehicles, overrides)
    print(f"   {len(vehicles)} Samsara vehicle(s).")
    print(f"   Matched {len(matched)} / {len(ts_names)}   "
          f"unmatched {len(unmatched)}   ambiguous {len(ambiguous)}\n")

    print(f"   {'Traumasoft':<26}{'prefix':<12}{'Samsara':<26}{'legs':>5}")
    print("   " + "-" * 69)
    for name in ts_names:
        prefix = str(SR.unit_prefix(name))
        if name in matched:
            target = matched[name].get("name", "?")
        elif name in ambiguous:
            target = "AMBIGUOUS: " + ", ".join(v.get("name", "?") for v in ambiguous[name])
        else:
            target = "-- NO MATCH --"
        print(f"   {name[:25]:<26}{prefix[:11]:<12}{target[:25]:<26}{leg_counts[name]:>5}")

    lost = sum(leg_counts[n] for n in ts_names if n not in matched)
    if lost:
        print(f"\n   {lost} leg(s) would be dropped for want of a vehicle match.")
        print(f"   Fix in {SR.VEHICLE_OVERRIDES_FILE}")

    out["vehicles"] = {
        "traumasoft": [
            {"name": n, "prefix": SR.unit_prefix(n), "legs": leg_counts[n]}
            for n in ts_names
        ],
        "samsara": [v.get("name") for v in vehicles],
        "matched": {k: v.get("name") for k, v in matched.items()},
        "unmatched": unmatched,
        "ambiguous": {k: [v.get("name") for v in hits] for k, hits in ambiguous.items()},
        "legs_lost_to_no_match": lost,
    }


# =============================
# MAIN
# =============================
def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--days", type=int, default=7,
                        help="Days of trips to sample, back from today. Max 31.")
    parser.add_argument("--end", help="Last day of the window, YYYY-MM-DD. Defaults to today.")
    parser.add_argument("--json", dest="json_path", help="Also write the findings as JSON.")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    days = max(1, min(args.days, 31))
    end = date.fromisoformat(args.end) if args.end else date.today()
    start = end - timedelta(days=days - 1)

    try:
        api = TraumasoftAPI()
        legs = api.get_trips(start.isoformat(), range_days=days)
    except TraumasoftAPIError as exc:
        log.error("Traumasoft API failed: %s", exc)
        return 2
    except RuntimeError as exc:
        log.error("%s", exc)
        return 2

    print("=" * 72)
    print(f"  Samsara readiness probe -- {start} to {end} ({days}d, {len(legs)} legs)")
    print("  Read-only. No patient-identifying value is printed.")
    print("=" * 72)

    if not legs:
        print("\n  No trips in this window. Try a wider --days or a different --end.")
        return 0

    out = {"window": {"start": start.isoformat(), "end": end.isoformat(), "days": days}}
    report_coverage(legs, out)
    offset = report_timezones(legs, out)
    report_vocabulary(legs, out)
    kept = report_funnel(legs, out)
    if kept:
        report_dropoff_sources(kept, out, offset)
        report_vehicle_join(kept, out)

    print("\n" + "=" * 72)
    print("  Safe to paste. Values shown are operational, never patient data.")
    print("=" * 72 + "\n")

    if args.json_path:
        with open(args.json_path, "w", encoding="utf-8") as handle:
            json.dump(out, handle, indent=2, default=str)
        print(f"  Written to {args.json_path}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
