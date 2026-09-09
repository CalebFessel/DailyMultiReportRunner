"""
Publish a day of Traumasoft trips to Samsara as routes.

One route per vehicle per day. Each scheduled leg contributes a pickup stop
and a drop-off stop, carrying the call type and level of service in the stop
notes, and the vehicle is joined to Samsara on the unit prefix the two systems
share.

    python push_samsara_routes.py                    # tomorrow, dry run
    python push_samsara_routes.py 2026-09-10         # that day, dry run
    python push_samsara_routes.py --publish          # actually create them
    python push_samsara_routes.py --replace --publish  # delete ours first, then create
    python push_samsara_routes.py --vehicle M-12     # one unit only
    python push_samsara_routes.py --json plan.json   # write the payloads for inspection

DRY RUN IS THE DEFAULT. Without --publish this reads from both systems, builds
the exact payloads it would send, prints the plan, and stops. Nothing in
Samsara changes until --publish is passed; the Samsara client is constructed
read-only until then, so a bug cannot write on a dry run.

WHAT GETS SKIPPED, AND WHY
    Legs with no scheduled pickup_time, no location on one end, no vehicle, or
    a call type in SAMSARA_EXCLUDED_CALL_TYPES. 911 and on-demand work has no
    schedule to route against. Every skipped leg is counted by reason in the
    plan, so a day that pushes fewer runs than expected explains itself.

    Vehicles that could not be joined to Samsara are listed too. Fix those in
    state/samsara_vehicle_overrides.json rather than renaming units.
"""

import argparse
import json
import logging
import sys
from collections import Counter
from datetime import date, datetime, timedelta

import samsara_routes as SR
import traumasoft_reports as R
from samsara_api import SamsaraAPIError, SamsaraClient, SamsaraReadOnlyError
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

log = logging.getLogger(__name__)

# Stamped on every route this tool creates so --replace can recognise its own
# work and leave routes built by hand or by another system alone.
ROUTE_NAME_PREFIX = "[TS]"


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Publish Traumasoft trips to Samsara as routes.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "trip_date",
        nargs="?",
        help="Day to publish, YYYY-MM-DD. Defaults to tomorrow.",
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Actually create the routes. Without this the run is a dry run.",
    )
    parser.add_argument(
        "--replace",
        action="store_true",
        help=f"Delete existing {ROUTE_NAME_PREFIX} routes for the day before creating new ones.",
    )
    parser.add_argument(
        "--vehicle",
        action="append",
        help="Limit to this Traumasoft vehicle name. Repeatable.",
    )
    parser.add_argument("--json", dest="json_path", help="Write the route payloads to this file.")
    parser.add_argument("--verbose", action="store_true", help="Debug logging.")
    return parser.parse_args(argv)


def resolve_date(text):
    if not text:
        return date.today() + timedelta(days=1)
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except ValueError:
        raise SystemExit(f"Not a date: {text!r}. Use YYYY-MM-DD.")


def print_plan(day, routes, notes, skipped, matched, unmatched, ambiguous, publishing):
    """The whole decision surface in one screen, before anything is sent."""
    total_stops = sum(len(r["stops"]) for r in routes)
    print()
    print("=" * 72)
    print(f"  Samsara route plan for {day.isoformat()}"
          f"{'  (DRY RUN -- nothing will be sent)' if not publishing else '  (PUBLISHING)'}")
    print("=" * 72)
    print(f"  Routes: {len(routes)}    Stops: {total_stops}    "
          f"Legs routed: {sum(n['legs'] for n in notes)}")
    print()

    if notes:
        print(f"  {'Unit':<16}{'Samsara vehicle':<26}{'Legs':>5}{'Stops':>7}{'Est. DO':>9}")
        print(f"  {'-' * 62}")
        for note in notes:
            print(f"  {note['vehicle_name'][:15]:<16}"
                  f"{(note['samsara_vehicle'] or '')[:25]:<26}"
                  f"{note['legs']:>5}{note['stops']:>7}{note['estimated_dropoff_times']:>9}")
        print()
        estimated = sum(n["estimated_dropoff_times"] for n in notes)
        if estimated:
            print(f"  {estimated} drop-off time(s) had no appointment time and were estimated at")
            print(f"  pickup + {SR.DEFAULT_TRANSPORT_MINUTES} min "
                  f"(SAMSARA_DEFAULT_TRANSPORT_MINUTES).")
            print()

    if skipped:
        print("  Legs not routed:")
        for reason, count in Counter(reason for _, reason in skipped).most_common():
            print(f"    {count:>5}  {reason}")
        print()

    if unmatched:
        print("  Traumasoft units with no Samsara vehicle "
              f"({len(unmatched)}) -- their trips were dropped:")
        for name in unmatched:
            print(f"    - {name}  (prefix {SR.unit_prefix(name)!r})")
        print(f"  Map these in {SR.VEHICLE_OVERRIDES_FILE}")
        print()

    if ambiguous:
        print("  Units whose prefix matched several Samsara vehicles -- dropped as unsafe:")
        for name, hits in ambiguous.items():
            print(f"    - {name} -> {', '.join(v.get('name', '?') for v in hits)}")
        print(f"  Disambiguate in {SR.VEHICLE_OVERRIDES_FILE}")
        print()

    if not routes:
        print("  Nothing to publish.")
        print()


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    day = resolve_date(args.trip_date)

    # --- Traumasoft: the day's legs -------------------------------------
    try:
        api = TraumasoftAPI()
        legs = api.get_trips(day.isoformat(), range_days=1)
    except TraumasoftAPIError as exc:
        log.error("Traumasoft API failed: %s", exc)
        return 2
    except RuntimeError as exc:
        log.error("%s", exc)
        return 2

    log.info("Traumasoft returned %d leg(s) for %s", len(legs), day)

    if args.vehicle:
        wanted = {v.strip().upper() for v in args.vehicle}
        legs = [
            leg for leg in legs
            if str(leg.get("vehicle_name") or "").strip().upper() in wanted
            or SR.unit_prefix(leg.get("vehicle_name")) in {SR.unit_prefix(v) for v in wanted}
        ]
        log.info("Filtered to %d leg(s) for %s", len(legs), ", ".join(sorted(wanted)))

    kept, skipped = SR.eligible_legs(legs, R.parse_ts_aware)
    log.info("%d leg(s) eligible, %d skipped", len(kept), len(skipped))

    # Trip stamps normally carry their own offset; this is the fallback for
    # any that do not, taken from the day's own data.
    default_offset = R.tenant_utc_offset(legs)

    # --- Samsara: vehicles, addresses -----------------------------------
    # Read-only unless the caller asked to publish, so a dry run cannot write
    # even if something below is wrong.
    try:
        samsara = SamsaraClient(read_only=not args.publish)
        vehicles = samsara.list_vehicles()
        addresses = samsara.list_addresses()
    except SamsaraAPIError as exc:
        log.error("Samsara API failed: %s\n%s", exc, exc.body or "")
        return 2
    except RuntimeError as exc:
        log.error("%s", exc)
        return 2

    log.info("Samsara returned %d vehicle(s) and %d address(es)", len(vehicles), len(addresses))

    ts_names = sorted({str(leg.get("vehicle_name") or "").strip() for leg in kept} - {""})
    overrides = SR.load_vehicle_overrides()
    matched, unmatched, ambiguous = SR.match_vehicles(ts_names, vehicles, overrides)

    routable = [leg for leg in kept if str(leg.get("vehicle_name") or "").strip() in matched]
    for leg in kept:
        name = str(leg.get("vehicle_name") or "").strip()
        if name not in matched:
            skipped.append((leg, f"vehicle {name!r} not found in Samsara"))

    routes, notes = SR.build_routes(
        routable,
        matched,
        day,
        R.parse_ts_aware,
        address_index=SR.index_addresses(addresses),
        default_offset=default_offset,
        name_prefix=ROUTE_NAME_PREFIX,
    )

    print_plan(day, routes, notes, skipped, matched, unmatched, ambiguous, args.publish)

    if args.json_path:
        with open(args.json_path, "w", encoding="utf-8") as handle:
            json.dump(routes, handle, indent=2)
        print(f"  Payloads written to {args.json_path}")
        print()

    if not args.publish:
        print("  Dry run. Re-run with --publish to create these in Samsara.")
        print()
        return 0
    if not routes:
        return 0

    # --- Publish ---------------------------------------------------------
    if args.replace:
        window_start = SR.rfc3339(datetime.combine(day, datetime.min.time()), default_offset)
        window_end = SR.rfc3339(
            datetime.combine(day + timedelta(days=1), datetime.min.time()), default_offset
        )
        try:
            existing = samsara.list_routes(window_start, window_end)
        except SamsaraAPIError as exc:
            log.error("Could not list existing routes: %s", exc)
            return 2
        ours = [r for r in existing if str(r.get("name", "")).startswith(ROUTE_NAME_PREFIX)]
        log.info("Replacing %d existing %s route(s)", len(ours), ROUTE_NAME_PREFIX)
        for route in ours:
            try:
                samsara.delete_route(route["id"])
            except SamsaraAPIError as exc:
                log.error("Could not delete route %s (%s): %s", route.get("id"),
                          route.get("name"), exc)
                return 2

    created, failed = 0, 0
    for route in routes:
        try:
            samsara.create_route(route)
            created += 1
            log.info("Created %s (%d stops)", route["name"], len(route["stops"]))
        except SamsaraReadOnlyError as exc:
            log.error("%s", exc)
            return 2
        except SamsaraAPIError as exc:
            failed += 1
            log.error("Failed %s: %s\n%s", route["name"], exc, exc.body or "")

    print(f"  Created {created} route(s), {failed} failed.")
    print()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
