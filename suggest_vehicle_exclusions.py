"""
Propose a vehicle exclusion list for the fleet reports.

The old SQL pruned dead vehicle records with
`status_reason NOT LIKE '%duplicate%' / '%retired%' / '%disposed%' / '%test%'`
plus a hard-coded id list. The ThirdParty API's vehicle allowlist has no
status_reason, so a truck that was scrapped years ago but left sitting at
"Out of Service" looks exactly like one that broke down this morning.

Two signals stand in for it:

  * the vehicle's name -- test rigs and check trucks announce themselves;
  * whether it has run a leg recently -- trip history backfills about 90 days,
    and a unit that has not moved in that whole window is almost certainly not
    a live asset.

This proposes a list from both and writes it for review. It does not exclude
anything on its own: a genuinely broken truck can sit idle for months, and that
is precisely what the out-of-service report exists to show. Read the file,
delete anything that should stay, then keep it.

Usage:
    python suggest_vehicle_exclusions.py [--days 90] [--apply]

Without --apply it writes a proposal alongside the real file and leaves the
real one untouched.
"""

import os
import sys
import json
import logging
import argparse
from pathlib import Path
from datetime import date, timedelta
from collections import defaultdict

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError
import traumasoft_reports as R

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("exclusions")

TRIP_CHUNK_DAYS = 31


def fetch_last_seen(api, days):
    """Last date each vehicle ran, over `days` of history, in 31-day chunks."""
    end = date.today() - timedelta(days=1)
    start = end - timedelta(days=days)
    by_day = defaultdict(list)

    cursor = start
    while cursor <= end:
        span = min(TRIP_CHUNK_DAYS, (end - cursor).days + 1)
        log.info("Fetching trips %s +%s days ...", cursor, span)
        try:
            legs = api.get_trips(cursor, range_days=span)
        except TraumasoftAPIError as exc:
            log.error("  failed: %s", exc)
            legs = []
        for leg in legs:
            pickup = R.parse_ts(leg.get("pickup_time"))
            if pickup:
                by_day[pickup.date()].append(leg)
        cursor += timedelta(days=span)

    return R.vehicle_last_seen(by_day), start, end


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=int, default=90, help="history to inspect (default 90)")
    parser.add_argument("--apply", action="store_true",
                        help="write state/vehicle_exclusions.json directly")
    args = parser.parse_args()

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("%s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Credentials rejected under every signing scheme.")
        return 1

    vehicles = api.list_vehicles()
    log.info("Fetched %s vehicles", len(vehicles))

    last_seen, start, end = fetch_last_seen(api, args.days)
    log.info("Trip activity covers %s .. %s (%s vehicles ran at least once)",
             start, end, len(last_seen))

    # Only out-of-service records are candidates. Retired and pending-delivery
    # are already dropped by status, and an in-service truck is by definition live.
    candidates = [
        v for v in vehicles
        if v.get("vehicle_status") in R.OUT_OF_SERVICE_STATUSES
    ]
    log.info("%s vehicles are currently out of service", len(candidates))

    by_name, by_dormancy, keep = [], [], []
    patterns = R.FLEET_EXCLUDED_NAME_PATTERNS
    for vehicle in candidates:
        name = (vehicle.get("name") or "").strip()
        lowered = name.lower()
        matched = next((p for p in patterns if p and p in lowered), None)
        ran_on = last_seen.get(str(vehicle.get("id")))

        entry = {
            "id": vehicle.get("id"),
            "name": name,
            "status": vehicle.get("vehicle_status"),
            "last_ran": ran_on.isoformat() if ran_on else None,
        }
        if matched:
            entry["reason"] = f"name matches '{matched}'"
            by_name.append(entry)
        elif ran_on is None:
            entry["reason"] = f"no trip activity in {args.days} days"
            by_dormancy.append(entry)
        else:
            entry["reason"] = "ran recently - keep"
            keep.append(entry)

    print()
    print("=" * 72)
    print(f"PROPOSED EXCLUSIONS  ({len(by_name) + len(by_dormancy)} of {len(candidates)} out-of-service)")
    print("=" * 72)

    if by_name:
        print(f"\nMatched a name pattern ({len(by_name)}) - safe to exclude:")
        for e in sorted(by_name, key=lambda e: e["name"].lower()):
            print(f"  {e['name']:<32} {e['reason']}")

    if by_dormancy:
        print(f"\nNo activity in {args.days} days ({len(by_dormancy)}) - REVIEW THESE:")
        print("  A truck genuinely down this long belongs in the report. Remove")
        print("  any of these from the file before keeping it.")
        for e in sorted(by_dormancy, key=lambda e: e["name"].lower()):
            print(f"  {e['name']:<32} id={e['id']}")

    if keep:
        print(f"\nRan within the window ({len(keep)}) - kept in the report:")
        for e in sorted(keep, key=lambda e: e["name"].lower()):
            print(f"  {e['name']:<32} last ran {e['last_ran']}")

    proposed = by_name + by_dormancy
    payload = {
        "note": (
            "Vehicles excluded from the fleet reports. exclude_ids and exclude_names "
            "are exact matches; exclude_name_patterns are case-insensitive substrings. "
            "Generated by suggest_vehicle_exclusions.py - review before use."
        ),
        "generated": date.today().isoformat(),
        "activity_window": f"{start} .. {end}",
        "exclude_ids": sorted(str(e["id"]) for e in proposed),
        "exclude_names": sorted(e["name"] for e in proposed if e["name"]),
        "_detail": sorted(proposed, key=lambda e: (e["name"] or "").lower()),
    }

    target = Path(R.VEHICLE_EXCLUSIONS_FILE)
    out_path = target if args.apply else target.with_suffix(".proposed.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    print()
    if args.apply:
        print(f"Written to {out_path} and now in effect.")
    else:
        print(f"Proposal written to {out_path}")
        print(f"Review it, then rename to {target.name} to put it in effect:")
        print(f"    Move-Item -Force '{out_path}' '{target}'")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
