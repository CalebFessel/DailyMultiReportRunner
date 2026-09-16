"""
Why can't OTP score a third of the legs, and would a different field fix it?

Roughly one completed leg in three carries no scheduled `pickup_time`, so OTP
cannot judge it late and drops it. Two explanations produce that same number
and they call for opposite responses:

  (a) The report reads the wrong field. TripLegSummary also carries
      `appt_time`, `requested_pickup_time` and `eta_time`. If one of those is
      populated on the legs `pickup_time` misses, the fix is configuration.

  (b) Those legs were never scheduled. An emergency call has no promised
      pickup, so there is nothing to be late against. Then the missing third
      is correct and any field that "recovers" it is inventing a deadline.

This probe separates them. For a single day it reports, per candidate field,
how many legs carry it, how many legs ONLY it can score, and what kind of call
those legs are -- then does the same for arrival stamps, since a leg needs both
halves to be scorable.

Read-only: GETs only, writes nothing but its own output file.

Usage:
    python probe_otp_coverage.py [YYYY-MM-DD] [--days N] [--out DIR]

Defaults to yesterday. --days widens the sample (capped at the API's 31-day
GetTrips range); a single busy weekday is usually enough to tell (a) from (b).
"""

import os
import sys
import json
import logging
from pathlib import Path
from collections import Counter
from datetime import datetime, timedelta, date

import traumasoft_reports as R
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("otp-coverage")

# Every field on TripLegSummary that could plausibly carry a scheduled pickup.
# `pickup_time` is what the report uses and what the old SQL compared against;
# the rest are here to be measured, not adopted.
PICKUP_CANDIDATES = [
    ("pickup_time", "CAD grid scheduled pickup -- what the report scores against"),
    ("appt_time", "appointment the trip must make; a different promise"),
    ("requested_pickup_time", "what the caller asked for, before dispatch scheduled it"),
    ("eta_time", "a projection, not a commitment"),
]

# Arrival stamps live in the timestamps map, not as top-level fields. Anything
# observed on the day is measured; these are the ones worth naming.
ARRIVAL_CANDIDATES = ["at_scene", "at_destination", "enroute"]


def parse_args(argv):
    args = {"day": None, "days": 1, "out": "api_probe"}
    rest = []
    i = 0
    while i < len(argv):
        token = argv[i]
        if token == "--days":
            i += 1
            args["days"] = max(1, min(31, int(argv[i])))
        elif token == "--out":
            i += 1
            args["out"] = argv[i]
        elif token.startswith("--"):
            raise SystemExit(f"Unknown option: {token}")
        else:
            rest.append(token)
        i += 1
    if rest:
        args["day"] = datetime.strptime(rest[0], "%Y-%m-%d").date()
    else:
        args["day"] = date.today() - timedelta(days=1)
    return args


def completed(leg):
    """Legs the OTP report would consider at all: cancellations are dropped."""
    status = (leg.get("trip_status") or "").strip().lower()
    return bool(status) and "cancel" not in status and "disregard" not in status


def field_coverage(legs, field):
    """How many legs carry `field` as a parseable timestamp."""
    return sum(1 for leg in legs if R.parse_ts(leg.get(field)))


def only_this_field(legs, field, baseline="pickup_time"):
    """Legs `field` can date that `baseline` cannot -- the recovery it offers."""
    return [
        leg for leg in legs
        if R.parse_ts(leg.get(field)) and not R.parse_ts(leg.get(baseline))
    ]


def call_type_breakdown(legs, limit=12):
    counts = Counter((leg.get("call_type") or "Unknown Call Type") for leg in legs)
    return counts.most_common(limit)


def arrival_coverage(legs):
    """Count every arrival-stamp name observed, not just the expected ones."""
    seen = Counter()
    for leg in legs:
        for name, value in R.timestamp_map(leg).items():
            if value:
                seen[name] += 1
    return seen


def main():
    args = parse_args(sys.argv[1:])
    out_dir = Path(args["out"])
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("Configuration error: %s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Could not authenticate. Check TS_API_KEY / TS_API_SECRET.")
        return 1

    log.info("Fetching trips for %s (+%s day(s)) ...", args["day"], args["days"])
    try:
        legs = api.get_trips(args["day"], range_days=args["days"])
    except TraumasoftAPIError as exc:
        log.error("GetTrips failed: %s", exc)
        return 1

    done = [leg for leg in legs if completed(leg)]
    log.info("%s legs returned, %s completed (cancellations dropped)", len(legs), len(done))
    if not done:
        log.error("No completed legs on this date -- pick a busier weekday.")
        return 1

    findings = {
        "date": str(args["day"]),
        "range_days": args["days"],
        "legs_returned": len(legs),
        "legs_completed": len(done),
        "pickup_fields": [],
        "arrival_stamps": {},
        "scorable": {},
    }

    log.info("")
    log.info("--- Scheduled-pickup candidates (of %s completed legs) ---", len(done))
    for field, note in PICKUP_CANDIDATES:
        carried = field_coverage(done, field)
        recovered = only_this_field(done, field) if field != "pickup_time" else []
        entry = {
            "field": field,
            "note": note,
            "legs_carrying": carried,
            "pct_carrying": round(100.0 * carried / len(done), 1),
            "legs_pickup_time_misses": len(recovered),
            "recovered_call_types": call_type_breakdown(recovered),
        }
        findings["pickup_fields"].append(entry)
        log.info("  %-22s %4s legs (%5.1f%%)  %s",
                 field, carried, entry["pct_carrying"],
                 f"recovers {len(recovered)} pickup_time misses" if field != "pickup_time" else "(baseline)")

    unscheduled = [leg for leg in done if not R.parse_ts(leg.get("pickup_time"))]
    findings["unscheduled_call_types"] = call_type_breakdown(unscheduled)
    log.info("")
    log.info("--- The %s completed legs with no pickup_time, by call type ---", len(unscheduled))
    for name, count in findings["unscheduled_call_types"]:
        log.info("  %5s  %s", count, name)

    stamps = arrival_coverage(done)
    findings["arrival_stamps"] = dict(stamps.most_common())
    log.info("")
    log.info("--- Timestamp names observed (of %s completed legs) ---", len(done))
    for name, count in stamps.most_common(20):
        marker = "  <-- used by OTP" if name in R.ARRIVAL_TIMESTAMP_KEYS else ""
        log.info("  %5s  %s%s", count, name, marker)
    for name in R.ARRIVAL_TIMESTAMP_KEYS:
        if name not in stamps:
            log.warning("  OTP is configured to read '%s', which no leg carries today.", name)

    # The number that actually matters: how many legs the report can judge.
    scorable = sum(
        1 for leg in done
        if R.scheduled_pickup_time(leg) and R.arrival_time(leg)
    )
    findings["scorable"] = {
        "pickup_keys": R.PICKUP_TIME_KEYS,
        "arrival_keys": R.ARRIVAL_TIMESTAMP_KEYS,
        "legs_scorable": scorable,
        "pct_scorable": round(100.0 * scorable / len(done), 1),
    }
    log.info("")
    log.info("--- Under the current configuration ---")
    log.info("  pickup keys : %s", ", ".join(R.PICKUP_TIME_KEYS))
    log.info("  arrival keys: %s", ", ".join(R.ARRIVAL_TIMESTAMP_KEYS))
    log.info("  scorable    : %s of %s completed legs (%.1f%%)",
             scorable, len(done), findings["scorable"]["pct_scorable"])

    path = out_dir / f"otp_coverage_{args['day']}.json"
    path.write_text(json.dumps(findings, indent=2), encoding="utf-8")
    log.info("")
    log.info("Wrote %s", path)

    log.info("")
    log.info("How to read this: if a candidate field recovers a large number of")
    log.info("pickup_time misses AND those legs are scheduled/transfer work, the")
    log.info("report is reading the wrong field. If the misses are emergency call")
    log.info("types, they were never scheduled and dropping them is correct.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
