"""
Rebuild OTP for a past window from the API, with the arrival stamp corrected.

Trips backfill roughly 90 days, so on-time performance -- unlike UHU and
staffing -- can be recomputed for days that have already gone by. That makes
this possible where a UHU rebuild is not: the numbers a run publishes depend
only on data the API still holds.

Why you would want it: until 2026-09-16 the arrival stamp defaulted to a chain
led by `at_scene: At Patient Bedside`, which is recorded on about 60% of legs
and lands a median of 7.8 minutes after `at_scene`. Legs carrying it were
scored at the patient and the rest at the scene -- one column, two definitions
-- and the result understated on-time performance by roughly fourteen points
against the pre-changeover series. Rebuilding with one stamp gives a clean
history on the definition now in use.

    python rebuild_otp_history.py 2026-08-01
    python rebuild_otp_history.py 2026-08-01 2026-09-15
    python rebuild_otp_history.py 2026-08-01 --out Reports

Defaults: end date is yesterday, output directory is Reports/.

WHAT THIS WRITES AND WHAT IT DOES NOT

It writes one standalone workbook. It deliberately does **not** touch the
append workbooks the daily runner maintains: those are that job's record of
what it published on the day, and a second writer backfilling them would mean
the file's contents depended on which process ran last. If the rebuilt series
is the one you want to keep, it belongs beside the appends, not inside them.

READ THE SUMMARY SHEET. Cost centers are attributed through the shift-name map
accumulated in state/, which reflects what shifts look like *now*. A profile
that moved cost centers since the window will carry its current one across the
whole rebuild. That is a real limitation of backfilling a value the API only
answers for the present, and it is stated on the sheet rather than buried here.
"""

import os
import sys
import logging
from pathlib import Path
from collections import defaultdict
from datetime import datetime, timedelta, date

import pandas as pd

import report_output as OUT
import traumasoft_reports as R
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("otp-rebuild")

# GetTrips takes an inclusive range and caps it at 31 days.
CHUNK_DAYS = 31


def parse_args(argv):
    args = {"start": None, "end": None, "out": "Reports"}
    rest = []
    i = 0
    while i < len(argv):
        token = argv[i]
        if token == "--out":
            i += 1
            args["out"] = argv[i]
        elif token.startswith("--out="):
            args["out"] = token.split("=", 1)[1]
        elif token.startswith("--"):
            raise SystemExit(f"Unknown option: {token}")
        else:
            rest.append(datetime.strptime(token, "%Y-%m-%d").date())
        i += 1

    if not rest:
        raise SystemExit(
            "A start date is required, e.g. python rebuild_otp_history.py 2026-08-01"
        )
    args["start"] = rest[0]
    args["end"] = rest[1] if len(rest) > 1 else date.today() - timedelta(days=1)
    if args["end"] < args["start"]:
        raise SystemExit(f"End date {args['end']} is before start date {args['start']}.")
    return args


def fetch_window(api, start, end):
    """Every leg in the window, in 31-day calls."""
    legs = []
    cursor = start
    while cursor <= end:
        span = min(CHUNK_DAYS, (end - cursor).days + 1)
        log.info("Fetching trips for %s +%s day(s) ...", cursor, span)
        try:
            legs.extend(api.get_trips(cursor, range_days=span))
        except TraumasoftAPIError as exc:
            # One failed chunk is a hole in the middle of the series, not a
            # reason to lose the rest of it. The Summary sheet counts the days
            # that came back empty, so the hole is visible rather than implied.
            log.error("  chunk starting %s failed: %s", cursor, exc)
        cursor += timedelta(days=span)
    return legs


def bucket_by_day(legs):
    """Group legs by scheduled pickup date -- the day the report calls theirs."""
    by_day = defaultdict(list)
    undated = 0
    for leg in legs:
        pickup = R.scheduled_pickup_time(leg)
        if pickup:
            by_day[pickup.date()].append(leg)
        else:
            undated += 1
    return by_day, undated


def company_totals(scored):
    """
    The whole-company line for one day.

    _otp_aggregate groups, and pandas refuses an empty key list, so the
    ungrouped total is counted directly rather than by grouping on a constant.
    Early counts as on time here exactly as it does there.
    """
    if scored.empty:
        return {"scored": 0, "early": 0, "on_time": 0, "late": 0, "pct": None}
    keep = scored[~scored["cost_center"].isin(R.OTP_EXCLUDED_COST_CENTERS)]
    if keep.empty:
        return {"scored": 0, "early": 0, "on_time": 0, "late": 0, "pct": None}
    early = int((keep["status"] == "Early").sum())
    on_time = int((keep["status"] == "On Time").sum())
    late = int((keep["status"] == "Late").sum())
    total = len(keep)
    return {
        "scored": total, "early": early, "on_time": on_time, "late": late,
        "pct": round(100.0 * (early + on_time) / total, 2),
    }


def build(by_day, start, end, cost_center_map):
    """Score every day in the window and roll the results up three ways."""
    daily_rows = []
    per_day_cc = []
    all_scored = []

    cursor = start
    while cursor <= end:
        legs = by_day.get(cursor, [])
        scored = R.scored_legs(legs, cost_center_map)
        if not scored.empty:
            scored = scored.assign(metrics_date=cursor)
            all_scored.append(scored)

            by_cc = R._otp_aggregate(scored, ["cost_center"])
            by_cc.insert(0, "metrics_date", cursor)
            per_day_cc.append(by_cc)

        # The day's own line goes in whether or not anything scored: a blank
        # row is how a gap in the backfill announces itself.
        totals = company_totals(scored)
        daily_rows.append({
            "metrics_date": cursor,
            "legs_returned": len(legs),
            "legs_scored": totals["scored"],
            "early_runs": totals["early"],
            "on_time_runs": totals["on_time"],
            "late_runs": totals["late"],
            "on_time_percentage": totals["pct"],
        })
        cursor += timedelta(days=1)

    scored_all = pd.concat(all_scored, ignore_index=True) if all_scored else pd.DataFrame()
    return {
        "daily": pd.DataFrame(daily_rows),
        "daily_by_cost_center": (
            pd.concat(per_day_cc, ignore_index=True) if per_day_cc else pd.DataFrame()
        ),
        "by_cost_center": R._otp_aggregate(scored_all, ["cost_center"]),
        "by_call_type": R._otp_aggregate(scored_all, ["cost_center", "call_type"]),
        "scored": scored_all,
    }


def summary_sheet(args, frames, legs, undated, empty_days, cost_center_map):
    """What the numbers rest on, in the first sheet rather than a footnote."""
    daily = frames["daily"]
    scored = int(daily["legs_scored"].sum())
    early = int(daily["early_runs"].sum())
    on_time = int(daily["on_time_runs"].sum())
    late = int(daily["late_runs"].sum())
    window_pct = round(100.0 * (early + on_time) / scored, 2) if scored else None

    def row(item, value, note=""):
        return {"item": item, "value": value, "note": note}

    rows = [
        row("Window", f"{args['start']} to {args['end']}",
            f"{(args['end'] - args['start']).days + 1} day(s)"),
        row("On-time percentage", window_pct,
            "Summed across the window: (early + on time) / scored. NOT the "
            "average of the daily percentages, which would weight a quiet "
            "Sunday like a full Monday."),
        row("Legs scored", scored, "Legs carrying both a scheduled pickup and an arrival stamp."),
        row("  early", early, "Counted as on time, matching the daily report."),
        row("  on time", on_time, f"Within +/-{R.OTP_ON_TIME_WINDOW_MINUTES} minutes."),
        row("  late", late, ""),
        row("Legs returned", len(legs), "Everything the API gave back for the window."),
        row("Legs with no scheduled pickup", undated,
            "Not scorable and not counted. A leg never promised a time cannot be late."),
        row("Days with no legs returned", len(empty_days),
            ", ".join(str(d) for d in empty_days[:10]) +
            (" ..." if len(empty_days) > 10 else "") if empty_days
            else "None -- the backfill covered every day in the window."),
        row("Arrival stamp", ", ".join(R.ARRIVAL_TIMESTAMP_KEYS),
            "TS_ARRIVAL_TIMESTAMP_KEYS. 'at_scene' is the stamp that reproduces "
            "the pre-changeover series; a chain would score some legs at the "
            "scene and others at the patient."),
        row("Scheduled pickup", ", ".join(R.PICKUP_TIME_KEYS), "TS_PICKUP_TIME_KEYS."),
        row("Excluded cost centers", ", ".join(R.OTP_EXCLUDED_COST_CENTERS) or "none",
            "As in the daily report."),
        row("Cost-center attribution", f"{len(cost_center_map.counts)} profile(s) mapped",
            "READ THIS. Cost center is not on a trip. It is resolved through "
            "shift_name -> crew -> employee, which the API answers only for the "
            "current shift window, so this rebuild uses the map accumulated in "
            "state/ as it stands TODAY. A profile that changed cost centers "
            "during the window carries its current one across the whole series."),
        row("Append workbooks", "not written",
            "This is a standalone rebuild. The daily runner's append workbooks "
            "are its record of what it published on the day and are left alone."),
    ]
    ambiguous = cost_center_map.ambiguous() if hasattr(cost_center_map, "ambiguous") else {}
    if ambiguous:
        rows.append(row("Contested profiles", len(ambiguous),
                        "Profiles seen under more than one cost center; the "
                        "dominant one won. " + ", ".join(list(ambiguous)[:5])))
    return pd.DataFrame(rows)


def main():
    args = parse_args(sys.argv[1:])

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("Configuration error: %s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Could not authenticate. Check TS_API_KEY / TS_API_SECRET.")
        return 1

    span = (args["end"] - args["start"]).days + 1
    log.info("Rebuilding OTP for %s to %s (%s day(s))", args["start"], args["end"], span)
    log.info("Arrival stamp: %s", ", ".join(R.ARRIVAL_TIMESTAMP_KEYS))
    if span > 90:
        log.warning("Trips backfill about 90 days; the earliest part of this "
                    "window may come back empty.")

    legs = fetch_window(api, args["start"], args["end"])
    if not legs:
        log.error("No legs returned for the whole window. Is the start date "
                  "beyond what the API still holds?")
        return 1
    log.info("Fetched %s legs", len(legs))

    by_day, undated = bucket_by_day(legs)
    cost_center_map = R.CostCenterMap()
    frames = build(by_day, args["start"], args["end"], cost_center_map)

    daily = frames["daily"]
    empty_days = [r["metrics_date"] for _, r in daily.iterrows() if r["legs_returned"] == 0]

    scored = int(daily["legs_scored"].sum())
    on_time_total = int(daily["early_runs"].sum() + daily["on_time_runs"].sum())
    window_pct = round(100.0 * on_time_total / scored, 2) if scored else None

    log.info("")
    log.info("--- %s to %s ---", args["start"], args["end"])
    log.info("  legs scored : %s", scored)
    log.info("  on time     : %s (%.2f%%)", on_time_total, window_pct or 0.0)
    if undated:
        log.info("  unscorable  : %s leg(s) with no scheduled pickup", undated)
    if empty_days:
        log.warning("  %s day(s) returned no legs: %s", len(empty_days),
                    ", ".join(str(d) for d in empty_days[:10]))

    sheets = {
        "Summary": summary_sheet(args, frames, legs, undated, empty_days, cost_center_map),
        "Daily": daily,
        "By Cost Center": frames["by_cost_center"],
        "By Call Type": frames["by_call_type"],
        "Daily by Cost Center": frames["daily_by_cost_center"],
    }

    out_dir = Path(args["out"])
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"CompanyWide_OTP_Rebuild_{args['start']}_to_{args['end']}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            OUT.write_df_sheet_with_table(
                writer, df if df is not None else pd.DataFrame(), name, name.replace(" ", "_")
            )

    log.info("")
    log.info("Wrote %s", out_path)
    log.info("Read the Summary sheet: cost centers come from today's map, and "
             "the append workbooks were deliberately not touched.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
