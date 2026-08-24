"""
Month-to-date reports for one region.

    python monthly_region_report.py Indiana                # 1st of this month -> yesterday
    python monthly_region_report.py Indiana 2026-08-23     # 1st of that month -> that date
    python monthly_region_report.py Indiana --days 30      # trailing 30 days instead
    python monthly_region_report.py --all-regions          # one workbook per region
    python monthly_region_report.py --list-cost-centers    # what the regions file could name

Which cost centers make up a region is read from state/regions.json. See
state/regions.example.json; nothing in the API answers that question.

The window is a calendar month to date -- it resets on the 1st -- because that
is how the count was asked for. `--days N` gives a trailing window instead.

WHAT THIS CAN AND CANNOT COVER

The bundle is honestly asymmetric, and the Summary sheet says so on every run:

  * OTP and run volume are rebuilt from the API across the whole window. Trips
    backfill 90 days, so these are complete from the first run.

  * UHU and staffing are read back out of the daily append workbooks, because
    they cannot be rebuilt. /Schedule/Shifts returns today-1..today+2 and
    ignores every filter, so the hours a unit was crewed on the 3rd exist only
    if the daily run wrote them down on the 3rd. A month's worth accrues a day
    at a time from changeover.

UHU is summed, not averaged. A month's ratio is sum(utilized) / sum(worked)
across the window -- averaging the daily percentages would weight a quiet
Sunday the same as a full Monday. The append sheets carry the raw hours
columns, which is what makes the exact figure available.
"""

import os
import sys
import logging
from pathlib import Path
from datetime import datetime, timedelta, date

import pandas as pd

import report_output as OUT
import traumasoft_reports as R
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

JOB_NAME = "Monthly Regional Report"

# Append workbook -> sheet, for the halves that cannot be rebuilt from the API.
UHU_COST_CENTER_APPEND = ("Daily_UHU_By_Cost_Center_APPEND.xlsx", "UHU by Cost Center")
UHU_PROFILE_APPEND = ("Daily_UHU_By_Shift_Profile_APPEND.xlsx", "UHU by Shift Profile")
STAFFING_APPEND = ("Staffing_Report_APPEND.xlsx", "Active Now")


# =============================
# WINDOW
# =============================
def month_to_date(end_date):
    """The calendar month containing end_date, up to and including it."""
    return end_date.replace(day=1), end_date


def trailing(end_date, days):
    """A trailing window of `days` ending on end_date, inclusive."""
    if days < 1:
        raise ValueError("--days must be at least 1")
    return end_date - timedelta(days=days - 1), end_date


def window_length(start, end):
    return (end - start).days + 1


# =============================
# CLI
# =============================
def parse_args(argv):
    """
    Region, end date and flags, in whatever order they were typed.

    A bare word is the region and a bare date is the end date, so
    `Indiana 2026-08-23` and `2026-08-23 Indiana` both work -- the alternative
    is a positional order nobody remembers between monthly runs.
    """
    args = {
        "region": None, "end_date": None, "days": None,
        "all_regions": False, "list_cost_centers": False, "no_email": True,
    }
    rest = list(argv[1:])
    while rest:
        token = rest.pop(0)
        lowered = token.lower()
        if lowered == "--all-regions":
            args["all_regions"] = True
        elif lowered == "--list-cost-centers":
            args["list_cost_centers"] = True
        elif lowered == "--days":
            if not rest:
                raise ValueError("--days needs a number, e.g. --days 30")
            args["days"] = int(rest.pop(0))
        elif lowered.startswith("--days="):
            args["days"] = int(token.split("=", 1)[1])
        elif lowered.startswith("-"):
            raise ValueError(f"Unrecognised option: {token}")
        else:
            try:
                args["end_date"] = datetime.strptime(token.strip(), "%Y-%m-%d").date()
                continue
            except ValueError:
                pass
            # Regions can be more than one word ("West Virginia"), so keep
            # joining bare words rather than taking only the first.
            args["region"] = f"{args['region']} {token}".strip() if args["region"] else token
    return args


# =============================
# APPEND-BACKED HALVES
# =============================
def _in_window(df, start, end, column="snapshot_date"):
    """
    Rows whose snapshot date falls in the window, with that date normalized.

    Excel hands the column back as a string on one run and a Timestamp on the
    next depending on how it was written, so it is parsed rather than compared
    as it arrives.
    """
    if df is None or getattr(df, "empty", True) or column not in df.columns:
        return pd.DataFrame()
    out = df.copy()
    parsed = pd.to_datetime(out[column], errors="coerce")
    out["_snapshot"] = parsed.dt.date
    keep = out["_snapshot"].map(
        lambda d: d is not None and d == d and start <= d <= end
    )
    return out.loc[keep].copy()


def days_present(df):
    """Which dates in the window actually have rows, so gaps are visible."""
    if df is None or df.empty or "_snapshot" not in df.columns:
        return []
    return sorted({d for d in df["_snapshot"] if d is not None and d == d})


def rollup_uhu(df, group_column):
    """
    A window's UHU for one grouping, summed rather than averaged.

    sum(utilized) / sum(worked) is the month's ratio. The mean of the daily
    ratios is a different and wrong number -- it gives a Sunday with two trucks
    out the same weight as a Monday with twenty.
    """
    columns = [
        group_column, "days_counted", "scheduled_hours", "worked_hours",
        "utilized_hours", "total_runs", "hours_per_run", "uhu_ratio",
    ]
    if df is None or df.empty or group_column not in df.columns:
        return pd.DataFrame(columns=columns)

    numeric = ("scheduled_hours", "worked_hours", "utilized_hours", "total_runs")
    frame = df.copy()
    for column in numeric:
        if column not in frame.columns:
            frame[column] = 0
        frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0)

    grouped = (
        frame.groupby(group_column, dropna=False)
        .agg(
            days_counted=("_snapshot", "nunique"),
            scheduled_hours=("scheduled_hours", "sum"),
            worked_hours=("worked_hours", "sum"),
            utilized_hours=("utilized_hours", "sum"),
            total_runs=("total_runs", "sum"),
        )
        .reset_index()
    )
    grouped["hours_per_run"] = grouped.apply(
        lambda r: round(r["utilized_hours"] / r["total_runs"], 3) if r["total_runs"] else 0,
        axis=1,
    )
    denominator = "worked_hours" if R.UHU_DENOMINATOR == "worked" else "scheduled_hours"
    grouped["uhu_ratio"] = grouped.apply(
        lambda r: round(r["utilized_hours"] / r[denominator], 3) if r[denominator] else 0,
        axis=1,
    )
    for column in ("scheduled_hours", "worked_hours", "utilized_hours"):
        grouped[column] = grouped[column].round(2)
    grouped["total_runs"] = grouped["total_runs"].astype(int)
    return grouped.sort_values("uhu_ratio", ascending=False)[columns]


def rollup_staffing(df):
    """
    How often each unit was short over the window.

    Counted in observations, not hours: the Active Now sheet is one look at the
    board per run, so a truck short at 07:45 and crewed by 09:00 reads as one
    short day. That makes this a measure of how often a shortfall was caught,
    which is the honest reading of a daily snapshot.
    """
    columns = [
        "cost_center", "shift_profile", "days_observed", "days_short",
        "worst_shortfall", "shortfall_crew_days", "crew_needed",
    ]
    if df is None or df.empty or "shift_profile" not in df.columns:
        return pd.DataFrame(columns=columns)

    frame = df.copy()
    if "cost_center" not in frame.columns:
        frame["cost_center"] = R.UNASSIGNED_COST_CENTER
    for column in ("crew_count", "crew_needed"):
        if column not in frame.columns:
            frame[column] = 0
        frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0)
    frame["_short_by"] = (frame["crew_needed"] - frame["crew_count"]).clip(lower=0)
    frame["_is_short"] = frame["_short_by"] > 0

    grouped = (
        frame.groupby(["cost_center", "shift_profile"], dropna=False)
        .agg(
            days_observed=("_snapshot", "nunique"),
            days_short=("_is_short", "sum"),
            worst_shortfall=("_short_by", "max"),
            shortfall_crew_days=("_short_by", "sum"),
            crew_needed=("crew_needed", "max"),
        )
        .reset_index()
    )
    for column in ("days_short", "worst_shortfall", "shortfall_crew_days", "crew_needed"):
        grouped[column] = grouped[column].astype(int)
    return grouped.sort_values(
        ["shortfall_crew_days", "days_short", "cost_center", "shift_profile"],
        ascending=[False, False, True, True],
    )[columns]


# =============================
# API-BACKED HALVES
# =============================
def fetch_window_legs(api, start, end):
    """
    Every leg in the window, in one call.

    A calendar month never exceeds the API's 31-day range cap, so this stays a
    single request; a longer --days window is split rather than refused.
    """
    legs = []
    cursor = start
    while cursor <= end:
        span = min(31, (end - cursor).days + 1)
        logging.info("Fetching trips for %s +%s day(s) ...", cursor, span)
        legs.extend(api.get_trips(cursor, range_days=span))
        cursor += timedelta(days=span)
    logging.info("Fetched %s legs across %s day(s)", len(legs), window_length(start, end))
    return legs


# =============================
# BUILD
# =============================
def build_region_bundle(region, regions, start, end, legs, cost_center_map, append_dir,
                        vehicles=None):
    """Every sheet for one region's window, plus the summary that qualifies them."""
    # Cut at the leg, not at the finished frame: a vehicle is filed under the
    # cost center it served most, so filtering rows would hand a truck that
    # split the month between two regions -- and all of its runs -- to one of
    # them. See legs_in_region.
    api_side, region_legs, unattributed = R.build_region_leg_reports(
        legs, vehicles, cost_center_map, regions, region
    )

    uhu_cc_raw = _in_window(
        OUT.read_append_sheet(os.path.join(append_dir, UHU_COST_CENTER_APPEND[0]),
                              UHU_COST_CENTER_APPEND[1]), start, end)
    uhu_sp_raw = _in_window(
        OUT.read_append_sheet(os.path.join(append_dir, UHU_PROFILE_APPEND[0]),
                              UHU_PROFILE_APPEND[1]), start, end)
    staffing_raw = _in_window(
        OUT.read_append_sheet(os.path.join(append_dir, STAFFING_APPEND[0]),
                              STAFFING_APPEND[1]), start, end)

    uhu_cc_raw = R.filter_frame_to_region(uhu_cc_raw, regions, region)
    uhu_sp_raw = R.filter_frame_to_region(uhu_sp_raw, regions, region)
    staffing_raw = R.filter_frame_to_region(staffing_raw, regions, region)

    sheets = {
        "OTP by Cost Center": api_side["otp_by_cost_center"],
        "OTP by Call Type": api_side["otp_by_call_type"],
        "Runs by Cost Center": api_side["runs_by_cost_center"],
        "Runs by Vehicle": api_side["runs_by_vehicle"],
        "UHU by Cost Center": rollup_uhu(uhu_cc_raw, "cost_center_name"),
        "UHU by Shift Profile": rollup_uhu(uhu_sp_raw, "shift_profile_name"),
        "Staffing Shortfalls": rollup_staffing(staffing_raw),
    }
    summary = build_summary(
        region, regions, start, end, region_legs,
        uhu_days=days_present(uhu_cc_raw),
        staffing_days=days_present(staffing_raw),
        sheets=sheets,
        unattributed=unattributed,
    )
    return {"Summary": summary, **sheets}


def build_summary(region, regions, start, end, legs, uhu_days, staffing_days, sheets,
                  unattributed=0):
    """
    What this bundle covers, and what it does not yet.

    First sheet in the book on purpose. A month-to-date UHU built from five
    days of accumulated hours is a useful number and a misleading one, and the
    difference is entirely whether the reader was told.
    """
    asked = window_length(start, end)
    rows = [
        ("Region", region),
        ("Window", f"{start} to {end}"),
        ("Days in window", asked),
        ("Cost centers configured", ", ".join(sorted(regions.configured.get(region, []))) or "(patterns only)"),
        ("", ""),
        ("OTP and run volume", f"complete -- rebuilt from the API across all {asked} day(s)"),
        ("Legs attributed to this region", len(legs)),
    ]
    if unattributed:
        rows.append((
            "Legs belonging to no cost center",
            f"{unattributed} company-wide, in no region's bundle. Mostly calls "
            "cancelled before they reached a unit, which have no unit to attribute "
            "them to. This is why the regional totals do not sum to the "
            "company-wide sheet.",
        ))
    rows += [
        ("", ""),
        ("UHU days accumulated", f"{len(uhu_days)} of {asked}"),
        ("Staffing days accumulated", f"{len(staffing_days)} of {asked}"),
    ]
    if len(uhu_days) < asked or len(staffing_days) < asked:
        rows.append((
            "Why they differ",
            "The shifts endpoint returns today-1..today+2 and ignores every filter, "
            "so crew hours for a past day exist only if the daily run recorded them "
            "that day. These sheets cover the days present, not the whole window, and "
            "fill in as the daily run accrues.",
        ))
    missing = sorted(set(_dates_between(start, end)) - set(uhu_days))
    if missing and len(missing) <= 15:
        rows.append(("UHU days missing", ", ".join(str(d) for d in missing)))
    elif missing:
        rows.append(("UHU days missing", f"{len(missing)}, from {missing[0]} to {missing[-1]}"))

    rows.append(("", ""))
    rows.append((
        "How UHU is totalled",
        "sum(utilized_hours) / sum(worked_hours) across the days present -- not the "
        "average of the daily ratios, which would weight every day equally.",
    ))
    rows.append((
        "Fleet sheets",
        "Not included. The API puts no cost center on a vehicle, so the in-service "
        "and out-of-service rosters cannot be scoped to a region. Runs by Vehicle is "
        "here because a leg is attributable through its shift profile.",
    ))
    for name, df in sheets.items():
        rows.append((f"Rows: {name}", 0 if df is None else len(df)))
    return pd.DataFrame(rows, columns=["item", "value"])


def _dates_between(start, end):
    day = start
    while day <= end:
        yield day
        day += timedelta(days=1)


# =============================
# MAIN
# =============================
def list_cost_centers(regions, append_dir):
    """
    Every cost center this deployment has seen, and which region claims it.

    The regions file has to name cost centers exactly, and nothing else prints
    the list -- so without this the first edit is a guess.
    """
    seen = set()
    for centers in R.CostCenterMap().counts.values():
        seen.update(centers)
    for filename, sheet in (UHU_COST_CENTER_APPEND, STAFFING_APPEND):
        df = OUT.read_append_sheet(os.path.join(append_dir, filename), sheet)
        column = R.cost_center_column(df)
        if column is not None and not df.empty:
            seen.update(str(v).strip() for v in df[column].dropna().unique())
    seen.discard("")
    seen.discard(R.UNASSIGNED_COST_CENTER)

    if not seen:
        print(
            "No cost centers known yet. They are learned from the daily run's shift\n"
            "window into state/shift_cost_center_map.json -- run the daily report at\n"
            "least once first."
        )
        return 1
    print(f"{len(seen)} cost center(s) known:\n")
    for name in sorted(seen):
        claimed = regions.resolve(name)
        print(f"  {name:<45} {claimed or '-- no region --'}")
    unmatched = regions.unmatched(seen)
    if unmatched:
        print(
            f"\n{len(unmatched)} cost center(s) belong to no region and will be missing\n"
            f"from every regional bundle. Add them to {regions.path}."
        )
    return 0


def main():
    try:
        args = parse_args(sys.argv)
    except ValueError as exc:
        print(f"{exc}\n\n{__doc__.strip().splitlines()[0]}\n", file=sys.stderr)
        print(__doc__.split("WHAT THIS")[0].strip(), file=sys.stderr)
        return 2

    end_date = args["end_date"] or (date.today() - timedelta(days=1))
    start_date, end_date = (
        trailing(end_date, args["days"]) if args["days"] else month_to_date(end_date)
    )

    output_dir = OUT.OUTPUT_DIR
    append_dir = OUT.APPEND_DIR
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    log_path = OUT.setup_logging(output_dir, f"regional_{end_date.isoformat()}")
    logging.info("=== %s ===", JOB_NAME)
    logging.info("Window: %s to %s (%s day(s))", start_date, end_date,
                 window_length(start_date, end_date))
    logging.info("Log: %s", log_path)

    regions = R.Regions()
    if args["list_cost_centers"]:
        return list_cost_centers(regions, append_dir)

    if not regions.names():
        logging.error(
            "No regions are defined. Copy state/regions.example.json to "
            "state/regions.json and name the cost centers in each region -- "
            "`--list-cost-centers` prints the ones this deployment has seen."
        )
        return 2

    if args["all_regions"]:
        wanted = regions.names()
    elif args["region"]:
        canonical = regions.canonical(args["region"])
        if canonical is None:
            logging.error(
                "No region named %r in %s. Defined: %s",
                args["region"], regions.path, ", ".join(regions.names()),
            )
            return 2
        wanted = [canonical]
    else:
        logging.error(
            "Name a region, or pass --all-regions. Defined: %s",
            ", ".join(regions.names()),
        )
        return 2

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        logging.error("Configuration error: %s", exc)
        return 2
    if api.detect_auth_mode() is None:
        logging.error("Credentials rejected under every signing scheme; aborting.")
        return 1

    cost_center_map = R.CostCenterMap()
    if not cost_center_map.counts:
        logging.error(
            "The cost-center map at %s is empty, so no leg can be attributed to a "
            "region and every sheet would come out blank. Run the daily report at "
            "least once first -- it is what learns the mapping.",
            cost_center_map.path,
        )
        return 2

    try:
        legs = fetch_window_legs(api, start_date, end_date)
        # Names and statuses for the vehicle rows. One call; a leg carries a
        # vehicle name too, so this only improves the sheet rather than
        # gating it.
        vehicles = api.list_vehicles()
    except TraumasoftAPIError as exc:
        logging.error("API failure: %s", exc, exc_info=True)
        return 1

    # Named once for the whole run rather than per region: an unclaimed cost
    # center is missing from every bundle, not just one.
    unmatched = regions.unmatched(
        centre for centers in cost_center_map.counts.values() for centre in centers
    )
    if unmatched:
        logging.warning(
            "%s cost center(s) belong to no region and appear in no bundle: %s. "
            "Add them to %s.",
            len(unmatched), ", ".join(unmatched), regions.path,
        )

    written = []
    failed = []
    for region in wanted:
        logging.info("--- Building: %s ---", region)
        try:
            sheets = build_region_bundle(
                region, regions, start_date, end_date, legs, cost_center_map,
                append_dir, vehicles,
            )
            slug = region.replace(" ", "_")
            label = "Trailing{}d".format(args["days"]) if args["days"] else "MTD"
            path = os.path.join(
                output_dir, f"{slug}_{label}_{start_date}_to_{end_date}.xlsx"
            )
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                for name, df in sheets.items():
                    OUT.write_df_sheet_with_table(
                        writer, df if df is not None else pd.DataFrame(), name, name
                    )
            rows = {name: (0 if df is None else len(df)) for name, df in sheets.items()}
            logging.info("--- OK: %s -> %s (%s)", region, os.path.basename(path), rows)
            written.append(path)
        except Exception as exc:  # one region must not lose the others
            logging.error("--- FAIL: %s: %s", region, exc, exc_info=True)
            failed.append((region, str(exc)))

    logging.info("")
    if written:
        logging.info("Files are in %s:", os.path.abspath(output_dir))
        for path in written:
            logging.info("  %s", os.path.basename(path))
    for region, error in failed:
        logging.error("%s failed: %s", region, error)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
