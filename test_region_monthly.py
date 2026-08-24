"""
Checks for the regional filter and the month-to-date rollups.

    python test_region_monthly.py

Plain asserts and no pytest, because the reporting machine has neither and
these have to be runnable there -- the whole point is to confirm a file copied
onto that machine by hand behaves.

Nothing here touches the network. The append workbooks are written to a
temporary directory with the real writer, then read back with the real reader,
so the round trip through Excel -- which is where the digit-string/integer
de-dupe bug lived -- is exercised rather than assumed.
"""

import os
import sys
import json
import shutil
import tempfile
from datetime import date

import pandas as pd

import report_output as OUT
import traumasoft_reports as R
import monthly_region_report as M

FAILURES = []


def check(name, condition, detail=""):
    if condition:
        print(f"  ok    {name}")
    else:
        print(f"  FAIL  {name}{(' -- ' + detail) if detail else ''}")
        FAILURES.append(name)


def regions_from(payload, tmpdir):
    path = os.path.join(tmpdir, "regions.json")
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle)
    return R.Regions(path=path)


# =============================
def test_region_matching(tmpdir):
    print("\nRegion matching")
    regions = regions_from({
        "regions": {
            "Indiana": {
                "cost_centers": ["Indianapolis", "South Bend"],
                "cost_center_patterns": ["indy"],
            },
            "Ohio": ["Columbus Admin", "Toledo Admin"],
        }
    }, tmpdir)

    check("exact name resolves", regions.resolve("Indianapolis") == "Indiana")
    check("exact match ignores case", regions.resolve("indianapolis") == "Indiana")
    check("exact match ignores surrounding space", regions.resolve("  South Bend ") == "Indiana")
    check("pattern resolves", regions.resolve("Indy Wheelchair") == "Indiana")
    check("bare list is shorthand for cost_centers", regions.resolve("Toledo Admin") == "Ohio")
    check("unknown cost center resolves to nothing", regions.resolve("Charleston") is None)
    check("blank resolves to nothing", regions.resolve("") is None)
    check("regions are listed", regions.names() == ["Indiana", "Ohio"])
    check("region name is matched case-insensitively",
          regions.canonical("indiana") == "Indiana")
    check("unknown region is not claimed", regions.canonical("Kentucky") is None)

    unmatched = regions.unmatched(
        ["Indianapolis", "Charleston", "Wheeling", R.UNASSIGNED_COST_CENTER, ""]
    )
    check("unmatched names the cost centers no region claims",
          unmatched == ["Charleston", "Wheeling"], str(unmatched))
    check("unmatched does not report the unassigned bucket",
          R.UNASSIGNED_COST_CENTER not in unmatched)


def test_missing_and_broken_files(tmpdir):
    print("\nRegions file that is missing or will not parse")
    absent = R.Regions(path=os.path.join(tmpdir, "nope.json"))
    check("a missing file defines no regions", absent.names() == [])
    check("a missing file claims nothing", absent.resolve("Indianapolis") is None)

    broken = os.path.join(tmpdir, "broken.json")
    with open(broken, "w", encoding="utf-8") as handle:
        handle.write("@'\n{ \"regions\": {} }\n'@\n")
    unparsable = R.Regions(path=broken)
    check("a here-string wrapper defines no regions", unparsable.names() == [])

    empty = os.path.join(tmpdir, "empty.json")
    open(empty, "w").close()
    check("an empty file defines no regions", R.Regions(path=empty).names() == [])


def test_frame_filtering(tmpdir):
    print("\nFiltering report frames")
    regions = regions_from(
        {"regions": {"Indiana": {"cost_centers": ["Indianapolis"]}}}, tmpdir
    )

    otp = pd.DataFrame({
        "cost_center": ["Indianapolis", "Columbus Admin"],
        "on_time_percentage": [91.0, 88.0],
    })
    runs = pd.DataFrame({
        "cost_center_name": ["Indianapolis", "Columbus Admin"],
        "total_runs": [40, 60],
    })
    fleet = pd.DataFrame({"vehicle_name": ["M-1", "M-2"], "vehicle_status": ["In Service"] * 2})

    check("the SQL's cost_center column is found",
          R.cost_center_column(otp) == "cost_center")
    check("the newer cost_center_name column is found",
          R.cost_center_column(runs) == "cost_center_name")
    check("a frame with neither reports none", R.cost_center_column(fleet) is None)

    reports = {"otp_by_cost_center": otp, "runs_by_cost_center": runs, "vehicles": fleet}
    filtered = R.filter_reports_to_region(reports, regions, "Indiana")
    check("OTP is filtered on cost_center",
          list(filtered["otp_by_cost_center"]["cost_center"]) == ["Indianapolis"])
    check("runs are filtered on cost_center_name",
          list(filtered["runs_by_cost_center"]["cost_center_name"]) == ["Indianapolis"])
    check("a frame naming no cost center passes through whole",
          len(filtered["vehicles"]) == 2)
    check("the original frame is not mutated", len(otp) == 2)

    check("cost_centers_in finds every name across both column spellings",
          R.cost_centers_in(reports) == ["Columbus Admin", "Indianapolis"])

    empty = R.filter_frame_to_region(
        pd.DataFrame(columns=["cost_center", "on_time_percentage"]), regions, "Indiana"
    )
    check("an empty frame filters to an empty frame, not an error", len(empty) == 0)


# =============================
def test_vehicle_working_two_regions(tmpdir):
    print("\nA vehicle that worked two regions")
    # The bug this guards: a vehicle is filed under the cost center it served
    # most, so filtering the finished frame hands the whole truck -- and every
    # run on it -- to one region. M-1 ran 3 legs for Ohio and 2 for Indiana;
    # Indiana's sheet must show 2, and its vehicle rows must sum to its cost
    # center rows.
    regions = regions_from({
        "regions": {
            "Indiana": {"cost_centers": ["Indianapolis"]},
            "Ohio": {"cost_centers": ["Cincinnati"]},
        }
    }, tmpdir)

    class Map:
        def resolve(self, profile):
            return {"INDY WC": "Indianapolis", "OH-A-CIN": "Cincinnati"}.get(profile)

    def leg(profile, n):
        return {
            "leg_id": f"{profile}-{n}", "shift_name": profile, "call_type": "BLS",
            "trip_status": "Clear", "vehicle_id": "1", "vehicle_name": "M-1",
            "pickup_time": "2026-08-01T09:00:00",
            "timestamps": [{"at_scene": "2026-08-01T09:02:00"}],
        }

    legs = [leg("OH-A-CIN", n) for n in range(3)] + [leg("INDY WC", n) for n in range(2)]
    legs.append({"leg_id": "orphan", "shift_name": None, "trip_status": "Canceled"})

    kept, unattributed = R.legs_in_region(legs, Map(), regions, "Indiana")
    check("only the region's legs are kept", len(kept) == 2, str(len(kept)))
    check("a leg no cost center claims is counted, not dropped silently",
          unattributed == 1, str(unattributed))

    frames, _, _ = R.build_region_leg_reports(legs, [], Map(), regions, "Indiana")
    vehicles = frames["runs_by_vehicle"]
    centres = frames["runs_by_cost_center"]
    check("the shared vehicle still appears in the region", len(vehicles) == 1, str(len(vehicles)))
    check("it carries only the region's runs, not the other region's",
          int(vehicles.iloc[0]["total_runs"]) == 2, str(vehicles["total_runs"].tolist()))
    check("vehicle runs sum to cost center runs",
          int(vehicles["total_runs"].sum()) == int(centres["total_runs"].sum()),
          f'{vehicles["total_runs"].sum()} vs {centres["total_runs"].sum()}')
    check("OTP counts only the region's legs",
          int(frames["otp_by_cost_center"]["total_runs"].sum()) == 2)

    # And the failure mode itself: filtering the finished frame gets it wrong.
    whole = R.build_runs_by_vehicle(legs, [], Map())
    filtered = R.filter_frame_to_region(whole, regions, "Indiana")
    check("filtering the finished frame is what loses the truck (why we do not)",
          len(filtered) == 0, str(len(filtered)))


def test_window():
    print("\nWindows")
    start, end = M.month_to_date(date(2026, 8, 23))
    check("month to date starts on the 1st", (start, end) == (date(2026, 8, 1), date(2026, 8, 23)))
    check("month to date counts inclusively", M.window_length(start, end) == 23)

    start, end = M.month_to_date(date(2026, 9, 1))
    check("the 1st is a one-day month to date", (start, end) == (date(2026, 9, 1), date(2026, 9, 1)))
    check("the window resets on the 1st", M.window_length(start, end) == 1)

    start, end = M.trailing(date(2026, 9, 1), 30)
    check("a trailing window crosses the month boundary",
          (start, end) == (date(2026, 8, 3), date(2026, 9, 1)))
    check("a trailing window is the length asked for", M.window_length(start, end) == 30)


def test_uhu_is_summed_not_averaged():
    print("\nMonthly UHU")
    # Two days, deliberately lopsided. A quiet day at 100% and a full day at
    # 50% average to 75%, but the month actually ran 11 of 21 hours -- 52.4%.
    df = pd.DataFrame({
        "_snapshot": [date(2026, 8, 1), date(2026, 8, 2)],
        "cost_center_name": ["Indianapolis", "Indianapolis"],
        "scheduled_hours": [1.0, 20.0],
        "worked_hours": [1.0, 20.0],
        "utilized_hours": [1.0, 10.0],
        "total_runs": [1, 20],
    })
    rolled = M.rollup_uhu(df, "cost_center_name")
    check("one row per cost center", len(rolled) == 1)
    row = rolled.iloc[0]
    check("worked hours are summed", row["worked_hours"] == 21.0, str(row["worked_hours"]))
    check("utilized hours are summed", row["utilized_hours"] == 11.0, str(row["utilized_hours"]))
    check("the ratio is sum over sum", abs(row["uhu_ratio"] - 0.524) < 0.001, str(row["uhu_ratio"]))
    check("the ratio is NOT the mean of the daily ratios",
          abs(row["uhu_ratio"] - 0.75) > 0.2, str(row["uhu_ratio"]))
    check("days counted reflects the days present", row["days_counted"] == 2)
    check("runs are summed", row["total_runs"] == 21)

    zero = pd.DataFrame({
        "_snapshot": [date(2026, 8, 1)],
        "cost_center_name": ["Indianapolis"],
        "scheduled_hours": [0.0], "worked_hours": [0.0],
        "utilized_hours": [0.0], "total_runs": [0],
    })
    check("a day with no worked hours does not divide by zero",
          M.rollup_uhu(zero, "cost_center_name").iloc[0]["uhu_ratio"] == 0)

    check("no rows gives an empty frame with the right columns",
          list(M.rollup_uhu(pd.DataFrame(), "cost_center_name").columns)[:2]
          == ["cost_center_name", "days_counted"])


def test_staffing_rollup():
    print("\nMonthly staffing shortfalls")
    df = pd.DataFrame({
        "_snapshot": [date(2026, 8, 1), date(2026, 8, 2), date(2026, 8, 3)],
        "cost_center": ["Indianapolis"] * 3,
        "shift_profile": ["IN-A-SBG-07-19"] * 3,
        "crew_count": [2, 1, 0],
        "crew_needed": [2, 2, 2],
    })
    rolled = M.rollup_staffing(df)
    row = rolled.iloc[0]
    check("every observed day is counted", row["days_observed"] == 3)
    check("only the short days are counted short", row["days_short"] == 2)
    check("the worst shortfall is the worst, not the last", row["worst_shortfall"] == 2)
    check("crew-days short is the total, not the count", row["shortfall_crew_days"] == 3)

    wheelchair = pd.DataFrame({
        "_snapshot": [date(2026, 8, 1)],
        "cost_center": ["Indianapolis"],
        "shift_profile": ["INDY WC"],
        "crew_count": [1], "crew_needed": [1],
    })
    check("a one-crew unit crewed to one is not short",
          M.rollup_staffing(wheelchair).iloc[0]["days_short"] == 0)

    over = pd.DataFrame({
        "_snapshot": [date(2026, 8, 1)],
        "cost_center": ["Indianapolis"], "shift_profile": ["IN-A-SBG-07-19"],
        "crew_count": [3], "crew_needed": [2],
    })
    check("a unit crewed above its minimum is not short by a negative",
          M.rollup_staffing(over).iloc[0]["shortfall_crew_days"] == 0)

    check("no rows gives an empty frame, not an error",
          len(M.rollup_staffing(pd.DataFrame())) == 0)


def test_window_filtering_and_gaps():
    print("\nWindow filtering and missing days")
    df = pd.DataFrame({
        "snapshot_date": ["2026-07-31", "2026-08-01", "2026-08-03", "2026-08-24"],
        "cost_center_name": ["Indianapolis"] * 4,
        "scheduled_hours": [1, 1, 1, 1], "worked_hours": [1, 1, 1, 1],
        "utilized_hours": [1, 1, 1, 1], "total_runs": [1, 1, 1, 1],
    })
    inside = M._in_window(df, date(2026, 8, 1), date(2026, 8, 23))
    check("the day before the window is excluded", len(inside) == 2, str(len(inside)))
    check("the day after the window is excluded",
          date(2026, 8, 24) not in list(inside["_snapshot"]))
    check("days present are reported",
          M.days_present(inside) == [date(2026, 8, 1), date(2026, 8, 3)])

    # The gap is the point: a month-to-date sheet built from two of 23 days is
    # useful and misleading, and the difference is whether the reader is told.
    summary = M.build_summary(
        "Indiana", R.Regions(path="does-not-exist.json"),
        date(2026, 8, 1), date(2026, 8, 23), legs=[],
        uhu_days=M.days_present(inside), staffing_days=[],
        sheets={"UHU by Cost Center": inside},
    )
    text = " ".join(str(v) for v in summary["value"])
    check("the summary states how many days accumulated", "2 of 23" in text, text[:200])
    check("the summary explains why they differ", "shifts endpoint" in text)
    check("the summary says UHU is summed, not averaged", "not the average" in text)
    check("the summary says the fleet sheets are absent", "no cost center on a vehicle" in text)

    whole = M.build_summary(
        "Indiana", R.Regions(path="does-not-exist.json"),
        date(2026, 8, 1), date(2026, 8, 2), legs=[],
        uhu_days=[date(2026, 8, 1), date(2026, 8, 2)],
        staffing_days=[date(2026, 8, 1), date(2026, 8, 2)],
        sheets={},
    )
    whole_text = " ".join(str(v) for v in whole["value"])
    check("a complete window does not warn about missing days",
          "shifts endpoint" not in whole_text)


def test_dependency_notes():
    print("\nDependency sheet")
    notes = R.build_dependency_notes(
        region="Indiana", window=(date(2026, 8, 1), date(2026, 8, 23)),
        uhu_days=[date(2026, 8, 1), date(2026, 8, 2)], staffing_days=[],
        unattributed=178,
    )
    check("the sheet has the documented columns",
          list(notes.columns) == R.DEPENDENCY_COLUMNS, str(list(notes.columns)))
    check("every row is filled in",
          not notes.isna().any().any() and (notes != "").all().all())

    text = " ".join(str(v) for v in notes.values.ravel())
    check("no format placeholder leaked into the prose", "%%" not in text)
    check("the UHU date limitation is stated", "THE DATE LIMITATION" in text)
    check("it names the endpoint that causes it", "today-1..today+2" in text)
    check("it says the days cannot be recovered",
          "Nothing can recover a day that was missed" in text)
    check("it says how many days actually accrued", "2 of 23 day(s)" in text)
    check("a report with no accumulated days says zero", "0 of 23 day(s)" in text)
    check("the summed-not-averaged rule is explained", "average of the daily ratios" in text)
    check("the region file is named as hand-maintained", "state/regions.json" in text)
    check("the unattributed legs are quantified", "178 leg(s)" in text)
    check("the fleet sheets are explained", "no cost center on a vehicle" in text)

    statuses = set(notes["status"])
    check("statuses are usable as a triage column",
          "Accruing -- cannot be backfilled" in statuses and "Complete" in statuses,
          str(sorted(statuses)))

    # The point of building it from the constants: change the span and the
    # sheet has to change with it, or it quietly describes a different report.
    original = R.UHU_SPAN
    try:
        R.UHU_SPAN = "task"
        task_text = " ".join(
            str(v) for v in R.build_dependency_notes().values.ravel())
        check("the biased default span is flagged as reading high",
              "tiles the shift" in task_text and "UHU_SPAN=transport" in task_text)
        R.UHU_SPAN = "transport"
        transport_text = " ".join(
            str(v) for v in R.build_dependency_notes().values.ravel())
        check("switching the span rewrites the note rather than repeating it",
              "tiles the shift" not in transport_text
              and "hit on arrival" in transport_text)
    finally:
        R.UHU_SPAN = original

    daily = R.build_dependency_notes(region="Indiana", metrics_date=date(2026, 8, 23))
    daily_text = " ".join(str(v) for v in daily.values.ravel())
    check("a single-day bundle says the run date, not a window",
          "2026-08-23" in daily_text and "day(s) in the window" not in daily_text)
    check("a run with nothing unattributed omits that row",
          "leg(s) in this window" not in daily_text)

    no_region = R.build_dependency_notes()
    check("without a region the region row is omitted",
          "state/regions.json" not in
          " ".join(str(v) for v in no_region.values.ravel()))


def test_append_round_trip(tmpdir):
    print("\nRound trip through a real append workbook")
    append_dir = os.path.join(tmpdir, "Append")
    path = os.path.join(append_dir, M.UHU_COST_CENTER_APPEND[0])
    sheet = M.UHU_COST_CENTER_APPEND[1]

    for day, worked, utilized in (("2026-08-01", 12.0, 6.0), ("2026-08-02", 24.0, 6.0)):
        OUT._append_to_workbook_xlsx(
            path, sheet,
            pd.DataFrame([{
                "cost_center_name": "Indianapolis",
                "scheduled_hours": worked, "worked_hours": worked,
                "utilized_hours": utilized, "total_runs": 5,
                "hours_per_run": 1.2, "uhu_ratio": utilized / worked,
            }]),
            dedupe_keys=["snapshot_date", "cost_center_name"],
            snapshot_date_value=day,
        )

    read_back = OUT.read_append_sheet(path, sheet)
    check("the sheet reads back", read_back is not None and len(read_back) == 2,
          str(None if read_back is None else len(read_back)))

    windowed = M._in_window(read_back, date(2026, 8, 1), date(2026, 8, 31))
    rolled = M.rollup_uhu(windowed, "cost_center_name")
    row = rolled.iloc[0]
    # 12/36, not the mean of 0.5 and 0.25.
    check("the month's ratio survives the Excel round trip",
          abs(row["uhu_ratio"] - 0.333) < 0.001, str(row["uhu_ratio"]))
    check("hours are summed across the round trip", row["worked_hours"] == 36.0)

    check("a missing workbook reads as nothing, not an error",
          OUT.read_append_sheet(os.path.join(append_dir, "nope.xlsx"), sheet) is None)
    check("a missing sheet reads as nothing, not an error",
          OUT.read_append_sheet(path, "No Such Sheet") is None)


def test_cli_parsing():
    print("\nCommand line")
    args = M.parse_args(["prog", "Indiana"])
    check("a bare word is the region", args["region"] == "Indiana")
    args = M.parse_args(["prog", "Indiana", "2026-08-23"])
    check("a bare date is the end date", args["end_date"] == date(2026, 8, 23))
    args = M.parse_args(["prog", "2026-08-23", "Indiana"])
    check("order does not matter",
          args["region"] == "Indiana" and args["end_date"] == date(2026, 8, 23))
    args = M.parse_args(["prog", "West", "Virginia"])
    check("a two-word region survives", args["region"] == "West Virginia")
    args = M.parse_args(["prog", "Indiana", "--days", "30"])
    check("--days is read", args["days"] == 30)
    args = M.parse_args(["prog", "Indiana", "--days=30"])
    check("--days=N is read", args["days"] == 30)
    args = M.parse_args(["prog", "--all-regions"])
    check("--all-regions is read", args["all_regions"] is True)
    try:
        M.parse_args(["prog", "--nonsense"])
        check("an unknown option is refused", False)
    except ValueError:
        check("an unknown option is refused", True)


def test_daily_runner_region_flag():
    print("\nDaily runner --region")
    import daily_report_runner_api as D
    check("--region NAME is read", D.parse_region(["prog", "--region", "Indiana"]) == "Indiana")
    check("--region=NAME is read", D.parse_region(["prog", "--region=Indiana"]) == "Indiana")
    check("no region is None", D.parse_region(["prog", "--zip"]) is None)
    check("a dangling --region is None", D.parse_region(["prog", "--region"]) is None)
    check("a region does not swallow the date",
          D.parse_cli_end_date(["prog", "--region", "Indiana", "2026-08-23"])
          == date(2026, 8, 23))
    check("the filename carries the region",
          D._out_name("CompanyWide_OTP", "_Indiana", "2026-08-23")
          == "CompanyWide_OTP_Indiana_2026-08-23.xlsx")
    check("the filename is unchanged without one",
          D._out_name("CompanyWide_OTP", "", "2026-08-23")
          == "CompanyWide_OTP_2026-08-23.xlsx")
    # A regional run must not write history the company-wide run owns.
    D._append(None, "whatever.xlsx", "Sheet", pd.DataFrame([{"a": 1}]), ["a"], "2026-08-23")
    check("a regional run appends nothing", not os.path.exists("whatever.xlsx"))


def main():
    tmpdir = tempfile.mkdtemp(prefix="region_tests_")
    try:
        test_region_matching(tmpdir)
        test_missing_and_broken_files(tmpdir)
        test_frame_filtering(tmpdir)
        test_vehicle_working_two_regions(tmpdir)
        test_window()
        test_uhu_is_summed_not_averaged()
        test_staffing_rollup()
        test_window_filtering_and_gaps()
        test_dependency_notes()
        test_append_round_trip(tmpdir)
        test_cli_parsing()
        test_daily_runner_region_flag()
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    print()
    if FAILURES:
        print(f"{len(FAILURES)} check(s) FAILED:")
        for name in FAILURES:
            print(f"  - {name}")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
