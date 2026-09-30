"""
Checks for the staffing review.

    python -m pytest test_staffing_review.py     # preferred
    python test_staffing_review.py               # same checks, no pytest needed

The review reads who was crewed out of the daily run's own append workbook,
so the checks that matter are: the crew ids survive the round trip through
Excel, an employee nobody crewed still gets a row, and a window with gaps is
reported as having gaps rather than as a fleet nobody staffed.

Crew are matched on user id, never on name. Two people share a name eventually,
and names change; the id in "First Last (ID 1234)" is the only stable handle.

Nothing here touches the network.
"""

import os
import sys
import shutil
import tempfile
from datetime import date, timedelta

import pandas as pd

import report_output as OUT
import employee_hours_report as H
import staffing_review as S
from test_employee_hours import employee

FAILURES = []


def check(name, condition, detail=""):
    """Assert, print, and record -- see test_region_monthly.check."""
    if condition:
        print(f"  ok    {name}")
        return
    message = f"{name}{(' -- ' + detail) if detail else ''}"
    print(f"  FAIL  {message}")
    FAILURES.append(name)
    raise AssertionError(message)


def staffing_row(snapshot, profile, crew, cost_center="Cincinnati", starts=None):
    """
    One staffing row as the daily runner writes it.

    `starts` is the day the shift actually began, which for a Tomorrow row is
    the day AFTER the snapshot. Defaults to the snapshot date, matching an
    Active Now row.
    """
    starts = starts or snapshot
    return {
        "snapshot_date": snapshot,
        "shift_profile": profile,
        "cost_center": cost_center,
        "start_time": f"{starts} 07:00:00",
        "end_time": f"{starts} 19:00:00",
        "crew_count": len(crew),
        "crew_needed": 2,
        "staffing_status": "OK" if len(crew) >= 2 else f"SHORT {2 - len(crew)}",
        "crew_members": "\n".join(f"A {name} (ID {uid})" for uid, name in crew),
    }


def write_append(tmpdir, rows, sheet=None, name="Append"):
    append_dir = os.path.join(tmpdir, name)
    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, S.STAFFING_APPEND), sheet or S.ACTIVE_SHEET,
        pd.DataFrame(rows),
        dedupe_keys=["snapshot_date", "cost_center", "shift_profile",
                     "start_time", "end_time"],
    )
    return append_dir


# =============================
# Parsing the crew cell
# =============================

def test_crew_ids_parse():
    print("\ntest_crew_ids_parse")

    check("one crew member", S.crew_ids("Jane Doe (ID 42)") == ["42"])
    check("two on separate lines",
          S.crew_ids("Jane Doe (ID 42)\nJohn Roe (ID 7)") == ["42", "7"])
    check("spacing inside the parens is tolerated",
          S.crew_ids("Jane Doe (ID  42 )") == ["42"])
    check("an empty cell yields nothing", S.crew_ids("") == [])
    check("a missing cell yields nothing", S.crew_ids(None) == [])
    check("a name containing digits is not mistaken for an id",
          S.crew_ids("Unit 7 Driver") == [],
          "only the (ID n) form counts")


# =============================
# The window
# =============================

def test_assignments_round_trip(tmpdir):
    print("\ntest_assignments_round_trip")

    append_dir = write_append(tmpdir, [
        staffing_row("2026-09-15", "OH-A-CIN-07-19", [(1, "One"), (2, "Two")]),
        staffing_row("2026-09-16", "OH-A-CIN-07-19", [(1, "One")]),
        staffing_row("2026-09-16", "OH-W-CIN-08-16", [(3, "Three")]),
    ])

    seen, days = S.assignments_in_window(append_dir, date(2026, 9, 1), date(2026, 9, 16))
    check("both snapshot days are found", days == [date(2026, 9, 15), date(2026, 9, 16)],
          f"got {days}")
    check("the crew member on both days has two", len(seen["1"]) == 2, f"got {seen.get('1')}")
    check("the one-day crew member has one", len(seen["2"]) == 1)
    check("units are tracked per day",
          seen["1"][date(2026, 9, 16)] == {"OH-A-CIN-07-19"}, f"got {seen['1']}")

    empty, days = S.assignments_in_window(append_dir, date(2026, 8, 1), date(2026, 8, 31))
    check("a window before anything was recorded is empty, not missing",
          empty == {} and days == [], f"got {empty}, {days}")


def test_rows_are_filed_by_shift_start_not_run_date(tmpdir):
    """
    A Tomorrow row is written on one day and describes the next.

    Filing it under snapshot_date would shift every one of those rows back a
    day, putting a Monday shift under Sunday -- and would make last_seen and
    days_since_last_seen wrong for everyone the Tomorrow sheet covers.
    """
    print("\ntest_rows_are_filed_by_shift_start_not_run_date")

    append_dir = write_append(tmpdir, [
        staffing_row("2026-09-15", "OH-A-CIN-07-19", [(1, "One")],
                     starts="2026-09-16"),
    ], sheet=S.TOMORROW_SHEET, name="ByStart")

    seen, days = S.assignments_in_window(append_dir, date(2026, 9, 1), date(2026, 9, 30))
    check("the row is filed under the day the shift started",
          days == [date(2026, 9, 16)], f"got {days}")
    check("and the crew member is credited with that day",
          list(seen["1"]) == [date(2026, 9, 16)], f"got {dict(seen['1'])}")


def test_both_sheets_are_read(tmpdir):
    """
    Active Now alone is a biased sample and must not be the only source.

    It holds units on shift at the instant the runner fired, so a crew whose
    shift did not span that moment never appears -- a morning run would put
    every night crew in the never-crewed list. Tomorrow carries the whole day.
    """
    print("\ntest_both_sheets_are_read")

    append_dir = os.path.join(tmpdir, "BothSheets", "Append")
    os.makedirs(append_dir, exist_ok=True)
    path = os.path.join(append_dir, S.STAFFING_APPEND)

    OUT._append_to_workbook_xlsx(
        path, S.TOMORROW_SHEET,
        pd.DataFrame([staffing_row("2026-09-15", "OH-A-CIN-19-07", [(2, "Night")],
                                   starts="2026-09-16")]),
        dedupe_keys=["snapshot_date", "shift_profile", "start_time"],
    )
    OUT._append_to_workbook_xlsx(
        path, S.ACTIVE_SHEET,
        pd.DataFrame([staffing_row("2026-09-16", "OH-A-CIN-07-19", [(1, "Day")])]),
        dedupe_keys=["snapshot_date", "shift_profile", "start_time"],
    )

    seen, days = S.assignments_in_window(append_dir, date(2026, 9, 1), date(2026, 9, 30))
    check("the day crew from Active Now is found", "1" in seen, f"got {sorted(seen)}")
    check("the night crew from Tomorrow is found too", "2" in seen,
          "reading only Active Now would report them as never crewed")
    check("both land on the same work date", days == [date(2026, 9, 16)], f"got {days}")


def test_missing_append_is_distinguished_from_empty(tmpdir):
    """
    No file at all and a file with no rows in range are different answers.

    The first means the daily runner has never run, which is a setup problem
    with a real remedy. The second means nobody was crewed. Reporting them the
    same way would send someone hunting for absent staff when the report simply
    was not there.
    """
    print("\ntest_missing_append_is_distinguished_from_empty")

    seen, days = S.assignments_in_window(os.path.join(tmpdir, "nothing-here"),
                                         date(2026, 9, 1), date(2026, 9, 16))
    check("a missing append reads as None", seen is None, f"got {seen}")

    append_dir = write_append(tmpdir, [
        staffing_row("2026-09-15", "OH-A-CIN-07-19", [(1, "One")]),
    ])
    seen, _ = S.assignments_in_window(append_dir, date(2026, 8, 1), date(2026, 8, 2))
    check("a present append with nothing in range reads as empty, not None",
          seen == {}, f"got {seen}")


# =============================
# The review itself
# =============================

def test_never_crewed_employees_are_reported(tmpdir):
    print("\ntest_never_crewed_employees_are_reported")

    roster_df = H.roster(
        [employee(1, "Works"), employee(2, "Never"), employee(3, "AlsoNever")],
        "Cincinnati", list(H.DEFAULT_LEVELS), include_inactive=False,
    )
    seen = {"1": {date(2026, 9, 15): {"OH-A-CIN-07-19"},
                  date(2026, 9, 16): {"OH-A-CIN-07-19"}}}

    result, never = S.review(roster_df, seen, {}, date(2026, 9, 1), date(2026, 9, 16))

    check("every employee has a row", len(result) == 3, f"got {len(result)}")
    check("the two never crewed are named",
          sorted(n for n, _ in never) == ["AlsoNever", "Never"], f"got {never}")
    check("they sort to the top", result.iloc[0]["days_crewed"] == 0)

    works = result[result["last_name"] == "Works"].iloc[0]
    check("the crewed employee has their days counted", works["days_crewed"] == 2)
    check("and their distinct units counted", works["units_crewed"] == 1)
    check("with a last-seen date", works["last_seen"] == date(2026, 9, 16))
    check("and days since, measured from the window end",
          works["days_since_last_seen"] == 0, f"got {works['days_since_last_seen']}")

    never_row = result[result["last_name"] == "Never"].iloc[0]
    check("a never-crewed employee has no last seen", never_row["last_seen"] is None)
    check("and no days-since rather than a misleading zero",
          pd.isna(never_row["days_since_last_seen"])
          and never_row["days_since_last_seen"] != 0,
          "pandas stores the absence as NaN, which Excel shows as blank; a 0 "
          "would read as 'crewed today'")


def test_new_hires_are_not_reported_as_never_crewed(tmpdir):
    """
    Someone hired mid-window has no history because they were not there.

    Left in the never-crewed list they sit beside people who genuinely are not
    being used, and a manager acting on that list would be asking the wrong
    question of the wrong person. They are marked and listed apart, not hidden:
    a brand-new hire with no assignments is still worth a look.
    """
    print("\ntest_new_hires_are_not_reported_as_never_crewed")

    roster_df = H.roster([
        employee(1, "Established", hire_date="2024-01-01"),
        employee(2, "JustHired", hire_date="2026-09-10"),
    ], "Cincinnati", list(H.DEFAULT_LEVELS), include_inactive=False)

    result, never = S.review(roster_df, {}, {}, date(2026, 7, 19), date(2026, 9, 16))

    check("the long-standing employee is reported as never crewed",
          [n for n, _ in never] == ["Established"], f"got {never}")
    check("the new hire is not", "JustHired" not in [n for n, _ in never])

    hired = result[result["last_name"] == "JustHired"].iloc[0]
    check("but is flagged so they can still be looked at",
          bool(hired["hired_during_window"]))
    established = result[result["last_name"] == "Established"].iloc[0]
    check("and the established employee is not flagged",
          not bool(established["hired_during_window"]))

    check("someone hired before the window counts normally",
          not S.hired_after("2024-01-01", date(2026, 7, 19)))
    check("someone hired inside it is flagged",
          S.hired_after("2026-09-10", date(2026, 7, 19)))
    check("an unparseable hire date does not flag",
          not S.hired_after("not a date", date(2026, 7, 19)),
          "an unreadable date must not quietly excuse someone from the list")
    check("a missing hire date does not flag",
          not S.hired_after(None, date(2026, 7, 19)))


def test_days_since_last_seen_flags_the_stale(tmpdir):
    print("\ntest_days_since_last_seen_flags_the_stale")

    roster_df = H.roster([employee(1, "Stale")], "Cincinnati",
                         list(H.DEFAULT_LEVELS), include_inactive=False)
    seen = {"1": {date(2026, 8, 1): {"OH-A-CIN-07-19"}}}
    result, never = S.review(roster_df, seen, {}, date(2026, 7, 19), date(2026, 9, 16))

    row = result.iloc[0]
    check("someone crewed once is not counted as never", not never)
    check("their gap is measured", row["days_since_last_seen"] == 46,
          f"got {row['days_since_last_seen']}")


def test_hours_column_is_supplementary(tmpdir):
    print("\ntest_hours_column_is_supplementary")

    roster_df = H.roster([employee(1, "Someone")], "Cincinnati",
                         list(H.DEFAULT_LEVELS), include_inactive=False)
    seen = {"1": {date(2026, 9, 16): {"OH-A-CIN-07-19"}}}

    result, _ = S.review(roster_df, seen, {}, date(2026, 9, 1), date(2026, 9, 16))
    check("with no hours recorded the column is blank, not zero",
          pd.isna(result.iloc[0]["hours_recorded"]),
          "a zero would claim they worked nothing; blank says we do not know")

    result, _ = S.review(roster_df, seen, {"1": 11.25},
                         date(2026, 9, 1), date(2026, 9, 16))
    check("when hours exist they are filled in",
          result.iloc[0]["hours_recorded"] == 11.25)


def test_summary_reports_gaps(tmpdir):
    print("\ntest_summary_reports_gaps")

    args = {"cost_center": "Cincinnati", "levels": list(H.DEFAULT_LEVELS),
            "all_levels": False, "days": 60}
    roster_df = H.roster([employee(1, "Someone")], "Cincinnati",
                         list(H.DEFAULT_LEVELS), include_inactive=False)
    days_present = [date(2026, 9, 15), date(2026, 9, 16)]

    result, _ = S.review(roster_df, {}, {}, date(2026, 7, 19), date(2026, 9, 16))
    sheet = S.summary_sheet(args, roster_df, days_present, date(2026, 7, 19),
                            date(2026, 9, 16), {}, [("Someone", "A")], {}, result)
    items = dict(zip(sheet["item"], sheet["value"]))
    notes = dict(zip(sheet["item"], sheet["note"]))

    check("days recorded is stated", items["Days actually recorded"] == 2)
    check("the gap is named", "58 day(s) have no record" in notes["Days actually recorded"],
          f"got {notes['Days actually recorded']}")
    check("never-crewed is surfaced as a headline", items["Never crewed in the window"] == 1)
    check("the days-are-not-hours caveat is stated",
          "not hours" in notes["What a 'day crewed' means"],
          "a manager reading days as hours is the likely misreading")
    check("the sheets read are named, with the sampling caveat",
          "point-in-time" in notes["Read from"], f"got {notes.get('Read from')}")
    check("the gap note warns it can create false never-crewed entries",
          "never crewed" in notes["Days actually recorded"],
          "someone who worked only on missing days looks like they never worked")

    full = S.summary_sheet(args, roster_df, days_present, date(2026, 9, 15),
                           date(2026, 9, 16), {}, [], {}, result)
    note = dict(zip(full["item"], full["note"]))["Days actually recorded"]
    check("a complete window says so", "COMPLETE" in note, f"got {note}")


def test_coverage_sheet_separates_the_two_kinds_of_gap():
    """
    Missing days are not one thing, and conflating them misleads.

    Days before the first record were never capturable -- the report did not
    exist yet, so nothing was lost through neglect. Days after it that have no
    record were capturable on the day and were not captured. Only the second
    kind is a process problem, and only the second kind can be prevented from
    recurring.
    """
    print("\ntest_coverage_sheet_separates_the_two_kinds_of_gap")

    present = [date(2026, 8, 19), date(2026, 8, 20), date(2026, 8, 22)]
    df = S.coverage_sheet(present, date(2026, 8, 17), date(2026, 8, 23))

    check("every day in the window has a row", len(df) == 7, f"got {len(df)}")
    by_date = df.set_index("date")

    check("days before the first record are named as such",
          by_date.loc[date(2026, 8, 17), "note"] == "before the first recorded day")
    check("a skipped run after it is named a missed run",
          by_date.loc[date(2026, 8, 21), "note"] == "run missed")
    check("and the two are not conflated",
          by_date.loc[date(2026, 8, 17), "note"]
          != by_date.loc[date(2026, 8, 21), "note"])
    check("recorded days carry no note", by_date.loc[date(2026, 8, 19), "note"] == "")
    check("recorded is yes/no per day",
          list(df["recorded"]) == ["no", "no", "yes", "yes", "no", "yes", "no"],
          f"got {list(df['recorded'])}")
    check("the weekday is shown, so a weekend pattern is visible",
          by_date.loc[date(2026, 8, 23), "weekday"] == "Sun")

    empty = S.coverage_sheet([], date(2026, 8, 17), date(2026, 8, 19))
    check("with no history at all every day says so",
          set(empty["note"]) == {"no history at all"}, f"got {set(empty['note'])}")
    check("and none are called missed runs",
          "run missed" not in set(empty["note"]),
          "blaming a missed run when the report never existed is the wrong story")


def test_cli():
    print("\ntest_cli")

    args = S.parse_args([])
    check("defaults to Cincinnati", args["cost_center"] == "Cincinnati")
    check("defaults to 60 days", args["days"] == 60)
    check("defaults to the three levels",
          args["levels"] == ["OH EMT - Driver", "OH EMT - Non Driver", "OH NEMT"],
          f"got {args['levels']}")
    check("levels are not all by default", not args["all_levels"])

    args = S.parse_args(["--all-levels", "--days", "14", "--cost-center", "Toledo"])
    check("--all-levels is read", args["all_levels"])
    check("--days is read", args["days"] == 14)
    check("--cost-center is read", args["cost_center"] == "Toledo")

    try:
        S.parse_args(["--bogus"])
        check("an unknown option is rejected", False, "it was accepted")
    except SystemExit:
        check("an unknown option is rejected", True)


def main():
    """The standalone runner, for a machine with no pytest."""
    tmpdir = tempfile.mkdtemp(prefix="staffing_review_tests_")
    tests = [
        (test_crew_ids_parse, ()),
        (test_assignments_round_trip, (tmpdir,)),
        (test_rows_are_filed_by_shift_start_not_run_date, (tmpdir,)),
        (test_both_sheets_are_read, (tmpdir,)),
        (test_missing_append_is_distinguished_from_empty, (tmpdir,)),
        (test_never_crewed_employees_are_reported, (tmpdir,)),
        (test_new_hires_are_not_reported_as_never_crewed, (tmpdir,)),
        (test_days_since_last_seen_flags_the_stale, (tmpdir,)),
        (test_hours_column_is_supplementary, (tmpdir,)),
        (test_summary_reports_gaps, (tmpdir,)),
        (test_coverage_sheet_separates_the_two_kinds_of_gap, ()),
        (test_cli, ()),
    ]
    try:
        for test, test_args in tests:
            try:
                test(*test_args)
            except AssertionError:
                pass
            except Exception as exc:  # a test that broke rather than failed
                print(f"  ERROR {test.__name__}: {exc.__class__.__name__}: {exc}")
                FAILURES.append(f"{test.__name__} (crashed)")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    print()
    if FAILURES:
        print(f"{len(FAILURES)} check(s) FAILED:")
        for name in FAILURES:
            print(f"  - {name}")
        return 1
    print(f"All checks passed ({len(tests)} tests).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
