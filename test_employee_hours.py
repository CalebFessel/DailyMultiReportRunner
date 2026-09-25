"""
Checks for the per-employee hours report.

    python -m pytest test_employee_hours.py     # preferred
    python test_employee_hours.py               # same checks, no pytest needed

This report exists to surface people who did NOT work, so the failure that
matters is not a wrong total -- it is a missing row. Group the punches and
every zero-hour employee disappears, which is precisely the population being
asked about. The join is therefore onto the roster, and these checks pin it.

The second failure that matters is a partial window presented as a whole one.
Hours only exist for days the report has already run, so a 60-day request on
day three must say three, not quietly average what it has.

Nothing here touches the network; the employees, shifts and punches are
synthetic and the append workbook is written to a temp directory with the real
writer, then read back with the real reader.
"""

import os
import sys
import shutil
import tempfile
from datetime import date, datetime, timedelta

import pandas as pd

import report_output as OUT
import employee_hours_report as E

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


def employee(user_id, last, level="OH EMT - Driver", cost_center="Cincinnati",
             **extra):
    row = {
        "user_id": user_id, "last_name": last, "first_name": "A",
        "employee_num": f"E{user_id}", "level": level, "license_level": level,
        "cost_center_name": cost_center, "station": "Main", "status": "Active",
        "hire_date": "2024-01-01", "disabled": False, "termination_date": None,
    }
    row.update(extra)
    return row


def shift(user_id, start, end, punches):
    """A shift row with bare UTC wall times, as /Schedule/Shifts returns them."""
    return {
        "user_id": user_id, "shift_name": "OH-A-CIN-07-19",
        "start_time": start, "end_time": end, "deleted": False,
        "punches": [
            {"start_time": p[0], "end_time": p[1], "deleted": False} for p in punches
        ],
    }


# =============================
# Matching
# =============================

# The tenant's real level vocabulary, copied from --list-levels output as
# (level, license_level). Levels are state-prefixed -- OH, WV, IN and MD each
# have their own EMT and NEMT entries -- and license_level carries a category
# prefix on top of the same name. The defaults were wrong once for exactly this
# reason -- the unprefixed "EMT - Driver" matches nobody here -- and an empty
# roster reads like "nobody works in Cincinnati" rather than "the filter is
# misconfigured", so the real vocabulary is pinned here.
TENANT_LEVELS = [
    ("OH EMT - Driver", "EMT - OH EMT - Driver"),
    ("OH EMT - Non Driver", "EMT - OH EMT - Non Driver"),
    ("OH NEMT", "EMT - OH NEMT"),
    ("WV EMT", "EMT - WV EMT"),
    ("WV EMVO", "Driver - WV EMVO"),
    ("IN EMT Driver", "EMT - IN EMT Driver"),
    ("IN NEMT", "EMT - IN NEMT"),
    ("IN SC DRIVER", "Driver - IN SC DRIVER"),
    ("IN WC DRIVER", "Driver - IN WC DRIVER"),
    ("MD EMT Driver", "EMT - MD EMT Driver"),
    ("MD EMT-Non Driver", "EMT - MD EMT-Non Driver"),
    ("MD NEMT", "EMT - MD NEMT"),
    ("MD SC DRIVER", "Driver - MD SC DRIVER"),
    ("OH WC Driver", "Driver - OH WC Driver"),
    ("OH SC DRIVER", "Driver - OH SC DRIVER"),
    ("OH - Advanced EMT", "OH - Advanced EMT - OH - Adv EMT"),
    ("Paramedic - Level I", "Paramedic - Level I"),
    ("Paramedic - Level II", "Paramedic - Level II"),
    ("Paramedic - Intro", "Paramedic - Intro"),
    ("Dispatcher", "Dispatcher - Level I"),
    ("Dispatch Supervisor", "Dispatcher - Level II"),
    ("Office - Level I", "Office - Level I"),
    ("Office - Level II", "Office - Level II - Level II"),
    ("Call Taker - Level I", "Call Taker - Level I"),
    ("PRN", "PRN - PRN"),
]


def test_defaults_match_the_tenants_real_levels():
    """
    The shipped defaults must select exactly the three Ohio field positions.

    Selecting none is the failure that hides: the roster comes back empty and
    reads as a staffing finding instead of a filter that matches nothing.
    """
    print("\ntest_defaults_match_the_tenants_real_levels")

    selected = [
        level for level, license_level in TENANT_LEVELS
        if E.matches_levels({"level": level, "license_level": license_level},
                            E.DEFAULT_LEVELS)
    ]
    check(
        "the defaults select exactly the three Ohio field positions",
        selected == ["OH EMT - Driver", "OH EMT - Non Driver", "OH NEMT"],
        f"got {selected}",
    )
    check("and they select something at all", selected,
          "a default matching nothing produces an empty roster that reads as "
          "a finding rather than a misconfiguration")

    check(
        "the unprefixed spelling matches nobody here",
        not any(
            E.matches_levels({"level": level, "license_level": license_level},
                             ["EMT - Driver", "EMT - Non Driver", "NEMT"])
            for level, license_level in TENANT_LEVELS
        ),
        "which is why the defaults are the state-prefixed names",
    )

    for other in ("WV EMT", "IN NEMT", "MD NEMT", "OH WC Driver", "OH SC DRIVER"):
        check(f"'{other}' is not swept in",
              not E.matches_levels({"level": other}, E.DEFAULT_LEVELS))


def test_level_matching_survives_spacing():
    print("\ntest_level_matching_survives_spacing")

    wanted = list(E.DEFAULT_LEVELS)
    for spelling in ("OH EMT - Driver", "OH EMT-Driver", "oh emt -  driver",
                     "OH EMT -Driver"):
        check(f"'{spelling}' matches OH EMT - Driver",
              E.matches_levels({"level": spelling}, wanted))

    check("OH NEMT matches", E.matches_levels({"level": "OH NEMT"}, wanted))
    check("Paramedic does not match",
          not E.matches_levels({"level": "Paramedic - Level I"}, wanted))
    check("OH EMT - Driver Trainee does not match",
          not E.matches_levels({"level": "OH EMT - Driver Trainee"}, wanted),
          "substring matching would wrongly include it")

    check("license_level is checked when level is blank",
          E.matches_levels({"level": "", "license_level": "OH NEMT"}, wanted))

    check("levels=None takes everyone, for --all-levels",
          all(E.matches_levels({"level": level}, None)
              for level, _ in TENANT_LEVELS))


def test_cost_center_matching():
    print("\ntest_cost_center_matching")

    check("exact name matches",
          E.matches_cost_center({"cost_center_name": "Cincinnati"}, "Cincinnati"))
    check("a longer name containing it matches",
          E.matches_cost_center({"cost_center_name": "Cincinnati North"}, "cincinnati"))
    check("another cost center does not",
          not E.matches_cost_center({"cost_center_name": "Columbus"}, "Cincinnati"))
    check("a blank cost center does not",
          not E.matches_cost_center({"cost_center_name": None}, "Cincinnati"))


def test_roster_includes_and_excludes():
    print("\ntest_roster_includes_and_excludes")

    people = [
        employee(1, "Keeps"),
        employee(2, "AlsoKeeps", level="OH NEMT"),
        employee(3, "WrongLevel", level="Paramedic - Level I"),
        employee(4, "WrongCostCenter", cost_center="Columbus"),
        employee(5, "Terminated", termination_date="2026-08-01"),
        employee(6, "Disabled", disabled=True),
    ]
    levels = list(E.DEFAULT_LEVELS)

    df = E.roster(people, "Cincinnati", levels, include_inactive=False)
    check("only matching active employees are listed",
          sorted(df["last_name"]) == ["AlsoKeeps", "Keeps"], f"got {list(df['last_name'])}")

    df = E.roster(people, "Cincinnati", levels, include_inactive=True)
    check("--include-inactive adds the terminated and disabled",
          sorted(df["last_name"]) == ["AlsoKeeps", "Disabled", "Keeps", "Terminated"],
          f"got {sorted(df['last_name'])}")
    inactive = df[df["last_name"] == "Terminated"]["status"].iloc[0]
    check("and marks them Inactive rather than hiding why", inactive == "Inactive")


# =============================
# Hours
# =============================

def test_hours_from_punches():
    print("\ntest_hours_from_punches")

    shifts = [
        shift(1, "2026-09-16T11:00:00", "2026-09-16T23:00:00",
              [("2026-09-16T11:05:00", "2026-09-16T22:55:00")]),      # 11.83 h
        shift(2, "2026-09-16T11:00:00", "2026-09-16T23:00:00",
              [("2026-09-16T11:00:00", "2026-09-16T15:00:00"),
               ("2026-09-16T16:00:00", "2026-09-16T20:00:00")]),      # 8.0 h
    ]
    totals, open_punches, running = E.hours_by_employee_day(shifts)

    one = round(totals[(1, date(2026, 9, 16))], 2)
    two = round(totals[(2, date(2026, 9, 16))], 2)
    check("a single punch is its own length", one == 11.83, f"got {one}")
    check("two punches on one shift add up", two == 8.0, f"got {two}")
    check("no punches were left open", not open_punches)


def test_open_punch_is_bounded_at_now():
    """
    A crew still on the clock must not be paid to the end of their shift.

    This is the same defect the UHU work fixed: an overnight unit read at 06:00
    against an 08:00 finish billed two hours nobody had worked yet.
    """
    print("\ntest_open_punch_is_bounded_at_now")

    shifts = [shift(1, "2026-09-16T20:00:00", "2026-09-17T08:00:00",
                    [("2026-09-16T20:00:00", None)])]

    # Naive, because that is what tenant_now(offset) hands the real caller.
    now = datetime(2026, 9, 17, 2, 0)
    totals, open_punches, running = E.hours_by_employee_day(shifts, now=now)
    hours = round(totals[(1, date(2026, 9, 16))], 2)
    check("an open punch is billed only up to now", hours == 6.0, f"got {hours}")
    check("it is counted as open", open_punches[1] == 1)
    check("and recognised as a shift still running", running[1] == 1)

    finished = datetime(2026, 9, 17, 12, 0)
    totals, _, running = E.hours_by_employee_day(shifts, now=finished)
    hours = round(totals[(1, date(2026, 9, 16))], 2)
    check("once the shift has ended it bills in full, not past it",
          hours == 12.0, f"got {hours}")
    check("and is no longer counted as still running", not running)

    try:
        E.hours_by_employee_day(shifts, now=datetime.now().astimezone())
        check("an aware `now` is rejected with a useful message", False,
              "it was accepted, and would fail later on a comparison instead")
    except ValueError as exc:
        check("an aware `now` is rejected with a useful message",
              "tenant_now" in str(exc), f"got {exc}")


def test_overnight_hours_stay_on_the_starting_day():
    print("\ntest_overnight_hours_stay_on_the_starting_day")

    shifts = [shift(1, "2026-09-16T20:00:00", "2026-09-17T08:00:00",
                    [("2026-09-16T22:00:00", "2026-09-17T06:00:00")])]
    totals, _, _ = E.hours_by_employee_day(shifts)
    check("the whole punch lands on the day it started",
          round(totals[(1, date(2026, 9, 16))], 2) == 8.0, f"got {dict(totals)}")
    check("and nothing lands on the following day",
          (1, date(2026, 9, 17)) not in totals)


def test_deleted_rows_are_ignored():
    print("\ntest_deleted_rows_are_ignored")

    gone = shift(1, "2026-09-16T11:00:00", "2026-09-16T23:00:00",
                 [("2026-09-16T11:00:00", "2026-09-16T23:00:00")])
    gone["deleted"] = True
    check("a deleted shift contributes nothing",
          not E.hours_by_employee_day([gone])[0])

    live = shift(1, "2026-09-16T11:00:00", "2026-09-16T23:00:00",
                 [("2026-09-16T11:00:00", "2026-09-16T15:00:00"),
                  ("2026-09-16T16:00:00", "2026-09-16T20:00:00")])
    live["punches"][1]["deleted"] = True
    totals, _, _ = E.hours_by_employee_day([live])
    check("a deleted punch is dropped but its siblings are kept",
          round(totals[(1, date(2026, 9, 16))], 2) == 4.0, f"got {dict(totals)}")


# =============================
# The whole point: zeros
# =============================

def test_zero_hour_employees_are_listed():
    print("\ntest_zero_hour_employees_are_listed")

    roster_df = E.roster(
        [employee(1, "Worked"), employee(2, "NeverWorked"), employee(3, "AlsoNever")],
        "Cincinnati", ["OH EMT - Driver"], include_inactive=False,
    )
    recorded = pd.DataFrame([
        {"work_date": date(2026, 9, 15), "user_id": 1, "hours_worked": 8.0},
        {"work_date": date(2026, 9, 16), "user_id": 1, "hours_worked": 4.5},
    ])

    summary, zero = E.summarize(roster_df, recorded,
                                date(2026, 9, 1), date(2026, 9, 16))

    check("every rostered employee has a row", len(summary) == 3, f"got {len(summary)}")
    check("the two who never worked are reported as zero",
          sorted(n for n, _ in zero) == ["AlsoNever", "NeverWorked"], f"got {zero}")

    worked = summary[summary["last_name"] == "Worked"].iloc[0]
    check("the one who worked has their hours summed",
          worked["hours_worked"] == 12.5, f"got {worked['hours_worked']}")
    check("and their days counted", worked["days_worked"] == 2)

    never = summary[summary["last_name"] == "NeverWorked"].iloc[0]
    check("a zero is a real 0.00, not blank", never["hours_worked"] == 0.0)
    check("with zero days", never["days_worked"] == 0)

    check("zero-hour employees sort to the top",
          summary.iloc[0]["hours_worked"] == 0.0,
          "the people being asked about should not be at the bottom of the sheet")


def test_no_recorded_days_still_lists_everyone():
    """Before the first run has accrued anything, the roster still stands."""
    print("\ntest_no_recorded_days_still_lists_everyone")

    roster_df = E.roster([employee(1, "Someone"), employee(2, "Another")],
                         "Cincinnati", ["OH EMT - Driver"], include_inactive=False)
    summary, zero = E.summarize(roster_df, None, date(2026, 7, 19), date(2026, 9, 16))

    check("both employees appear", len(summary) == 2)
    check("both read zero", list(summary["hours_worked"]) == [0.0, 0.0])
    check("and both are reported as zero-hour", len(zero) == 2)


# =============================
# Honesty about the window
# =============================

def test_summary_states_the_real_coverage():
    print("\ntest_summary_states_the_real_coverage")

    args = {"cost_center": "Cincinnati", "levels": ["OH EMT - Driver"], "days": 60,
            "append_dir": "Reports/Append", "no_append": False}
    roster_df = E.roster([employee(1, "Someone")], "Cincinnati",
                         ["OH EMT - Driver"], include_inactive=False)
    recorded = pd.DataFrame([
        {"work_date": date(2026, 9, 15), "user_id": 1, "hours_worked": 8.0},
        {"work_date": date(2026, 9, 16), "user_id": 1, "hours_worked": 8.0},
    ])
    start, end = date(2026, 7, 19), date(2026, 9, 16)
    days_present = [date(2026, 9, 15), date(2026, 9, 16)]

    sheet = E.summary_sheet(args, roster_df, recorded, start, end, days_present,
                            pd.DataFrame())
    items = dict(zip(sheet["item"], sheet["value"]))
    notes = dict(zip(sheet["item"], sheet["note"]))

    check("the window asked for is stated",
          items["Window asked for"] == "2026-07-19 to 2026-09-16")
    check("the days actually recorded are stated", items["Days of hours actually recorded"] == 2)
    check("and the shortfall is named in the note",
          "58 day(s) missing" in notes["Days of hours actually recorded"],
          f"got {notes['Days of hours actually recorded']}")
    check("the note says why it cannot be backfilled",
          "backfill" in notes["Days of hours actually recorded"].lower())
    check("employees listed is reported", items["Employees listed"] == 1)

    full = E.summary_sheet(args, roster_df, recorded, date(2026, 9, 15),
                           date(2026, 9, 16), days_present, pd.DataFrame())
    full_note = dict(zip(full["item"], full["note"]))["Days of hours actually recorded"]
    check("a complete window says so instead of warning",
          "COMPLETE" in full_note and "backfill" not in full_note.lower(),
          f"got {full_note}")


def test_no_append_is_flagged_as_lossy():
    print("\ntest_no_append_is_flagged_as_lossy")

    args = {"cost_center": "Cincinnati", "levels": ["OH EMT - Driver"], "days": 60,
            "append_dir": "Reports/Append", "no_append": True}
    roster_df = E.roster([employee(1, "Someone")], "Cincinnati",
                         ["OH EMT - Driver"], include_inactive=False)
    sheet = E.summary_sheet(args, roster_df, None, date(2026, 9, 1),
                            date(2026, 9, 16), [], pd.DataFrame())
    note = dict(zip(sheet["item"], sheet["note"]))["Append workbook"]
    check("the sheet says today's hours were not recorded",
          "not retrievable" in note.lower(), f"got {note}")


# =============================
# Round trip through Excel
# =============================

def test_append_round_trip(tmpdir):
    """
    The window is read back out of Excel, so the round trip is what matters.

    Excel reads a column of digit strings back as integers, which is how an
    earlier de-dupe bug in this project went unnoticed -- a key written as "7"
    never matched the 7 read back, and a re-run appended instead of replacing.
    """
    print("\ntest_append_round_trip")

    append_dir = os.path.join(tmpdir, "Append")
    rows = pd.DataFrame([
        {"work_date": "2026-09-15", "user_id": 1, "employee_num": "E1",
         "last_name": "Someone", "first_name": "A", "level": "OH EMT - Driver",
         "cost_center_name": "Cincinnati", "hours_worked": 8.0,
         "open_punches": 0, "still_on_shift": 0},
    ])
    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, E.APPEND_FILE), E.DAILY_SHEET, rows,
        dedupe_keys=["work_date", "user_id"],
    )

    back = E.window_rows(append_dir, date(2026, 9, 1), date(2026, 9, 16))
    check("the recorded day reads back", back is not None and len(back) == 1,
          f"got {back}")
    check("as a real date", back["work_date"].iloc[0] == date(2026, 9, 15))
    check("stored as a string so the de-dupe key survives Excel",
          isinstance(rows["work_date"].iloc[0], str),
          "a date object comes back as a Timestamp and never matches again")

    # Same day again, different hours: it must replace, not duplicate.
    rows.loc[0, "hours_worked"] = 9.5
    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, E.APPEND_FILE), E.DAILY_SHEET, rows,
        dedupe_keys=["work_date", "user_id"],
    )
    back = E.window_rows(append_dir, date(2026, 9, 1), date(2026, 9, 16))
    check("re-running the same day replaces rather than duplicates",
          len(back) == 1, f"got {len(back)} rows")
    check("with the newer value", back["hours_worked"].iloc[0] == 9.5)

    outside = E.window_rows(append_dir, date(2026, 8, 1), date(2026, 8, 31))
    check("a window with nothing in it comes back empty, not wrong",
          outside is not None and outside.empty, f"got {outside}")


def test_cli_defaults():
    print("\ntest_cli_defaults")

    args = E.parse_args([])
    check("defaults to Cincinnati", args["cost_center"] == "Cincinnati")
    check("defaults to 60 days", args["days"] == 60)
    check("defaults to the three requested levels",
          args["levels"] == ["OH EMT - Driver", "OH EMT - Non Driver", "OH NEMT"],
          f"got {args['levels']}")

    args = E.parse_args(["--cost-center", "Toledo", "--days", "30",
                         "--levels", "NEMT,Paramedic"])
    check("the cost center can be changed", args["cost_center"] == "Toledo")
    check("the window can be changed", args["days"] == 30)
    check("the levels can be changed", args["levels"] == ["NEMT", "Paramedic"])

    try:
        E.parse_args(["--nonsense"])
        check("an unknown option is rejected", False, "it was accepted")
    except SystemExit:
        check("an unknown option is rejected", True)


def main():
    """The standalone runner, for a machine with no pytest."""
    tmpdir = tempfile.mkdtemp(prefix="emp_hours_tests_")
    tests = [
        (test_defaults_match_the_tenants_real_levels, ()),
        (test_level_matching_survives_spacing, ()),
        (test_cost_center_matching, ()),
        (test_roster_includes_and_excludes, ()),
        (test_hours_from_punches, ()),
        (test_open_punch_is_bounded_at_now, ()),
        (test_overnight_hours_stay_on_the_starting_day, ()),
        (test_deleted_rows_are_ignored, ()),
        (test_zero_hour_employees_are_listed, ()),
        (test_no_recorded_days_still_lists_everyone, ()),
        (test_summary_states_the_real_coverage, ()),
        (test_no_append_is_flagged_as_lossy, ()),
        (test_append_round_trip, (tmpdir,)),
        (test_cli_defaults, ()),
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
