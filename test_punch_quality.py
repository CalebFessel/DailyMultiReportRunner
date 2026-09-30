"""
Checks for the punch-quality probe's UHU denominator figure.

    python -m pytest test_punch_quality.py     # preferred
    python test_punch_quality.py               # same checks, no pytest needed

This probe reports one number that someone will quote in an argument about
whether crews are closing their punches: the share of the UHU denominator that
comes from the shift-end fallback rather than a measured clock-out.

That number was wrong, and wrong in the direction that manufactures alarm. It
pooled every day in the rolling window, including today -- where every crew
currently on the road has an open punch, correctly bounded at `now`. On a real
run it read 41.1% when the finished day in the same window read 6.4%. The daily
runner reports on yesterday, so the finished day is the only one that describes
anything published.

These checks pin the separation. Nothing here touches the network.
"""

import sys
from datetime import datetime, date

import probe_punch_quality as P

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


def shift(start, end, deleted=False):
    """A shift row with bare UTC wall times, as /Schedule/Shifts returns them."""
    return {
        "user_id": 1, "shift_name": "OH-A-CIN-07-19",
        "start_time": start, "end_time": end, "deleted": deleted,
        "punches": [],
    }


def test_finished_day_detection():
    print("\ntest_finished_day_detection")

    now = datetime(2026, 9, 18, 16, 30)
    shifts = [
        shift("2026-09-17T07:00:00", "2026-09-17T19:00:00"),   # yesterday, over
        shift("2026-09-18T07:00:00", "2026-09-18T19:00:00"),   # today, running
    ]

    check("a day whose last shift has ended is finished",
          P.day_is_finished(shifts, date(2026, 9, 17), None, now))
    check("a day with a shift still running is not",
          not P.day_is_finished(shifts, date(2026, 9, 18), None, now),
          "its open punches are crews on the road, not missed punch-outs")

    check("a day with no shifts at all is not called finished",
          not P.day_is_finished(shifts, date(2026, 9, 19), None, now),
          "nothing to judge is different from judged and clean")


def test_overnight_shift_keeps_its_day_open():
    """
    A day is finished only when its LAST shift has ended, not its first.

    An overnight unit starting 19:00 on the 17th runs to 07:00 on the 18th.
    Reading at 06:00 on the 18th, the 17th is not over -- and its open punches
    are that crew, still out.
    """
    print("\ntest_overnight_shift_keeps_its_day_open")

    shifts = [
        shift("2026-09-17T07:00:00", "2026-09-17T19:00:00"),
        shift("2026-09-17T19:00:00", "2026-09-18T07:00:00"),
    ]

    early = datetime(2026, 9, 18, 6, 0)
    check("the day is still open while its overnight unit is out",
          not P.day_is_finished(shifts, date(2026, 9, 17), None, early),
          "the day-shift ending does not finish the day")

    later = datetime(2026, 9, 18, 9, 0)
    check("and finished once that unit is in",
          P.day_is_finished(shifts, date(2026, 9, 17), None, later))


def test_deleted_shifts_do_not_hold_a_day_open():
    print("\ntest_deleted_shifts_do_not_hold_a_day_open")

    shifts = [
        shift("2026-09-17T07:00:00", "2026-09-17T19:00:00"),
        shift("2026-09-17T19:00:00", "2026-09-18T23:00:00", deleted=True),
    ]
    now = datetime(2026, 9, 18, 9, 0)
    check("a deleted row cannot keep the day in progress",
          P.day_is_finished(shifts, date(2026, 9, 17), None, now),
          "it is not a crew that is out")


def test_offset_is_applied():
    """Shift times are bare UTC; the comparison happens on the tenant clock."""
    print("\ntest_offset_is_applied")

    from datetime import timedelta
    shifts = [shift("2026-09-17T23:00:00", "2026-09-18T03:00:00")]
    offset = timedelta(hours=-4)

    # 03:00 UTC is 23:00 on the 17th locally, so the shift starts AND ends on
    # the 17th once shifted -- and at 23:30 local it has just finished.
    now = datetime(2026, 9, 17, 23, 30)
    check("the shift is filed under its tenant-local start date",
          P.day_is_finished(shifts, date(2026, 9, 17), offset, now),
          "without the offset this row would land on the 18th and be missed")


def main():
    """The standalone runner, for a machine with no pytest."""
    tests = [
        test_finished_day_detection,
        test_overnight_shift_keeps_its_day_open,
        test_deleted_shifts_do_not_hold_a_day_open,
        test_offset_is_applied,
    ]
    for test in tests:
        try:
            test()
        except AssertionError:
            pass
        except Exception as exc:  # a test that broke rather than failed
            print(f"  ERROR {test.__name__}: {exc.__class__.__name__}: {exc}")
            FAILURES.append(f"{test.__name__} (crashed)")

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
