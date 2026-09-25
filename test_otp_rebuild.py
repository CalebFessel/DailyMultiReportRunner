"""
Checks for the backfilled OTP rebuild.

    python -m pytest test_otp_rebuild.py     # preferred
    python test_otp_rebuild.py               # same checks, no pytest needed

The rebuild exists to restate a published series, so the ways it can be wrong
are the ways a restated series misleads: a window percentage that averages the
days instead of summing them, a day the API never returned quietly reading as a
day with no runs, and legs excluded from the daily report creeping back in.

Nothing here touches the network -- the legs are synthetic and the window is
built by hand, so the arithmetic is checkable against numbers written out in
the test rather than against whatever the API happened to return.
"""

import sys
from datetime import date, timedelta

import pandas as pd

import traumasoft_reports as R
import rebuild_otp_history as B

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


class FixedMap:
    """A cost-center map with no state file behind it."""

    def __init__(self, mapping):
        self.mapping = mapping
        self.counts = mapping

    def resolve(self, shift_name):
        return self.mapping.get(shift_name)

    def ambiguous(self):
        return {}


def leg(day, hour, minute_late, shift="UNIT-A", call_type="BLS"):
    """One leg, arriving `minute_late` minutes after its scheduled pickup."""
    pickup = f"{day}T{hour:02d}:00:00-04:00"
    arrived = f"{day}T{hour:02d}:{minute_late:02d}:00-04:00"
    return {
        "leg_id": f"{day}-{hour}", "run_number": f"R{hour}",
        "trip_status": "Completed", "shift_name": shift, "call_type": call_type,
        "pickup_time": pickup, "timestamps": [{"at_scene": arrived}],
    }


def test_day_bucketing():
    print("\ntest_day_bucketing")

    legs = [leg("2026-08-01", 9, 2), leg("2026-08-02", 9, 2)]
    legs.append({"leg_id": "x", "trip_status": "Completed", "shift_name": "UNIT-A",
                 "timestamps": [{"at_scene": "2026-08-01T09:05:00-04:00"}]})

    by_day, undated = B.bucket_by_day(legs)
    check("legs land on their scheduled pickup date",
          sorted(by_day) == [date(2026, 8, 1), date(2026, 8, 2)], f"got {sorted(by_day)}")
    check("a leg with no scheduled pickup is counted apart, not dropped silently",
          undated == 1, f"got {undated}")


def test_window_is_summed_not_averaged():
    """
    The window percentage must weight days by volume.

    Twenty legs on a good day and two on a bad one is not the same as one good
    day and one bad day, and averaging the two daily percentages says it is.
    """
    print("\ntest_window_is_summed_not_averaged")

    busy = [leg("2026-08-01", 8 + i, 2) for i in range(10)]          # 10 on time
    quiet = [leg("2026-08-02", 9, 45), leg("2026-08-02", 10, 45)]    # 2 late

    by_day, _ = B.bucket_by_day(busy + quiet)
    frames = B.build(by_day, date(2026, 8, 1), date(2026, 8, 2),
                     FixedMap({"UNIT-A": "Columbus"}))
    daily = frames["daily"]

    check("each day is scored separately", list(daily["legs_scored"]) == [10, 2],
          f"got {list(daily['legs_scored'])}")
    check("the busy day is 100%", daily["on_time_percentage"].iloc[0] == 100.0)
    check("the quiet day is 0%", daily["on_time_percentage"].iloc[1] == 0.0)

    scored = int(daily["legs_scored"].sum())
    on_time = int((daily["early_runs"] + daily["on_time_runs"]).sum())
    summed = round(100.0 * on_time / scored, 2)
    averaged = round(daily["on_time_percentage"].mean(), 2)

    check("the window sums to 10 of 12", summed == 83.33, f"got {summed}")
    check("which is not the average of the daily percentages",
          averaged == 50.0 and summed != averaged,
          f"averaging would have published {averaged}%")


def test_empty_days_are_visible():
    """A day the API returned nothing for must not read as a day with no runs."""
    print("\ntest_empty_days_are_visible")

    by_day, _ = B.bucket_by_day([leg("2026-08-01", 9, 2), leg("2026-08-03", 9, 2)])
    frames = B.build(by_day, date(2026, 8, 1), date(2026, 8, 3),
                     FixedMap({"UNIT-A": "Columbus"}))
    daily = frames["daily"]

    check("every day in the window gets a row", len(daily) == 3, f"got {len(daily)}")
    gap = daily[daily["metrics_date"] == date(2026, 8, 2)].iloc[0]
    check("the missing day returned no legs", gap["legs_returned"] == 0)
    check("and its percentage is blank, not zero",
          pd.isna(gap["on_time_percentage"]),
          "a 0% day and a day with no data are different claims")


def test_excluded_cost_centers_stay_excluded():
    print("\ntest_excluded_cost_centers_stay_excluded")

    excluded = R.OTP_EXCLUDED_COST_CENTERS[0]
    legs = [leg("2026-08-01", 9, 2, shift="UNIT-A"),
            leg("2026-08-01", 10, 45, shift="UNIT-B")]
    by_day, _ = B.bucket_by_day(legs)
    frames = B.build(by_day, date(2026, 8, 1), date(2026, 8, 1),
                     FixedMap({"UNIT-A": "Columbus", "UNIT-B": excluded}))

    day = frames["daily"].iloc[0]
    check(f"the {excluded} leg is not scored", day["legs_scored"] == 1,
          f"got {day['legs_scored']}")
    check("so the day reads 100%, as the daily report would publish it",
          day["on_time_percentage"] == 100.0, f"got {day['on_time_percentage']}")


def test_company_totals_match_the_daily_report():
    """
    Early counts as on time, the same way the daily aggregation counts it.

    If these two ever disagree, the rebuilt series and the live one are
    measuring different things and neither can be compared to the other.
    """
    print("\ntest_company_totals_match_the_daily_report")

    legs = [
        leg("2026-08-01", 9, 2),      # on time
        leg("2026-08-01", 11, 45),    # late
    ]
    early = leg("2026-08-01", 12, 0)
    early["timestamps"] = [{"at_scene": "2026-08-01T11:30:00-04:00"}]  # 30 min early
    legs.append(early)

    cost_centers = FixedMap({"UNIT-A": "Columbus"})
    scored = R.scored_legs(legs, cost_centers)
    totals = B.company_totals(scored)

    check("all three legs scored", totals["scored"] == 3, f"got {totals}")
    check("the early leg is recognised as early", totals["early"] == 1, f"got {totals}")
    check("early counts toward on time", totals["pct"] == 66.67, f"got {totals}")

    grouped = R._otp_aggregate(scored, ["cost_center"])
    check(
        "the ungrouped total agrees with the grouped aggregation",
        int(grouped["total_runs"].sum()) == totals["scored"]
        and float(grouped["on_time_percentage"].iloc[0]) == totals["pct"],
        f"grouped {grouped.to_dict('records')} vs {totals}",
    )

    check("an empty frame does not divide by zero",
          B.company_totals(pd.DataFrame())["pct"] is None)


def test_cli():
    print("\ntest_cli")

    args = B.parse_args(["2026-08-01", "2026-09-15"])
    check("both dates are read", args["start"] == date(2026, 8, 1)
          and args["end"] == date(2026, 9, 15), f"got {args}")

    args = B.parse_args(["2026-08-01"])
    check("the end date defaults to yesterday",
          args["end"] == date.today() - timedelta(days=1), f"got {args}")

    args = B.parse_args(["2026-08-01", "--out", "Reports/Rebuild"])
    check("--out is honoured", args["out"] == "Reports/Rebuild", f"got {args}")

    for bad, why in ((["2026-09-15", "2026-08-01"], "end before start"),
                     ([], "no start date")):
        try:
            B.parse_args(bad)
            check(f"rejects {why}", False, "it was accepted")
        except SystemExit:
            check(f"rejects {why}", True)


def main():
    """The standalone runner, for a machine with no pytest."""
    tests = [
        test_day_bucketing,
        test_window_is_summed_not_averaged,
        test_empty_days_are_visible,
        test_excluded_cost_centers_stay_excluded,
        test_company_totals_match_the_daily_report,
        test_cli,
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
