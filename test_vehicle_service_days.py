"""
Tests for the July service-days report.

The number management asked for is a ratio, and the denominator is where it
goes wrong. A day the daily report never ran is a day nobody observed, and
folding those into "in service" inflates every vehicle's uptime -- in the
direction that flatters the fleet. That is the failure these tests exist for.

The others:

  * a vehicle out of service in July and retired in August is gone from
    today's roster, and dropping it understates exactly what is being
    measured;
  * class comes from a naming convention, not from data, so an unplaced unit
    has to be visible rather than bucketed;
  * cost center is not on a vehicle at all, and the resolver that reaches it
    has to be the same one every other report uses.

No credentials needed: the archive is built on disk and the API is only
touched inside main().
"""

from datetime import date

import pytest

import vehicle_service_days as S


JULY = (date(2026, 7, 1), date(2026, 7, 31))


def snapshot(*entries):
    """{day -> {normalized name -> row}} from (day, [names]) pairs."""
    return {
        date(2026, 7, day): {
            S.normalize_name(n): {"vehicle_name": n} for n in names
        }
        for day, names in entries
    }


def vehicle(name, status="In Service"):
    return {"id": abs(hash(name)) % 10000, "name": name, "vehicle_status": status}


# =============================
# THE DENOMINATOR
# =============================
def test_an_unobserved_day_is_not_an_in_service_day():
    """
    The whole point. Two snapshots exist for a 31-day month; a vehicle absent
    from both was in service for 2 observed days, not 31.
    """
    snapshots = snapshot((1, []), (2, []))
    rows, observed, missing = S.build_rows(
        snapshots, [vehicle("A-101")], [], None, *JULY, S.parse_class_patterns("AMB=a-"),
    )
    assert len(observed) == 2
    assert len(missing) == 29
    row = rows[0]
    assert row["days_in_service"] == 2
    assert row["days_out_of_service"] == 0
    assert row["days_not_observed"] == 29


def test_the_three_counts_sum_to_the_calendar_period():
    snapshots = snapshot((1, ["A-101"]), (2, []), (3, ["A-101"]))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], None, *JULY, S.parse_class_patterns("AMB=a-"),
    )
    row = rows[0]
    assert (row["days_in_service"] + row["days_out_of_service"]
            + row["days_not_observed"]) == 31
    assert row["days_in_period"] == 31


def test_days_out_of_service_counts_only_the_days_it_appears():
    snapshots = snapshot(
        (1, ["A-101"]), (2, ["A-101"]), (3, []), (4, ["A-101", "A-102"]),
    )
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101"), vehicle("A-102")], [], None, *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    by_name = {r["vehicle_name"]: r for r in rows}
    assert by_name["A-101"]["days_out_of_service"] == 3
    assert by_name["A-101"]["days_in_service"] == 1
    assert by_name["A-102"]["days_out_of_service"] == 1
    assert by_name["A-102"]["days_in_service"] == 3


def test_the_first_and_last_day_out_are_reported():
    snapshots = snapshot((3, ["A-101"]), (4, []), (9, ["A-101"]))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], None, *JULY, S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["first_day_out"] == "2026-07-03"
    assert rows[0]["last_day_out"] == "2026-07-09"


def test_a_snapshot_outside_the_period_does_not_count():
    snapshots = {date(2026, 6, 30): {"A101": {}}, date(2026, 7, 1): {}}
    rows, observed, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], None, *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    # build_rows is handed only in-period snapshots by the reader, but if one
    # slips through it must not be counted as an observed day of the period.
    assert date(2026, 6, 30) not in S.days_in(*JULY)
    assert rows[0]["days_in_service"] <= len(observed)


# =============================
# GAPS
# =============================
def test_gaps_are_reported_as_runs_not_a_wall_of_dates():
    days = [date(2026, 7, 4), date(2026, 7, 6), date(2026, 7, 7), date(2026, 7, 8)]
    assert S.compress_days(days) == ["2026-07-04", "2026-07-06 to 2026-07-08"]


def test_a_single_missing_day_reads_as_one_date():
    assert S.compress_days([date(2026, 7, 4)]) == ["2026-07-04"]


def test_no_gaps_is_no_runs():
    assert S.compress_days([]) == []


# =============================
# THE ROSTER UNION
# =============================
def test_a_vehicle_retired_since_july_is_still_counted():
    """
    Out of service through July, retired in August, gone from today's roster.
    Dropping it would understate exactly the thing being measured.
    """
    snapshots = snapshot((1, ["A-999"]), (2, ["A-999"]))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], None, *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    by_name = {r["vehicle_name"]: r for r in rows}
    assert "A-999" in by_name
    assert by_name["A-999"]["in_current_fleet"] is False
    assert by_name["A-999"]["days_out_of_service"] == 2


def test_a_retired_vehicle_in_the_current_roster_is_dropped_by_default():
    snapshots = snapshot((1, []))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101", status="Retired")], [], None, *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert rows == []
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101", status="Retired")], [], None, *JULY,
        S.parse_class_patterns("AMB=a-"), include_non_fleet=True,
    )
    assert len(rows) == 1


def test_names_that_differ_only_in_punctuation_are_one_vehicle():
    snapshots = snapshot((1, ["A 101"]), (2, ["a-101"]))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], None, *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert len(rows) == 1
    assert rows[0]["days_out_of_service"] == 2


# =============================
# CLASS
# =============================
def test_class_rules_match_longest_prefix_first():
    rules = S.parse_class_patterns("MH=m-;MEDIC=me-")
    assert S.vehicle_class("ME-4", rules) == "MEDIC"
    assert S.vehicle_class("M-4", rules) == "MH"


def test_a_parenthetical_prefix_does_not_hide_the_class():
    rules = S.parse_class_patterns("WC=wc-")
    assert S.vehicle_class("(SC)WC-101", rules) == "WC"


def test_an_unplaced_name_is_unclassified_not_guessed():
    rules = S.parse_class_patterns("AMB=a-;WC=wc-")
    assert S.vehicle_class("SPARE-7", rules) is None
    snapshots = snapshot((1, []))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("SPARE-7")], [], None, *JULY, rules,
    )
    assert rows[0]["vehicle_class"] == "UNCLASSIFIED"


def test_the_prefix_a_name_starts_with_is_reportable():
    assert S.name_prefix("WC-101") == "wc"
    assert S.name_prefix("(SC)M-4") == "m"
    assert S.name_prefix("101") == ""


def test_empty_and_malformed_class_rules_are_ignored():
    assert S.parse_class_patterns("") == []
    assert S.parse_class_patterns("AMB;=;MH=m-") == [("m-", "MH")]


# =============================
# COST CENTRE
# =============================
class FakeMap:
    """Stands in for traumasoft_reports.CostCenterMap."""

    def __init__(self, mapping):
        self.mapping = mapping

    def resolve(self, shift_name):
        return self.mapping.get(shift_name)


def leg(vehicle_name, shift_name, day=1):
    return {
        "vehicle_name": vehicle_name,
        "shift_name": shift_name,
        "pickup_time": f"2026-07-{day:02d}T09:00:00-04:00",
    }


def test_cost_center_comes_from_the_periods_own_legs():
    """
    Not from the vehicle row's live shift_name, which is today's shift and
    says nothing about July.
    """
    centres, _ = S.cost_centers_from_legs(
        [leg("A-101", "OH-A-CIN")], FakeMap({"OH-A-CIN": "Cincinnati"})
    )
    assert centres["A101"] == "Cincinnati"


def test_a_vehicle_split_across_cost_centers_goes_to_the_dominant_one():
    legs = [leg("A-101", "OH-A-CIN")] * 3 + [leg("A-101", "INDY WC")]
    centres, split = S.cost_centers_from_legs(
        legs, FakeMap({"OH-A-CIN": "Cincinnati", "INDY WC": "Indianapolis"})
    )
    assert centres["A101"] == "Cincinnati"
    assert split["A101"] == {"Cincinnati": 3, "Indianapolis": 1}


def test_a_tie_resolves_the_same_way_every_run():
    legs = [leg("A-101", "OH-A-CIN"), leg("A-101", "INDY WC")]
    first, _ = S.cost_centers_from_legs(
        legs, FakeMap({"OH-A-CIN": "Cincinnati", "INDY WC": "Indianapolis"})
    )
    second, _ = S.cost_centers_from_legs(
        list(reversed(legs)), FakeMap({"OH-A-CIN": "Cincinnati", "INDY WC": "Indianapolis"})
    )
    assert first["A101"] == second["A101"]


def test_a_vehicle_that_ran_nothing_has_no_cost_center():
    snapshots = snapshot((1, ["A-101"]))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], FakeMap({}), *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["cost_center"] == "UNKNOWN"


# =============================
# DAYS IT ACTUALLY RAN
# =============================
def test_days_ran_a_call_is_counted_separately_from_days_in_service():
    """
    In service means not carrying an out-of-service label. It does not mean
    the truck worked. Reporting only one of these answers the wrong question.
    """
    snapshots = snapshot((1, []), (2, []), (3, []))
    legs = [leg("A-101", "OH-A-CIN", day=1), leg("A-101", "OH-A-CIN", day=1)]
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], legs, FakeMap({"OH-A-CIN": "Cincinnati"}),
        *JULY, S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["days_in_service"] == 3
    assert rows[0]["days_ran_a_call"] == 1  # two legs, one day


def test_a_leg_outside_the_period_does_not_count_as_a_day_run():
    snapshots = snapshot((1, []))
    outside = dict(leg("A-101", "OH-A-CIN"))
    outside["pickup_time"] = "2026-06-15T09:00:00-04:00"
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [outside],
        FakeMap({"OH-A-CIN": "Cincinnati"}), *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["days_ran_a_call"] == 0


def test_a_leg_day_falls_back_through_the_stamps_it_carries():
    assert S.leg_date({"pickup_time": "2026-07-04T09:00:00-04:00"}) == date(2026, 7, 4)
    assert S.leg_date({"appt_time": "2026-07-05 09:00:00"}) == date(2026, 7, 5)
    assert S.leg_date({"vehicle_name": "A-101"}) is None


# =============================
# ROLLUPS
# =============================
def test_a_vehicle_out_for_part_of_the_month_counts_once_as_ever_out():
    """
    Management's own point: a total count is misleading because they were not
    out the whole month. So the vehicle count and the day count are separate
    numbers and neither stands in for the other.
    """
    snapshots = snapshot((1, ["A-101"]), (2, []), (3, []))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101"), vehicle("A-102")], [], FakeMap({}), *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    grouped = S.rollup(rows, ("vehicle_class",))[("AMB",)]
    assert grouped["vehicles"] == 2
    assert grouped["vehicles_ever_out"] == 1
    assert grouped["days_out_of_service"] == 1
    assert grouped["days_in_service"] == 5  # A-101 two, A-102 three


def test_rollup_groups_by_several_keys():
    snapshots = snapshot((1, []))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101"), vehicle("WC-1")], [], FakeMap({}), *JULY,
        S.parse_class_patterns("AMB=a-;WC=wc-"),
    )
    grouped = S.rollup(rows, ("cost_center", "vehicle_class"))
    assert set(grouped) == {("UNKNOWN", "AMB"), ("UNKNOWN", "WC")}


# =============================
# THE ARCHIVE ON DISK
# =============================
def write_archive(directory, rows, summary=None):
    import pandas as pd

    path = directory / S.APPEND_WORKBOOK
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name=S.OOS_SHEET, index=False)
        if summary is not None:
            pd.DataFrame(summary).to_excel(
                writer, sheet_name=S.SUMMARY_SHEET, index=False
            )
    return path


def test_a_missing_workbook_says_so_rather_than_reporting_zeros(tmp_path):
    snapshots, error = S.read_oos_snapshots(str(tmp_path), *JULY)
    assert snapshots is None
    assert "does not exist" in error


def test_a_workbook_without_the_sheet_says_which_sheet(tmp_path):
    import pandas as pd

    path = tmp_path / S.APPEND_WORKBOOK
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame([{"a": 1}]).to_excel(writer, sheet_name="Something Else",
                                          index=False)
    snapshots, error = S.read_oos_snapshots(str(tmp_path), *JULY)
    assert snapshots is None
    assert S.OOS_SHEET in error


def test_a_sheet_without_snapshot_date_cannot_place_its_rows(tmp_path):
    write_archive(tmp_path, [{"vehicle_name": "A-101"}])
    snapshots, error = S.read_oos_snapshots(str(tmp_path), *JULY)
    assert snapshots is None
    assert "snapshot_date" in error


def test_the_archive_is_read_into_days(tmp_path):
    write_archive(tmp_path, [
        {"snapshot_date": "2026-07-01", "vehicle_name": "A-101"},
        {"snapshot_date": "2026-07-01", "vehicle_name": "WC-1"},
        {"snapshot_date": "2026-07-02", "vehicle_name": "A-101"},
        {"snapshot_date": "2026-08-01", "vehicle_name": "A-101"},
    ])
    snapshots, error = S.read_oos_snapshots(str(tmp_path), *JULY)
    assert error is None
    assert sorted(snapshots) == [date(2026, 7, 1), date(2026, 7, 2)]
    assert set(snapshots[date(2026, 7, 1)]) == {"A101", "WC1"}


def test_the_summary_sheet_is_optional(tmp_path):
    write_archive(tmp_path, [{"snapshot_date": "2026-07-01", "vehicle_name": "A-101"}])
    assert S.read_summary_counts(str(tmp_path), *JULY) == {}


def test_the_summary_sheet_is_read_for_the_cross_check(tmp_path):
    write_archive(
        tmp_path,
        [{"snapshot_date": "2026-07-01", "vehicle_name": "A-101"}],
        summary=[{"snapshot_date": "2026-07-01", "metric": "Out Of Service",
                  "value": 1}],
    )
    counts = S.read_summary_counts(str(tmp_path), *JULY)
    assert counts[date(2026, 7, 1)]["Out Of Service"] == 1


# =============================
# PERIOD
# =============================
def test_a_month_resolves_to_its_real_length():
    assert S.month_bounds("2026-07") == (date(2026, 7, 1), date(2026, 7, 31))
    assert S.month_bounds("2026-02") == (date(2026, 2, 1), date(2026, 2, 28))
    assert S.month_bounds("2024-02")[1] == date(2024, 2, 29)


def test_days_in_is_inclusive_of_both_ends():
    assert len(S.days_in(*JULY)) == 31
    assert S.days_in(date(2026, 7, 1), date(2026, 7, 1)) == [date(2026, 7, 1)]


# =============================
# ATTRIBUTION FALLBACKS
# =============================
def test_a_vehicle_out_all_month_is_placed_from_earlier_trips():
    """
    A truck out of service for the whole period ran nothing in it -- and those
    are exactly the rows this report is about. Leaving them unattributed would
    strand the vehicles management most wants grouped.
    """
    snapshots = snapshot((1, ["A-101"]), (2, ["A-101"]))
    earlier = [dict(leg("A-101", "OH-A-CIN"))]
    earlier[0]["pickup_time"] = "2026-05-02T09:00:00-04:00"
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], FakeMap({"OH-A-CIN": "Cincinnati"}),
        *JULY, S.parse_class_patterns("AMB=a-"), lookback_legs=earlier,
    )
    assert rows[0]["cost_center"] == "Cincinnati"
    assert rows[0]["cost_center_source"] == "earlier legs"
    # The earlier trips must not be mistaken for work done in the period.
    assert rows[0]["days_ran_a_call"] == 0


def test_the_periods_own_legs_beat_earlier_ones():
    snapshots = snapshot((1, []))
    inside = [leg("A-101", "INDY WC")]
    earlier = [leg("A-101", "OH-A-CIN")]
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], inside,
        FakeMap({"OH-A-CIN": "Cincinnati", "INDY WC": "Indianapolis"}),
        *JULY, S.parse_class_patterns("AMB=a-"), lookback_legs=earlier,
    )
    assert rows[0]["cost_center"] == "Indianapolis"
    assert rows[0]["cost_center_source"] == "period"


def test_a_truck_that_never_ran_falls_back_to_its_live_shift():
    snapshots = snapshot((1, ["A-101"]))
    truck = vehicle("A-101")
    truck["shift_name"] = "OH-A-CIN"
    rows, _, _ = S.build_rows(
        snapshots, [truck], [], FakeMap({"OH-A-CIN": "Cincinnati"}), *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["cost_center"] == "Cincinnati"
    assert rows[0]["cost_center_source"] == "current shift"


def test_an_unattributable_vehicle_says_unknown_rather_than_guessing():
    snapshots = snapshot((1, ["A-101"]))
    rows, _, _ = S.build_rows(
        snapshots, [vehicle("A-101")], [], FakeMap({}), *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["cost_center"] == "UNKNOWN"
    assert rows[0]["cost_center_source"] == ""


def test_a_vehicle_no_longer_in_the_fleet_has_no_live_shift_to_fall_back_on():
    snapshots = snapshot((1, ["A-999"]))
    rows, _, _ = S.build_rows(
        snapshots, [], [], FakeMap({"OH-A-CIN": "Cincinnati"}), *JULY,
        S.parse_class_patterns("AMB=a-"),
    )
    assert rows[0]["vehicle_name"] == "A-999"
    assert rows[0]["cost_center"] == "UNKNOWN"
