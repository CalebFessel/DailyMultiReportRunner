"""
Tests for the per-day assignment report.

Three ways this can be quietly wrong, and all three flatter or slander the
station it is about:

  * a day the backfill never reached looks exactly like a day nobody was
    dispatched. Trips reach roughly 90 days and July's first days are right at
    that edge, so this is not hypothetical -- it is the likely case;
  * a truck loaned to another station looks idle unless "assigned elsewhere"
    is its own state. Two states would have somebody asking why Boardman had a
    truck doing nothing on the 12th;
  * the not-assigned list is only as complete as the fleet behind it, and a
    truck that sat still all month ran no legs to be found by.

No credentials needed: the API is only touched inside main().
"""

from datetime import date

import pytest

import vehicle_assignment_by_day as A
import vehicle_service_days as VS


JULY = [date(2026, 7, d) for d in range(1, 8)]


class FakeMap:
    """Stands in for traumasoft_reports.CostCenterMap."""

    def __init__(self, mapping):
        self.mapping = mapping

    def resolve(self, shift_name):
        return self.mapping.get(str(shift_name or "").strip())


MAP = FakeMap({"BOARD-A-1": "Boardman", "BOARD-WC": "Boardman",
               "CIN-A-1": "Cincinnati"})


def leg(vehicle, profile, day):
    return {
        "vehicle_name": vehicle,
        "shift_name": profile,
        "pickup_time": f"2026-07-{day:02d}T09:00:00-04:00",
    }


def grid_for(legs, fleet_names, days=None, centres=("Boardman",)):
    days = days or JULY
    activity, covered, _ = A.legs_by_vehicle_day(legs, set(centres), MAP)
    fleet = {VS.normalize_name(n): "test" for n in fleet_names}
    return A.build_grid(fleet, activity, days, covered, set(centres)), covered


# =============================
# NO DATA IS NOT IDLE
# =============================
def test_a_day_the_backfill_never_reached_is_not_an_idle_day():
    """
    The failure this report exists to avoid. July's first days sit at the edge
    of the 90-day trip window; reading them as "nothing was assigned" would
    invent an idle fleet out of an API limit.
    """
    legs = [leg("B-1", "BOARD-A-1", 5)]
    grid, covered = grid_for(legs, ["B-1"])
    assert covered == {date(2026, 7, 5)}
    row = grid["B1"]
    assert row[date(2026, 7, 5)] == A.ASSIGNED
    assert row[date(2026, 7, 1)] == A.NO_DATA
    assert row[date(2026, 7, 4)] == A.NO_DATA


def test_a_day_with_legs_for_other_stations_is_a_real_zero():
    """
    The API answered for that day; this station simply had nothing. That is a
    genuine idle day and must not be hidden as 'no data'.
    """
    legs = [leg("OTHER-9", "CIN-A-1", 3), leg("B-1", "BOARD-A-1", 5)]
    grid, covered = grid_for(legs, ["B-1"])
    assert date(2026, 7, 3) in covered
    assert grid["B1"][date(2026, 7, 3)] == A.IDLE


def test_coverage_comes_from_every_leg_not_just_this_stations():
    legs = [leg("OTHER-9", "CIN-A-1", 2)]
    _, covered, _ = A.legs_by_vehicle_day(legs, {"Boardman"}, MAP)
    assert covered == {date(2026, 7, 2)}


# =============================
# ASSIGNED ELSEWHERE
# =============================
def test_a_truck_loaned_to_another_station_is_not_idle():
    legs = [leg("B-1", "BOARD-A-1", 1), leg("B-1", "CIN-A-1", 2)]
    grid, _ = grid_for(legs, ["B-1"])
    assert grid["B1"][date(2026, 7, 1)] == A.ASSIGNED
    assert grid["B1"][date(2026, 7, 2)] == A.ELSEWHERE


def test_a_day_with_both_counts_as_assigned_here():
    """It ran for this station that day. That it also ran elsewhere does not
    take the day away."""
    legs = [leg("B-1", "CIN-A-1", 1), leg("B-1", "BOARD-A-1", 1)]
    grid, _ = grid_for(legs, ["B-1"])
    assert grid["B1"][date(2026, 7, 1)] == A.ASSIGNED


def test_a_leg_whose_profile_maps_nowhere_is_counted_as_elsewhere_not_here():
    legs = [leg("B-1", "UNKNOWN-PROFILE", 1)]
    grid, _ = grid_for(legs, ["B-1"])
    assert grid["B1"][date(2026, 7, 1)] == A.ELSEWHERE


def test_unattributable_legs_are_counted_and_reported():
    legs = [leg("B-1", "UNKNOWN-PROFILE", 1), leg("B-1", "BOARD-A-1", 2)]
    _, _, unattributed = A.legs_by_vehicle_day(legs, {"Boardman"}, MAP)
    assert unattributed == 1


# =============================
# THE COST CENTRE NAME
# =============================
def test_an_exact_name_wins():
    matched, how = A.matching_centres("Boardman", {"Boardman", "Boardman Annex"})
    assert matched == ["Boardman"]
    assert how == "exact"


def test_the_legal_entity_wrapper_is_seen_through():
    matched, how = A.matching_centres(
        "Boardman", {"Lynx EMS LLC dba Lynx Boardman"}
    )
    assert matched == ["Lynx EMS LLC dba Lynx Boardman"]
    assert "wrapper" in how


def test_a_substring_match_is_flagged_loudly():
    """
    'Columbus' matches Columbus Ohio and Columbus Indiana alike. Folding them
    together silently is how a station's numbers get somebody else's trucks.
    """
    matched, how = A.matching_centres("Columbus", {"Columbus OH", "Columbus IN"})
    assert set(matched) == {"Columbus OH", "Columbus IN"}
    assert "SUBSTRING" in how


def test_nothing_matching_says_so_rather_than_returning_everything():
    matched, how = A.matching_centres("Nowhere", {"Boardman", "Cincinnati"})
    assert matched == []
    assert how == "no cost center matched"


def test_an_empty_ask_matches_nothing():
    assert A.matching_centres("", {"Boardman"})[0] == []
    assert A.matching_centres("   ", {"Boardman"})[0] == []


def test_case_and_spacing_do_not_decide_the_station():
    matched, _ = A.matching_centres("  boardman  ", {"Boardman"})
    assert matched == ["Boardman"]


# =============================
# THE FLEET
# =============================
def test_a_vehicle_that_ran_here_is_in_the_fleet():
    activity, _, _ = A.legs_by_vehicle_day([leg("B-1", "BOARD-A-1", 1)],
                                           {"Boardman"}, MAP)
    fleet = A.build_fleet(activity, {}, set(), [], {"Boardman"}, MAP)
    assert fleet == {"B1": "ran for it this period"}


def test_a_vehicle_that_only_ran_for_someone_else_is_not_in_the_fleet():
    activity, _, _ = A.legs_by_vehicle_day([leg("X-9", "CIN-A-1", 1)],
                                           {"Boardman"}, MAP)
    fleet = A.build_fleet(activity, {}, set(), [], {"Boardman"}, MAP)
    assert fleet == {}


def test_a_truck_idle_all_period_is_found_by_the_lookback():
    """
    The vehicle management is actually asking about. It ran nothing in July,
    so July's own legs can never find it.
    """
    earlier, _, _ = A.legs_by_vehicle_day([leg("B-2", "BOARD-A-1", 1)],
                                          {"Boardman"}, MAP)
    fleet = A.build_fleet({}, earlier, set(), [], {"Boardman"}, MAP)
    assert fleet == {"B2": "ran for it before the period"}


def test_the_periods_own_legs_beat_the_lookback():
    activity, _, _ = A.legs_by_vehicle_day([leg("B-1", "BOARD-A-1", 1)],
                                           {"Boardman"}, MAP)
    earlier = dict(activity)
    fleet = A.build_fleet(activity, earlier, set(), [], {"Boardman"}, MAP)
    assert fleet["B1"] == "ran for it this period"


def test_a_live_shift_places_a_truck_that_never_ran():
    fleet = A.build_fleet(
        {}, {}, set(),
        [{"name": "B-3", "shift_name": "BOARD-WC"}],
        {"Boardman"}, MAP,
    )
    assert fleet == {"B3": "its live shift says so"}


def test_the_roster_file_catches_what_nothing_else_can():
    fleet = A.build_fleet({}, {}, {"B4"}, [], {"Boardman"}, MAP)
    assert fleet == {"B4": "named in the roster file"}


def test_the_roster_never_downgrades_a_stronger_source():
    activity, _, _ = A.legs_by_vehicle_day([leg("B-1", "BOARD-A-1", 1)],
                                           {"Boardman"}, MAP)
    fleet = A.build_fleet(activity, {}, {"B1"}, [], {"Boardman"}, MAP)
    assert fleet["B1"] == "ran for it this period"


def test_a_roster_file_is_read_from_either_shape(tmp_path):
    import json

    plain = tmp_path / "plain.json"
    plain.write_text(json.dumps(["B-1", "B-2"]))
    assert A.load_roster(str(plain)) == {"B1", "B2"}

    wrapped = tmp_path / "wrapped.json"
    wrapped.write_text(json.dumps({"vehicles": ["B-3"]}))
    assert A.load_roster(str(wrapped)) == {"B3"}


def test_a_missing_roster_is_not_fatal():
    assert A.load_roster("does/not/exist.json") == set()
    assert A.load_roster(None) == set()


def test_a_fleet_vehicle_with_no_legs_reads_as_idle_on_covered_days():
    """The whole point of assembling the fleet: this truck must appear."""
    legs = [leg("B-1", "BOARD-A-1", 1)]
    grid, _ = grid_for(legs, ["B-1", "B-2"])
    assert grid["B2"][date(2026, 7, 1)] == A.IDLE


# =============================
# THE ROLLUPS
# =============================
def names_for(grid):
    return {n: n for n in grid}


def test_the_day_row_names_the_idle_vehicles():
    legs = [leg("B-1", "BOARD-A-1", 1)]
    grid, _ = grid_for(legs, ["B-1", "B-2"], days=[date(2026, 7, 1)])
    rows = A.per_day_rows(grid, [date(2026, 7, 1)], names_for(grid))
    assert rows[0]["assigned"] == 1
    assert rows[0]["not_assigned"] == 1
    assert rows[0]["not_assigned_vehicles"] == "B2"
    assert rows[0]["fleet"] == 2


def test_the_vehicle_row_splits_its_legs_by_who_they_were_for():
    legs = [leg("B-1", "BOARD-A-1", 1), leg("B-1", "BOARD-A-1", 1),
            leg("B-1", "CIN-A-1", 2)]
    days = [date(2026, 7, 1), date(2026, 7, 2)]
    grid, _ = grid_for(legs, ["B-1"], days=days)
    activity, _, _ = A.legs_by_vehicle_day(legs, {"Boardman"}, MAP)
    rows = A.per_vehicle_rows(grid, days, names_for(grid),
                              {"B1": "test"}, activity, {"Boardman"})
    assert rows[0]["legs_for_this_cost_center"] == 2
    assert rows[0]["legs_for_others"] == 1
    assert rows[0]["days_assigned"] == 1
    assert rows[0]["days_assigned_elsewhere"] == 1


def test_days_observed_excludes_the_days_with_no_data():
    legs = [leg("B-1", "BOARD-A-1", 5)]
    grid, _ = grid_for(legs, ["B-1"])
    rows = A.per_vehicle_rows(grid, JULY, names_for(grid), {"B1": "test"},
                              {}, {"Boardman"})
    assert rows[0]["days_no_data"] == 6
    assert rows[0]["days_observed"] == 1


def test_the_matrix_uses_one_mark_per_state():
    # The 3rd needs a leg from somewhere, or it is 'no data' rather than a
    # genuine idle day -- which is exactly the distinction being marked.
    legs = [leg("B-1", "BOARD-A-1", 1), leg("B-1", "CIN-A-1", 2),
            leg("OTHER-9", "CIN-A-1", 3)]
    days = [date(2026, 7, d) for d in (1, 2, 3, 4)]
    grid, _ = grid_for(legs, ["B-1"], days=days)
    row = A.matrix_rows(grid, days, names_for(grid))[0]
    assert row["2026-07-01"] == "A"
    assert row["2026-07-02"] == "E"
    assert row["2026-07-03"] == "."
    assert row["2026-07-04"] == "?"


def test_every_state_has_a_mark():
    assert set(A.MATRIX_MARKS) == {A.ASSIGNED, A.ELSEWHERE, A.IDLE, A.NO_DATA}
    assert len(set(A.MATRIX_MARKS.values())) == 4


# =============================
# NAMES
# =============================
def test_the_live_roster_spelling_wins_over_the_legs():
    legs = [leg("b 1", "BOARD-A-1", 1)]
    names = A.display_names({"B1": "test"}, {}, {},
                            [{"name": "B-1"}], legs)
    assert names["B1"] == "B-1"


def test_a_vehicle_only_the_legs_know_still_gets_a_name():
    legs = [leg("B-7", "BOARD-A-1", 1)]
    names = A.display_names({"B7": "test"}, {}, {}, [], legs)
    assert names["B7"] == "B-7"


def test_a_vehicle_nothing_names_falls_back_to_its_key():
    assert A.display_names({"B9": "test"}, {}, {}, [], [])["B9"] == "B9"


def test_an_unattributable_leg_is_not_counted_as_another_stations_work():
    """
    A leg whose profile is not in the map may well have been this station's.
    Filing it under 'for others' would undercount the days this station had a
    truck out -- the direction that makes a station look idler than it was.
    """
    legs = [leg("B-1", "UNKNOWN-PROFILE", 1), leg("B-1", "CIN-A-1", 2),
            leg("B-1", "BOARD-A-1", 3)]
    days = [date(2026, 7, d) for d in (1, 2, 3)]
    grid, _ = grid_for(legs, ["B-1"], days=days)
    activity, _, _ = A.legs_by_vehicle_day(legs, {"Boardman"}, MAP)
    row = A.per_vehicle_rows(grid, days, names_for(grid), {"B1": "test"},
                             activity, {"Boardman"})[0]
    assert row["legs_for_this_cost_center"] == 1
    assert row["legs_for_others"] == 1
    assert row["legs_unattributed"] == 1


def test_the_never_assigned_list_is_sorted_the_same_way_every_run():
    legs = [leg("OTHER-9", "CIN-A-1", 1)]
    grid, _ = grid_for(legs, ["B-9", "B-1", "B-5"], days=[date(2026, 7, 1)])
    never = sorted((n for n in grid
                    if not any(s == A.ASSIGNED for s in grid[n].values())),
                   key=lambda n: names_for(grid)[n])
    assert never == ["B1", "B5", "B9"]
