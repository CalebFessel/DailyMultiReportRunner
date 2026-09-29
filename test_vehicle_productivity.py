"""
Tests for the vehicle productivity classifier.

The classifier answers a question upper management will act on, so the ways
it can be quietly wrong matter more than the ways it can be right. Three of
them are load-bearing and each has tests here:

  * a factor that could not be read must never read as False. A silent
    telematics gateway is not a parked truck, and an absent ePCR route is not
    an incomplete ePCR;
  * a vehicle-day the three definitions do not cover must not be rounded into
    the nearest one;
  * a shift stamp arrives as bare UTC and the window is tenant-local, so an
    overnight roster lands on the wrong day if that is mishandled.

Import-safe without credentials: the module only constructs API clients
inside main().
"""

from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import pytest

import vehicle_productivity as V


EASTERN = ZoneInfo("America/New_York")
DAY_START = datetime(2026, 9, 22, 0, 0, tzinfo=EASTERN)
DAY_END = DAY_START + timedelta(hours=24)


def factors(**overrides):
    """A vehicle-day that is Productive but for the ePCR nobody can read."""
    base = {
        "engine_on": True,
        "moved": True,
        "has_crew": True,
        "has_scheduled_pickup": True,
        "epcr_complete": None,
        "out_of_service": False,
    }
    base.update(overrides)
    return base


# =============================
# THE RULES AS WRITTEN
# =============================
def test_all_conditions_and_a_completed_epcr_is_productive():
    verdict, _ = V.classify(factors(epcr_complete=True))
    assert verdict == V.PRODUCTIVE


def test_all_conditions_but_an_incomplete_epcr_is_indetermined():
    verdict, _ = V.classify(factors(epcr_complete=False))
    assert verdict == V.INDETERMINATE


def test_no_crew_is_non_productive():
    verdict, reason = V.classify(factors(has_crew=False))
    assert verdict == V.NON_PRODUCTIVE
    assert "crew" in reason


def test_out_of_service_is_non_productive():
    verdict, reason = V.classify(factors(out_of_service=True))
    assert verdict == V.NON_PRODUCTIVE
    assert "out of service" in reason


def test_out_of_service_wins_even_when_everything_else_holds():
    """
    The definition states Non-Productive as an OR, so it is decided first.
    A truck labelled out of service that still ran calls is a data problem
    worth seeing, not a productive truck.
    """
    verdict, _ = V.classify(factors(out_of_service=True, epcr_complete=True))
    assert verdict == V.NON_PRODUCTIVE


def test_non_productive_does_not_need_telematics():
    """Both of its terms come from Traumasoft, so Samsara being down cannot
    stop a vehicle being called Non-Productive."""
    verdict, _ = V.classify(
        factors(engine_on=None, moved=None, has_crew=False)
    )
    assert verdict == V.NON_PRODUCTIVE


# =============================
# UNREADABLE IS NOT FALSE
# =============================
def test_unreadable_epcr_is_undetermined_not_productive():
    verdict, reason = V.classify(factors(epcr_complete=None))
    assert verdict == V.UNDETERMINED
    assert reason.startswith(V.PENDING_EPCR_REASON_PREFIX)


def test_silent_gateway_is_undetermined_not_unclassified():
    verdict, reason = V.classify(factors(engine_on=None, moved=None))
    assert verdict == V.UNDETERMINED
    assert "engine state" in reason and "distance" in reason


def test_a_failed_condition_beats_an_unreadable_one():
    """
    No call was assigned, so no ePCR and no telematics reading could make
    this Productive. Say the decisive thing, not the missing one.
    """
    verdict, reason = V.classify(
        factors(engine_on=None, moved=None, has_scheduled_pickup=False)
    )
    assert verdict == V.UNCLASSIFIED
    assert "no call was assigned" in reason


# =============================
# THE GAP IN THE DEFINITION
# =============================
def test_crewed_moving_truck_with_no_call_matches_no_rule():
    """The moved-but-never-dispatched case: posting, repositioning, a
    maintenance run. It is not Non-Productive as defined -- it has crew and
    is in service -- and it is not Productive or Indetermined either."""
    verdict, reason = V.classify(factors(has_scheduled_pickup=False))
    assert verdict == V.UNCLASSIFIED
    assert "no call was assigned" in reason


def test_crewed_in_service_truck_that_never_started_matches_no_rule():
    verdict, reason = V.classify(factors(engine_on=False, moved=False))
    assert verdict == V.UNCLASSIFIED
    assert "engine never ran" in reason


# =============================
# ENGINE STATE
# =============================
def event(hour, value, minute=0):
    return {
        "time": datetime(2026, 9, 22, hour, minute, tzinfo=EASTERN)
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ"),
        "value": value,
    }


def test_no_events_at_all_is_unknown_not_off():
    assert V.engine_on_in_window([], DAY_START, DAY_END) is None


def test_engine_running_inside_the_window():
    assert V.engine_on_in_window([event(9, "On")], DAY_START, DAY_END) is True


def test_engine_off_all_day():
    assert V.engine_on_in_window(
        [event(3, "Off"), event(15, "Off")], DAY_START, DAY_END
    ) is False


def test_idle_counts_as_engine_on():
    """An EMS unit idles for climate control and equipment power. That is a
    truck in use, not a parked one."""
    assert V.engine_on_in_window([event(2, "Idle")], DAY_START, DAY_END) is True


def test_state_carried_in_across_midnight_counts():
    """A truck already running at midnight was running in the window even if
    the gateway says nothing until morning."""
    before = {
        "time": (DAY_START - timedelta(hours=3))
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ"),
        "value": "On",
    }
    assert V.engine_on_in_window([before], DAY_START, DAY_END) is True


def test_state_carried_in_off_is_still_a_reading():
    before = {
        "time": (DAY_START - timedelta(hours=3))
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ"),
        "value": "Off",
    }
    assert V.engine_on_in_window([before], DAY_START, DAY_END) is False


def test_an_unrecognised_state_counts_as_on():
    """Only "off" means not in use. Anything else is treated as running
    rather than silently dropped into the parked bucket."""
    assert V.engine_on_in_window(
        [event(9, "SomethingNew")], DAY_START, DAY_END
    ) is True


def test_events_after_the_window_do_not_count():
    late = {
        "time": (DAY_END + timedelta(hours=2))
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ"),
        "value": "On",
    }
    assert V.engine_on_in_window([late], DAY_START, DAY_END) is None


# =============================
# DISTANCE
# =============================
def odo(hour, meters, key="obdOdometerMeters"):
    return key, {
        "time": datetime(2026, 9, 22, hour, tzinfo=EASTERN)
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ"),
        "value": meters,
    }


def odo_row(*entries):
    row = {}
    for key, entry in entries:
        row.setdefault(key, []).append(entry)
    return row


def test_one_odometer_reading_says_nothing_about_the_window():
    """One reading says where the odometer stood, not that it stayed there.
    Zero would be a claim the data does not support."""
    row = odo_row(odo(9, 1_000_000))
    assert V.miles_from_odometer(row, DAY_START, DAY_END) is None


def test_odometer_delta_in_miles():
    row = odo_row(odo(6, 1_000_000), odo(18, 1_016_093))
    miles = V.miles_from_odometer(row, DAY_START, DAY_END)
    assert miles == pytest.approx(10.0, abs=0.01)


def test_odometer_baseline_is_the_last_reading_before_midnight():
    """A truck that moved before its first in-window sample is not credited
    from zero."""
    before = ("obdOdometerMeters", {
        "time": (DAY_START - timedelta(hours=2))
        .astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "value": 1_000_000,
    })
    row = odo_row(before, odo(9, 1_008_047))
    miles = V.miles_from_odometer(row, DAY_START, DAY_END)
    assert miles == pytest.approx(5.0, abs=0.01)


def test_the_two_odometer_series_are_never_mixed():
    """OBD and GPS odometers are different instruments with different
    absolute values; a delta across both is noise, not distance."""
    row = odo_row(odo(6, 1_000_000), odo(18, 1_001_609))
    row["gpsOdometerMeters"] = [
        {"time": "2026-09-22T12:00:00Z", "value": 9_000_000}
    ]
    miles = V.miles_from_odometer(row, DAY_START, DAY_END)
    assert miles == pytest.approx(1.0, abs=0.01)


def test_no_odometer_at_all_is_none():
    assert V.miles_from_odometer({}, DAY_START, DAY_END) is None


def fix(hour, lat, lon):
    return {
        "time": datetime(2026, 9, 22, hour, tzinfo=EASTERN)
        .astimezone(timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ"),
        "latitude": lat,
        "longitude": lon,
    }


def test_gps_track_length():
    # One degree of latitude is about 69 miles.
    row = {"gps": [fix(6, 40.0, -75.0), fix(7, 41.0, -75.0)]}
    miles = V.miles_from_gps(row, DAY_START, DAY_END)
    assert miles == pytest.approx(69.0, abs=0.5)


def test_one_gps_fix_is_not_zero_miles():
    row = {"gps": [fix(6, 40.0, -75.0)]}
    assert V.miles_from_gps(row, DAY_START, DAY_END) is None


def test_gps_fixes_outside_the_window_are_ignored():
    row = {"gps": [fix(6, 40.0, -75.0)]}
    row["gps"].append({
        "time": (DAY_END + timedelta(hours=1))
        .astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "latitude": 45.0, "longitude": -75.0,
    })
    assert V.miles_from_gps(row, DAY_START, DAY_END) is None


def test_distance_source_reports_what_it_used():
    row = odo_row(odo(6, 1_000_000), odo(18, 1_003_219))
    miles, source = V.distance_miles(row, DAY_START, DAY_END, "odometer")
    assert source == "odometer" and miles == pytest.approx(2.0, abs=0.01)
    assert V.distance_miles({}, DAY_START, DAY_END, "odometer") == (None, "none")


# =============================
# CREW
# =============================
def shift(vehicle_name, start, end, punched=False, deleted=False):
    """A shift row as Traumasoft returns it: bare UTC, no offset."""
    return {
        "vehicle_name": vehicle_name,
        "start_time": start,
        "end_time": end,
        "deleted": deleted,
        "punches": [{"start_time": start}] if punched else [],
    }


def test_two_medics_on_one_truck_are_two_crew_on_one_unit():
    shifts = [
        shift("A-101", "2026-09-22T12:00:00", "2026-09-23T00:00:00"),
        shift("A-101", "2026-09-22T12:00:00", "2026-09-23T00:00:00"),
    ]
    crew = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN)
    assert crew["A101"]["crew"] == 2
    assert crew["A101"]["shifts"] == 2


def test_shift_times_are_read_as_utc_not_local():
    """
    08:00 bare-UTC is 04:00 Eastern -- inside the 22nd. Read as local it would
    be 08:00 on the 22nd, which happens to land in the same day here; the
    overnight case below is the one that actually separates them.
    """
    shifts = [shift("A-101", "2026-09-22T08:00:00", "2026-09-22T20:00:00")]
    crew = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN)
    assert crew["A101"]["crew"] == 1


def test_an_overnight_shift_lands_on_the_right_day():
    """
    20:00 to 08:00 bare-UTC is 16:00 on the 21st to 04:00 on the 22nd Eastern.
    It overlaps the 22nd. Read as tenant-local it would run 20:00 on the 21st
    to 08:00 on the 22nd -- also overlapping, so the discriminating case is a
    shift that only overlaps under one reading.
    """
    shifts = [shift("A-101", "2026-09-22T01:00:00", "2026-09-22T03:00:00")]
    # Bare UTC: 21:00-23:00 Eastern on the 21st. Does NOT touch the 22nd.
    crew = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN)
    assert "A101" not in crew
    # Read as local it would sit squarely inside the 22nd.
    local = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN,
                              shifts_are_utc=False)
    assert local["A101"]["crew"] == 1


def test_a_deleted_shift_is_not_crew():
    shifts = [shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00",
                    deleted=True)]
    assert V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN) == {}


def test_a_deleted_flag_of_string_zero_is_not_deleted():
    shifts = [shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00")]
    shifts[0]["deleted"] = "0"
    crew = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN)
    assert crew["A101"]["crew"] == 1


def test_punched_is_carried_alongside_assigned_not_instead_of_it():
    """The definition says assigned. A crew that no-showed is a different
    finding from a truck nobody was rostered to."""
    shifts = [
        shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00", punched=True),
        shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00", punched=False),
    ]
    crew = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN)
    assert crew["A101"] == {"crew": 2, "punched": 1, "shifts": 2}


def test_a_shift_with_no_vehicle_contributes_to_nothing():
    shifts = [shift("", "2026-09-22T12:00:00", "2026-09-22T20:00:00")]
    assert V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN) == {}


def test_unit_names_that_differ_only_in_punctuation_are_one_unit():
    shifts = [
        shift(" A-101 ", "2026-09-22T12:00:00", "2026-09-22T20:00:00"),
        shift("A101", "2026-09-22T12:00:00", "2026-09-22T20:00:00"),
    ]
    crew = V.crew_by_vehicle(shifts, DAY_START, DAY_END, EASTERN)
    assert crew["A101"]["crew"] == 2


# =============================
# CALLS
# =============================
def test_a_cancelled_call_still_had_a_scheduled_pickup():
    legs = [
        {"vehicle_id": 7, "trip_status": "Transported"},
        {"vehicle_id": 7, "trip_status": "Canceled"},
    ]
    counts = V.legs_by_vehicle(legs)
    assert counts["7"] == {"legs": 2, "cancelled": 1}


def test_cancelled_calls_can_be_excluded():
    legs = [{"vehicle_id": 7, "trip_status": "Canceled"}]
    assert V.legs_by_vehicle(legs, exclude_cancelled=True) == {}


def test_a_leg_with_no_vehicle_is_nobody_s_call():
    assert V.legs_by_vehicle([{"vehicle_id": None}]) == {}


# =============================
# OUT OF SERVICE
# =============================
def test_out_of_service_status():
    assert V.out_of_service({"vehicle_status": "Out of Service"}) is True
    assert V.out_of_service({"vehicle_status": "Out of Service - Collision"}) is True
    assert V.out_of_service({"vehicle_status": "In Service"}) is False


def test_disabled_or_deleted_counts_as_out_of_service():
    assert V.out_of_service({"vehicle_status": "In Service", "disabled": "1"}) is True
    assert V.out_of_service({"vehicle_status": "In Service", "deleted": True}) is True


def test_a_string_zero_flag_does_not_take_a_truck_out_of_service():
    assert V.out_of_service(
        {"vehicle_status": "In Service", "disabled": "0", "deleted": "0"}
    ) is False


# =============================
# THE SAMSARA JOIN
# =============================
def test_vin_is_preferred_over_name():
    ts = [{"id": 1, "name": "A-101", "vin": "1FDUF5HT7KDA12345"}]
    sam = [
        {"id": "s1", "name": "Totally Different", "vin": "1fduf5ht7kda12345"},
        {"id": "s2", "name": "A101", "vin": "1FDUF5HT7KDA99999"},
    ]
    index, _ = V.build_samsara_index(ts, sam)
    assert index["1"] == ("s1", "vin")


def test_name_is_the_fallback_when_traumasoft_has_no_vin():
    ts = [{"id": 1, "name": "A-101", "vin": ""}]
    sam = [{"id": "s2", "name": "A 101", "vin": "1FDUF5HT7KDA99999"}]
    index, _ = V.build_samsara_index(ts, sam)
    assert index["1"] == ("s2", "name")


def test_a_short_vin_is_not_a_vin():
    """A 17-character VIN or nothing. A placeholder joins on nothing."""
    ts = [{"id": 1, "name": "A-101", "vin": "N/A"}]
    sam = [{"id": "s1", "name": "B-202", "vin": "N/A"}]
    index, _ = V.build_samsara_index(ts, sam)
    assert index == {}


def test_an_ambiguous_match_is_left_unmatched_and_reported():
    ts = [{"id": 1, "name": "A-101", "vin": ""}]
    sam = [
        {"id": "s1", "name": "A-101", "vin": ""},
        {"id": "s2", "name": "A101", "vin": ""},
    ]
    index, ambiguous = V.build_samsara_index(ts, sam)
    assert index == {}
    assert len(ambiguous) == 1


# =============================
# END TO END
# =============================
def vehicle(vid, name, status="In Service", vin=""):
    return {"id": vid, "name": name, "vehicle_status": status, "vin": vin}


def test_no_samsara_leaves_a_crewed_dispatched_truck_undetermined():
    """
    Samsara down must not turn working trucks into Non-Productive ones, and
    must not hide the Traumasoft half of the answer either.
    """
    rows, _ = V.build_rows(
        [vehicle(7, "A-101")],
        [{"vehicle_id": 7, "trip_status": "Transported"}],
        [shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00")],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(),
    )
    assert rows[0]["classification"] == V.UNDETERMINED
    assert rows[0]["samsara_join"] == "unmatched"
    assert rows[0]["crew_assigned"] == 1
    assert rows[0]["legs_assigned"] == 1


def test_an_uncrewed_truck_is_non_productive_without_any_telematics():
    rows, _ = V.build_rows(
        [vehicle(7, "A-101")], [], [],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(),
    )
    assert rows[0]["classification"] == V.NON_PRODUCTIVE
    assert rows[0]["reason"] == "no crew assigned"


def test_retired_vehicles_are_not_a_productivity_question():
    rows, _ = V.build_rows(
        [vehicle(7, "A-101", status="Retired")], [], [],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(),
    )
    assert rows == []
    rows, _ = V.build_rows(
        [vehicle(7, "A-101", status="Retired")], [], [],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(), include_non_fleet=True,
    )
    assert len(rows) == 1


def test_everything_but_epcr_reads_as_pending_not_productive():
    stats = {
        "id": "s1",
        "engineStates": [event(9, "On")],
        "obdOdometerMeters": [odo(6, 1_000_000)[1], odo(18, 1_016_093)[1]],
    }
    rows, _ = V.build_rows(
        [vehicle(7, "A-101")],
        [{"vehicle_id": 7, "trip_status": "Transported"}],
        [shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00")],
        stats_rows=[stats], sam_index={"7": ("s1", "vin")},
        window_start=DAY_START, window_end=DAY_END, zone=EASTERN,
        day=DAY_START.date(),
    )
    row = rows[0]
    assert row["classification"] == V.UNDETERMINED
    assert row["reason"].startswith(V.PENDING_EPCR_REASON_PREFIX)
    assert row["engine_on"] is True
    assert row["moved"] is True
    assert row["miles"] == pytest.approx(10.0, abs=0.01)
    assert row["epcr_complete"] is None


def test_a_truck_that_moved_under_a_mile_did_not_move():
    stats = {
        "id": "s1",
        "engineStates": [event(9, "On")],
        "obdOdometerMeters": [odo(6, 1_000_000)[1], odo(18, 1_000_800)[1]],
    }
    rows, _ = V.build_rows(
        [vehicle(7, "A-101")],
        [{"vehicle_id": 7, "trip_status": "Transported"}],
        [shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00")],
        stats_rows=[stats], sam_index={"7": ("s1", "vin")},
        window_start=DAY_START, window_end=DAY_END, zone=EASTERN,
        day=DAY_START.date(),
    )
    assert rows[0]["moved"] is False
    assert rows[0]["classification"] == V.UNCLASSIFIED


def test_a_shift_naming_an_unknown_unit_is_reported_not_dropped():
    _, orphans = V.build_rows(
        [vehicle(7, "A-101")], [],
        [shift("Z-999", "2026-09-22T12:00:00", "2026-09-22T20:00:00")],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(),
    )
    assert orphans == ["Z999"]


def test_observed_history_overrides_todays_status_and_says_so():
    """
    Traumasoft reports one current status per vehicle. For a past day that is
    an anachronism, so where accumulated observations cover the vehicle they
    win -- and the row records which answer it used.
    """
    rows, _ = V.build_rows(
        [vehicle(7, "A-101", status="In Service")], [],
        [shift("A-101", "2026-09-22T12:00:00", "2026-09-22T20:00:00")],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(), oos_history={"7": True},
    )
    assert rows[0]["out_of_service"] is True
    assert rows[0]["out_of_service_source"] == "observed history"
    assert rows[0]["classification"] == V.NON_PRODUCTIVE


def test_without_history_the_row_admits_it_used_todays_status():
    rows, _ = V.build_rows(
        [vehicle(7, "A-101")], [], [],
        stats_rows=[], sam_index={}, window_start=DAY_START, window_end=DAY_END,
        zone=EASTERN, day=DAY_START.date(),
    )
    assert rows[0]["out_of_service_source"] == "current status"


# =============================
# THE WINDOW
# =============================
def test_the_window_is_tenant_local_not_utc():
    _, _, start, end = V.day_window(DAY_START.date(), EASTERN)
    assert start == DAY_START
    assert end == DAY_END


def test_the_fetch_reaches_back_further_than_the_window():
    fetch_from, _, start, _ = V.day_window(DAY_START.date(), EASTERN,
                                           lookback_hours=24)
    assert V.parse_ts(fetch_from) == start - timedelta(hours=24)


def test_a_narrower_x_window_shortens_the_window_only():
    fetch_from, fetch_to, start, end = V.day_window(
        DAY_START.date(), EASTERN, window_hours=12
    )
    assert end - start == timedelta(hours=12)
    assert V.parse_ts(fetch_from) == start - timedelta(hours=24)
