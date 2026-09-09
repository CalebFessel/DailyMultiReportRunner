"""
Tests for the Traumasoft -> Samsara route mapping.

Everything here is pure: fixture legs in, payloads out, no token and no
network. Run with `python -m pytest test_samsara_routes.py`.

The cases that matter are the ones where a silent wrong answer is plausible:
the unit-prefix join, drop-off times that would sort ahead of their own
pickup, and legs that must not reach Samsara at all.
"""

import json
from datetime import date, timedelta

import pytest

import samsara_routes as SR
from traumasoft_reports import parse_ts_aware


# =============================
# FIXTURES
# =============================
def leg(**overrides):
    """A routable leg. Override only what a test is about."""
    base = {
        "leg_id": 1,
        "run_number": "24-10871",
        "call_type": "BLS Transport",
        "los": "BLS",
        "vehicle_name": "M-12",
        "pickup_time": "2026-09-10T07:30:00-04:00",
        "appt_time": "2026-09-10T08:05:00-04:00",
        "deleted": "false",
        "pu_facility_name": "Riverside Manor",
        "pu_address1": "100 River Rd",
        "pu_city": "Charleston",
        "pu_state": "WV",
        "pu_zipcode": "25301",
        "pu_lat": 38.3498,
        "pu_lon": -81.6326,
        "do_facility_name": "St. Mary's Dialysis",
        "do_address1": "2900 First Ave",
        "do_city": "Huntington",
        "do_state": "WV",
        "do_zipcode": "25702",
        "do_lat": 38.4192,
        "do_lon": -82.4452,
        "patient_first_name": "Jane",
        "patient_last_name": "Doe",
        "response_priority": "3",
        "transport_priority": "3",
    }
    base.update(overrides)
    return base


def vehicle(vid, name):
    return {"id": vid, "name": name}


# =============================
# UNIT PREFIX
# =============================
@pytest.mark.parametrize(
    "name,expected",
    [
        ("M-12", "M12"),
        ("M12", "M12"),
        ("m 12", "M12"),
        ("M_12", "M12"),
        ("M-12 Ford E450", "M12"),
        ("M-012", "M12"),          # leading zeros collapse
        ("WC-3", "WC3"),
        ("Medic 12", "MEDIC12"),
        ("M12A", "M12A"),          # a suffix letter is part of the unit
        ("M12B", "M12B"),
        ("  M-12  ", "M12"),
    ],
)
def test_unit_prefix_normalises_the_common_spellings(name, expected):
    assert SR.unit_prefix(name) == expected


def test_unit_prefix_keeps_suffixed_units_distinct():
    """M12A and M12B are two ambulances, not one typed twice."""
    assert SR.unit_prefix("M-12A") != SR.unit_prefix("M-12B")


def test_unit_prefix_falls_back_to_first_token_without_a_designator():
    assert SR.unit_prefix("Supervisor Truck") == "SUPERVISOR"


def test_unit_prefix_handles_empty():
    assert SR.unit_prefix(None) is None
    assert SR.unit_prefix("   ") is None


# =============================
# VEHICLE MATCHING
# =============================
def test_match_joins_on_the_shared_prefix_despite_different_formatting():
    matched, unmatched, ambiguous = SR.match_vehicles(
        ["M-12", "WC-3"],
        [vehicle("281", "M12 - Medic 12"), vehicle("282", "WC3 Wheelchair Van")],
    )
    assert matched["M-12"]["id"] == "281"
    assert matched["WC-3"]["id"] == "282"
    assert not unmatched and not ambiguous


def test_match_prefers_an_exact_name():
    matched, _, _ = SR.match_vehicles(
        ["M-12"], [vehicle("1", "M-12"), vehicle("2", "M12 Spare")]
    )
    assert matched["M-12"]["id"] == "1"


def test_ambiguous_prefix_is_reported_not_guessed():
    """Two Samsara vehicles on one prefix must not silently pick one."""
    matched, unmatched, ambiguous = SR.match_vehicles(
        ["M-12"], [vehicle("1", "M12 Front"), vehicle("2", "M-12 Reserve")]
    )
    assert "M-12" not in matched
    assert "M-12" in ambiguous
    assert len(ambiguous["M-12"]) == 2


def test_unmatched_is_reported():
    matched, unmatched, _ = SR.match_vehicles(["M-99"], [vehicle("1", "M12")])
    assert not matched
    assert unmatched == ["M-99"]


def test_override_wins_over_the_prefix_rule():
    matched, _, _ = SR.match_vehicles(
        ["WC-3"],
        [vehicle("1", "WC3 Something"), vehicle("2", "Wheelchair 3")],
        overrides={"WC-3": "Wheelchair 3"},
    )
    assert matched["WC-3"]["id"] == "2"


def test_override_resolves_an_otherwise_ambiguous_prefix():
    matched, _, ambiguous = SR.match_vehicles(
        ["M-12"],
        [vehicle("1", "M12 Front"), vehicle("2", "M-12 Reserve")],
        overrides={"M-12": "M12 Front"},
    )
    assert matched["M-12"]["id"] == "1"
    assert not ambiguous


def test_override_keys_are_case_insensitive():
    matched, _, _ = SR.match_vehicles(
        ["wc-3"], [vehicle("2", "Wheelchair 3")], overrides={"WC-3": "Wheelchair 3"}
    )
    assert matched["wc-3"]["id"] == "2"


# =============================
# LEG SELECTION
# =============================
def test_eligible_keeps_a_normal_scheduled_leg():
    kept, skipped = SR.eligible_legs([leg()], parse_ts_aware)
    assert len(kept) == 1 and not skipped


@pytest.mark.parametrize(
    "overrides,reason_fragment",
    [
        ({"deleted": "true"}, "deleted"),
        ({"pickup_time": None}, "no scheduled pickup"),
        ({"pickup_time": ""}, "no scheduled pickup"),
        ({"vehicle_name": ""}, "no vehicle"),
        ({"call_type": "Standby"}, "excluded"),
        ({"call_type": "Cancelled En Route"}, "excluded"),
        ({"pu_lat": None, "pu_lon": None, "pu_address1": ""}, "no pickup location"),
        ({"do_lat": None, "do_lon": None, "do_address1": ""}, "no drop-off location"),
    ],
)
def test_ineligible_legs_are_skipped_with_a_reason(overrides, reason_fragment):
    kept, skipped = SR.eligible_legs([leg(**overrides)], parse_ts_aware)
    assert not kept
    assert reason_fragment in skipped[0][1]


def test_a_leg_with_only_an_address_is_still_routable():
    """Coordinates are preferred but a postal address alone is enough."""
    kept, _ = SR.eligible_legs(
        [leg(pu_lat=None, pu_lon=None, do_lat=None, do_lon=None)], parse_ts_aware
    )
    assert len(kept) == 1


def test_zero_coordinates_do_not_count_as_a_location():
    """0,0 is the Atlantic, not a facility."""
    kept, skipped = SR.eligible_legs(
        [leg(pu_lat=0, pu_lon=0, pu_address1="")], parse_ts_aware
    )
    assert not kept
    assert "no pickup location" in skipped[0][1]


# =============================
# STOPS
# =============================
def test_a_leg_makes_a_pickup_then_a_dropoff():
    stops, _ = SR.leg_stops(leg(), parse_ts_aware)
    assert len(stops) == 2
    assert stops[0]["name"].startswith("PU ·")
    assert stops[1]["name"].startswith("DO ·")
    assert stops[0]["scheduledArrivalTime"] < stops[1]["scheduledArrivalTime"]


def test_dropoff_uses_the_appointment_time():
    stops, source = SR.leg_stops(leg(), parse_ts_aware)
    assert source == "appt_time"
    assert stops[1]["scheduledArrivalTime"] == "2026-09-10T12:05:00Z"  # 08:05 -04:00


def test_dropoff_falls_back_to_the_eta_then_to_an_estimate():
    stops, source = SR.leg_stops(
        leg(appt_time=None, dropoff_eta="2026-09-10T08:20:00-04:00"), parse_ts_aware
    )
    assert source == "dropoff_eta"
    assert stops[1]["scheduledArrivalTime"] == "2026-09-10T12:20:00Z"

    stops, source = SR.leg_stops(leg(appt_time=None, dropoff_eta=None), parse_ts_aware)
    assert source == "estimated"
    # 07:30 -04:00 is 11:30Z, plus the 45-minute default.
    assert stops[1]["scheduledArrivalTime"] == "2026-09-10T12:15:00Z"


def test_a_dropoff_before_its_own_pickup_is_pushed_after_it():
    """
    Samsara orders stops by scheduledArrivalTime. An appointment time earlier
    than the pickup -- a return leg whose appt_time is the outbound one --
    would put the drop-off first and send the driver backwards.
    """
    stops, source = SR.leg_stops(
        leg(appt_time="2026-09-10T06:00:00-04:00"), parse_ts_aware
    )
    assert stops[0]["scheduledArrivalTime"] < stops[1]["scheduledArrivalTime"]
    assert "adjusted" in source


def test_equal_pickup_and_dropoff_times_are_separated():
    stops, _ = SR.leg_stops(leg(appt_time=leg()["pickup_time"]), parse_ts_aware)
    assert stops[0]["scheduledArrivalTime"] < stops[1]["scheduledArrivalTime"]


def test_stop_carries_coordinates_when_present():
    stops, _ = SR.leg_stops(leg(), parse_ts_aware)
    loc = stops[0]["singleUseLocation"]
    assert loc["latitude"] == pytest.approx(38.3498)
    assert loc["longitude"] == pytest.approx(-81.6326)
    assert "100 River Rd" in loc["address"]


def test_a_registered_samsara_address_is_preferred_over_a_single_use_location():
    """A registered address brings the geofence somebody already drew."""
    index = SR.index_addresses([{"id": "9001", "name": "Riverside Manor"}])
    stops, _ = SR.leg_stops(leg(), parse_ts_aware, address_index=index)
    assert stops[0]["addressId"] == "9001"
    assert "singleUseLocation" not in stops[0]
    # The drop-off is not registered, so it stays single-use.
    assert "singleUseLocation" in stops[1]


def test_notes_carry_the_call_type_and_level_of_service():
    stops, _ = SR.leg_stops(leg(), parse_ts_aware)
    notes = stops[0]["notes"]
    assert "BLS Transport" in notes
    assert "LOS BLS" in notes
    assert "24-10871" in notes


def test_patient_name_is_off_by_default():
    """
    A name plus a pickup address is PHI. It does not go to a third-party
    system unless somebody deliberately turns it on.
    """
    stops, _ = SR.leg_stops(leg(), parse_ts_aware)
    assert "Jane" not in stops[0]["notes"]
    assert "Doe" not in stops[0]["notes"]


def test_patient_name_can_be_enabled_and_is_then_only_on_the_pickup(monkeypatch):
    monkeypatch.setattr(SR, "INCLUDE_PATIENT_NAME", True)
    stops, _ = SR.leg_stops(leg(), parse_ts_aware)
    assert "Jane Doe" in stops[0]["notes"]
    assert "Jane Doe" not in stops[1]["notes"]


def test_format_address_skips_missing_parts():
    assert SR.format_address(leg(pu_address2="Suite 4"), "pu") == (
        "100 River Rd, Suite 4, Charleston, WV 25301"
    )
    assert SR.format_address(leg(pu_city="", pu_state="", pu_zipcode=""), "pu") == "100 River Rd"


# =============================
# TIME CONVERSION
# =============================
def test_rfc3339_converts_an_offset_to_utc():
    assert SR.rfc3339(parse_ts_aware("2026-09-10T07:30:00-04:00")) == "2026-09-10T11:30:00Z"


def test_a_naive_stamp_uses_the_tenant_offset_rather_than_being_read_as_utc():
    """
    Without this a 07:30 pickup on Eastern daylight time lands in Samsara as
    07:30Z -- four hours early, and the driver is dispatched before dawn.
    """
    naive = parse_ts_aware("2026-09-10T07:30:00")
    assert SR.rfc3339(naive, default_offset=timedelta(hours=-4)) == "2026-09-10T11:30:00Z"


# =============================
# ROUTES
# =============================
def test_one_route_per_vehicle_with_stops_in_pickup_order():
    legs = [
        leg(leg_id=2, run_number="B", pickup_time="2026-09-10T13:00:00-04:00",
            appt_time="2026-09-10T13:40:00-04:00"),
        leg(leg_id=1, run_number="A", pickup_time="2026-09-10T07:30:00-04:00"),
    ]
    matched = {"M-12": vehicle("281", "M12")}
    routes, notes = SR.build_routes(legs, matched, date(2026, 9, 10), parse_ts_aware)

    assert len(routes) == 1
    route = routes[0]
    assert route["vehicleId"] == "281"
    assert len(route["stops"]) == 4
    times = [s["scheduledArrivalTime"] for s in route["stops"]]
    assert times == sorted(times), "stops must be emitted in chronological order"
    assert notes[0]["legs"] == 2 and notes[0]["stops"] == 4
    assert notes[0]["run_numbers"] == ["A", "B"]


def test_two_vehicles_make_two_routes():
    legs = [leg(), leg(vehicle_name="M-14", run_number="24-10999")]
    matched = {"M-12": vehicle("281", "M12"), "M-14": vehicle("282", "M14")}
    routes, _ = SR.build_routes(legs, matched, date(2026, 9, 10), parse_ts_aware)
    assert {r["vehicleId"] for r in routes} == {"281", "282"}


def test_legs_for_an_unmatched_vehicle_produce_no_route():
    routes, notes = SR.build_routes([leg()], {}, date(2026, 9, 10), parse_ts_aware)
    assert routes == [] and notes == []


def test_route_name_carries_the_prefix_and_the_unit():
    routes, _ = SR.build_routes(
        [leg()], {"M-12": vehicle("281", "M12")}, date(2026, 9, 10),
        parse_ts_aware, name_prefix="[TS]",
    )
    assert routes[0]["name"].startswith("[TS] M-12")
    assert "Sep 10" in routes[0]["name"]


def test_estimated_dropoffs_are_counted_for_the_plan():
    legs = [leg(), leg(leg_id=2, appt_time=None, dropoff_eta=None,
                       pickup_time="2026-09-10T14:00:00-04:00")]
    _, notes = SR.build_routes(
        legs, {"M-12": vehicle("281", "M12")}, date(2026, 9, 10), parse_ts_aware
    )
    assert notes[0]["estimated_dropoff_times"] == 1


def test_route_payload_is_json_serialisable_and_shaped_for_samsara():
    """Every stop needs a time and exactly one kind of location."""
    routes, _ = SR.build_routes(
        [leg()], {"M-12": vehicle("281", "M12")}, date(2026, 9, 10), parse_ts_aware
    )
    json.dumps(routes)  # must not raise
    for route in routes:
        assert route["name"] and route["vehicleId"] and route["stops"]
        for stop in route["stops"]:
            assert stop["scheduledArrivalTime"]
            assert ("addressId" in stop) ^ ("singleUseLocation" in stop)


# =============================
# OVERRIDES FILE
# =============================
def test_overrides_file_round_trips(tmp_path):
    path = tmp_path / "overrides.json"
    path.write_text(json.dumps({"note": ["ignored"], "overrides": {"WC-3": "Wheelchair 3"}}))
    assert SR.load_vehicle_overrides(str(path)) == {"WC-3": "Wheelchair 3"}


def test_missing_overrides_file_is_not_an_error(tmp_path):
    assert SR.load_vehicle_overrides(str(tmp_path / "nope.json")) == {}


def test_unreadable_overrides_file_is_not_fatal(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text("{not json")
    assert SR.load_vehicle_overrides(str(path)) == {}


def test_the_shipped_example_overrides_file_parses():
    assert isinstance(
        SR.load_vehicle_overrides("state/samsara_vehicle_overrides.example.json"), dict
    )
