"""
Tests for building a vehicle -> cost center map from Samsara tags.

The point of this map is that it states where a truck BELONGS. Every other
source in the repository infers it from where the truck worked, so the ways
this can be quietly wrong are the ways it stops being a statement and goes
back to being a guess:

  * a vehicle whose tags name two stations must not be assigned to one of
    them. A station roster that looks authoritative and is partly invented is
    worse than no roster;
  * a short tag must not substring-match a cost center. 'Columbus' claiming
    two states is exactly the failure the regions file was written to avoid;
  * a vehicle the map places elsewhere must not be dragged back into a
    station's fleet by a week of covering for it -- that is the inference the
    map exists to replace.

No credentials needed: the APIs are only touched inside main().
"""

import json
from datetime import date

import pytest

import build_vehicle_cost_centers as BV
import vehicle_assignment_by_day as A
import vehicle_service_days as VS


CENTRES = ["Boardman", "Cincinnati", "Lynx EMS LLC dba Lynx Parma Heights"]


def sam(vid, name, tags=(), attrs=(), vin=""):
    return {
        "id": vid,
        "name": name,
        "vin": vin,
        "tags": [{"name": t} for t in tags],
        "attributes": [{"name": "Station", "stringValues": list(attrs)}] if attrs else [],
    }


def ts(vid, name, vin=""):
    return {"id": vid, "name": name, "vin": vin, "vehicle_status": "In Service"}


# =============================
# READING THE LABELS
# =============================
def test_tags_are_read_however_the_payload_spells_them():
    assert BV.tag_names({"tags": [{"name": "Boardman"}, "Cincinnati"]}) == \
        ["Boardman", "Cincinnati"]


def test_a_tag_with_only_an_id_still_counts():
    assert BV.tag_names({"tags": [{"id": "123"}]}) == ["123"]


def test_no_tags_is_an_empty_list_not_an_error():
    assert BV.tag_names({}) == []
    assert BV.tag_names({"tags": None}) == []


def test_attributes_are_read_too():
    """Some fleets put the station in an attribute. Only looking at tags would
    report an empty answer on a tenant that had the data all along."""
    vehicle = sam("s1", "A-101", attrs=["Boardman"])
    assert BV.attribute_values(vehicle) == ["Boardman"]


def test_a_malformed_attribute_is_skipped_not_fatal():
    assert BV.attribute_values({"attributes": ["nonsense", {"name": "x"}]}) == []


# =============================
# MATCHING A LABEL TO A STATION
# =============================
def test_an_exact_name_matches():
    centre, how = BV.resolve_label("Boardman", CENTRES, {})
    assert centre == "Boardman"
    assert how == "exact name"


def test_the_legal_entity_wrapper_is_seen_through():
    centre, how = BV.resolve_label("Parma Heights", CENTRES, {})
    assert centre == "Lynx EMS LLC dba Lynx Parma Heights"
    assert "wrapper" in how


def test_a_tag_map_entry_beats_the_name_match():
    """A decision beats an observation, the same precedence the cost-center
    overrides use."""
    centre, how = BV.resolve_label("YO-1", CENTRES, {"yo-1": "Boardman"})
    assert centre == "Boardman"
    assert how == "tag map"


def test_there_is_no_substring_pass():
    """
    A tag is a short label, and substring-matching short labels is how
    'Columbus' claims Columbus Ohio and Columbus Indiana at once.
    """
    centre, why = BV.resolve_label("Board", CENTRES, {})
    assert centre is None
    assert why == "matches no cost center"


def test_an_empty_label_matches_nothing():
    assert BV.resolve_label("", CENTRES, {})[0] is None
    assert BV.resolve_label("   ", CENTRES, {})[0] is None


def test_a_tag_map_is_read_from_either_shape(tmp_path):
    plain = tmp_path / "plain.json"
    plain.write_text(json.dumps({"YO-1": "Boardman"}))
    assert BV.load_tag_map(str(plain)) == {"yo-1": "Boardman"}

    wrapped = tmp_path / "wrapped.json"
    wrapped.write_text(json.dumps({"tags": {"YO-2": "Cincinnati"}}))
    assert BV.load_tag_map(str(wrapped)) == {"yo-2": "Cincinnati"}


def test_a_missing_tag_map_is_not_fatal():
    assert BV.load_tag_map("does/not/exist.json") == {}
    assert BV.load_tag_map(None) == {}


# =============================
# RESOLVING A VEHICLE
# =============================
def test_one_matching_tag_resolves_the_vehicle():
    centre, source, note = BV.resolve_vehicle(["Boardman", "Ambulance"], CENTRES, {})
    assert centre == "Boardman"
    assert "Boardman" in source
    assert note is None


def test_two_stations_on_one_vehicle_is_left_unset():
    """
    Genuinely ambiguous. Picking one would produce a station roster that looks
    authoritative and is partly invented.
    """
    centre, _, note = BV.resolve_vehicle(["Boardman", "Cincinnati"], CENTRES, {})
    assert centre is None
    assert "Boardman" in note and "Cincinnati" in note


def test_the_same_station_named_twice_is_not_ambiguous():
    centre, _, _ = BV.resolve_vehicle(
        ["Parma Heights", "Lynx EMS LLC dba Lynx Parma Heights"], CENTRES, {}
    )
    assert centre == "Lynx EMS LLC dba Lynx Parma Heights"


def test_no_tags_at_all_says_so():
    centre, _, note = BV.resolve_vehicle([], CENTRES, {})
    assert centre is None
    assert note == "no tags or attributes"


def test_tags_that_match_nothing_say_something_different():
    centre, _, note = BV.resolve_vehicle(["Ambulance", "2019"], CENTRES, {})
    assert centre is None
    assert note == "no label matched a cost center"


# =============================
# THE WHOLE MAP
# =============================
def test_a_tagged_vehicle_joined_by_vin_resolves():
    vin = "1FDUF5HT7KDA12345"
    resolved, unresolved, _ = BV.build_map(
        [ts(1, "A-101", vin)], [sam("s1", "Whatever", tags=["Boardman"], vin=vin)],
        CENTRES, {},
    )
    assert resolved["A101"]["cost_center"] == "Boardman"
    assert resolved["A101"]["samsara_join"] == "vin"
    assert unresolved == []


def test_a_vehicle_with_no_samsara_match_is_reported_not_dropped():
    resolved, unresolved, _ = BV.build_map([ts(1, "A-101")], [], CENTRES, {})
    assert resolved == {}
    assert unresolved[0][0] == "A-101"
    assert unresolved[0][1] == "no Samsara match"


def test_an_untagged_vehicle_is_reported_with_its_own_reason():
    resolved, unresolved, _ = BV.build_map(
        [ts(1, "A-101")], [sam("s1", "A-101")], CENTRES, {},
    )
    assert resolved == {}
    assert unresolved[0][1] == "no tags or attributes"


def test_writing_and_reading_the_map_round_trips(tmp_path):
    path = tmp_path / "vehicle_cost_centers.json"
    BV.write_map(str(path), {
        "A101": {"vehicle_name": "A-101", "cost_center": "Boardman",
                 "source": "exact name: Boardman"},
    }, "test")
    assert BV.load_vehicle_cost_centers(str(path)) == {"A101": "Boardman"}


def test_the_written_file_records_when_it_was_built(tmp_path):
    """A tag is current, not historical. Using it for July assumes nothing
    moved since, and a reader can only judge that with a date."""
    path = tmp_path / "m.json"
    payload = BV.write_map(str(path), {}, "test")
    assert payload["built"]
    assert any("historical" in line for line in payload["note"])


def test_a_missing_map_is_not_an_error():
    assert BV.load_vehicle_cost_centers("does/not/exist.json") == {}
    assert BV.load_vehicle_cost_centers(None) == {}


def test_a_hand_edited_map_reads_back():
    """The file is plain JSON precisely so anything unresolved can be typed
    in. It has to read back the same way as a generated one."""
    import tempfile, os
    handle = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False)
    json.dump({"vehicles": {"B-7": "Boardman"}}, handle)
    handle.close()
    try:
        assert BV.load_vehicle_cost_centers(handle.name) == {"B7": "Boardman"}
    finally:
        os.unlink(handle.name)


# =============================
# THE FLEET, ONCE THE MAP EXISTS
# =============================
class FakeMap:
    def __init__(self, mapping):
        self.mapping = mapping

    def resolve(self, shift_name):
        return self.mapping.get(str(shift_name or "").strip())


MAP = FakeMap({"BOARD-A-1": "Boardman", "CIN-A-1": "Cincinnati"})


def leg(vehicle, profile, day=1):
    return {"vehicle_name": vehicle, "shift_name": profile,
            "pickup_time": f"2026-07-{day:02d}T09:00:00-04:00"}


def test_the_map_places_a_truck_that_never_ran():
    fleet = A.build_fleet({}, {}, set(), [], {"Boardman"}, MAP,
                          vehicle_centres={"B1": "Boardman"})
    assert fleet == {"B1": "its cost center map entry says so"}


def test_the_map_beats_every_behaviour_derived_source():
    activity, _, _ = A.legs_by_vehicle_day([leg("B-1", "BOARD-A-1")],
                                           {"Boardman"}, MAP)
    fleet = A.build_fleet(activity, {}, {"B1"},
                          [{"name": "B-1", "shift_name": "BOARD-A-1"}],
                          {"Boardman"}, MAP, vehicle_centres={"B1": "Boardman"})
    assert fleet["B1"] == "its cost center map entry says so"


def test_a_truck_the_map_puts_elsewhere_stays_out_despite_running_here():
    """
    The whole reason for the map. A Cincinnati truck that covered Boardman for
    a week is Cincinnati's, and letting its behaviour override the map would
    put us back where we started.
    """
    activity, _, _ = A.legs_by_vehicle_day([leg("C-9", "BOARD-A-1")],
                                           {"Boardman"}, MAP)
    fleet = A.build_fleet(activity, {}, set(), [], {"Boardman"}, MAP,
                          vehicle_centres={"C9": "Cincinnati"})
    assert fleet == {}


def test_a_truck_the_map_puts_elsewhere_is_not_rescued_by_the_roster():
    fleet = A.build_fleet({}, {}, {"C9"},
                          [{"name": "C-9", "shift_name": "BOARD-A-1"}],
                          {"Boardman"}, MAP, vehicle_centres={"C9": "Cincinnati"})
    assert fleet == {}


def test_a_truck_the_map_does_not_mention_still_falls_back_to_behaviour():
    """The map is authoritative where it speaks, not where it is silent."""
    activity, _, _ = A.legs_by_vehicle_day([leg("B-2", "BOARD-A-1")],
                                           {"Boardman"}, MAP)
    fleet = A.build_fleet(activity, {}, set(), [], {"Boardman"}, MAP,
                          vehicle_centres={"B1": "Boardman"})
    assert fleet["B2"] == "ran for it this period"


def test_no_map_at_all_behaves_exactly_as_before():
    activity, _, _ = A.legs_by_vehicle_day([leg("B-1", "BOARD-A-1")],
                                           {"Boardman"}, MAP)
    assert A.build_fleet(activity, {}, set(), [], {"Boardman"}, MAP) == \
        A.build_fleet(activity, {}, set(), [], {"Boardman"}, MAP,
                      vehicle_centres={})


# =============================
# THE COMPARISON
# =============================
def test_agreement_and_disagreement_are_told_apart():
    resolved = {
        "B1": {"vehicle_name": "B-1", "cost_center": "Boardman"},
        "C9": {"vehicle_name": "C-9", "cost_center": "Cincinnati"},
    }
    legs = [leg("B-1", "BOARD-A-1"), leg("C-9", "BOARD-A-1")]
    agree, differ, tag_only, legs_only = BV.compare_with_crew_map(
        resolved, legs, MAP)
    assert agree == ["B-1"]
    assert differ == [("C-9", "Cincinnati", "Boardman")]
    assert tag_only == [] and legs_only == []


def test_a_truck_the_tags_place_and_the_legs_cannot_is_the_whole_point():
    resolved = {"B2": {"vehicle_name": "B-2", "cost_center": "Boardman"}}
    agree, differ, tag_only, _ = BV.compare_with_crew_map(resolved, [], MAP)
    assert tag_only == ["B-2"]
    assert agree == [] and differ == []


def test_the_wrapper_does_not_read_as_a_disagreement():
    resolved = {"P1": {"vehicle_name": "P-1",
                       "cost_center": "Lynx EMS LLC dba Lynx Parma Heights"}}
    parma = FakeMap({"P-PROF": "Parma Heights"})
    agree, differ, _, _ = BV.compare_with_crew_map(
        resolved, [leg("P-1", "P-PROF")], parma)
    assert agree == ["P-1"] and differ == []


def test_a_note_in_a_wrapped_tag_map_is_not_mistaken_for_a_station(tmp_path):
    path = tmp_path / "noted.json"
    path.write_text(json.dumps({
        "note": ["why this file exists"],
        "tags": {"YO-1": "Boardman"},
    }))
    assert BV.load_tag_map(str(path)) == {"yo-1": "Boardman"}


def test_a_bare_tag_map_ignores_non_string_values(tmp_path):
    path = tmp_path / "bare.json"
    path.write_text(json.dumps({"note": ["prose"], "YO-1": "Boardman"}))
    assert BV.load_tag_map(str(path)) == {"yo-1": "Boardman"}


def test_an_unrelated_attribute_is_not_fed_to_the_station_matcher():
    """
    This tenant's attributes are Asset Status, Unit Type and Vehicle Status.
    Feeding 'Active' or 'Secure Car' into a cost-center matcher is asking for
    the day a station is named something that collides with one of them.
    """
    vehicle = {"attributes": [
        {"name": "Asset Status", "stringValues": ["Out of Service"]},
        {"name": "Unit Type", "stringValues": ["Secure Car"]},
        {"name": "Vehicle Status", "stringValues": ["Retired"]},
    ]}
    assert BV.attribute_values(vehicle) == []


def test_a_station_shaped_attribute_is_still_read():
    for label in ("Station", "Home Base", "Cost Center", "Primary Location"):
        vehicle = {"attributes": [{"name": label, "stringValues": ["Boardman"]}]}
        assert BV.attribute_values(vehicle) == ["Boardman"], label


def test_the_real_near_misses_on_this_fleet_resolve_through_the_tag_map():
    """
    Ellicott City / Ellicott, Newburg / Newburgh, Parma / Parma Heights are
    what probe_samsara_tags section 6 actually found. None of them matches on
    name, and none of them should be reached by a substring rule.
    """
    centres = ["Ellicott", "Newburgh", "Parma Heights"]
    for tag in ("Ellicott City", "Newburg", "Parma"):
        assert BV.resolve_label(tag, centres, {})[0] is None, tag
    tag_map = BV.load_tag_map("state/tag_cost_center_map.example.json")
    for tag, expected in (("Ellicott City", "Ellicott"),
                          ("Newburg", "Newburgh"),
                          ("Parma", "Parma Heights")):
        assert BV.resolve_label(tag, centres, tag_map) == (expected, "tag map")


def test_a_state_tag_does_not_claim_a_station_inside_it():
    """'Indiana' must not resolve to 'Indianapolis'."""
    assert BV.resolve_label("Indiana", ["Indianapolis"], {})[0] is None
    assert BV.resolve_label("Ohio", ["Boardman", "Cincinnati"], {})[0] is None
