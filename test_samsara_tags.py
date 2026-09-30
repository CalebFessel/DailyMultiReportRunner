"""
Tests for the tag probe's own judgement.

The probe prints rules for a human to paste into a config file, which makes a
wrong suggestion worse than no suggestion. Live data caught exactly that: it
offered `"Indiana": "Indianapolis"` because one name contains the other, and
pasting that line would have handed Salem's and Sellersburg's trucks to
Indianapolis. The tag hierarchy rules it out where a string comparison cannot.

Import-safe without credentials: the probe only builds clients inside main().
"""

import pytest

import probe_samsara_tags as P


# The real shape on this tenant: stations hang off a state, and the top-level
# tags name states, classes or the whole fleet.
PARENTS = {
    "ALL": None,
    "Ohio": None,
    "Indiana": None,
    "West Virginia": None,
    "BLS": None,
    "In Service (TraumaSoft)": None,
    "Indianapolis": "5816349",
    "Salem": "5816349",
    "Sellersburg": "5816349",
    "Newburg": "5816349",
    "Boardman": "5816352",
    "Parma": "5816352",
    "Ellicott City": "5816370",
}


# =============================
# STATION vs STATE
# =============================
def test_a_station_tag_hangs_off_a_state():
    assert P.is_child_tag("Indianapolis", PARENTS) is True
    assert P.is_child_tag("Boardman", PARENTS) is True


def test_a_state_tag_is_not_a_station():
    """
    The one that matters. `Indiana` contains `Indianapolis` as a substring,
    so a name comparison offers it as a rule -- and it is the parent of
    Salem, Sellersburg and Newburg too.
    """
    assert P.is_child_tag("Indiana", PARENTS) is False


def test_class_and_fleet_wide_tags_are_not_stations():
    for tag in ("ALL", "BLS", "In Service (TraumaSoft)", "Ohio"):
        assert P.is_child_tag(tag, PARENTS) is False, tag


def test_a_tag_nothing_knows_about_is_not_a_station():
    assert P.is_child_tag("Johnstown", {}) is False
    assert P.is_child_tag("Johnstown", None) is False


# =============================
# THE STRICT MATCHER
# =============================
def test_an_exact_name_matches():
    assert P.station_matches("Boardman", "Boardman") is True


def test_case_and_spacing_do_not_decide_it():
    assert P.station_matches("  boardman ", "Boardman") is True


def test_the_legal_entity_wrapper_is_seen_through():
    assert P.station_matches("Parma Heights",
                             "Lynx EMS LLC dba Lynx Parma Heights") is True


def test_a_state_does_not_match_a_station_inside_it():
    """This is why the resolver was safe even while the suggestion was not."""
    assert P.station_matches("Indiana", "Indianapolis") is False
    assert P.station_matches("Ohio", "Boardman") is False


def test_the_real_near_misses_do_not_match_on_name():
    for tag, centre in (("Ellicott City", "Ellicott"),
                        ("Newburg", "Newburgh"),
                        ("Parma", "Parma Heights")):
        assert P.station_matches(tag, centre) is False, tag


def test_nothing_matches_nothing():
    assert P.station_matches("", "Boardman") is False
    assert P.station_matches("Boardman", "") is False
    assert P.station_matches(None, None) is False
