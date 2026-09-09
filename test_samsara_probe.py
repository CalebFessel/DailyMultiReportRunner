"""
Tests for the readiness probe's own arithmetic.

The probe recommends configuration, so its numbers have to be right for the
same reason the mapping's do -- a confident wrong recommendation is worse than
no recommendation. Import-safe without credentials: the probe only constructs
an API client inside main().
"""

from datetime import timedelta

import pytest

import probe_samsara_readiness as P
from test_samsara_routes import leg


# =============================
# PERCENTILE
# =============================
def test_percentile_endpoints_and_middle():
    values = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    assert P.percentile(values, 0.0) == 1
    assert P.percentile(values, 1.0) == 10
    assert P.percentile(values, 0.5) in (5, 6)


def test_percentile_is_order_independent():
    assert P.percentile([9, 1, 5, 3, 7], 0.5) == 5


def test_percentile_of_nothing_is_none():
    assert P.percentile([], 0.5) is None


def test_percentile_of_one_value():
    assert P.percentile([42], 0.9) == 42


# =============================
# COLLISION COUNTING
# =============================
def unscheduled(pickup, vehicle="A-101"):
    """A leg with no appointment time, so the estimate is actually consulted."""
    return leg(pickup_time=pickup, appt_time=None, dropoff_eta=None,
               vehicle_name=vehicle, timestamps=[])


def test_a_long_estimate_collides_with_a_close_next_pickup():
    legs = [
        unscheduled("2026-09-10T07:30:00-04:00"),
        unscheduled("2026-09-10T08:00:00-04:00"),   # 30 min later
    ]
    collisions, considered = P.count_collisions(legs, lambda l: 45)
    assert considered == 1        # only the first leg has a next pickup
    assert collisions == 1


def test_a_short_estimate_does_not_collide():
    legs = [
        unscheduled("2026-09-10T07:30:00-04:00"),
        unscheduled("2026-09-10T08:00:00-04:00"),
    ]
    collisions, considered = P.count_collisions(legs, lambda l: 20)
    assert considered == 1
    assert collisions == 0


def test_legs_with_a_real_appointment_time_are_never_counted():
    """The model is not consulted for a leg that already has a real time."""
    legs = [
        leg(pickup_time="2026-09-10T07:30:00-04:00", vehicle_name="A-101", timestamps=[]),
        leg(pickup_time="2026-09-10T08:00:00-04:00", vehicle_name="A-101", timestamps=[]),
    ]
    collisions, considered = P.count_collisions(legs, lambda l: 999)
    assert (collisions, considered) == (0, 0)


def test_the_last_leg_of_a_day_cannot_collide():
    collisions, considered = P.count_collisions(
        [unscheduled("2026-09-10T07:30:00-04:00")], lambda l: 999
    )
    assert (collisions, considered) == (0, 0)


def test_different_vehicles_do_not_collide_with_each_other():
    """Two units running the same hour are two routes, not one."""
    legs = [
        unscheduled("2026-09-10T07:30:00-04:00", vehicle="A-101"),
        unscheduled("2026-09-10T07:40:00-04:00", vehicle="WC-400"),
    ]
    collisions, considered = P.count_collisions(legs, lambda l: 45)
    assert (collisions, considered) == (0, 0)


def test_the_same_unit_on_different_days_does_not_collide():
    legs = [
        unscheduled("2026-09-10T22:00:00-04:00"),
        unscheduled("2026-09-11T07:00:00-04:00"),
    ]
    collisions, considered = P.count_collisions(legs, lambda l: 45)
    assert (collisions, considered) == (0, 0)


def test_collisions_are_counted_across_a_longer_day():
    legs = [unscheduled(f"2026-09-10T{h:02d}:00:00-04:00") for h in (7, 8, 9, 10)]
    # 45 minutes never reaches the next pickup an hour later.
    assert P.count_collisions(legs, lambda l: 45) == (0, 3)
    # 90 minutes always does.
    assert P.count_collisions(legs, lambda l: 90) == (3, 3)
