"""
Checks for where OTP takes its two timestamps from.

    python -m pytest test_otp_sources.py     # preferred
    python test_otp_sources.py               # same checks, no pytest needed

Two things are being defended here.

The first is scoring. OTP compares an arrival stamp against a scheduled pickup,
and both halves are now configurable -- which means both halves can be
misconfigured. These checks pin the defaults (`at_scene` against `pickup_time`,
with nothing chained in behind either) and prove the chain is honoured in
preference order when it is set, so a leg is never scored against a field
nobody chose.

The second is that probe_epcr_huly.py cannot write. It sweeps an undocumented
surface with guessed action names, which is fine while every one of them is a
read and catastrophic the moment one is not. The guard walks its AST rather
than reading its text, so a write-shaped name cannot arrive in a rename, a
list comprehension or a string that merely mentions it in a comment.
"""

import ast
import sys
import inspect
from pathlib import Path

import traumasoft_reports as R

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


def leg(**fields):
    """A trip leg carrying only what a check cares about."""
    base = {"leg_id": 1, "run_number": "R1", "trip_status": "Completed"}
    base.update(fields)
    return base


def stamps(**pairs):
    """The timestamps array shape GetTrips returns: a list of one-key maps."""
    return [{name: value} for name, value in pairs.items()]


# =============================
# Defaults
# =============================

def test_default_keys():
    print("\ntest_default_keys")

    check(
        "arrival defaults to at_scene alone",
        R.ARRIVAL_TIMESTAMP_KEYS == ["at_scene"],
        f"got {R.ARRIVAL_TIMESTAMP_KEYS}",
    )
    check(
        "the bedside stamp is not in the default chain",
        not any("bedside" in key.lower() for key in R.ARRIVAL_TIMESTAMP_KEYS),
        "this tenant does not capture At Patient Bedside; a stamp nobody "
        "records can only ever be a miss",
    )
    check(
        "pickup defaults to pickup_time alone",
        R.PICKUP_TIME_KEYS == ["pickup_time"],
        f"got {R.PICKUP_TIME_KEYS}",
    )
    check(
        "requested_pickup_time is not chained in behind the scheduled pickup",
        "requested_pickup_time" not in R.PICKUP_TIME_KEYS,
        "what the caller asked for is not what dispatch promised; scoring "
        "against it measures the call taker, not the crew",
    )


# =============================
# Reading the scheduled pickup
# =============================

def test_scheduled_pickup_reads_the_configured_field():
    print("\ntest_scheduled_pickup_reads_the_configured_field")

    row = leg(
        pickup_time="2026-08-19T09:00:00-04:00",
        requested_pickup_time="2026-08-19T08:00:00-04:00",
        appt_time="2026-08-19T10:00:00-04:00",
    )
    found = R.scheduled_pickup_time(row)
    check(
        "a leg with every candidate populated is read from pickup_time",
        found is not None and found.hour == 9,
        f"got {found}",
    )

    check(
        "a leg with only requested_pickup_time has no scheduled pickup",
        R.scheduled_pickup_time(leg(requested_pickup_time="2026-08-19T08:00:00-04:00")) is None,
        "it must not be silently substituted",
    )

    found = R.scheduled_pickup_time(
        leg(requested_pickup_time="2026-08-19T08:00:00-04:00"),
        keys=["pickup_time", "requested_pickup_time"],
    )
    check(
        "an explicit chain falls through to the next field",
        found is not None and found.hour == 8,
        f"got {found}",
    )

    check(
        "a blank string is treated as absent, not as a parse failure",
        R.scheduled_pickup_time(leg(pickup_time="")) is None,
    )


# =============================
# Reading the arrival stamp
# =============================

def test_arrival_reads_the_configured_stamp():
    print("\ntest_arrival_reads_the_configured_stamp")

    row = leg(timestamps=stamps(
        enroute="2026-08-19T08:50:00-04:00",
        at_scene="2026-08-19T09:05:00-04:00",
    ))
    found = R.arrival_time(row)
    check(
        "at_scene is the arrival stamp",
        found is not None and (found.hour, found.minute) == (9, 5),
        f"got {found}",
    )

    bedside = leg(timestamps=[{"at_scene: At Patient Bedside": "2026-08-19T09:20:00-04:00"}])
    check(
        "a bedside-only leg no longer resolves by default",
        R.arrival_time(bedside) is None,
        "the default chain is at_scene; bedside returns only when asked for",
    )
    found = R.arrival_time(bedside, keys=["at_scene: At Patient Bedside", "at_scene"])
    check(
        "...but it still resolves when the chain names it",
        found is not None and found.hour == 9,
        f"got {found}",
    )


# =============================
# Scoring
# =============================

def test_scoring_needs_both_halves():
    print("\ntest_scoring_needs_both_halves")

    on_time = leg(
        pickup_time="2026-08-19T09:00:00-04:00",
        timestamps=stamps(at_scene="2026-08-19T09:05:00-04:00"),
    )
    status, delta = R.score_leg(on_time)
    check("five minutes late is On Time inside the window", status == "On Time",
          f"got {status} ({delta})")

    late = leg(
        pickup_time="2026-08-19T09:00:00-04:00",
        timestamps=stamps(at_scene="2026-08-19T09:40:00-04:00"),
    )
    status, delta = R.score_leg(late)
    check("forty minutes is Late", status == "Late" and round(delta) == 40,
          f"got {status} ({delta})")

    unscheduled = leg(
        requested_pickup_time="2026-08-19T09:00:00-04:00",
        timestamps=stamps(at_scene="2026-08-19T09:40:00-04:00"),
    )
    status, _ = R.score_leg(unscheduled)
    check(
        "a leg with no scheduled pickup is Missing Data, not Late",
        status == "Missing Data",
        "an unscheduled call was never promised a time; calling it late "
        "invents a deadline nobody gave the crew",
    )

    status, _ = R.score_leg(unscheduled, pickup_keys=["requested_pickup_time"])
    check(
        "...and becomes scorable only when the chain is changed deliberately",
        status == "Late",
        f"got {status}",
    )


def test_scored_legs_passes_the_chain_through():
    print("\ntest_scored_legs_passes_the_chain_through")

    legs = [
        leg(leg_id=1, pickup_time="2026-08-19T09:00:00-04:00",
            timestamps=stamps(at_scene="2026-08-19T09:02:00-04:00")),
        leg(leg_id=2, requested_pickup_time="2026-08-19T09:00:00-04:00",
            timestamps=stamps(at_scene="2026-08-19T09:02:00-04:00")),
    ]
    cost_centers = R.CostCenterMap()

    default = R.scored_legs(legs, cost_centers)
    check("only the scheduled leg is scored by default", len(default) == 1,
          f"got {len(default)} rows")

    widened = R.scored_legs(legs, cost_centers,
                            pickup_keys=["pickup_time", "requested_pickup_time"])
    check("both are scored when the chain is widened", len(widened) == 2,
          f"got {len(widened)} rows")


# =============================
# The ePCR probe cannot write
# =============================

WRITE_SHAPED = ("update", "set", "create", "delete", "post", "put", "insert",
                "remove", "save", "write", "patch", "sync")


def test_epcr_probe_is_read_only():
    print("\ntest_epcr_probe_is_read_only")

    path = Path(__file__).resolve().parent / "probe_epcr_huly.py"
    check("the probe exists", path.exists(), str(path))
    tree = ast.parse(path.read_text(encoding="utf-8"))

    # Every string constant that sits in one of the candidate lists. Walking
    # the AST rather than grepping means a comment mentioning HulyUpdateTrip
    # -- the docstring does, to say why it is excluded -- is not a finding,
    # and a name smuggled in through a list comprehension still is.
    candidates = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if not isinstance(target, ast.Name):
                continue
            if target.id not in ("HULY_RTYPES", "NEIGHBOURING_PATHS"):
                continue
            for inner in ast.walk(node.value):
                if isinstance(inner, ast.Constant) and isinstance(inner.value, str):
                    candidates.append(inner.value)

    check("candidate actions were found to inspect", len(candidates) > 5,
          f"got {candidates}")

    offenders = [
        name for name in candidates
        if any(word in name.lower() for word in WRITE_SHAPED)
    ]
    check(
        "no candidate action is write-shaped",
        not offenders,
        f"these would mutate ePCR data: {offenders}",
    )
    check(
        "HulyUpdateTrip is never a candidate",
        not any("hulyupdate" in name.lower() for name in candidates),
        "the one action the spec names as a write must never be sent",
    )

    # The probe may only reach the network through api.get(). A call to
    # .request(), .post() or the session directly would bypass that guarantee.
    forbidden = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if isinstance(func, ast.Attribute) and func.attr in (
            "post", "put", "patch", "delete", "request"
        ):
            forbidden.append(func.attr)
    check(
        "the probe issues no write-method calls",
        not forbidden,
        f"found {sorted(set(forbidden))}; a probe may only call api.get()",
    )


def main():
    """The standalone runner, for a machine with no pytest."""
    tests = [
        test_default_keys,
        test_scheduled_pickup_reads_the_configured_field,
        test_arrival_reads_the_configured_stamp,
        test_scoring_needs_both_halves,
        test_scored_legs_passes_the_chain_through,
        test_epcr_probe_is_read_only,
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
