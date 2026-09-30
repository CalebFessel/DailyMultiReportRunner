"""
Whether a vehicle was productive, against the definition upper management gave.

The definition, verbatim:

    Productive      Engine On AND Moved >1mile in X time window AND Has
                    assigned crew AND Has scheduled pickup AND Has Completed
                    ePCR AND Do not have Out of service label.
    Non-Productive  No crew assigned OR Known out of service
    Indetermined    Engine on, Moving, crewed, Run assigned but ePCR not
                    complete

Six inputs. This repository can reach four of them today:

    Engine on            Samsara     fleet/vehicles/stats/history engineStates
    Moved > 1 mile       Samsara     odometer delta, or GPS track length
    Assigned crew        Traumasoft  Schedule/Shifts, joined on vehicle_name
    Scheduled pickup     Traumasoft  CAD trip legs assigned to the vehicle
    Completed ePCR       NOT REACHABLE -- see below
    Out of service       Traumasoft  Vehicle vehicle_status / disabled / deleted

ePCR IS THE PROBLEM, AND IT IS NOT A SMALL ONE. It appears in `Productive`
(must be complete) and in `Indetermined` (must be incomplete), so it is the
single condition that separates the two positive categories from each other.
The published ThirdParty API does not expose it: `ThirdParty/Data/Epcr/Huly`
is listed in the spec under "Not included in this spec -- private or
non-partner integrations", and the only place an ePCR id appears in the whole
schema is as a write-side attachment target. Whether a named `rtype` action
can still read it is what probe_epcr_huly.py is for, and that is unsettled.
Until it is, this report can say with confidence which vehicles were NOT
productive, and cannot certify that any vehicle WAS. That is reported as
`undetermined`, never quietly as productive -- and the count of rows that
meet every other condition is the size of the prize. The `epcr_complete`
key is carried through every row as an explicit None, so one assignment
turns it on the day the surface opens. See docs/VEHICLE_PRODUCTIVITY.md.

THE DEFINITION AS FIRST WRITTEN LEFT TWO CASES UNCOVERED: a crewed,
in-service truck that moved with no call assigned, and one that never
started. Management settled both as Non-Productive, so the rule implemented
here is "any Productive condition known to fail". The specific condition is
carried in every row's reason and broken out in the report, because a truck
that burned crew hours and fuel with nothing dispatched is a different
problem from one that never turned a wheel -- and the first is precisely the
case the existing Daily Vehicle Overview is blind to.

Unit of analysis is one vehicle, one tenant-local calendar day. Nothing in the
definition says so; it is the unit the fleet sheets already use and the one
"X time window" most plausibly means. `--window-hours` narrows it if that is
wrong.

STRICTLY READ-ONLY. Both clients are constructed read-only and nothing here
passes write=True.

PHI: no patient-identifying value is read or printed. Vehicle names, crew
counts and call counts are operational.

Usage:
    python vehicle_productivity.py
    python vehicle_productivity.py --day 2026-09-22 --csv productivity.csv
    python vehicle_productivity.py --distance-source gps --min-miles 1
"""

import argparse
import csv
import json
import logging
import math
import os
import re
import sys
import textwrap
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone

log = logging.getLogger(__name__)

STATS_PATH = "fleet/vehicles/stats/history"

# Samsara's engine-state vocabulary. Only "off" means the vehicle was not in
# use: an EMS unit idles with the engine running for climate control and
# equipment power, and that is a truck in use, not a parked one. Anything
# outside this set counts as engine-on and is reported by name, never
# silently bucketed.
OFF_VALUES = {"off"}

# Classifications. The first three are the definition as given; the last two
# exist because the definition does not cover every vehicle-day and because
# some inputs cannot be read at all.
PRODUCTIVE = "productive"
NON_PRODUCTIVE = "non-productive"
INDETERMINATE = "indetermined"
UNDETERMINED = "undetermined"      # a required input could not be read

CLASS_ORDER = [PRODUCTIVE, INDETERMINATE, UNDETERMINED, NON_PRODUCTIVE]

# Why a vehicle-day is Non-Productive. The verdict is one word; these are the
# operational stories behind it, and they are not the same story. A truck that
# burned crew hours and fuel with no call is a dispatch or staffing question;
# a truck that never started is a readiness one. Kept on every row so the
# report can break the total down instead of just reporting it.
NON_PRODUCTIVE_REASONS = {
    "out_of_service": "out of service",
    "has_crew": "no crew assigned",
    "has_scheduled_pickup": "no call was assigned to it",
    "engine_on": "the engine never ran",
    "moved": "it did not move far enough",
}

# The order the chain actually breaks in: a truck has to be available, then
# staffed, then given work, then actually go. A vehicle usually fails several
# of these at once -- an unstaffed truck gets no calls and never starts -- and
# reporting all four would fragment the grouping into one bucket per
# combination while burying the thing to act on. So the EARLIEST break is the
# headline reason, and every failing condition is carried in its own column
# for anyone who wants the rest. You cannot fix "never started" when the real
# problem is that nobody dispatched it.
CONDITION_PRIORITY = ["has_crew", "has_scheduled_pickup", "engine_on", "moved"]

MILES_PER_METER = 0.000621371
EARTH_RADIUS_MILES = 3958.7613

# Traumasoft vehicle_status values, mirroring traumasoft_reports so the two
# reports never disagree about which trucks are down.
OUT_OF_SERVICE_STATUSES = {"Out of Service", "Out of Service - Collision"}
NON_FLEET_STATUSES = {"Retired", "New - Waiting for Delivery", "Waiting for Inspection"}


# =============================
# TIME
# =============================
def tenant_zone():
    """
    The tenant's zone, for calendar-day boundaries.

    A calendar day is the window productivity is measured over, so getting
    this wrong shifts every boundary by the offset and silently reassigns a
    night shift's work to the wrong date.
    """
    name = os.getenv("SAMSARA_TENANT_TIMEZONE", "").strip()
    if not name:
        raise RuntimeError(
            "SAMSARA_TENANT_TIMEZONE is not set. A calendar-day window cannot "
            "be drawn without it -- set it in .env (e.g. America/New_York)."
        )
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo(name), name
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"SAMSARA_TENANT_TIMEZONE={name!r} is not a zone: {exc}")


def day_window(day, zone, window_hours=24, lookback_hours=24):
    """
    (fetch_from, fetch_to, window_start, window_end) for one analysed day.

    `window_hours` is "X" from the definition, measured from local midnight.
    `lookback_hours` extends only the fetch, never the window: without it a
    truck shut down at 20:00 yesterday has no state until its first event
    today, and a vehicle that was already moving at midnight loses its
    odometer baseline.
    """
    start_local = datetime(day.year, day.month, day.day, tzinfo=zone)
    end_local = start_local + timedelta(hours=window_hours)
    fmt = lambda m: m.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")  # noqa: E731
    return (
        fmt(start_local - timedelta(hours=lookback_hours)),
        fmt(end_local),
        start_local,
        end_local,
    )


def parse_ts(text):
    """An RFC3339 stamp as an aware datetime, or None."""
    if not text:
        return None
    raw = str(text).strip().replace("Z", "+00:00")
    try:
        moment = datetime.fromisoformat(raw)
    except ValueError:
        return None
    return moment if moment.tzinfo else moment.replace(tzinfo=timezone.utc)


# =============================
# JOIN
# =============================
def normalize_vin(value):
    """
    A VIN reduced to what is comparable between the two systems.

    Case and stray punctuation differ between hand-entered and telematics-fed
    records; nothing else about a VIN is safe to alter. Same rule as
    probe_samsara_tags.report_vin_join, so both agree on what matches.
    """
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def normalize_name(value):
    """A unit name reduced to what is comparable. Case and separators drift."""
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def build_samsara_index(ts_vehicles, sam_vehicles):
    """
    Traumasoft vehicle id -> (samsara vehicle id, how it matched).

    VIN first: it is the only identifier both systems hold independently, so
    it cannot drift when somebody renames a unit. Name second, because some
    Traumasoft records carry no VIN and dropping them silently would shrink
    the fleet without saying so. How each one matched is carried through to
    the output so a name match can be audited.

    A VIN or name that resolves to more than one Samsara vehicle is left
    unmatched rather than guessed at -- a wrong join produces a number that
    looks right.
    """
    sam_by_vin, sam_by_name = defaultdict(list), defaultdict(list)
    for vehicle in sam_vehicles:
        vin = normalize_vin(vehicle.get("vin"))
        if len(vin) == 17:
            sam_by_vin[vin].append(vehicle)
        name = normalize_name(vehicle.get("name"))
        if name:
            sam_by_name[name].append(vehicle)

    index, ambiguous = {}, []
    for vehicle in ts_vehicles:
        key = str(vehicle.get("id"))
        vin = normalize_vin(vehicle.get("vin"))
        if len(vin) == 17 and vin in sam_by_vin:
            rows = sam_by_vin[vin]
            if len(rows) == 1:
                index[key] = (str(rows[0].get("id")), "vin")
                continue
            ambiguous.append((vehicle.get("name"), f"VIN {vin} matches {len(rows)} Samsara vehicles"))
            continue
        name = normalize_name(vehicle.get("name"))
        if name and name in sam_by_name:
            rows = sam_by_name[name]
            if len(rows) == 1:
                index[key] = (str(rows[0].get("id")), "name")
                continue
            ambiguous.append((vehicle.get("name"), f"name matches {len(rows)} Samsara vehicles"))
    return index, ambiguous


# =============================
# SAMSARA FACTS
# =============================
def fetch_stats(samsara, start, end, types):
    """One stats-history query over a window, all vehicles."""
    return list(
        samsara.paginate(
            STATS_PATH, params={"startTime": start, "endTime": end, "types": types}
        )
    )


def engine_on_in_window(events, window_start, window_end):
    """
    Whether the engine was on at any point in the window.

    Returns True, False, or None. None is not False: a vehicle with no events
    and no carried-in state is a gateway that said nothing, and counting that
    as engine-off would quietly mark working trucks unproductive. The state in
    force at the window's start counts -- a truck already running at midnight
    was running in the window even if it reports nothing until 06:00.
    """
    timeline = []
    for event in events or []:
        moment = parse_ts(event.get("time"))
        if moment:
            timeline.append((moment, str(event.get("value") or "").lower()))
    timeline.sort()

    carried = None
    for moment, value in timeline:
        if moment <= window_start:
            carried = value
        else:
            break
    inside = [value for moment, value in timeline if window_start < moment < window_end]

    if carried is None and not inside:
        return None
    states = ([carried] if carried is not None else []) + inside
    return any(state not in OFF_VALUES for state in states)


def _odometer_readings(row):
    """(time, meters) pairs from whichever odometer series the row carries."""
    readings = []
    for key in ("obdOdometerMeters", "gpsOdometerMeters"):
        for entry in row.get(key) or []:
            moment = parse_ts(entry.get("time"))
            value = entry.get("value")
            if moment is not None and isinstance(value, (int, float)):
                readings.append((moment, float(value)))
        if readings:
            # Don't mix the two series: they are different instruments and
            # their absolute values differ, so a delta across both is noise.
            break
    readings.sort()
    return readings


def miles_from_odometer(row, window_start, window_end):
    """
    Miles covered in the window from the odometer, or None.

    The baseline is the last reading at or before the window opens, so a truck
    that moved before its first in-window sample is not credited from zero.
    Fewer than two usable points means the odometer said nothing about this
    window -- that is None, never 0.0.
    """
    readings = _odometer_readings(row)
    if not readings:
        return None
    baseline = None
    for moment, value in readings:
        if moment <= window_start:
            baseline = value
        else:
            break
    inside = [v for m, v in readings if window_start < m <= window_end]
    points = ([baseline] if baseline is not None else []) + inside
    if len(points) < 2:
        return None
    # Odometers are monotonic; max-min survives an out-of-order sample in a
    # way that last-minus-first does not.
    return (max(points) - min(points)) * MILES_PER_METER


def haversine_miles(lat1, lon1, lat2, lon2):
    """Great-circle distance between two fixes, in miles."""
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    d_phi = phi2 - phi1
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    return 2 * EARTH_RADIUS_MILES * math.asin(min(1.0, math.sqrt(a)))


def miles_from_gps(row, window_start, window_end):
    """
    Miles covered in the window as the length of the GPS track, or None.

    This is a lower bound and should be read as one: the track is a polyline
    through the samples, so a gap between fixes cuts the corner off every turn
    taken inside it. Fewer than two in-window fixes is None, not 0.0 -- one
    fix says where a truck was, not that it stayed there.
    """
    fixes = []
    for entry in row.get("gps") or []:
        moment = parse_ts(entry.get("time"))
        lat, lon = entry.get("latitude"), entry.get("longitude")
        if moment is None or lat is None or lon is None:
            continue
        if window_start < moment <= window_end:
            fixes.append((moment, float(lat), float(lon)))
    if len(fixes) < 2:
        return None
    fixes.sort()
    total = 0.0
    for (_, lat1, lon1), (_, lat2, lon2) in zip(fixes, fixes[1:]):
        total += haversine_miles(lat1, lon1, lat2, lon2)
    return total


def distance_miles(row, window_start, window_end, source):
    """(miles, source_used). Returns (None, "none") when nothing could be read."""
    if source == "gps":
        miles = miles_from_gps(row, window_start, window_end)
        return (miles, "gps") if miles is not None else (None, "none")
    miles = miles_from_odometer(row, window_start, window_end)
    return (miles, "odometer") if miles is not None else (None, "none")


# =============================
# TRAUMASOFT FACTS
# =============================
def _is_truthy_flag(value):
    """
    Whether an API boolean is set.

    These come back as real booleans on some endpoints and as "0"/"1" or
    "true"/"false" strings on others, so a bare truth test would treat the
    string "0" as set. Same rule as traumasoft_reports._is_truthy_flag.
    """
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in ("1", "true", "yes", "y")


def out_of_service(vehicle):
    """
    Whether the vehicle carries an out-of-service label.

    IMPORTANT: this is the status the fleet list reports NOW, not the status
    the vehicle had on the analysed day. Traumasoft's ThirdParty API offers no
    status history, so for any day but today this is an anachronism -- a truck
    that went down this morning reads as out of service for last Tuesday too.
    `--oos-history` reads the observations traumasoft_reports accumulates in
    state/vehicle_oos_history.json, which is the only per-day answer available
    and only reaches back to when that file started.
    """
    if _is_truthy_flag(vehicle.get("deleted")) or _is_truthy_flag(vehicle.get("disabled")):
        return True
    return str(vehicle.get("vehicle_status") or "").strip() in OUT_OF_SERVICE_STATUSES


def localize_shift_ts(value, zone, shifts_are_utc=True):
    """
    A shift stamp as an aware instant, comparable with the analysed window.

    Traumasoft returns /Schedule/Shifts start and end as BARE UTC wall time,
    while trips return local time with an explicit offset -- the discrepancy
    traumasoft_reports documents at SHIFT_TIMES_ARE_UTC. Left unreconciled,
    every shift lands in the wrong day by the tenant's offset, which on an
    overnight roster moves whole crews across the date boundary.

    So a bare stamp is read as UTC, and only read as tenant-local if the
    tenant is known to have changed (TS_SHIFT_TIMES_ARE_UTC=0). A stamp that
    already carries an offset is trusted as it stands.
    """
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        moment = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if moment.tzinfo is not None:
        return moment
    return moment.replace(tzinfo=timezone.utc) if shifts_are_utc \
        else moment.replace(tzinfo=zone)


def crew_by_vehicle(shifts, window_start, window_end, zone, shifts_are_utc=True):
    """
    normalized vehicle name -> {"crew": n, "punched": n, "shifts": n}.

    The shift feed returns one row per crew member, so a two-medic truck
    arrives twice; that is counted as two crew on one unit, which is what
    "has assigned crew" asks. `punched` counts the crew who actually clocked
    in, carried alongside rather than substituted for the assignment: the
    definition says assigned, and a crew that no-showed is a different
    finding from a truck nobody was rostered to.

    /Schedule/Shifts takes no date filter, so the whole feed comes back and
    is narrowed here by overlap with the window.
    """
    counts = defaultdict(lambda: {"crew": 0, "punched": 0, "shifts": 0})
    for shift in shifts or []:
        if _is_truthy_flag(shift.get("deleted")):
            continue
        name = normalize_name(shift.get("vehicle_name"))
        if not name:
            continue
        start = localize_shift_ts(shift.get("start_time"), zone, shifts_are_utc)
        end = localize_shift_ts(shift.get("end_time"), zone, shifts_are_utc)
        if start is None or end is None or end <= start:
            continue
        if end <= window_start or start >= window_end:
            continue
        entry = counts[name]
        entry["shifts"] += 1
        entry["crew"] += 1
        if any(
            p.get("start_time") and not _is_truthy_flag(p.get("deleted"))
            for p in shift.get("punches") or []
        ):
            entry["punched"] += 1
    return dict(counts)


CANCELLED_STATUSES = {"canceled", "cancelled", "disregard", "no transport"}


def legs_by_vehicle(legs, exclude_cancelled=False):
    """
    Traumasoft vehicle id -> {"legs": n, "cancelled": n}.

    "Has scheduled pickup" is read as "a call was assigned to this unit". A
    cancelled call still had a scheduled pickup, so it counts by default and
    is carried in its own column; `--exclude-cancelled` reads the definition
    the other way for whoever asks.
    """
    counts = defaultdict(lambda: {"legs": 0, "cancelled": 0})
    for leg in legs or []:
        vehicle_id = leg.get("vehicle_id")
        if not vehicle_id:
            continue
        cancelled = str(leg.get("trip_status") or "").strip().lower() in CANCELLED_STATUSES
        if cancelled and exclude_cancelled:
            continue
        entry = counts[str(vehicle_id)]
        entry["legs"] += 1
        if cancelled:
            entry["cancelled"] += 1
    return dict(counts)


# =============================
# CLASSIFICATION
# =============================
def failed_conditions(factors):
    """
    Every Productive condition known to be false, earliest break first.

    Separate from `classify` so the verdict can name one thing while the row
    still carries all of them.
    """
    return [name for name in CONDITION_PRIORITY if factors.get(name) is False]


def classify(factors):
    """
    (classification, reason) for one vehicle-day.

    Three-valued throughout: every factor is True, False, or None for "could
    not be read", and None never collapses into False.

    A KNOWN FAILURE BEATS AN UNKNOWN. Any Productive condition that is
    positively false makes the vehicle Non-Productive, whatever else could not
    be read -- a truck with no call assigned cannot become Productive on the
    strength of telematics nobody has, so the verdict does not wait for them.
    That is what keeps Samsara being down from turning the whole fleet into
    undetermined rows.

    The definition as first written stated Non-Productive as only "no crew or
    out of service", which left two real cases matching no rule at all: a
    crewed, in-service truck that moved with no call assigned, and one that
    never started. Management settled both as Non-Productive, so the rule is
    now "any Productive condition known to fail". The specific condition is
    carried in the reason, because those two cases are different operational
    problems and the totals should still be able to tell them apart.
    """
    if factors.get("out_of_service") is True:
        return NON_PRODUCTIVE, NON_PRODUCTIVE_REASONS["out_of_service"]

    failed = failed_conditions(factors)
    if failed:
        return NON_PRODUCTIVE, NON_PRODUCTIVE_REASONS[failed[0]]

    unreadable = [
        name for name in CONDITION_PRIORITY if factors.get(name) is None
    ]
    if unreadable:
        return UNDETERMINED, "could not read " + ", ".join(
            {
                "engine_on": "engine state",
                "moved": "distance travelled",
                "has_crew": "crew assignment",
                "has_scheduled_pickup": "call assignment",
            }[name]
            for name in unreadable
        )

    # All four hold. ePCR alone decides between Productive and Indetermined,
    # and it is the one input this API key cannot reach.
    epcr = factors.get("epcr_complete")
    if epcr is True:
        return PRODUCTIVE, "all conditions met"
    if epcr is False:
        return INDETERMINATE, "all conditions met except a completed ePCR"
    return UNDETERMINED, (
        "productive or indetermined -- every other condition holds, and ePCR "
        "completion is not readable through this API"
    )


PENDING_EPCR_REASON_PREFIX = "productive or indetermined"


# =============================
# ASSEMBLY
# =============================
def build_rows(ts_vehicles, legs, shifts, stats_rows, sam_index, window_start,
               window_end, zone, min_miles=1.0, distance_source="odometer",
               shifts_are_utc=True, exclude_cancelled=False,
               oos_history=None, day=None, include_non_fleet=False):
    """
    One row per in-scope vehicle, with every factor and its provenance.

    Retired and not-yet-delivered vehicles are dropped by default: they are
    not a productivity question, and leaving them in buries the fleet in
    non-productive rows that nobody can act on.
    """
    crew = crew_by_vehicle(shifts, window_start, window_end, zone, shifts_are_utc)
    calls = legs_by_vehicle(legs, exclude_cancelled)
    stats_by_id = {str(row.get("id")): row for row in stats_rows or []}
    # A shift naming a unit the fleet list does not carry is a real finding,
    # not a row to drop silently.
    seen_crew_names = set()

    rows, collisions = [], defaultdict(list)
    for vehicle in ts_vehicles:
        status = str(vehicle.get("vehicle_status") or "").strip()
        if not include_non_fleet and status in NON_FLEET_STATUSES:
            continue
        # The list call asks the server to omit deleted and disabled records
        # and the server does not honour it -- probe_samsara_tags measured 31
        # deleted records coming back on this tenant despite
        # include_deleted=false. Left in, each one became a second row for a
        # truck that already had one: A-101 appearing as both In Service and
        # Out of Service, counted twice, one of them non-productive. Same
        # filter chain as traumasoft_reports._fleet_partition.
        if _is_truthy_flag(vehicle.get("deleted")) \
                or _is_truthy_flag(vehicle.get("disabled")):
            continue

        ts_id = str(vehicle.get("id"))
        name = vehicle.get("name")
        norm = normalize_name(name)
        if norm in collisions:
            # Two live records for one unit name. The filter above resolves
            # every case this tenant has, so reaching here means a new one --
            # reported rather than silently doubling the fleet.
            collisions[norm].append(name)
            continue
        collisions[norm].append(name)
        seen_crew_names.add(norm)

        crew_entry = crew.get(norm, {"crew": 0, "punched": 0, "shifts": 0})
        call_entry = calls.get(ts_id, {"legs": 0, "cancelled": 0})

        sam_id, join = sam_index.get(ts_id, (None, "unmatched"))
        stats = stats_by_id.get(sam_id) if sam_id else None
        if stats is None:
            engine, miles, source = None, None, "none"
        else:
            engine = engine_on_in_window(stats.get("engineStates"), window_start, window_end)
            miles, source = distance_miles(stats, window_start, window_end, distance_source)

        oos = out_of_service(vehicle)
        oos_source = "current status"
        if oos_history is not None and day is not None:
            historical = oos_history.get(ts_id)
            if historical is not None:
                oos, oos_source = historical, "observed history"

        factors = {
            "engine_on": engine,
            "moved": None if miles is None else miles > min_miles,
            "has_crew": crew_entry["crew"] > 0,
            "has_scheduled_pickup": call_entry["legs"] > 0,
            # Not readable. Kept as an explicit key so the day credentials for
            # the Epcr/Huly surface arrive, one assignment turns it on.
            "epcr_complete": None,
            "out_of_service": oos,
        }
        classification, reason = classify(factors)
        failed = failed_conditions(factors)
        if factors["out_of_service"]:
            failed = ["out_of_service"] + failed

        rows.append({
            "day": day.isoformat() if day else None,
            "vehicle_id": ts_id,
            "vehicle_name": name,
            "vehicle_status": status,
            "classification": classification,
            "reason": reason,
            "failed_conditions": ";".join(failed),
            "engine_on": factors["engine_on"],
            "miles": None if miles is None else round(miles, 2),
            "moved": factors["moved"],
            "miles_source": source,
            "crew_assigned": crew_entry["crew"],
            "crew_punched": crew_entry["punched"],
            "legs_assigned": call_entry["legs"],
            "legs_cancelled": call_entry["cancelled"],
            "epcr_complete": None,
            "out_of_service": oos,
            "out_of_service_source": oos_source,
            "samsara_join": join,
            "samsara_id": sam_id,
        })

    orphan_crew = sorted(set(crew) - seen_crew_names)
    duplicated = {k: v for k, v in collisions.items() if len(v) > 1}
    if duplicated:
        log.warning(
            "%s unit name(s) had more than one live vehicle record; the first "
            "was kept and the rest dropped: %s",
            len(duplicated),
            ", ".join(sorted(duplicated)),
        )
    return rows, orphan_crew


def load_oos_history(path, day):
    """
    vehicle id -> whether it was observed out of service on `day`.

    Reads the file traumasoft_reports accumulates. Only vehicles the history
    actually covers are returned, so a vehicle missing from it falls back to
    current status rather than being assumed in service.
    """
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, ValueError) as exc:
        log.warning("Could not read %s (%s); falling back to current status.", path, exc)
        return None

    since = payload.get("since") if isinstance(payload, dict) else None
    if not isinstance(since, dict):
        log.warning("%s has no 'since' map; falling back to current status.", path)
        return None

    target = day.isoformat()
    observed = {}
    for vehicle_id, started in since.items():
        if isinstance(started, str) and started <= target:
            observed[str(vehicle_id)] = True
    return observed or None


# =============================
# OUTPUT
# =============================
def print_report(rows, orphan_crew, ambiguous, day, min_miles, distance_source,
                 window_hours):
    counts = Counter(row["classification"] for row in rows)
    pending = [
        row for row in rows
        if row["classification"] == UNDETERMINED
        and row["reason"].startswith(PENDING_EPCR_REASON_PREFIX)
    ]

    print("\n1. WHAT THE RULES SAY ABOUT %s" % day.isoformat())
    print("   " + "-" * 66)
    print(f"   {'classification':<20}{'vehicles':>9}")
    print("   " + "-" * 31)
    for name in CLASS_ORDER:
        print(f"   {name:<20}{counts.get(name, 0):>9}")
    print("   " + "-" * 31)
    print(f"   {'total':<20}{len(rows):>9}")

    print("\n2. THE ePCR HOLE")
    print("   " + "-" * 66)
    if counts.get(PRODUCTIVE, 0) or counts.get(INDETERMINATE, 0):
        print("   ePCR completion is being read from somewhere. Check where --")
        print("   this repository has no route to it.")
    else:
        print("   No vehicle can be called Productive and none can be called")
        print("   Indetermined, because both verdicts turn on whether the ePCR")
        print("   was completed and this API key cannot read that.")
    print(f"\n   {len(pending)} vehicle(s) meet every other Productive condition.")
    print("   Each is Productive or Indetermined; which one is unknowable")
    print("   until credentials for ThirdParty/Data/Epcr/Huly exist.")
    for row in pending[:20]:
        print(f"      {str(row['vehicle_name'])[:20]:<22}"
              f"{row['crew_assigned']} crew  "
              f"{row['legs_assigned']} call(s)  "
              f"{row['miles']} mi")
    if len(pending) > 20:
        print(f"      ... and {len(pending) - 20} more")

    print("\n3. WHY THE NON-PRODUCTIVE ONES ARE NON-PRODUCTIVE")
    print("   " + "-" * 66)
    unproductive = [r for r in rows if r["classification"] == NON_PRODUCTIVE]
    if not unproductive:
        print("   None.")
    else:
        print("   One verdict, several different problems. Grouped by why:\n")
        for reason, count in Counter(
            r["reason"] for r in unproductive
        ).most_common():
            wrapped = textwrap.wrap(reason, 56)
            print(f"      {count:>4}  {wrapped[0]}")
            for line in wrapped[1:]:
                print(f"            {line}")
        moved_undispatched = [
            r for r in unproductive
            if NON_PRODUCTIVE_REASONS["has_scheduled_pickup"] in r["reason"]
            and r["moved"] is True and r["crew_assigned"] > 0
        ]
        if moved_undispatched:
            plural = "was" if len(moved_undispatched) == 1 else "were"
            print(f"\n   {len(moved_undispatched)} of those {plural} crewed and "
                  f"moved more than {min_miles} mile")
            print("   with no call assigned -- posting, repositioning or a")
            print("   maintenance run. Crew hours and fuel spent, nothing")
            print("   dispatched. The current Daily Vehicle Overview cannot see")
            print("   these at all; it reads them as simply unused.")
            for row in moved_undispatched[:15]:
                print(f"      {str(row['vehicle_name'])[:20]:<22}"
                      f"{row['crew_assigned']} crew  {row['miles']} mi")
            if len(moved_undispatched) > 15:
                print(f"      ... and {len(moved_undispatched) - 15} more")

    print("\n4. WHAT COULD NOT BE READ")
    print("   " + "-" * 66)
    undetermined = [
        r for r in rows if r["classification"] == UNDETERMINED
        and not r["reason"].startswith(PENDING_EPCR_REASON_PREFIX)
    ]
    if not undetermined:
        print("   Nothing beyond ePCR.")
    for reason, count in Counter(r["reason"] for r in undetermined).most_common():
        wrapped = textwrap.wrap(reason, 56)
        print(f"      {count:>4}  {wrapped[0]}")
        for line in wrapped[1:]:
            print(f"            {line}")

    no_telematics = [r for r in rows if r["samsara_join"] == "unmatched"]
    by_name = [r for r in rows if r["samsara_join"] == "name"]
    no_distance = [r for r in rows if r["samsara_id"] and r["miles_source"] == "none"]
    print(f"\n   Vehicles with no Samsara match:   {len(no_telematics)}")
    if no_telematics:
        print("      " + ", ".join(
            str(r["vehicle_name"]) for r in no_telematics[:12]
        )[:200])
        print("      These need a VIN in Traumasoft, or they can never be")
        print("      Productive -- engine state and distance are unreadable.")
    print(f"   Matched on name, not VIN:         {len(by_name)}")
    if by_name:
        print("      Auditable but drift-prone. A renamed unit silently")
        print("      unmatches; a VIN does not.")
    print(f"   Matched but no {distance_source} signal: {len(no_distance)}")
    if no_distance and distance_source == "odometer":
        print("      Try --distance-source gps. Run probe_samsara_movement.py")
        print("      to see which series this fleet actually populates.")

    if ambiguous:
        print(f"\n   Ambiguous joins, left unmatched ({len(ambiguous)}):")
        for name, why in ambiguous[:12]:
            print(f"      {str(name)[:22]:<24}{why}")

    if orphan_crew:
        print(f"\n   Shifts naming a unit the fleet list does not carry "
              f"({len(orphan_crew)}):")
        print("      " + ", ".join(orphan_crew[:12])[:200])
        print("      Crew rostered to these contributes to no vehicle's row.")

    print("\n5. HOW TO READ THIS")
    print("   " + "-" * 66)
    print(f"   Unit of analysis  one vehicle, one tenant-local day")
    print(f"   Window            {window_hours}h from local midnight")
    print(f"   Movement          > {min_miles} mile, from {distance_source}")
    print("   Out of service    the status Traumasoft reports; for a past day")
    print("                     this is an anachronism unless --oos-history")
    print("                     covers the vehicle. See the column.")
    print("   ePCR              never read. Always unknown.")


def write_csv(path, rows):
    if not rows:
        log.warning("No rows to write to %s.", path)
        return
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


# =============================
# CLI
# =============================
def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--day", help="Day to classify, YYYY-MM-DD. Default yesterday.")
    parser.add_argument("--window-hours", type=float, default=24.0,
                        help='"X" from the definition, hours from local midnight '
                             "(default 24, a full calendar day).")
    parser.add_argument("--min-miles", type=float, default=1.0,
                        help="Miles a vehicle must exceed to count as moved "
                             "(default 1).")
    parser.add_argument("--distance-source", choices=("odometer", "gps"),
                        default="odometer",
                        help="Which Samsara series measures distance. Run "
                             "probe_samsara_movement.py to see which this fleet "
                             "populates.")
    parser.add_argument("--exclude-cancelled", action="store_true",
                        help="Do not count a cancelled call as a scheduled pickup.")
    parser.add_argument("--include-non-fleet", action="store_true",
                        help="Also classify retired and undelivered vehicles.")
    parser.add_argument("--oos-history", nargs="?", const="state/vehicle_oos_history.json",
                        help="Read out-of-service status for the day from the "
                             "accumulated observations instead of today's status.")
    parser.add_argument("--csv", dest="csv_path", help="Write every row as CSV.")
    parser.add_argument("--json", dest="json_path", help="Write every row as JSON.")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    day = date.fromisoformat(args.day) if args.day else date.today() - timedelta(days=1)

    try:
        zone, zone_label = tenant_zone()
    except RuntimeError as exc:
        log.error("%s", exc)
        return 2

    fetch_from, fetch_to, window_start, window_end = day_window(
        day, zone, window_hours=args.window_hours
    )

    print("=" * 72)
    print("  Vehicle productivity -- against the definition as given")
    print(f"  Day {day}   zone {zone_label}   window {args.window_hours}h")
    print("  Read-only. GET only, no writes. No patient data.")
    print("=" * 72)

    # -- Traumasoft: crew, calls, fleet, status --
    try:
        import traumasoft_reports as R
        from traumasoft_api import TraumasoftAPI
        api = TraumasoftAPI()
        legs = api.get_trips(day.isoformat(), range_days=1)
        shifts = api.list_shifts()
        ts_vehicles = api.list_vehicles()
    except Exception as exc:  # noqa: BLE001 -- reporting the failure IS the result
        log.error("Traumasoft unavailable (%s: %s). Nothing can be classified.",
                  type(exc).__name__, exc)
        return 2

    # -- Samsara: engine state and distance --
    types = "engineStates,gps" if args.distance_source == "gps" \
        else "engineStates,obdOdometerMeters,gpsOdometerMeters"
    stats_rows, sam_vehicles = [], []
    try:
        from samsara_api import SamsaraClient
        samsara = SamsaraClient(read_only=True)
        sam_vehicles = samsara.list_vehicles()
        stats_rows = fetch_stats(samsara, fetch_from, fetch_to, types)
    except Exception as exc:  # noqa: BLE001
        log.error("Samsara unavailable (%s: %s).", type(exc).__name__, exc)
        log.error("Engine state and distance will read as unknown for every "
                  "vehicle, so nothing can be called Productive. Traumasoft's "
                  "half of the definition is still reported.")

    sam_index, ambiguous = build_samsara_index(ts_vehicles, sam_vehicles)
    oos_history = load_oos_history(args.oos_history, day) if args.oos_history else None

    rows, orphan_crew = build_rows(
        ts_vehicles, legs, shifts, stats_rows, sam_index, window_start, window_end,
        zone, min_miles=args.min_miles, distance_source=args.distance_source,
        shifts_are_utc=R.SHIFT_TIMES_ARE_UTC,
        exclude_cancelled=args.exclude_cancelled,
        oos_history=oos_history, day=day, include_non_fleet=args.include_non_fleet,
    )

    print_report(rows, orphan_crew, ambiguous, day, args.min_miles,
                 args.distance_source, args.window_hours)

    print("\n" + "=" * 72)
    print("  Safe to paste. Operational values only, no patient data.")
    print("=" * 72 + "\n")

    if args.csv_path:
        write_csv(args.csv_path, rows)
        print(f"  Written to {args.csv_path}")
    if args.json_path:
        with open(args.json_path, "w", encoding="utf-8") as handle:
            json.dump(rows, handle, indent=2, default=str)
        print(f"  Written to {args.json_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
