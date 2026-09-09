"""
Turn Traumasoft trip legs into Samsara route payloads.

Every function here is pure: legs in, payloads out, no network. The client
lives in samsara_api.py and the CLI in push_samsara_routes.py, so the mapping
can be tested against fixture legs without a token.

THE SHAPE
    One Samsara route per vehicle per day. Each eligible leg contributes two
    stops -- pickup, then drop-off -- and Samsara orders the whole list by
    scheduledArrivalTime, so the driver sees the day in sequence.

THE JOIN
    Traumasoft `vehicle_name` and the Samsara vehicle name share a unit
    prefix, but not a format: 'M-12', 'M12' and 'm 12' are one unit typed
    three ways. unit_prefix() collapses all of them to 'M12' so the two
    systems agree, and anything that still will not match is reported rather
    than silently dropped.
"""

import json
import logging
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone

try:  # stdlib on 3.9+, but a machine can be missing the tzdata package
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None

log = logging.getLogger(__name__)


def normalize_los(value):
    """
    A level of service reduced to a comparable key.

    Real data carries 'Non Emergency' and 'Non-Emergency' as one thing typed
    two ways, and 'Non Amnbulatory' as a straight typo. Case, hyphens and
    runs of whitespace are collapsed so a rule written once catches the
    variants; genuine misspellings still need naming individually.
    """
    text = str(value or "").strip().lower().replace("-", " ")
    return " ".join(text.split())


STATE_DIR = os.getenv("TS_STATE_DIR", "state")
VEHICLE_OVERRIDES_FILE = os.path.join(STATE_DIR, "samsara_vehicle_overrides.json")

# When a leg has no appointment time and no drop-off ETA there is nothing that
# says when it should arrive, but Samsara requires a scheduledArrivalTime on
# every stop after the first. Rather than drop the leg, the drop-off is placed
# this many minutes after the pickup and the plan reports it as estimated.
#
# Two thirds of real legs land here, so the number matters. Calibrate it from
# the legs that DO carry an appointment time --
# `probe_samsara_readiness.py --days 30` prints the observed interval.
DEFAULT_TRANSPORT_MINUTES = int(os.getenv("SAMSARA_DEFAULT_TRANSPORT_MINUTES", "45"))

# Per level of service, as 'ambulatory=25,non ambulatory=50'. A wheelchair
# van's day and an ALS transport's are not the same shape, and one number for
# both is wrong for both. Keys are normalised the same way as EXCLUDED_LOS,
# so the Non Emergency / Non-Emergency split does not need writing twice.
TRANSPORT_MINUTES_BY_LOS = {}
for _pair in os.getenv("SAMSARA_TRANSPORT_MINUTES_BY_LOS", "").split(","):
    if "=" not in _pair:
        continue
    _key, _, _value = _pair.partition("=")
    try:
        TRANSPORT_MINUTES_BY_LOS[normalize_los(_key)] = int(_value.strip())
    except ValueError:
        log.warning("Ignoring unparseable SAMSARA_TRANSPORT_MINUTES_BY_LOS entry %r", _pair)

# Samsara sorts stops by scheduledArrivalTime. A drop-off stamped at or before
# its own pickup would jump ahead of it, so the drop-off is pushed at least
# this far past the pickup.
MIN_STOP_GAP_MINUTES = 1

# Whether the patient's name rides along in the pickup stop notes. OFF by
# default: a name plus a pickup address is PHI, and pushing it into Samsara
# discloses it to a third party, which is only defensible if Samsara is
# covered by a business associate agreement. Crews often do want it to confirm
# they have the right patient -- so it is one env var away, as a decision
# somebody makes rather than a default they inherit.
INCLUDE_PATIENT_NAME = os.getenv("SAMSARA_INCLUDE_PATIENT_NAME", "").strip().lower() in (
    "1", "true", "yes", "y",
)

# The tenant's time zone as an IANA name, e.g. 'America/New_York'. Preferred
# over a fixed offset because it follows daylight saving on its own, resolved
# against each trip's own date. Needed because this tenant sends naive trip
# times and leaves the per-leg `timezone` field empty.
TENANT_TIMEZONE = os.getenv("SAMSARA_TENANT_TIMEZONE", "").strip()

# A fixed UTC offset, '-04:00' or '-4'. Works, but is wrong for half the year
# unless somebody remembers to change it twice. Prefer SAMSARA_TENANT_TIMEZONE.
TENANT_UTC_OFFSET = os.getenv("SAMSARA_TENANT_UTC_OFFSET", "").strip()

# Levels of service never pushed. Emergency work is dispatched in the moment,
# so a route built for it the night before describes nothing. Empty by
# default: whether that is true of a given operation is its own call, and
# probe_samsara_readiness.py prints the real vocabulary to decide against.
EXCLUDED_LOS = tuple(
    normalize_los(v) for v in os.getenv("SAMSARA_EXCLUDED_LOS", "").split(",") if v.strip()
)

# Call types that describe something other than a transport a driver can be
# routed to. Matched case-insensitively as substrings of call_type.
EXCLUDED_CALL_TYPE_PATTERNS = tuple(
    p.strip().lower()
    for p in os.getenv("SAMSARA_EXCLUDED_CALL_TYPES", "standby,cancel,no transport,dry run").split(",")
    if p.strip()
)

# A unit designator: letters, an optional separator, then digits.
# 'M-12' / 'M 12' / 'M12' -> ('M', '12'). Trailing letters are kept so
# 'M12A' and 'M12B' stay distinct units.
_UNIT_RE = re.compile(r"^([A-Z]+)[\s\-_./]*(\d+)([A-Z]*)")


# =============================
# VEHICLE MATCHING
# =============================
def unit_prefix(name):
    """
    The comparable unit designator inside a vehicle name.

    'M-12 Ford E450' -> 'M12'; 'Medic 12' -> 'MEDIC12'; 'WC-3' -> 'WC3'.
    Returns None when the name carries no letter+digit designator at all,
    which is the signal to fall back to an override rather than guess.
    """
    if not name:
        return None
    text = str(name).strip().upper()
    if not text:
        return None
    match = _UNIT_RE.match(text)
    if match:
        letters, digits, suffix = match.groups()
        # Drop leading zeros so 'M-012' and 'M12' agree.
        return f"{letters}{int(digits)}{suffix}"
    # No designator: fall back to the first token, which at least lets an
    # exact name match work.
    token = re.split(r"[\s\-_./]+", text, maxsplit=1)[0]
    return token or None


def load_vehicle_overrides(path=None):
    """
    Explicit Traumasoft-name -> Samsara-name pairs for units the prefix rule
    cannot join, kept in state/ because it is local knowledge the API cannot
    reproduce.
    """
    path = path or VEHICLE_OVERRIDES_FILE
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            raw = json.load(handle)
    except (OSError, ValueError) as exc:
        log.warning("Could not read %s (%s); continuing without overrides", path, exc)
        return {}
    mapping = raw.get("overrides", raw) if isinstance(raw, dict) else {}
    return {
        str(k).strip().upper(): str(v).strip()
        for k, v in mapping.items()
        if k and v and not str(k).startswith("_")
    }


def index_samsara_vehicles(vehicles):
    """
    {unit prefix: [vehicle, ...]} plus an exact-name index.

    Prefixes that land on more than one Samsara vehicle are kept as lists so
    match_vehicles() can report the ambiguity instead of picking one.
    """
    by_prefix = defaultdict(list)
    by_name = {}
    for vehicle in vehicles:
        name = (vehicle.get("name") or "").strip()
        if not name:
            continue
        by_name[name.upper()] = vehicle
        prefix = unit_prefix(name)
        if prefix:
            by_prefix[prefix].append(vehicle)
    return dict(by_prefix), by_name


def match_vehicles(traumasoft_names, samsara_vehicles, overrides=None):
    """
    Join Traumasoft vehicle names to Samsara vehicles.

    Returns (matched, unmatched, ambiguous):
        matched   {traumasoft name: samsara vehicle}
        unmatched [traumasoft name, ...]        -- no Samsara vehicle found
        ambiguous {traumasoft name: [vehicle]}  -- prefix hit several
    """
    overrides = overrides or {}
    by_prefix, by_name = index_samsara_vehicles(samsara_vehicles)

    matched, unmatched, ambiguous = {}, [], {}
    for ts_name in traumasoft_names:
        if not ts_name:
            continue
        key = str(ts_name).strip().upper()

        # 1. An explicit override always wins.
        target = overrides.get(key)
        if target:
            vehicle = by_name.get(target.upper())
            if vehicle:
                matched[ts_name] = vehicle
                continue
            hits = by_prefix.get(unit_prefix(target) or "", [])
            if len(hits) == 1:
                matched[ts_name] = hits[0]
                continue
            unmatched.append(ts_name)
            continue

        # 2. Identical names.
        if key in by_name:
            matched[ts_name] = by_name[key]
            continue

        # 3. The shared unit prefix.
        prefix = unit_prefix(ts_name)
        hits = by_prefix.get(prefix or "", [])
        if len(hits) == 1:
            matched[ts_name] = hits[0]
        elif len(hits) > 1:
            ambiguous[ts_name] = hits
        else:
            unmatched.append(ts_name)

    return matched, sorted(unmatched), ambiguous


# =============================
# TIME ZONE
# =============================
def parse_offset_string(text):
    """'-04:00', '-4', '+5:30' -> a timedelta. None if unparseable."""
    text = str(text or "").strip()
    if not text:
        return None
    match = re.fullmatch(r"([+-]?)(\d{1,2})(?::?(\d{2}))?", text)
    if not match:
        return None
    sign, hours, minutes = match.groups()
    delta = timedelta(hours=int(hours), minutes=int(minutes or 0))
    return -delta if sign == "-" else delta


def format_offset(delta):
    """A timedelta as '-04:00'. str() renders -4h as '-1 day, 20:00:00'."""
    if delta is None:
        return "unknown"
    total = int(delta.total_seconds())
    sign = "-" if total < 0 else "+"
    total = abs(total)
    return f"{sign}{total // 3600:02d}:{(total % 3600) // 60:02d}"


def resolve_tenant_offset(legs, parse_ts_aware):
    """
    The UTC offset to stamp on a naive trip time, and where it came from.

    This matters more than it looks. On this tenant `pickup_time` is naive on
    every leg -- not one carries an offset -- so something has to supply it,
    and getting it wrong is not a subtle error: it dispatches a 07:30 pickup
    at 03:30 local.

    The obvious source, the offset on the status timestamps, is exactly the
    one that fails for the job's main use. A trip scheduled for tomorrow has
    no status stamps yet, because nothing has happened to it, so a day of
    future work carries no offset anywhere in it. Hence the chain:

        1. SAMSARA_TENANT_TIMEZONE, an IANA zone resolved against the trip's
           own date, so daylight saving is handled without anyone's diary
        2. SAMSARA_TENANT_UTC_OFFSET, a fixed offset
        3. an offset on pickup_time itself, if this tenant ever sends one
        4. the leg's own `timezone` field -- empty on this tenant, but it
           costs nothing to prefer real data where it exists
        5. the status timestamps, which only works for days already worked

    Returns (offset, source). A None offset means refuse to publish rather
    than guess -- see push_samsara_routes.py.
    """
    sample_date = _first_pickup(legs, parse_ts_aware)

    if TENANT_TIMEZONE and ZoneInfo is not None:
        try:
            zone = ZoneInfo(TENANT_TIMEZONE)
        except Exception:  # noqa: BLE001
            log.warning("SAMSARA_TENANT_TIMEZONE=%r is not a known zone", TENANT_TIMEZONE)
        else:
            return (
                sample_date.replace(tzinfo=zone).utcoffset(),
                f"SAMSARA_TENANT_TIMEZONE ({TENANT_TIMEZONE})",
            )

    explicit = parse_offset_string(TENANT_UTC_OFFSET)
    if explicit is not None:
        return explicit, "SAMSARA_TENANT_UTC_OFFSET"

    offsets = Counter()
    for leg in legs:
        parsed = parse_ts_aware(leg.get("pickup_time"))
        if parsed is not None and parsed.tzinfo is not None:
            offsets[parsed.utcoffset()] += 1
    if offsets:
        return offsets.most_common(1)[0][0], "pickup_time offset"

    if ZoneInfo is not None:
        zones = Counter(
            str(leg.get("timezone") or "").strip()
            for leg in legs
            if str(leg.get("timezone") or "").strip()
        )
        for zone_name, _ in zones.most_common():
            try:
                zone = ZoneInfo(zone_name)
            except Exception:  # noqa: BLE001 -- unknown zone, try the next
                continue
            return (
                sample_date.replace(tzinfo=zone).utcoffset(),
                f"timezone field ({zone_name})",
            )

    from_stamps = _offset_from_status_stamps(legs, parse_ts_aware)
    if from_stamps is not None:
        return from_stamps, "status timestamps"

    return None, "unresolved"


def _first_pickup(legs, parse_ts_aware):
    """
    A naive date to resolve a zone against.

    It has to be a trip's own date, not today: a run the far side of a
    daylight-saving boundary sits at a different offset, and resolving
    November's trips against an August clock is an hour wrong.
    """
    for leg in legs:
        parsed = parse_ts_aware(leg.get("pickup_time"))
        if parsed is not None:
            return parsed.replace(tzinfo=None)
    return datetime.now()


def _offset_from_status_stamps(legs, parse_ts_aware):
    """The most common offset across every status stamp on these legs."""
    offsets = Counter()
    for leg in legs:
        for entry in leg.get("timestamps") or []:
            if not isinstance(entry, dict):
                continue
            for value in entry.values():
                parsed = parse_ts_aware(value)
                if parsed is not None and parsed.tzinfo is not None:
                    offsets[parsed.utcoffset()] += 1
    return offsets.most_common(1)[0][0] if offsets else None


# =============================
# LEG SELECTION
# =============================
def _is_truthy_deleted(value):
    return str(value).strip().lower() in ("1", "true", "yes", "y", "t")


def has_location(leg, side):
    """Whether a leg carries something Samsara can route to on this side."""
    lat, lon = leg.get(f"{side}_lat"), leg.get(f"{side}_lon")
    if lat not in (None, "") and lon not in (None, ""):
        try:
            if float(lat) or float(lon):
                return True
        except (TypeError, ValueError):
            pass
    return bool(str(leg.get(f"{side}_address1") or "").strip())


def transport_minutes(los):
    """How long to allow for a transport at this level of service."""
    return TRANSPORT_MINUTES_BY_LOS.get(normalize_los(los), DEFAULT_TRANSPORT_MINUTES)


def is_excluded_los(los):
    """Whether this level of service is one the operation never pre-routes."""
    if not EXCLUDED_LOS:
        return False
    return normalize_los(los) in EXCLUDED_LOS


def is_excluded_call_type(call_type):
    text = str(call_type or "").strip().lower()
    if not text:
        return False
    return any(pattern in text for pattern in EXCLUDED_CALL_TYPE_PATTERNS)


def eligible_legs(legs, parse_ts_aware):
    """
    Legs a driver can actually be routed to, and why the rest were dropped.

    Requires a scheduled pickup time and a location on both ends. 911 and
    on-demand work has no schedule to route against, so it falls out here
    rather than landing in Samsara with a guessed arrival time.

    Returns (kept, skipped) where skipped is [(leg, reason), ...].
    """
    kept, skipped = [], []
    for leg in legs:
        if _is_truthy_deleted(leg.get("deleted")):
            skipped.append((leg, "deleted"))
            continue
        if is_excluded_call_type(leg.get("call_type")):
            skipped.append((leg, f"call type {leg.get('call_type')!r} excluded"))
            continue
        if is_excluded_los(leg.get("los")):
            skipped.append((leg, f"level of service {leg.get('los')!r} excluded"))
            continue
        if not str(leg.get("vehicle_name") or "").strip():
            skipped.append((leg, "no vehicle assigned"))
            continue
        if not parse_ts_aware(leg.get("pickup_time")):
            skipped.append((leg, "no scheduled pickup time"))
            continue
        if not has_location(leg, "pu"):
            skipped.append((leg, "no pickup location"))
            continue
        if not has_location(leg, "do"):
            skipped.append((leg, "no drop-off location"))
            continue
        kept.append(leg)
    return kept, skipped


# =============================
# STOP CONSTRUCTION
# =============================
def rfc3339(moment, default_offset=None):
    """
    Samsara wants an RFC3339 instant.

    A naive moment with no offset to apply raises rather than defaulting to
    UTC. Silently reading tenant-local time as UTC is the failure that
    dispatches an ambulance four hours early, and it looks like valid output
    all the way to the driver's phone -- so it is refused at the point where
    the information is missing.
    """
    if moment is None:
        return None
    if moment.tzinfo is None:
        if default_offset is None:
            raise ValueError(
                "Refusing to convert a naive timestamp without a tenant UTC "
                "offset: it would be read as UTC and land hours out. Set "
                "SAMSARA_TENANT_UTC_OFFSET."
            )
        moment = moment.replace(tzinfo=timezone(default_offset))
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def format_address(leg, side):
    """A single-line postal address for one end of a leg."""
    parts = [
        str(leg.get(f"{side}_address1") or "").strip(),
        str(leg.get(f"{side}_address2") or "").strip(),
    ]
    street = ", ".join(p for p in parts if p)
    city = str(leg.get(f"{side}_city") or "").strip()
    state = str(leg.get(f"{side}_state") or "").strip()
    zipcode = str(leg.get(f"{side}_zipcode") or "").strip()
    tail = " ".join(p for p in (city + "," if city and state else city, state, zipcode) if p)
    return ", ".join(p for p in (street, tail) if p).strip(", ").strip()


def stop_location(leg, side, address_index=None):
    """
    The location half of a stop payload.

    Prefers a registered Samsara address: those carry the geofence somebody
    already drew around the facility, where a single-use location is always a
    300m circle and will fire arrival on the wrong side of a hospital campus.
    Falls back to coordinates, then to the postal address alone.
    """
    facility = str(leg.get(f"{side}_facility_name") or "").strip()
    if address_index and facility:
        registered = address_index.get(facility.upper())
        if registered and registered.get("id"):
            return {"addressId": str(registered["id"])}

    single = {"address": format_address(leg, side)}
    lat, lon = leg.get(f"{side}_lat"), leg.get(f"{side}_lon")
    try:
        if lat not in (None, "") and lon not in (None, ""):
            single["latitude"] = float(lat)
            single["longitude"] = float(lon)
    except (TypeError, ValueError):
        pass
    if not single.get("address"):
        single["address"] = facility or "Unknown location"
    return {"singleUseLocation": single}


def stop_name(leg, side):
    """What the driver sees in the stop list."""
    label = "PU" if side == "pu" else "DO"
    facility = str(leg.get(f"{side}_facility_name") or "").strip()
    if not facility:
        facility = str(leg.get(f"{side}_address1") or "").strip() or "Unknown"
    return f"{label} · {facility}"[:255]


def stop_notes(leg, side):
    """
    Call type and level of service on the stop itself, so a driver opening a
    stop sees what the call is and what the unit is expected to be, not just
    an address.

    The patient name is omitted unless SAMSARA_INCLUDE_PATIENT_NAME is set --
    see INCLUDE_PATIENT_NAME.
    """
    bits = []
    run = str(leg.get("run_number") or leg.get("trip_number") or "").strip()
    if run:
        bits.append(f"Run {run}")
    call_type = str(leg.get("call_type") or "").strip()
    if call_type:
        bits.append(call_type)
    los = str(leg.get("los") or "").strip()
    if los:
        bits.append(f"LOS {los}")
    priority_key = "response_priority" if side == "pu" else "transport_priority"
    priority = str(leg.get(priority_key) or "").strip()
    if priority:
        bits.append(f"Priority {priority}")
    if side == "pu" and INCLUDE_PATIENT_NAME:
        patient = " ".join(
            str(leg.get(k) or "").strip()
            for k in ("patient_first_name", "patient_last_name")
        ).strip()
        if patient:
            bits.append(patient)
    return " · ".join(bits)[:1000]


def leg_stops(leg, parse_ts_aware, address_index=None, default_offset=None,
              next_pickup=None):
    """
    The two stops for one leg, with a drop-off time that always follows the
    pickup.

    Drop-off time is the appointment time where the trip has one -- that is
    the hour the patient is actually due -- then the drop-off ETA, and only
    then an estimate off the pickup.

    `next_pickup` is when this vehicle is next due somewhere, and it exists
    because Samsara sorts every stop on a route by arrival time. A drop-off
    invented at pickup + 45 minutes will sort after the next leg's pickup
    whenever the two are closer together than that, and the driver is then
    told to collect the next patient before delivering the one on board. A
    guessed time must never reorder real ones, so the estimate is capped just
    short of the next pickup. A genuine appointment time is left alone: that
    is data, and if it overlaps, the overlap is real.
    """
    pickup = parse_ts_aware(leg.get("pickup_time"))
    if pickup is None:
        return [], None

    dropoff = None
    dropoff_source = None
    for key in ("appt_time", "dropoff_eta"):
        candidate = parse_ts_aware(leg.get(key))
        if candidate is not None:
            dropoff, dropoff_source = candidate, key
            break

    floor = pickup + timedelta(minutes=MIN_STOP_GAP_MINUTES)
    if dropoff is None or dropoff <= floor:
        if dropoff is None:
            dropoff_source = "estimated"
            dropoff = pickup + timedelta(minutes=transport_minutes(leg.get("los")))
            if next_pickup is not None and dropoff >= next_pickup:
                capped = next_pickup - timedelta(minutes=MIN_STOP_GAP_MINUTES)
                dropoff = max(floor, capped)
                dropoff_source = "estimated (capped at next pickup)"
        else:
            # A real time that lands before its own pickup would make Samsara
            # sort the drop-off ahead of it. Keep the ordering, flag the data.
            dropoff_source = f"{dropoff_source} (adjusted, was before pickup)"
            dropoff = floor

    stops = [
        {
            "name": stop_name(leg, "pu"),
            "scheduledArrivalTime": rfc3339(pickup, default_offset),
            "notes": stop_notes(leg, "pu"),
            **stop_location(leg, "pu", address_index),
        },
        {
            "name": stop_name(leg, "do"),
            "scheduledArrivalTime": rfc3339(dropoff, default_offset),
            "notes": stop_notes(leg, "do"),
            **stop_location(leg, "do", address_index),
        },
    ]
    return stops, dropoff_source


# =============================
# ROUTE ASSEMBLY
# =============================
def route_name(vehicle_name, day, prefix=None):
    """
    Stable, human-readable, and identifiable as ours.

    The unit and the date are enough for a dispatcher to find it, and the
    prefix lets push_samsara_routes.py recognise a route it created earlier
    when replacing the day.
    """
    label = f"{vehicle_name} — {day.strftime('%a %b %-d')}" if hasattr(day, "strftime") else str(day)
    return f"{prefix} {label}" if prefix else label


def build_routes(
    legs,
    matched_vehicles,
    day,
    parse_ts_aware,
    address_index=None,
    default_offset=None,
    name_prefix=None,
):
    """
    Group eligible legs by vehicle and emit one Samsara route payload each.

    Returns (routes, notes) where each route is
        {"name", "vehicleId", "stops": [...]}
    and notes carries per-route provenance for the dry-run report.
    """
    by_vehicle = defaultdict(list)
    for leg in legs:
        by_vehicle[str(leg.get("vehicle_name") or "").strip()].append(leg)

    routes, notes = [], []
    for vehicle_name in sorted(by_vehicle):
        vehicle = matched_vehicles.get(vehicle_name)
        if not vehicle:
            continue
        vehicle_legs = sorted(
            by_vehicle[vehicle_name],
            key=lambda l: parse_ts_aware(l.get("pickup_time")) or datetime.max.replace(
                tzinfo=timezone.utc
            ),
        )

        # Each leg needs to know when this vehicle is next due somewhere, so a
        # guessed drop-off cannot sort past a real pickup. See leg_stops.
        pickups = [parse_ts_aware(l.get("pickup_time")) for l in vehicle_legs]

        stops, estimated, capped, run_numbers = [], 0, 0, []
        for index, leg in enumerate(vehicle_legs):
            next_pickup = pickups[index + 1] if index + 1 < len(pickups) else None
            leg_pair, source = leg_stops(
                leg, parse_ts_aware, address_index, default_offset,
                next_pickup=next_pickup,
            )
            if not leg_pair:
                continue
            if source and source.startswith("estimated (capped"):
                capped += 1
            stops.extend(leg_pair)
            if source and source != "appt_time":
                estimated += 1
            run = str(leg.get("run_number") or leg.get("trip_number") or "").strip()
            if run:
                run_numbers.append(run)

        if not stops:
            continue

        routes.append(
            {
                "name": route_name(vehicle_name, day, name_prefix),
                "vehicleId": str(vehicle.get("id")),
                "stops": stops,
            }
        )
        notes.append(
            {
                "vehicle_name": vehicle_name,
                "samsara_vehicle": vehicle.get("name"),
                "samsara_vehicle_id": str(vehicle.get("id")),
                "legs": len(vehicle_legs),
                "stops": len(stops),
                "estimated_dropoff_times": estimated,
                "capped_dropoff_times": capped,
                "run_numbers": run_numbers,
            }
        )
    return routes, notes


def index_addresses(addresses):
    """{upper-cased address name: address} for facility-name lookup."""
    index = {}
    for address in addresses or []:
        name = (address.get("name") or "").strip()
        if name:
            index[name.upper()] = address
    return index
