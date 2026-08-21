"""
Report data layer backed by the Traumasoft ThirdParty API.

Produces the same DataFrames the SQL queries in `Script` produce, so the Excel,
append/retention, email, and status_logger layers stay untouched. Each builder
here replaces one `pd.read_sql_query(...)` call.

What the live probes established, and what it forces:

  * Trips backfill. GetTrips returned data 90 days back, so OTP and the UHU
    numerator keep full history and `--date` backfill still works for them.

  * Shifts do not. /Schedule/Shifts ignores every filter and always returns the
    same rolling window anchored to now (751 rows, today-1..today+2, identical
    for requests at -30, +0 and +30 days). Anything derived from scheduled
    hours or crew therefore describes only the present. Historical staffing and
    the UHU denominator cannot be reconstructed after the fact -- they have to
    be captured daily as they pass.

  * Cost center is absent from trips and vehicles. It reaches a trip only via
    shift_name -> the crew on that shift -> employee.cost_center_name. Since
    shifts are present-only, that mapping is accumulated to disk daily and
    reused for historical legs (see CostCenterMap).

  * Vehicle status history does not exist. `oos_since` is likewise accumulated
    locally (see OutOfServiceHistory).
"""

import os
import json
import logging
from pathlib import Path
from collections import Counter, defaultdict
from datetime import datetime, timedelta, date, timezone

import pandas as pd

log = logging.getLogger(__name__)

# =============================
# CONFIG
# =============================
# State that has to survive between runs because the API cannot reproduce it.
STATE_DIR = os.getenv("TS_STATE_DIR", "state")
COST_CENTER_MAP_FILE = os.path.join(STATE_DIR, "shift_cost_center_map.json")
COST_CENTER_OVERRIDES_FILE = os.path.join(STATE_DIR, "shift_cost_center_overrides.json")
OOS_HISTORY_FILE = os.path.join(STATE_DIR, "vehicle_oos_history.json")

# Trip timestamp names this tenant emits, in preference order, for the moment a
# unit reached the patient. The current SQL scores against ePCR field 549,
# which this API does not expose; `at_scene` is the closest CAD equivalent and
# the bedside variant is closer still where it is recorded.
ARRIVAL_TIMESTAMP_KEYS = [
    k.strip() for k in os.getenv(
        "TS_ARRIVAL_TIMESTAMP_KEYS", "at_scene: At Patient Bedside,at_scene"
    ).split(",") if k.strip()
]

# Minutes either side of scheduled pickup that still count as on time.
OTP_ON_TIME_WINDOW_MINUTES = int(os.getenv("OTP_ON_TIME_WINDOW_MINUTES", "10"))

OTP_EXCLUDED_COST_CENTERS = ("CPR",)

# Fleet status strings, which replace the status id + status_reason LIKE
# filtering the SQL used. Observed on this tenant: In Service,
# New - Waiting for Delivery, Out of Service, Out of Service - Collision,
# Retired, Waiting for Inspection.
IN_SERVICE_STATUSES = {"In Service"}
OUT_OF_SERVICE_STATUSES = {"Out of Service", "Out of Service - Collision"}
NON_FLEET_STATUSES = {"Retired", "New - Waiting for Delivery", "Waiting for Inspection"}

# The SQL pruned dead records with `status_reason NOT LIKE '%duplicate%' /
# '%retired%' / '%disposed%' / '%test%'` and a hard-coded id list. The API's
# vehicle allowlist has no status_reason, so the equivalent is name patterns
# plus an explicit list -- see VehicleExclusions. Vehicles left sitting at
# "Out of Service" for years because nobody set them to Retired are
# indistinguishable from genuinely broken ones through this API.
FLEET_EXCLUDED_NAME_PATTERNS = tuple(
    p.strip().lower()
    for p in os.getenv(
        "FLEET_EXCLUDED_NAME_PATTERNS",
        "test,check test,do not use,dnu,duplicate,disposed,retired,spare - out",
    ).split(",")
    if p.strip()
)

VEHICLE_EXCLUSIONS_FILE = os.path.join(STATE_DIR, "vehicle_exclusions.json")

# How far back to look for trip activity when flagging dormant vehicles.
# Capped at the API's 31-day range limit so it costs a single call.
FLEET_ACTIVITY_LOOKBACK_DAYS = int(os.getenv("FLEET_ACTIVITY_LOOKBACK_DAYS", "30"))

# Staffing: cost centers excluded by name. What a unit needs on the clock is
# no longer a flat number here -- it comes from state/unit_staffing_rules.json,
# the same rules the UHU denominator uses, so a wheelchair or secure car needs
# one where an ambulance needs two.
STAFFING_EXCLUDED_COST_CENTER_PATTERNS = ("dispatch", "cpr", "training", "admin")


# The SQL excluded certification templates 101-104. The API exposes a
# license_level string instead of those ids, so the exclusion is by name and
# must be confirmed against Lists/User/LicenseLevels before trusting counts.
STAFFING_EXCLUDED_LICENSE_LEVELS = {
    lv.strip().lower()
    for lv in os.getenv("STAFFING_EXCLUDED_LICENSE_LEVELS", "").split(",")
    if lv.strip()
}

# Which span of a trip counts as utilized time for UHU. "task" is enroute ->
# clear (the whole committed period); "loaded" is transporting -> at_destination
# (patient on board only). The SQL used an estimated duration, so neither is a
# like-for-like match -- see docs/API_MIGRATION.md.
UHU_SPAN = os.getenv("UHU_SPAN", "task")
UHU_SPANS = {
    # enroute -> clear. The whole committed period in principle, but only as
    # good as the crew's discipline about pressing Clear. On this tenant the
    # next leg's enroute lands seconds after the previous leg's clear, which
    # means clear is being set when the next call is assigned rather than when
    # the last one ended -- so this tiles the shift and pushes UHU toward 100%.
    "task": ("enroute", "clear"),
    # enroute -> at_destination. Dispatch to drop-off, stopping at a stamp the
    # crew hits on arrival rather than one they get around to later. Drops the
    # post-drop-off tail, which is where the sloppiness lives.
    "transport": ("enroute", "at_destination"),
    # transporting -> at_destination. Patient on board only; excludes response
    # and on-scene time, so it understates a unit's real commitment.
    "loaded": ("transporting", "at_destination"),
    # enroute -> at_scene. Response only. Not a utilization measure; useful as
    # a floor when comparing candidates.
    "response": ("enroute", "at_scene"),
}

# The UHU SQL excluded schedules named like Dispatch / Comm / Call, which are
# communications rosters rather than transport units and would otherwise sink
# every ratio with scheduled hours that can never be utilized.
UHU_EXCLUDED_PROFILE_PATTERNS = ("dispatch", "comm", "call")

# Trip statuses that are not runs. The SQL restricted loaded hours to
# last_status_id IN (1..7); these are the equivalent status strings observed on
# this tenant for legs that never ran.
UHU_EXCLUDED_TRIP_STATUSES = {"canceled", "cancelled", "disregard", "no transport"}

# /Schedule/Shifts returns start_time and end_time as bare UTC wall time, while
# trips return local time with an explicit offset. Left unreconciled, every
# comparison between a shift and a trip -- or between a shift and "now" -- is
# wrong by the tenant's offset. Set TS_SHIFT_TIMES_ARE_UTC=0 if Traumasoft ever
# starts returning these as local, and TS_SHIFT_UTC_OFFSET_HOURS to pin the
# offset when a day's trips carry none to read it from.
SHIFT_TIMES_ARE_UTC = os.getenv("TS_SHIFT_TIMES_ARE_UTC", "1").strip().lower() not in (
    "0", "false", "no",
)
_offset_override = os.getenv("TS_SHIFT_UTC_OFFSET_HOURS", "").strip()
SHIFT_UTC_OFFSET_HOURS = float(_offset_override) if _offset_override else None

# What the UHU denominator measures. "worked" is unit hours the truck was
# actually crewed to minimum staffing, read from shift punches; "scheduled" is
# hours rostered. A profile rostered with one person is not a unit that ran --
# anything but a wheelchair or secure car needs two -- so rostered hours count
# trucks that never turned a wheel and understate every ratio they appear in.
UHU_DENOMINATOR = os.getenv("UHU_DENOMINATOR", "worked").strip().lower()

# Crew needed simultaneously on the clock for a unit to be in service. Two for
# BLS and ALS alike (an EMT or medic plus a driver); one for wheelchair and
# secure car. Profiles that need one are named in the rules file, since only
# the operation knows which its names refer to.
UHU_DEFAULT_MIN_CREW = int(os.getenv("UHU_DEFAULT_MIN_CREW", "2"))
UNIT_STAFFING_RULES_FILE = os.path.join(STATE_DIR, "unit_staffing_rules.json")

# The SQL used HAVING COUNT(DISTINCT user_id) > 2, which this inherited as
# STAFFING_MIN_CREW=3. A flat floor cannot express "one for a wheelchair, two
# for an ambulance", and at 3 it dropped nearly every properly crewed truck.
# Anyone who set it deserves to be told it stopped applying rather than to
# find the number quietly doing nothing.
if os.getenv("STAFFING_MIN_CREW"):
    log.warning(
        "STAFFING_MIN_CREW=%s is set but no longer used. Crew minimums now come "
        "from %s, per shift profile. Units below their own minimum are listed "
        "and marked SHORT rather than dropped.",
        os.getenv("STAFFING_MIN_CREW"), UNIT_STAFFING_RULES_FILE,
    )


UNASSIGNED_COST_CENTER = "No Cost Center Assigned"
UNASSIGNED_VEHICLE = "No Vehicle Assigned"

# A trip row with a blank trip_status. Excluded from run counts by default --
# see is_run.
COUNT_STATUSLESS_LEGS = os.getenv("TS_COUNT_STATUSLESS_LEGS", "0").strip().lower() in (
    "1", "true", "yes",
)


# =============================
# PARSING HELPERS
# =============================
def parse_ts(value):
    """
    Parse the timestamp shapes this API returns.

    Shifts come back as '2026-08-17T09:00' (no zone); trip timestamps are
    ISO-8601 with an offset. Offsets are dropped to naive local time so the
    two can be compared, which matches how the SQL treated these columns.
    """
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
        else:
            return None
    if parsed.tzinfo is not None:
        parsed = parsed.replace(tzinfo=None)
    return parsed


def timestamp_map(leg):
    """Flatten a leg's `timestamps` array of {name: iso} maps into one dict."""
    flat = {}
    for entry in leg.get("timestamps") or []:
        if isinstance(entry, dict):
            for name, value in entry.items():
                if value:
                    flat[name] = value
    return flat


def arrival_time(leg, keys=None):
    """The first configured arrival stamp present on this leg."""
    stamps = timestamp_map(leg)
    for key in (keys or ARRIVAL_TIMESTAMP_KEYS):
        if stamps.get(key):
            parsed = parse_ts(stamps[key])
            if parsed:
                return parsed
    return None


def span_minutes(leg, start_key, end_key):
    """Minutes between two named stamps on a leg, or None if either is absent."""
    stamps = timestamp_map(leg)
    start = parse_ts(stamps.get(start_key))
    end = parse_ts(stamps.get(end_key))
    if not start or not end or end <= start:
        return None
    return (end - start).total_seconds() / 60.0


def profile_name(record):
    """
    A shift profile name usable as a join key.

    Some rows carry leading or trailing spaces -- '  WV-A-MCD-08-20' and
    'WV-A-MCD-08-20' are one unit typed twice -- and left alone they split into
    two profiles that each look half-staffed. Stripped on the way in, on both
    the shift and the trip side, so the two always agree.
    """
    name = record.get("shift_name")
    if name is None:
        return None
    name = str(name).strip()
    return name or None


def parse_ts_aware(value):
    """Like parse_ts, but keeps the offset when the value carries one."""
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def tenant_utc_offset(legs):
    """
    The tenant's current UTC offset, read from the trip timestamps themselves.

    Trip stamps arrive as local time with an explicit offset ('-04:00'); shift
    start/end arrive as bare UTC wall time with none. Comparing the two without
    reconciling them is wrong by the offset -- four hours on Eastern daylight
    time -- which silently misplaces every shift.

    Taking the offset from the response rather than a config value means it
    follows the tenant across daylight saving without anyone remembering to
    change it, and needs no timezone database on the machine.
    """
    counts = Counter()
    for leg in legs:
        for value in timestamp_map(leg).values():
            parsed = parse_ts_aware(value)
            if parsed is not None and parsed.tzinfo is not None:
                counts[parsed.utcoffset()] += 1
    if not counts:
        return None
    return counts.most_common(1)[0][0]


def resolve_shift_offset(legs):
    """The offset to add to a shift's UTC time to reach tenant-local time."""
    if not SHIFT_TIMES_ARE_UTC:
        return timedelta(0)
    if SHIFT_UTC_OFFSET_HOURS is not None:
        return timedelta(hours=SHIFT_UTC_OFFSET_HOURS)
    offset = tenant_utc_offset(legs)
    if offset is None:
        log.warning(
            "No trip timestamp carried a UTC offset, so shift times cannot be "
            "converted to local. Scheduled hours and staffing will be wrong by "
            "the tenant's offset. Set TS_SHIFT_UTC_OFFSET_HOURS to fix it."
        )
        return timedelta(0)
    return offset


def parse_shift_ts(value, offset=None):
    """
    Parse a shift start/end into tenant-local naive time.

    The feed returns these as bare UTC; every other timestamp in the reports is
    local, so they are shifted into line here rather than at each comparison.
    """
    parsed = parse_ts(value)
    if parsed is None or not offset:
        return parsed
    return parsed + offset


def tenant_now(shift_offset=None):
    """
    Now, on the same clock the shifts and punches are on.

    Derived from UTC plus the offset the trips reported rather than from the
    machine's own clock, so a reporting box in the wrong timezone -- or one
    that follows daylight saving differently from the tenant -- does not shift
    the answer. Falls back to local time when no offset could be resolved.
    """
    if shift_offset is None:
        return datetime.now()
    return datetime.now(timezone.utc).replace(tzinfo=None) + shift_offset


def overlap_minutes(start, end, window_start, window_end):
    """Minutes of [start, end) that fall inside [window_start, window_end)."""
    if not start or not end:
        return 0.0
    latest_start = max(start, window_start)
    earliest_end = min(end, window_end)
    if earliest_end <= latest_start:
        return 0.0
    return (earliest_end - latest_start).total_seconds() / 60.0


# =============================
# ACCUMULATED STATE
# =============================
def _describe_unparsable(path):
    """Say what is actually in a state file that would not parse."""
    try:
        with open(path, "rb") as handle:
            raw = handle.read()
    except OSError:
        return

    if not raw.strip():
        log.warning("  It is %s. Nothing was ever written to it -- check "
                    "whether an editor saved alongside it as '%s.txt' instead.",
                    "empty" if not raw else "blank (%s bytes of whitespace)" % len(raw),
                    os.path.basename(path))
        return

    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        log.warning("  It is UTF-16. A PowerShell '>' redirect writes UTF-16; "
                    "use Set-Content -Encoding UTF8, or copy the file in "
                    "rather than retyping it.")
        return

    first = ""
    for line in raw.decode("utf-8", "replace").splitlines():
        if line.strip():
            first = line.strip()
            break
    log.warning("  It is %s bytes and starts with: %s",
                len(raw), first[:70] + ("..." if len(first) > 70 else ""))
    if first.startswith(("@'", '@"', "'@", '"@')) or "Set-Content" in first:
        log.warning("  That is the PowerShell command that was meant to write "
                    "the file, not the file's contents. Only the JSON between "
                    "the quotes belongs here -- it must start with '{'.")
    elif not first.startswith(("{", "[")):
        log.warning("  JSON has to start with '{'. Anything before it -- a "
                    "shell prompt, a heading, a stray character -- has to go.")


def load_state_file(path, what):
    """
    Read one of the state/ JSON files, or None if it is not usable.

    A file that exists but does not parse is a typo someone just made, not an
    absence. Swallowing that silently means an edit appears to have no effect,
    which is the hardest kind of change to debug -- so say so.
    """
    try:
        with open(path, "r", encoding="utf-8-sig") as handle:
            return json.load(handle)
    except FileNotFoundError:
        return None
    except ValueError as exc:
        # "Expecting value: line 1 column 1" is the same message whether the
        # file is empty or starts with something that is not JSON at all, and
        # the two are fixed differently. Say which, and show it, rather than
        # leaving the reader to guess from a parser's column number.
        log.warning(
            "%s is not valid JSON and is being ignored: %s. "
            "The %s it defines will not take effect until it parses.",
            os.path.abspath(path), exc, what,
        )
        _describe_unparsable(path)
        return None
    except OSError as exc:
        log.warning("Could not read %s: %s", path, exc)
        return None


class CostCenterMap:
    """
    shift_name -> cost center, accumulated across runs.

    Cost center is not on trips or vehicles, and the only route to it --
    shift_name -> crew -> employee.cost_center_name -- is available solely for
    the rolling shift window. Historical legs would otherwise be
    unattributable, so every run folds what it can see into a file on disk.

    Crew counts per (shift_name, cost_center) are kept rather than a single
    winner, because a profile picks up a second cost center whenever someone
    from a neighbouring station fills in. The dominant cost center wins and the
    contested ones stay visible via `ambiguous()`.
    """

    def __init__(self, path=COST_CENTER_MAP_FILE, overrides_path=COST_CENTER_OVERRIDES_FILE):
        self.path = path
        self.overrides_path = overrides_path
        self.counts = defaultdict(Counter)
        self.override_names = {}
        self.override_prefixes = []
        self._load()
        self._load_overrides()

    def _load(self):
        stored = load_state_file(self.path, "learned cost centers")
        if not stored:
            return
        for shift_name, centers in (stored.get("counts") or {}).items():
            self.counts[shift_name] = Counter(centers)

    def save(self):
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "updated": datetime.now().isoformat(timespec="seconds"),
            "counts": {name: dict(counter) for name, counter in self.counts.items()},
        }
        with open(self.path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, sort_keys=True)

    def update(self, shifts, employees):
        """Fold today's visible shift window into the accumulated mapping."""
        employee_cc = {
            str(emp.get("user_id")): emp.get("cost_center_name")
            for emp in employees
            if emp.get("user_id") and emp.get("cost_center_name")
        }
        learned = 0
        for shift in shifts:
            name = profile_name(shift)
            cost_center = employee_cc.get(str(shift.get("user_id")))
            if name and cost_center:
                self.counts[name][cost_center] += 1
                learned += 1
        log.info(
            "Cost-center map: folded in %s crew rows, %s profiles known",
            learned,
            len(self.counts),
        )
        return learned

    def _load_overrides(self):
        """
        Hand-written mappings, for profiles the crew route can never reach.

        Some profiles carry real runs but no scheduled crew -- outsourced work
        is the obvious case -- so no amount of accumulating will attribute them.
        A deliberate mapping is the only answer, and it wins over what was
        learned, being a decision rather than an observation.
        """
        stored = load_state_file(self.overrides_path, "cost center overrides")
        if not stored:
            return
        self.override_names = {
            str(k).strip().lower(): v
            for k, v in (stored.get("by_name") or {}).items() if v
        }
        # Longest prefix first, so a specific rule beats a general one.
        self.override_prefixes = sorted(
            ((str(k).strip().lower(), v) for k, v in (stored.get("by_prefix") or {}).items() if v),
            key=lambda kv: -len(kv[0]),
        )
        if self.override_names or self.override_prefixes:
            log.info(
                "Cost-center overrides: %s exact, %s prefix rule(s) from %s",
                len(self.override_names), len(self.override_prefixes), self.overrides_path,
            )

    def resolve(self, shift_name):
        """Dominant cost center for a shift profile, or None if unknown."""
        if not shift_name:
            return None
        lowered = str(shift_name).strip().lower()
        if lowered in self.override_names:
            return self.override_names[lowered]
        for prefix, centre in self.override_prefixes:
            if lowered.startswith(prefix):
                return centre
        counter = self.counts.get(shift_name)
        if not counter:
            # Trailing spaces and casing drift between the trip and shift feeds;
            # fall back to a normalized match before giving up.
            for name, stored in self.counts.items():
                if str(name).strip().lower() == lowered:
                    counter = stored
                    break
        if not counter:
            return None
        # most_common breaks ties by insertion order; sort for determinism.
        return sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]

    def ambiguous(self):
        """Profiles seen against more than one cost center."""
        return {
            name: dict(counter)
            for name, counter in self.counts.items()
            if len(counter) > 1
        }


class UnitStaffingRules:
    """
    Shift profile -> the crew it needs on the clock at once to be a unit.

    Two people for BLS and ALS -- a licensed clinician plus someone to drive --
    and one for wheelchair and secure car work. Which profile names mean which
    is knowledge the API does not carry and only the operation has, so the
    exceptions are listed rather than inferred from a naming convention that
    could change.
    """

    def __init__(self, path=UNIT_STAFFING_RULES_FILE, default=None):
        self.path = path
        self.default = default if default is not None else UHU_DEFAULT_MIN_CREW
        self.patterns = []
        self._load()

    def _load(self):
        stored = load_state_file(self.path, "unit staffing rules")
        if not stored:
            return
        if stored.get("default_min_crew") is not None:
            self.default = int(stored["default_min_crew"])
        # Longest pattern first, so a specific name beats a general one.
        self.patterns = sorted(
            ((str(k).strip().lower(), int(v))
             for k, v in (stored.get("by_pattern") or {}).items()),
            key=lambda kv: -len(kv[0]),
        )
        if self.patterns:
            log.info("Unit staffing rules: %s pattern(s), default %s crew, from %s",
                     len(self.patterns), self.default, self.path)

    def min_crew(self, shift_name):
        lowered = str(shift_name or "").strip().lower()
        for pattern, needed in self.patterns:
            if pattern and pattern in lowered:
                return needed
        return self.default


class VehicleExclusions:
    """
    Vehicles to drop from the fleet reports entirely.

    Replaces the SQL's `v.id NOT IN (...)` list and its status_reason LIKE
    filters. Two mechanisms, because they catch different things:

      * name patterns catch obvious rigs -- test vehicles, check trucks;
      * an explicit list of ids or names catches units that were simply
        abandoned at "Out of Service" instead of being set to Retired, which
        this API gives no way to detect.

    Edit state/vehicle_exclusions.json by hand, or seed it with
    suggest_vehicle_exclusions.py.
    """

    def __init__(self, path=VEHICLE_EXCLUSIONS_FILE, patterns=None):
        self.path = path
        self.patterns = tuple(patterns if patterns is not None else FLEET_EXCLUDED_NAME_PATTERNS)
        self.ids = set()
        self.names = set()
        self._load()

    def _load(self):
        stored = load_state_file(self.path, "vehicle exclusions")
        if not stored:
            return
        self.ids = {str(i) for i in (stored.get("exclude_ids") or [])}
        self.names = {str(n).strip().lower() for n in (stored.get("exclude_names") or [])}
        extra = stored.get("exclude_name_patterns")
        if extra:
            self.patterns = tuple(list(self.patterns) + [str(p).strip().lower() for p in extra])

    def save(self, note=None):
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "note": note or (
                "Vehicles excluded from the fleet reports. exclude_ids and "
                "exclude_names are exact matches; exclude_name_patterns are "
                "substrings matched case-insensitively."
            ),
            "updated": datetime.now().isoformat(timespec="seconds"),
            "exclude_ids": sorted(self.ids),
            "exclude_names": sorted(self.names),
        }
        with open(self.path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)

    def excludes(self, vehicle):
        if str(vehicle.get("id")) in self.ids:
            return "explicit id"
        name = (vehicle.get("name") or "").strip()
        if name.lower() in self.names:
            return "explicit name"
        lowered = name.lower()
        for pattern in self.patterns:
            if pattern and pattern in lowered:
                return f"name matches '{pattern}'"
        return None


class OutOfServiceHistory:
    """
    vehicle id -> the first date it was seen out of service, accumulated daily.

    The API exposes only a vehicle's current status; there is no status log and
    no work-order endpoint, so `oos_since` and days-out have to be observed
    rather than queried. A vehicle returning to service clears its entry, so
    the next spell starts a fresh clock.
    """

    def __init__(self, path=OOS_HISTORY_FILE):
        self.path = path
        self.since = {}
        self._load()

    def _load(self):
        stored = load_state_file(self.path, "out-of-service history")
        if not stored:
            return
        self.since = dict(stored.get("since") or {})

    def save(self):
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "updated": datetime.now().isoformat(timespec="seconds"),
            "since": self.since,
        }
        with open(self.path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, sort_keys=True)

    def update(self, vehicles, run_date):
        """Record newly out-of-service vehicles; clear ones that came back."""
        run_date_str = run_date.isoformat() if hasattr(run_date, "isoformat") else str(run_date)
        for vehicle in vehicles:
            key = str(vehicle.get("id"))
            status = vehicle.get("vehicle_status")
            if status in OUT_OF_SERVICE_STATUSES:
                self.since.setdefault(key, run_date_str)
            elif key in self.since:
                del self.since[key]

    def days_out(self, vehicle_id, as_of):
        since = self.since.get(str(vehicle_id))
        if not since:
            return None, None
        since_date = parse_ts(since)
        if not since_date:
            return None, None
        as_of_date = as_of if isinstance(as_of, date) else parse_ts(as_of).date()
        return since, (as_of_date - since_date.date()).days


# =============================
# ON-TIME PERFORMANCE
# =============================
def score_leg(leg, arrival_keys=None, window=OTP_ON_TIME_WINDOW_MINUTES):
    """
    Classify one leg as Early / On Time / Late / Missing Data.

    Mirrors the SQL's CASE: the delta is arrival minus scheduled pickup, and
    anything inside +/- `window` minutes is on time.
    """
    pickup = parse_ts(leg.get("pickup_time"))
    arrived = arrival_time(leg, arrival_keys)
    if not pickup or not arrived:
        return "Missing Data", None
    delta = (arrived - pickup).total_seconds() / 60.0
    if -window <= delta <= window:
        return "On Time", delta
    if delta < -window:
        return "Early", delta
    return "Late", delta


def scored_legs(legs, cost_center_map, arrival_keys=None):
    """
    Score every leg that OTP can actually judge.

    Only legs carrying both a scheduled pickup and an arrival stamp are
    scorable. On this tenant that is exactly the population where shift_name is
    always present, so cost-center attribution is complete for the rows that
    survive this filter; cancellations and disregards, which supply the
    unattributable legs, are dropped here as the SQL dropped them via
    last_status_id.
    """
    rows = []
    for leg in legs:
        status, delta = score_leg(leg, arrival_keys)
        if status == "Missing Data":
            continue
        shift_name = profile_name(leg)
        rows.append(
            {
                "leg_id": leg.get("leg_id"),
                "run_number": leg.get("run_number"),
                "shift_name": shift_name,
                "call_type": leg.get("call_type") or "Unknown Call Type",
                "cost_center": cost_center_map.resolve(shift_name) or UNASSIGNED_COST_CENTER,
                "status": status,
                "delta_minutes": round(delta, 2) if delta is not None else None,
            }
        )
    return pd.DataFrame(rows)


def _otp_aggregate(scored, group_cols):
    if scored.empty:
        return pd.DataFrame(
            columns=group_cols
            + ["total_runs", "early_runs", "on_time_runs", "late_runs", "on_time_percentage"]
        )

    keep = ~scored["cost_center"].isin(OTP_EXCLUDED_COST_CENTERS)
    scored = scored[keep]

    grouped = (
        scored.assign(
            early_runs=(scored["status"] == "Early").astype(int),
            on_time_runs=(scored["status"] == "On Time").astype(int),
            late_runs=(scored["status"] == "Late").astype(int),
        )
        .groupby(group_cols, dropna=False)
        .agg(
            total_runs=("status", "size"),
            early_runs=("early_runs", "sum"),
            on_time_runs=("on_time_runs", "sum"),
            late_runs=("late_runs", "sum"),
        )
        .reset_index()
    )
    grouped["on_time_percentage"] = (
        100.0 * (grouped["on_time_runs"] + grouped["early_runs"]) / grouped["total_runs"]
    ).round(2)
    return grouped


def build_otp_by_call_type(scored):
    df = _otp_aggregate(scored, ["cost_center", "call_type"])
    return df.sort_values(["cost_center", "on_time_percentage"], ascending=[True, False])


def build_otp_by_cost_center(scored):
    df = _otp_aggregate(scored, ["cost_center"])
    return df.sort_values("on_time_percentage", ascending=False)


# =============================
# STAFFING
# =============================
def _excluded_cost_center(name):
    if not name:
        return False
    lowered = name.lower()
    return any(pattern in lowered for pattern in STAFFING_EXCLUDED_COST_CENTER_PATTERNS)


def _excluded_uhu_profile(name):
    if not name:
        return False
    lowered = name.lower()
    return any(pattern in lowered for pattern in UHU_EXCLUDED_PROFILE_PATTERNS)


def _staffing_rows(shifts, employees, predicate, shift_offset=None,
                   staffing_rules=None):
    """
    Collapse crew-level shift rows into one row per unit, staffed or not.

    The API returns one row per crew member per shift, which is the shape the
    SQL grouped over. Rows are keyed on the shift profile and its start/end so
    two crews on the same profile at different times stay separate.

    What a unit needs comes from the same rules the UHU denominator uses, so an
    ambulance needs two and a wheelchair or secure car needs one. A unit short
    of its own minimum stays in the sheet and is marked, because that is the
    staffing problem the report exists to surface -- dropping it is what let it
    go unnoticed.
    """
    rules = staffing_rules if staffing_rules is not None else UnitStaffingRules()
    by_user = {str(emp.get("user_id")): emp for emp in employees if emp.get("user_id")}
    units = defaultdict(
        lambda: {"crew": {}, "cost_centers": Counter(), "start": None, "end": None}
    )

    for shift in shifts:
        start = parse_shift_ts(shift.get("start_time"), shift_offset)
        end = parse_shift_ts(shift.get("end_time"), shift_offset)
        if not start or not end or not predicate(start, end):
            continue
        if shift.get("deleted"):
            continue

        level = (shift.get("license_level") or "").strip().lower()
        if level and level in STAFFING_EXCLUDED_LICENSE_LEVELS:
            continue

        user_id = str(shift.get("user_id") or "")
        if not user_id:
            continue

        key = (profile_name(shift), start, end)
        unit = units[key]
        unit["start"] = start
        unit["end"] = end

        employee = by_user.get(user_id, {})
        first = employee.get("first_name") or ""
        last = employee.get("last_name") or ""
        label = f"{first} {last} (ID {user_id})".strip()
        unit["crew"][user_id] = label
        if employee.get("cost_center_name"):
            unit["cost_centers"][employee["cost_center_name"]] += 1

    rows = []
    for (shift_name, start, end), unit in units.items():
        cost_center = None
        if unit["cost_centers"]:
            cost_center = sorted(
                unit["cost_centers"].items(), key=lambda kv: (-kv[1], kv[0])
            )[0][0]
        if _excluded_cost_center(cost_center):
            continue
        crew_count = len(unit["crew"])
        needed = rules.min_crew(shift_name)
        short_by = max(0, needed - crew_count)
        rows.append(
            {
                "shift_profile": shift_name,
                "cost_center": cost_center or UNASSIGNED_COST_CENTER,
                "start_time": start,
                "end_time": end,
                "crew_count": crew_count,
                "crew_needed": needed,
                "staffing_status": "OK" if not short_by else f"SHORT {short_by}",
                "crew_members": "\n".join(sorted(unit["crew"].values())),
            }
        )

    columns = [
        "shift_profile", "cost_center", "start_time", "end_time",
        "crew_count", "crew_needed", "staffing_status", "crew_members",
    ]
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=columns)
    # Short units first inside each cost center, so the problems are at the top
    # of the block rather than scattered through it.
    df = df.reindex(columns=columns)
    df["_short"] = df["staffing_status"] != "OK"
    df = df.sort_values(
        ["cost_center", "_short", "shift_profile", "start_time"],
        ascending=[True, False, True, True],
    )
    return df.drop(columns="_short")


def build_staffing_active_now(shifts, employees, as_of, shift_offset=None,
                              staffing_rules=None):
    """Units on shift at a moment in time, whether or not they are crewed."""
    return _staffing_rows(shifts, employees, lambda s, e: s <= as_of <= e,
                          shift_offset, staffing_rules)


def build_staffing_for_date(shifts, employees, target_date, shift_offset=None,
                            staffing_rules=None):
    """Units whose shift starts on a given calendar day."""
    return _staffing_rows(shifts, employees, lambda s, e: s.date() == target_date,
                          shift_offset, staffing_rules)


# =============================
# VEHICLES
# =============================
def _is_truthy_flag(value):
    """
    Whether an API boolean is set.

    These come back as real booleans on some endpoints and as "0"/"1" or
    "true"/"false" strings on others, so a bare truth test would treat the
    string "0" as set.
    """
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in ("1", "true", "yes", "y")


def _fleet_partition(vehicles, exclusions=None):
    """
    Split the live fleet into in-service and out-of-service.

    Exclusions apply to both halves: a decommissioned truck should not inflate
    the in-service count any more than the out-of-service list.
    """
    in_service, out_of_service = [], []
    for vehicle in vehicles:
        # Belt and braces: the list call already asks the server to omit
        # deleted and disabled records, but the fields are returned so there is
        # no reason to trust that over the data in hand.
        if _is_truthy_flag(vehicle.get("deleted")) or _is_truthy_flag(vehicle.get("disabled")):
            continue
        status = vehicle.get("vehicle_status")
        if status in NON_FLEET_STATUSES:
            continue
        if exclusions and exclusions.excludes(vehicle):
            continue
        if status in IN_SERVICE_STATUSES:
            in_service.append(vehicle)
        elif status in OUT_OF_SERVICE_STATUSES:
            out_of_service.append(vehicle)
    return in_service, out_of_service


def vehicle_last_seen(legs_by_day):
    """
    vehicle id -> the most recent date it ran a leg.

    Built from a trip lookback, which backfills fine. It is the only signal
    this API offers for telling a truck that is genuinely down from one that
    was abandoned in the system years ago.
    """
    last_seen = {}
    for day, legs in legs_by_day.items():
        for leg in legs:
            vehicle_id = leg.get("vehicle_id")
            if not vehicle_id:
                continue
            key = str(vehicle_id)
            if key not in last_seen or day > last_seen[key]:
                last_seen[key] = day
    return last_seen


def used_vehicle_ids(legs):
    """Vehicles that actually ran a leg on the day."""
    return {str(leg.get("vehicle_id")) for leg in legs if leg.get("vehicle_id")}


def build_vehicle_summary(vehicles, legs, run_date, exclusions=None):
    in_service, out_of_service = _fleet_partition(vehicles, exclusions)
    used = used_vehicle_ids(legs)
    in_service_ids = {str(v.get("id")) for v in in_service}
    used_in_service = in_service_ids & used

    rows = [
        {"metric": "Total Fleet (excluding retired/pending)", "value": len(in_service) + len(out_of_service)},
        {"metric": "In Service", "value": len(in_service)},
        {"metric": "Out Of Service", "value": len(out_of_service)},
        {"metric": "Vehicles Used", "value": len(used)},
        {"metric": "In-Service Vehicles Used", "value": len(used_in_service)},
        {"metric": "Unused In-Service Vehicles", "value": len(in_service_ids - used)},
        {"metric": "Report Date", "value": str(run_date)},
    ]
    return pd.DataFrame(rows)


def build_vehicles_in_use(vehicles, legs, exclusions=None):
    used = used_vehicle_ids(legs)
    leg_counts = Counter(str(leg.get("vehicle_id")) for leg in legs if leg.get("vehicle_id"))
    rows = [
        {
            "vehicle_id": v.get("id"),
            "vehicle_name": v.get("name"),
            "vehicle_status": v.get("vehicle_status"),
            "shift_name": v.get("shift_name"),
            "legs_run": leg_counts.get(str(v.get("id")), 0),
        }
        for v in vehicles
        if str(v.get("id")) in used and not (exclusions and exclusions.excludes(v))
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


def build_vehicles_all_in_service(vehicles, exclusions=None):
    in_service, _ = _fleet_partition(vehicles, exclusions)
    rows = [
        {
            "vehicle_id": v.get("id"),
            "vehicle_name": v.get("name"),
            "vehicle_status": v.get("vehicle_status"),
            "odometer": v.get("odometer"),
            "shift_name": v.get("shift_name"),
        }
        for v in in_service
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


def build_vehicles_unused_in_service(vehicles, legs, exclusions=None):
    in_service, _ = _fleet_partition(vehicles, exclusions)
    used = used_vehicle_ids(legs)
    rows = [
        {
            "vehicle_id": v.get("id"),
            "vehicle_name": v.get("name"),
            "vehicle_status": v.get("vehicle_status"),
            "odometer": v.get("odometer"),
        }
        for v in in_service
        if str(v.get("id")) not in used
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


def build_vehicles_out_of_service(vehicles, oos_history, run_date, exclusions=None, last_seen=None):
    """
    Out-of-service sheet.

    Four columns the SQL produced have no API equivalent: status_reason, the
    work-order odometers, and the work-order station. `oos_since` and days-out
    come from locally accumulated observations instead of the status log, so
    they are blank for any vehicle first seen out of service before this
    history began.
    """
    _, out_of_service = _fleet_partition(vehicles, exclusions)
    last_seen = last_seen or {}
    as_of = run_date if isinstance(run_date, date) else parse_ts(run_date).date()

    rows = []
    for vehicle in out_of_service:
        since, days = oos_history.days_out(vehicle.get("id"), run_date)
        ran_on = last_seen.get(str(vehicle.get("id")))
        days_since_run = (as_of - ran_on).days if ran_on else None
        rows.append(
            {
                "vehicle_id": vehicle.get("id"),
                "vehicle_name": vehicle.get("name"),
                "vehicle_status": vehicle.get("vehicle_status"),
                "current_odometer": vehicle.get("odometer"),
                "oos_since": since,
                "total_days_out_of_service": days,
                "last_ran": ran_on.isoformat() if ran_on else None,
                "days_since_last_run": days_since_run,
            }
        )
    df = pd.DataFrame(rows)
    return df.sort_values("vehicle_name") if not df.empty else df


# =============================
# UNIT-HOUR UTILIZATION
# =============================
def shift_instances(shifts, shift_offset=None):
    """
    shift_name -> the distinct unit-shifts the feed describes.

    The feed returns one row per crew member, so the same unit-shift arrives two
    or three times; identical (start, end) pairs collapse to one instance. This
    is the unit-hour denominator: two medics on one truck for twelve hours is
    twelve unit hours, not twenty-four.
    """
    instances = defaultdict(set)
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = profile_name(shift)
        if not name or _excluded_uhu_profile(name):
            continue
        start = parse_shift_ts(shift.get("start_time"), shift_offset)
        end = parse_shift_ts(shift.get("end_time"), shift_offset)
        if not start or not end or end <= start:
            continue
        instances[name].add((start, end))
    return {name: _merge_spans(spans) for name, spans in instances.items()}


def _merge_spans(spans):
    """
    Collapse overlapping unit-shifts into the period the unit was covered.

    Crew on one truck rarely start and finish together -- a partner leaving
    early is two rows with the same start and different ends, which is one
    truck, not two. Summing them would bill the unit twice for the same hour.
    Shifts that merely meet end-to-end are left alone; those are a handover.
    """
    merged = []
    for start, end in sorted(spans):
        if merged and start < merged[-1][1]:
            merged[-1] = [merged[-1][0], max(merged[-1][1], end)]
        else:
            merged.append([start, end])
    return [tuple(span) for span in merged]


def leg_anchor(leg, start_key):
    """When a leg committed its unit -- the span's own start, or the pickup."""
    stamps = timestamp_map(leg)
    return parse_ts(stamps.get(start_key)) or parse_ts(leg.get("pickup_time"))


def assign_leg(leg, spans, start_key, end_key):
    """
    Which unit-shift a leg belongs to, or None.

    A leg belongs to the instance running when it committed the unit. Falling
    back to any instance the leg's span overlaps catches a call taken moments
    before the shift's recorded start.
    """
    if not spans:
        return None
    anchor = leg_anchor(leg, start_key)
    if anchor:
        for span in spans:
            if span[0] <= anchor < span[1]:
                return span

    stamps = timestamp_map(leg)
    leg_start = parse_ts(stamps.get(start_key)) or anchor
    leg_end = parse_ts(stamps.get(end_key))
    if leg_start and leg_end:
        overlapping = [
            span for span in spans
            if overlap_minutes(leg_start, leg_end, span[0], span[1]) > 0
        ]
        if overlapping:
            return max(
                overlapping,
                key=lambda s: overlap_minutes(leg_start, leg_end, s[0], s[1]),
            )
    return None


# =============================
# RUN VOLUME
# =============================
def has_status(leg):
    """Whether a leg says anything at all about how it ended."""
    return bool((leg.get("trip_status") or "").strip())


def is_run(leg):
    """
    Whether a leg actually ran, by the same rule UHU counts by.

    A leg carrying no status at all is missing data, not evidence of a
    completed transport, and on this tenant those legs carry no shift either --
    the same signature as a call cancelled before it was ever assigned. Counting
    them would quietly inflate run volume, so they are excluded by default.
    Set TS_COUNT_STATUSLESS_LEGS=1 to count them instead.
    """
    status = (leg.get("trip_status") or "").strip().lower()
    if not status:
        return COUNT_STATUSLESS_LEGS
    return status not in UHU_EXCLUDED_TRIP_STATUSES


def build_runs_by_cost_center(legs, cost_center_map):
    """
    Transport volume per cost center for the day.

    A count, not a duration, so unlike UHU it does not depend on crews clearing
    a call promptly -- it is unaffected by whatever the utilized-time span
    settles on. Cancellations are counted separately rather than dropped, so a
    quiet day and a day full of cancelled calls do not look the same.
    """
    runs, cancelled, no_status = Counter(), Counter(), Counter()
    vehicles_seen = defaultdict(set)
    statusless = sum(1 for leg in legs if not has_status(leg))
    if statusless:
        log.info(
            "Run volume: %s leg(s) carry no trip_status and are %s. "
            "Set TS_COUNT_STATUSLESS_LEGS=1 to change that.",
            statusless, "counted" if COUNT_STATUSLESS_LEGS else "not counted as runs",
        )
    for leg in legs:
        centre = cost_center_map.resolve(profile_name(leg)) or UNASSIGNED_COST_CENTER
        if is_run(leg):
            runs[centre] += 1
            if leg.get("vehicle_id"):
                vehicles_seen[centre].add(str(leg.get("vehicle_id")))
        elif has_status(leg):
            cancelled[centre] += 1
        else:
            # Neither a run nor a cancellation -- the row says nothing. Calling
            # it cancelled would invent a fact.
            no_status[centre] += 1

    rows = [
        {
            "cost_center_name": centre,
            "total_runs": runs.get(centre, 0),
            "vehicles_used": len(vehicles_seen.get(centre, ())),
            "runs_per_vehicle": round(runs.get(centre, 0) / len(vehicles_seen[centre]), 2)
            if vehicles_seen.get(centre) else 0,
            "cancelled_legs": cancelled.get(centre, 0),
            "no_status_legs": no_status.get(centre, 0),
        }
        for centre in sorted(set(runs) | set(cancelled) | set(no_status))
    ]
    df = pd.DataFrame(rows)
    return df.sort_values("total_runs", ascending=False) if not df.empty else df


def build_runs_by_vehicle(legs, vehicles, cost_center_map):
    """
    Transport volume per vehicle for the day.

    Every vehicle that ran appears, including ones the fleet reports exclude: an
    exclusion means a truck is decommissioned, and one that ran legs today
    plainly is not. Dropping it here would also stop the vehicle rows summing to
    the cost-center rows, which is the first thing anyone checks.

    A vehicle can work for more than one cost center in a day, so the dominant
    one is named and `cost_centers_served` says whether to trust it.
    """
    by_id = {str(v.get("id")): v for v in vehicles if v.get("id")}

    runs, cancelled, no_status = Counter(), Counter(), Counter()
    centres = defaultdict(Counter)
    names = {}
    for leg in legs:
        # A leg with no vehicle still happened. Bucketing it rather than
        # skipping it keeps the vehicle rows summing to the cost-center rows,
        # and puts the gap on the sheet where someone can see it.
        key = str(leg.get("vehicle_id") or UNASSIGNED_VEHICLE)
        # Trip rows carry a name too, which is the only one left for a vehicle
        # deleted since the leg ran.
        if leg.get("vehicle_name"):
            names.setdefault(key, leg.get("vehicle_name"))
        if is_run(leg):
            runs[key] += 1
            centres[key][cost_center_map.resolve(profile_name(leg)) or UNASSIGNED_COST_CENTER] += 1
        elif has_status(leg):
            cancelled[key] += 1
        else:
            no_status[key] += 1

    rows = []
    for key in sorted(set(runs) | set(cancelled) | set(no_status)):
        vehicle = by_id.get(key, {})
        if key == UNASSIGNED_VEHICLE:
            # Not a missing vehicle record -- a leg that named no vehicle.
            vehicle = {"id": "", "name": UNASSIGNED_VEHICLE, "vehicle_status": "n/a"}
        served = centres.get(key) or Counter()
        rows.append(
            {
                "vehicle_id": vehicle.get("id", key),
                "vehicle_name": vehicle.get("name") or names.get(key) or key,
                "vehicle_status": vehicle.get("vehicle_status") or "Not in vehicle list",
                "cost_center_name": (
                    # Ties break on name so a re-run cannot reorder the sheet.
                    sorted(served.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
                    if served else UNASSIGNED_COST_CENTER
                ),
                "cost_centers_served": len(served),
                "total_runs": runs.get(key, 0),
                "cancelled_legs": cancelled.get(key, 0),
                "no_status_legs": no_status.get(key, 0),
            }
        )
    df = pd.DataFrame(rows)
    return df.sort_values("total_runs", ascending=False) if not df.empty else df


def staffed_hours(intervals, min_crew):
    """
    Hours during which at least `min_crew` of the given intervals overlap.

    A unit is in service when enough of its crew are on the clock together, so
    this is a coverage question rather than a sum. Summing punches would make
    one person working twelve hours look identical to two working six, and only
    the second is a staffed truck.
    """
    if min_crew <= 0:
        return 0.0
    events = []
    for start, end in intervals:
        if start and end and end > start:
            events.append((start, 1))
            events.append((end, -1))
    if not events:
        return 0.0
    # Starts before ends at the same instant, so a clean handover stays covered.
    events.sort(key=lambda e: (e[0], -e[1]))

    total = 0.0
    live = 0
    previous = None
    for moment, delta in events:
        if previous is not None and live >= min_crew:
            total += (moment - previous).total_seconds()
        live += delta
        previous = moment
    return total / 3600.0


def unit_punches_by_instance(shifts, shift_offset=None, metrics_date=None, now=None):
    """
    shift_name -> {unit-shift instance: [(punch start, punch end), ...]}.

    Punches belong to a crew member; a unit-shift is the several crew rows that
    share an instance. Each punch is clipped to its instance, because a punch
    left open would otherwise run until whenever someone noticed, and time
    outside the shift is not that shift's unit hours.

    A punch with no end is bounded at the end of its shift or at `now`,
    whichever comes first. Pass `now` for any run that can overlap a shift
    still in progress, which every morning run does; leave it None to bound
    only at the shift end.

    Returns the grouping, a count of open punches per profile, and how many of
    those are open because the shift has not finished yet -- a crew still on
    the road and a crew who forgot to punch out look identical in the feed but
    mean different things.
    """
    all_instances = shift_instances(shifts, shift_offset)
    by_instance = defaultdict(lambda: defaultdict(list))
    open_punches = Counter()
    still_running = Counter()
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = profile_name(shift)
        if not name or _excluded_uhu_profile(name):
            continue
        start = parse_shift_ts(shift.get("start_time"), shift_offset)
        if not start:
            continue
        instance = next(
            (span for span in all_instances.get(name, ())
             if span[0] <= start < span[1]),
            None,
        )
        if instance is None:
            continue
        if metrics_date is not None and instance[0].date() != metrics_date:
            continue

        for punch in shift.get("punches") or []:
            if punch.get("deleted"):
                continue
            punch_start = parse_shift_ts(punch.get("start_time"), shift_offset)
            punch_end = parse_shift_ts(punch.get("end_time"), shift_offset)
            if not punch_start:
                continue
            if not punch_end:
                # Still on the clock as far as the record goes. Bound it at the
                # end of the shift rather than letting it run away -- but never
                # past now, or a shift still running bills hours nobody has
                # worked yet. An overnight unit probed at 06:00 would otherwise
                # be credited to its 08:00 end.
                open_punches[name] += 1
                punch_end = instance[1]
                if now is not None and now < punch_end:
                    punch_end = now
                    still_running[name] += 1
            punch_start = max(punch_start, instance[0])
            punch_end = min(punch_end, instance[1])
            if punch_end > punch_start:
                by_instance[name][instance].append((punch_start, punch_end))

    return by_instance, open_punches, still_running


def unit_worked_hours(shifts, metrics_date, shift_offset=None, staffing_rules=None,
                      now=None):
    """shift_name -> unit hours actually crewed to minimum staffing."""
    rules = staffing_rules if staffing_rules is not None else UnitStaffingRules()
    by_instance, open_punches, still_running = unit_punches_by_instance(
        shifts, shift_offset, metrics_date, now
    )

    worked = {}
    for name, instances in by_instance.items():
        needed = rules.min_crew(name)
        worked[name] = sum(
            staffed_hours(punches, needed) for punches in instances.values()
        )
    if open_punches:
        running = sum(still_running.values())
        log.info(
            "UHU: %s punch(es) across %s profile(s) were still open. %s were "
            "bounded at now because the shift is still running; %s at the end "
            "of a shift that has already finished, which means a missed "
            "punch-out rather than a crew on the road.",
            sum(open_punches.values()), len(open_punches),
            running, sum(open_punches.values()) - running,
        )
    return worked


def build_uhu(shifts, legs, cost_center_map, metrics_date, span=UHU_SPAN,
              shift_offset=None, staffing_rules=None, now=None):
    """
    Scheduled vs utilized hours per shift profile, attributed by shift instance.

    A unit-shift is the unit of account, not the calendar day. A truck working
    19:00 to 07:00 is one twelve-hour instance belonging to the day it started,
    and the calls it runs after midnight belong to it too -- so `legs` must
    cover the metrics date and the day after, or an overnight unit's tail is
    lost. Utilized time is clipped to the instance for the same reason the
    scheduled side is bounded by it: a leg whose clear stamp never fired would
    otherwise contribute its full wall-clock span, and a unit cannot be utilized
    after its crew went home.

    Legs whose profile has no instance on this date still appear, with zero
    scheduled hours and a meaningless ratio, rather than vanishing. That is the
    normal state for a backfilled date -- the shift feed describes only a
    rolling window around today and cannot be steered to an older one.
    """
    start_key, end_key = UHU_SPANS.get(span, UHU_SPANS["task"])
    all_instances = shift_instances(shifts, shift_offset)
    worked = unit_worked_hours(shifts, metrics_date, shift_offset, staffing_rules,
                               now if now is not None else tenant_now(shift_offset))

    # Worked hours are the honest denominator, but only where punches exist. A
    # feed carrying none at all -- a tenant that does not use them, or a date
    # outside whatever window they are kept for -- would otherwise make every
    # ratio zero and look like a fleet that never turned a wheel. Fall back
    # rather than publish that, and say so.
    denominator_is_worked = UHU_DENOMINATOR == "worked"
    if denominator_is_worked and not any(shift.get("punches") for shift in shifts):
        denominator_is_worked = False
        log.warning(
            "UHU: no shift row carries punches, so unit hours cannot be measured "
            "from them. Falling back to scheduled hours for %s -- ratios will "
            "count units that were rostered but may not have run.",
            metrics_date,
        )

    # A unit-shift belongs to the day it starts on.
    todays = {
        name: [s for s in spans if s[0].date() == metrics_date]
        for name, spans in all_instances.items()
    }
    todays = {name: spans for name, spans in todays.items() if spans}

    scheduled = {
        name: sum((e - s).total_seconds() for s, e in spans) / 3600.0
        for name, spans in todays.items()
    }

    utilized = defaultdict(float)
    runs = Counter()
    adjacent = orphaned = 0
    for leg in legs:
        name = profile_name(leg)
        if not name or _excluded_uhu_profile(name):
            continue
        # A cancelled or disregarded leg was never a run and must not dilute
        # hours-per-run, even though it still carries a shift assignment.
        if not is_run(leg):
            continue

        spans = todays.get(name)
        instance = assign_leg(leg, spans, start_key, end_key)
        if spans and instance is None:
            # Legs are fetched for the next day as well, to catch an overnight
            # unit's tail, so most of these belong to an adjacent day's instance
            # and are meant to be skipped. One belonging to no instance at all
            # is the interesting case.
            if assign_leg(leg, all_instances.get(name), start_key, end_key):
                adjacent += 1
            else:
                orphaned += 1
            continue

        runs[name] += 1
        stamps = timestamp_map(leg)
        leg_start = parse_ts(stamps.get(start_key))
        leg_end = parse_ts(stamps.get(end_key))
        if not leg_start or not leg_end or leg_end <= leg_start:
            continue
        if instance:
            minutes = overlap_minutes(leg_start, leg_end, instance[0], instance[1])
        else:
            # No schedule to bound it; the row carries no ratio either way.
            minutes = (leg_end - leg_start).total_seconds() / 60.0
        utilized[name] += minutes / 60.0

    attributed = sum(runs.values())
    total = attributed + adjacent + orphaned
    if adjacent or orphaned:
        log.info(
            "UHU %s: %s leg(s) counted, %s belong to an adjacent day's unit "
            "(expected -- the next day is fetched for the overnight tail), "
            "%s match no shift instance at all.",
            metrics_date, attributed, adjacent, orphaned,
        )
    # Only an orphan is a real signal. A leg landing outside every instance of
    # its own profile, on any date, is what a shift/trip clock mismatch looks
    # like from the inside.
    if orphaned > 20 and orphaned > total * 0.25:
        log.warning(
            "UHU: %s of %s legs match no shift instance on any date. That is "
            "the signature of a shift/trip clock mismatch rather than real "
            "scheduling -- check the offset applied to shift times "
            "(TS_SHIFT_UTC_OFFSET_HOURS).",
            orphaned, total,
        )

    rows = []
    unstaffed = []
    for name in sorted(set(scheduled) | set(utilized) | set(runs) | set(worked)):
        scheduled_hours = round(scheduled.get(name, 0.0), 2)
        worked_hours = round(worked.get(name, 0.0), 2)
        utilized_hours = round(utilized.get(name, 0.0), 2)
        total_runs = runs.get(name, 0)
        # A profile rostered but never crewed to minimum staffing did not run.
        denominator = worked_hours if denominator_is_worked else scheduled_hours
        if scheduled_hours and not worked_hours:
            unstaffed.append(name)
        rows.append(
            {
                "cost_center_name": cost_center_map.resolve(name) or UNASSIGNED_COST_CENTER,
                "shift_profile_name": name,
                "scheduled_hours": scheduled_hours,
                "worked_hours": worked_hours,
                "utilized_hours": utilized_hours,
                "total_runs": total_runs,
                "hours_per_run": round(utilized_hours / total_runs, 3) if total_runs else 0,
                "uhu_ratio": round(utilized_hours / denominator, 3) if denominator else 0,
            }
        )

    if denominator_is_worked and unstaffed:
        log.info(
            "UHU: %s profile(s) were rostered on %s but never reached minimum "
            "crew, so they contribute no unit hours: %s",
            len(unstaffed), metrics_date, ", ".join(sorted(unstaffed)[:8]),
        )
    return pd.DataFrame(rows)


def build_uhu_by_shift_profile(shifts, legs, cost_center_map, metrics_date, span=UHU_SPAN, staffing_rules=None,
                               shift_offset=None, now=None):
    df = build_uhu(shifts, legs, cost_center_map, metrics_date, span, shift_offset,
                   staffing_rules, now)
    if df.empty:
        return df
    return df.sort_values("uhu_ratio", ascending=False)


def build_uhu_by_cost_center(shifts, legs, cost_center_map, metrics_date, span=UHU_SPAN, staffing_rules=None,
                             shift_offset=None, by_profile=None, now=None):
    # by_profile lets a caller that already built the per-profile frame roll it
    # up rather than rebuilding it, which also stops the diagnostics being
    # logged twice for one run.
    df = by_profile if by_profile is not None else build_uhu(
        shifts, legs, cost_center_map, metrics_date, span, shift_offset,
        staffing_rules, now
    )
    if df.empty:
        return pd.DataFrame(
            columns=[
                "cost_center_name", "scheduled_hours", "worked_hours",
                "utilized_hours", "total_runs", "hours_per_run", "uhu_ratio",
            ]
        )
    grouped = (
        df.groupby("cost_center_name", dropna=False)
        .agg(
            scheduled_hours=("scheduled_hours", "sum"),
            worked_hours=("worked_hours", "sum"),
            utilized_hours=("utilized_hours", "sum"),
            total_runs=("total_runs", "sum"),
        )
        .reset_index()
    )
    grouped["hours_per_run"] = grouped.apply(
        lambda r: round(r["utilized_hours"] / r["total_runs"], 3) if r["total_runs"] else 0,
        axis=1,
    )
    denominator_col = "worked_hours" if UHU_DENOMINATOR == "worked" else "scheduled_hours"
    grouped["uhu_ratio"] = grouped.apply(
        lambda r: round(r["utilized_hours"] / r[denominator_col], 3) if r[denominator_col] else 0,
        axis=1,
    )
    grouped["worked_hours"] = grouped["worked_hours"].round(2)
    grouped["scheduled_hours"] = grouped["scheduled_hours"].round(2)
    grouped["utilized_hours"] = grouped["utilized_hours"].round(2)
    return grouped.sort_values("uhu_ratio", ascending=False)


# =============================
# ORCHESTRATION
# =============================
def fetch_day(api, metrics_date):
    """
    One pull of everything a day's reports need.

    Trips are fetched for the target date and honour backfill. Shifts,
    employees and vehicles describe the present regardless of the date asked
    for, which is why the accumulated state files exist.
    """
    if isinstance(metrics_date, str):
        metrics_date = datetime.strptime(metrics_date, "%Y-%m-%d").date()

    log.info("Fetching trips for %s ...", metrics_date)
    legs = api.get_trips(metrics_date, range_days=1)

    # UHU attributes a leg to the unit-shift that was running when it started,
    # and an overnight shift's calls land on the following date. Those legs
    # belong to this date's report but not to any other sheet, so they are
    # fetched separately rather than widening `legs` under OTP and the fleet
    # reports, which are strictly one-day.
    log.info("Fetching trips for %s (overnight tail, UHU only) ...",
             metrics_date + timedelta(days=1))
    uhu_legs = legs + api.get_trips(metrics_date + timedelta(days=1), range_days=1)

    log.info("Fetching shifts (rolling window, not %s) ...", metrics_date)
    shifts = api.list_shifts()

    log.info("Fetching employees ...")
    employees = api.list_employees()

    log.info("Fetching vehicles ...")
    vehicles = api.list_vehicles()

    log.info(
        "Fetched %s legs (%s including the overnight tail), %s shift rows, "
        "%s employees, %s vehicles",
        len(legs), len(uhu_legs), len(shifts), len(employees), len(vehicles),
    )
    return {
        "metrics_date": metrics_date,
        "legs": legs,
        "uhu_legs": uhu_legs,
        "shifts": shifts,
        "employees": employees,
        "vehicles": vehicles,
    }


def fetch_fleet_activity(api, metrics_date, lookback_days=None):
    """
    Last date each vehicle ran a leg, over a trailing window.

    One GetTrips call: the lookback is capped at the API's 31-day range limit.
    Used to flag out-of-service vehicles that have not moved in a long time,
    which is how abandoned records give themselves away without status_reason.
    """
    lookback = min(int(lookback_days or FLEET_ACTIVITY_LOOKBACK_DAYS), 31)
    start = metrics_date - timedelta(days=lookback - 1)
    legs = api.get_trips(start, range_days=lookback)

    by_day = defaultdict(list)
    for leg in legs:
        pickup = parse_ts(leg.get("pickup_time"))
        if pickup:
            by_day[pickup.date()].append(leg)
    return vehicle_last_seen(by_day)


def build_all(data, cost_center_map=None, oos_history=None, now=None,
              exclusions=None, fleet_activity=None, staffing_rules=None):
    """Build every report DataFrame from one fetch, refreshing accumulated state."""
    metrics_date = data["metrics_date"]
    legs = data["legs"]
    # Absent when a caller built the fetch dict itself; UHU then loses only the
    # overnight tail rather than failing.
    uhu_legs = data.get("uhu_legs") or legs
    shifts = data["shifts"]
    employees = data["employees"]
    vehicles = data["vehicles"]
    # Shift times are UTC while everything else is local; resolve the gap once
    # from the trips themselves and hand it to everything that reads a shift.
    shift_offset = resolve_shift_offset(uhu_legs)
    if shift_offset:
        log.info("Shift times shifted by %s to reach tenant-local time.", shift_offset)

    # "Now" has to be on the tenant's clock, not the reporting machine's, since
    # everything it is compared against -- shift instances, punches -- has been
    # moved onto that clock above.
    if now is None:
        now = tenant_now(shift_offset)

    cost_center_map = cost_center_map or CostCenterMap()
    oos_history = oos_history or OutOfServiceHistory()
    exclusions = exclusions if exclusions is not None else VehicleExclusions()
    staffing_rules = staffing_rules if staffing_rules is not None else UnitStaffingRules()

    # Learn from the present before attributing the past.
    cost_center_map.update(shifts, employees)
    oos_history.update(
        [v for v in vehicles if not exclusions.excludes(v)], metrics_date
    )

    dropped = [v for v in vehicles if exclusions.excludes(v)]
    if dropped:
        log.info(
            "Excluded %s vehicle(s) from the fleet reports: %s",
            len(dropped),
            ", ".join(sorted(str(v.get("name")) for v in dropped)[:10]),
        )

    scored = scored_legs(legs, cost_center_map)
    uhu_by_profile = build_uhu_by_shift_profile(
        shifts, uhu_legs, cost_center_map, metrics_date,
        shift_offset=shift_offset, staffing_rules=staffing_rules, now=now,
    )

    reports = {
        "otp_by_call_type": build_otp_by_call_type(scored),
        "otp_by_cost_center": build_otp_by_cost_center(scored),
        "staffing_active_now": build_staffing_active_now(
            shifts, employees, now, shift_offset, staffing_rules
        ),
        "staffing_tomorrow": build_staffing_for_date(
            shifts, employees, now.date() + timedelta(days=1), shift_offset,
            staffing_rules
        ),
        "vehicle_summary": build_vehicle_summary(vehicles, legs, metrics_date, exclusions),
        "vehicles_in_use": build_vehicles_in_use(vehicles, legs, exclusions),
        "vehicles_unused_in_service": build_vehicles_unused_in_service(vehicles, legs, exclusions),
        "vehicles_all_in_service": build_vehicles_all_in_service(vehicles, exclusions),
        "vehicles_out_of_service": build_vehicles_out_of_service(
            vehicles, oos_history, metrics_date, exclusions, fleet_activity
        ),
        "runs_by_cost_center": build_runs_by_cost_center(legs, cost_center_map),
        "runs_by_vehicle": build_runs_by_vehicle(legs, vehicles, cost_center_map),
        "uhu_by_cost_center": build_uhu_by_cost_center(
            shifts, uhu_legs, cost_center_map, metrics_date,
            shift_offset=shift_offset, by_profile=uhu_by_profile,
        ),
        "uhu_by_shift_profile": uhu_by_profile,
    }

    cost_center_map.save()
    oos_history.save()
    return reports


# =============================
# NOT A SCRIPT
# =============================
if __name__ == "__main__":
    # Running a library file does nothing and says nothing, which reads exactly
    # like a report that ran and produced no output. Say which file to run.
    import sys as _sys
    _sys.stderr.write(
        "\n{me} is a library, not a script -- it defines how the reports are\n"
        "built and does nothing on its own.\n\n"
        "To produce the reports, from the folder holding these files:\n"
        "    .venv\\Scripts\\python.exe daily_report_runner_api.py --zip\n\n"
        "With no date that reports yesterday, which is what the daily run wants.\n"
        "Add a date to redo one: daily_report_runner_api.py 2026-08-19 --zip\n\n"
        .format(me=__file__)
    )
    raise SystemExit(2)
