"""
Find out whether Samsara knows which market a vehicle belongs to.

The Traumasoft ThirdParty API puts no cost center and no vehicle class on a
vehicle -- the allowlist is id, name, vehicle_status, vin, odometer, disabled,
deleted plus live-shift enrichment, and probe_traumasoft_api.py has already
confirmed against the live API that cost_center_id, cost_center_name,
status_reason and oos_since are simply absent. That is why every fleet sheet
is company-wide rather than regional.

Samsara returns tags on its vehicle list, and nothing in this repo has ever
read them: the route publisher joins Traumasoft to Samsara by unit prefix and
ignores everything else on the record. So "the tags carry the station" is an
assumption, and this answers it before anyone builds on it.

What it reports:

    1. Fleet record shape   -- which top-level fields Samsara actually
                               populates, so a field carrying the class
                               (AMB / MH / WC) is spotted if one exists
    2. Tag vocabulary       -- every distinct tag, how many vehicles carry it,
                               and any parent/child structure
    3. Attribute vocabulary -- Samsara's other free-form slot, for the same
                               reason
    4. The join             -- how many Traumasoft units reach a Samsara
                               vehicle at all, since an unmatched unit gets no
                               tag no matter how good the tags are
    5. Tags per unit        -- the table that actually answers the question
    6. Tags vs cost centers -- whether the tag vocabulary can be mapped onto
                               the cost center names Traumasoft uses
    7. Vehicle status       -- the current out-of-service picture, for context

STRICTLY READ-ONLY. Every Samsara call goes through SamsaraClient's default
read_only=True, whose guard refuses a write on the caller's own declaration
rather than inferring it from the HTTP method. Nothing here passes write=True.

Vehicle names, tags and cost centers are operational values, not patient data,
and are printed in full. The output is safe to paste into a chat or an issue.

Usage:
    python probe_samsara_tags.py
    python probe_samsara_tags.py --json tags.json
"""

import argparse
import hashlib
import json
import logging
import re
import sys
from collections import Counter, defaultdict

import samsara_routes as SR
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

log = logging.getLogger(__name__)


def build_id():
    """
    A short hash of the code actually running.

    Same reason as the readiness probe: these files are copied onto the
    reporting machine by hand, and a cached download can quietly hand back the
    previous version. A report that does not say which build produced it can
    be read as current when it is not.
    """
    digest = hashlib.sha256()
    for path in (__file__, SR.__file__):
        try:
            with open(path, "rb") as handle:
                digest.update(handle.read())
        except OSError:
            return "unknown"
    return digest.hexdigest()[:8]


def pct(count, total):
    return f"{(100.0 * count / total):5.1f}%" if total else "    -"


def populated(value):
    if value is None:
        return False
    if isinstance(value, (list, dict)):
        return bool(value)
    text = str(value).strip()
    return bool(text) and text.lower() not in ("none", "null")


def tag_entries(vehicle):
    """
    A vehicle's tags as (name, parent id) pairs.

    Samsara returns tags as objects carrying id/name/parentTagId, but a bare
    list of strings is cheap to tolerate and costs nothing to allow for.
    """
    entries = []
    for tag in vehicle.get("tags") or []:
        if isinstance(tag, dict):
            name = str(tag.get("name") or tag.get("id") or "").strip()
            if name:
                entries.append((name, tag.get("parentTagId")))
        elif str(tag).strip():
            entries.append((str(tag).strip(), None))
    return entries


def normalize(text):
    """
    Collapse a name to comparable letters and digits.

    'Lynx EMS LLC dba Lynx Boardman' and 'Boardman' have to be recognisable as
    the same place, so the legal-entity wrapper is dropped along with spacing
    and punctuation.
    """
    text = str(text or "").lower()
    text = re.sub(r"\b(lynx|ems|llc|dba|inc|the)\b", " ", text)
    return re.sub(r"[^a-z0-9]+", "", text)


# =============================
# SECTIONS
# =============================
def report_fleet_shape(vehicles, out):
    """
    Which fields Samsara populates, and how often.

    Printed before the tags because the class (AMB / MH / WC) has to come from
    somewhere too, and if Samsara carries it on a dedicated field that is a
    better answer than reading it out of a tag.
    """
    print("\n1. SAMSARA FLEET RECORD SHAPE")
    print("   " + "-" * 66)
    total = len(vehicles)
    print(f"   {total} Samsara vehicle(s).\n")
    if not total:
        out["fleet_shape"] = {}
        return

    counts = Counter()
    for vehicle in vehicles:
        for key, value in vehicle.items():
            if populated(value):
                counts[key] += 1

    print(f"   {'field':<32}{'populated':>10}   {'sample value'}")
    print("   " + "-" * 66)
    shape = {}
    for key in sorted(counts, key=lambda k: (-counts[k], k)):
        sample = next(
            (v.get(key) for v in vehicles if populated(v.get(key))), None
        )
        sample_text = json.dumps(sample, default=str)
        if len(sample_text) > 28:
            sample_text = sample_text[:25] + "..."
        print(f"   {key[:31]:<32}{pct(counts[key], total):>10}   {sample_text}")
        shape[key] = counts[key]
    out["fleet_shape"] = shape


def report_tag_vocabulary(vehicles, out):
    """Every distinct tag and how much of the fleet carries it."""
    print("\n2. TAG VOCABULARY")
    print("   " + "-" * 66)

    per_tag = Counter()
    parents = {}
    untagged = 0
    for vehicle in vehicles:
        entries = tag_entries(vehicle)
        if not entries:
            untagged += 1
        for name, parent in entries:
            per_tag[name] += 1
            if parent is not None:
                parents[name] = parent

    if not per_tag:
        print("   No tags on any vehicle. Samsara cannot supply the cost")
        print("   center this way -- see section 6 for what is left.")
        out["tags"] = {}
        return per_tag

    print(f"   {len(per_tag)} distinct tag(s). "
          f"{untagged} of {len(vehicles)} vehicle(s) carry none.\n")
    print(f"   {'tag':<38}{'vehicles':>9}{'parent':>12}")
    print("   " + "-" * 60)
    for name, count in per_tag.most_common():
        parent = parents.get(name)
        print(f"   {name[:37]:<38}{count:>9}{str(parent or '-'):>12}")

    # A tag applied to every vehicle says nothing about which market a unit is
    # in -- it is a fleet-wide label. Worth naming so nobody maps one.
    blanket = [n for n, c in per_tag.items() if c == len(vehicles)]
    if blanket:
        print(f"\n   Applied to the whole fleet (carries no market signal): "
              f"{', '.join(blanket)}")

    out["tags"] = dict(per_tag)
    out["tag_parents"] = {name: parents.get(name) for name in per_tag}
    out["untagged_vehicles"] = untagged
    return per_tag


def report_attributes(vehicles, out):
    """Samsara's other free-form slot, checked for the same reason as tags."""
    print("\n3. ATTRIBUTE VOCABULARY")
    print("   " + "-" * 66)

    per_attr = defaultdict(Counter)
    for vehicle in vehicles:
        for attr in vehicle.get("attributes") or []:
            if not isinstance(attr, dict):
                continue
            name = str(attr.get("name") or "").strip()
            if not name:
                continue
            values = attr.get("stringValues") or attr.get("numberValues") or []
            if not values:
                per_attr[name]["(no value)"] += 1
            for value in values:
                per_attr[name][str(value)] += 1

    if not per_attr:
        print("   No attributes on any vehicle.")
        out["attributes"] = {}
        return

    for name, values in per_attr.items():
        print(f"\n   {name}")
        for value, count in values.most_common(20):
            print(f"      {value[:44]:<46}{count:>5}")
    out["attributes"] = {k: dict(v) for k, v in per_attr.items()}


def report_join(ts_vehicles, sam_vehicles, out):
    """
    How much of the Traumasoft fleet reaches a Samsara vehicle.

    This bounds everything else: a unit with no Samsara match gets no tag
    however well the tags are maintained, and it is the whole fleet roster
    being joined here rather than the units that happened to run legs, because
    a vehicle sitting out of service all month is exactly the one the report
    is asking about.
    """
    print("\n4. TRAUMASOFT x SAMSARA JOIN")
    print("   " + "-" * 66)

    ts_names = sorted({str(v.get("name") or "").strip() for v in ts_vehicles} - {""})
    overrides = SR.load_vehicle_overrides()
    matched, unmatched, ambiguous = SR.match_vehicles(ts_names, sam_vehicles, overrides)

    print(f"   {len(ts_names)} Traumasoft unit(s), {len(sam_vehicles)} Samsara vehicle(s).")
    print(f"   Matched {len(matched)}   unmatched {len(unmatched)}   "
          f"ambiguous {len(ambiguous)}   "
          f"({pct(len(matched), len(ts_names)).strip()} joined)")

    # That percentage is over the RAW list, which on this tenant is mostly
    # iPads and retired trucks -- it reads far worse than the join really is.
    # The number that matters is the join across the roster the fleet sheets
    # actually see, so it is computed here rather than left to be inferred
    # from section 9.
    try:
        import traumasoft_reports as R
        exclusions = R.VehicleExclusions()
        live = [
            v for v in ts_vehicles
            if not (R._is_truthy_flag(v.get("deleted"))
                    or R._is_truthy_flag(v.get("disabled")))
            and v.get("vehicle_status") not in R.NON_FLEET_STATUSES
            and not exclusions.excludes(v)
        ]
        live_names = sorted({str(v.get("name") or "").strip() for v in live} - {""})
        live_matched = [n for n in live_names if n in matched]
        print(f"\n   Across the {len(live_names)} unit(s) that reach the fleet sheets "
              f"(section 9):")
        print(f"   Matched {len(live_matched)}   unmatched "
              f"{len(live_names) - len(live_matched)}   "
              f"({pct(len(live_matched), len(live_names)).strip()} joined)")
        missing = [n for n in live_names if n not in matched]
        if missing:
            print(f"   These are the ones that actually matter:")
            print("      " + ", ".join(missing)[:220])
    except Exception as exc:  # noqa: BLE001 -- the raw figure above still stands
        log.info("Could not scope the join to the fleet sheets (%s).", exc)
        live_names, live_matched = [], []
    print()

    if unmatched:
        print(f"   No Samsara vehicle ({len(unmatched)}):")
        for name in unmatched:
            print(f"      {name[:40]:<42}prefix {SR.unit_prefix(name)}")
    if ambiguous:
        print(f"\n   Prefix hit several ({len(ambiguous)}):")
        for name, hits in ambiguous.items():
            print(f"      {name[:30]:<32}"
                  f"{', '.join(v.get('name', '?') for v in hits)[:34]}")
    if unmatched or ambiguous:
        print(f"\n   Fix in {SR.VEHICLE_OVERRIDES_FILE}")

    out["join"] = {
        "fleet_sheet_units": len(live_names),
        "fleet_sheet_matched": len(live_matched),
        "traumasoft_units": len(ts_names),
        "samsara_vehicles": len(sam_vehicles),
        "matched": {k: v.get("name") for k, v in matched.items()},
        "unmatched": unmatched,
        "ambiguous": {k: [v.get("name") for v in hits] for k, hits in ambiguous.items()},
    }
    return matched


def station_matches(tag, centre):
    """
    Whether a tag names this cost center.

    Exact, then ignoring the legal-entity wrapper. NO substring pass: on this
    tenant that would have `Indiana` claim `Indianapolis`, and `Columbus`
    claim two states. The near-misses this tenant actually has --
    Ellicott/Ellicott City, Newburgh/Newburg, Parma Heights/Parma -- are
    deliberately left to state/tag_cost_center_map.json, because each is a
    decision about which station a tag means and a substring rule would make
    the same call for a pair that is genuinely two places.
    """
    tag, centre = str(tag or "").strip(), str(centre or "").strip()
    if not tag or not centre:
        return False
    if tag.lower() == centre.lower():
        return True
    return normalize(tag) == normalize(centre) and bool(normalize(tag))


def report_tags_per_unit(ts_vehicles, matched, out, centres=None):
    """
    The table the whole probe exists for: each unit and the station its tags name.

    The first version of this printed the raw tag list truncated to fit, and
    that hid the answer. Samsara returns tags alphabetically and this tenant
    tags almost everything `In Service (TraumaSoft)`, so any station later in
    the alphabet -- Massillon, Ohio, Parma, Sandusky, Shelby, Toledo -- was
    pushed off the end of the line. Half the fleet read as having no station
    tag when it had one all along.

    So the station is resolved into its own column, by the same rule
    build_vehicle_cost_centers uses, and the remaining tags follow. A unit
    whose tags name two stations says so rather than showing the first one
    that fits.

    Ordered by Traumasoft name, so a naming convention -- if the tags follow
    one -- is visible down the column rather than having to be inferred.
    """
    print("\n5. TAGS PER TRAUMASOFT UNIT")
    print("   " + "-" * 66)
    print(f"   {'Traumasoft unit':<22}{'status':<18}{'station from tags':<20}{'other tags'}")
    print("   " + "-" * 78)

    centres = sorted(centres or ())
    rows, resolved_count, ambiguous, unstationed = [], 0, [], []
    for vehicle in sorted(ts_vehicles, key=lambda v: str(v.get("name") or "")):
        name = str(vehicle.get("name") or "").strip()
        if not name:
            continue
        status = str(vehicle.get("vehicle_status") or "-")
        sam = matched.get(name)
        if sam is None:
            print(f"   {name[:21]:<22}{status[:17]:<18}{'-- no Samsara match --':<20}")
            rows.append({"name": name, "vehicle_status": vehicle.get("vehicle_status"),
                         "samsara_name": None, "tags": None, "station": None})
            continue

        tags = [t for t, _ in tag_entries(sam)]
        hits = sorted({c for c in centres for t in tags if station_matches(t, c)})
        if len(hits) == 1:
            station, resolved_count = hits[0], resolved_count + 1
        elif hits:
            station = "AMBIGUOUS: " + "/".join(hits)
            ambiguous.append((name, hits))
        else:
            station = "-- none --"
            unstationed.append(name)
        others = [t for t in tags
                  if not any(station_matches(t, c) for c in centres)]
        print(f"   {name[:21]:<22}{status[:17]:<18}{station[:19]:<20}"
              f"{', '.join(others)[:28]}")
        rows.append({"name": name, "vehicle_status": vehicle.get("vehicle_status"),
                     "samsara_name": sam.get("name"), "tags": tags,
                     "station": hits[0] if len(hits) == 1 else None,
                     "station_candidates": hits})

    print(f"\n   Resolved to exactly one station: {resolved_count}")
    if ambiguous:
        print(f"   Tags name more than one ({len(ambiguous)}) -- these cannot be")
        print("   placed without a decision:")
        for name, hits in ambiguous:
            print(f"      {name[:22]:<24}{'/'.join(hits)}")
    if unstationed:
        print(f"   Matched in Samsara but no station tag ({len(unstationed)}):")
        print("      " + ", ".join(unstationed[:25])[:220])

    out["units"] = rows
    out["station_summary"] = {
        "resolved": resolved_count,
        "ambiguous": {name: hits for name, hits in ambiguous},
        "no_station_tag": unstationed,
    }


def fetch_cost_centers(api):
    """
    Every cost center name Traumasoft knows, from where they actually live.

    Hoisted out of section 6 so section 5 can resolve a station too: without
    it that table could only print raw tags, which is what hid the answer.
    """
    centres = set()
    try:
        for employee in api.list_employees():
            name = employee.get("cost_center_name")
            if name:
                centres.add(str(name).strip())
    except TraumasoftAPIError as exc:
        log.warning("Employee roster unavailable: %s", exc)

    # The dedicated endpoint exists in the spec but nothing in this repo has
    # needed it; try it too, since it would be the better source if populated.
    try:
        for row in api.list_cost_centers():
            name = row.get("name") or row.get("cost_center_name")
            if name:
                centres.add(str(name).strip())
    except (TraumasoftAPIError, AttributeError) as exc:
        log.info("Cost center list unavailable (%s); employee roster only.", exc)
    return centres


def is_child_tag(tag, parents):
    """
    Whether a tag sits under another tag in Samsara's hierarchy.

    This is what separates a station from a state. On this tenant the
    top-level tags are ALL, Ohio, West Virginia, Indiana, Maryland, Colorado,
    BLS, Secure Car, Wheelchair and In Service (TraumaSoft) -- none of which
    names a station -- while every station tag hangs off a state.

    It matters because a state name can be a substring of a station inside it.
    `Indiana` is a parent of Indianapolis, Salem, Sellersburg and Newburg, and
    a near-miss rule mapping `Indiana` to `Indianapolis` would hand Salem's
    and Sellersburg's trucks to Indianapolis. The hierarchy rules that out
    where a string comparison cannot.
    """
    return bool((parents or {}).get(tag))


def report_tags_vs_cost_centers(api, tag_counts, out, centres=None, parents=None):
    """
    Whether the tag vocabulary can be mapped onto Traumasoft's cost centers.

    The cost center names come from the employee roster, which is where they
    actually live -- there is no vehicle-side source, which is the whole
    reason for this probe. The matching here is deliberately crude: it exists
    to show whether a mapping is plausible, not to be the mapping. That file
    should be written by hand, the way the region and override files are,
    because a tag quietly claiming the wrong station is the failure mode.
    """
    print("\n6. TAGS vs TRAUMASOFT COST CENTERS")
    print("   " + "-" * 66)

    centres = set(centres or ()) or fetch_cost_centers(api)
    if not centres:
        print("   No cost center names available to compare against.")
        out["cost_centers"] = []
        return

    print(f"   {len(centres)} cost center name(s) in Traumasoft:\n")
    for name in sorted(centres):
        print(f"      {name}")

    if not tag_counts:
        print("\n   No tags to map. Section 2 found none.")
        out["cost_centers"] = sorted(centres)
        return

    # A tag that matches outright is usable as it stands. A tag that only
    # matches loosely is a DECISION -- which station does `Parma` mean, is
    # `Newburg` the same place as `Newburgh` -- and those go in a
    # hand-written map, so they are separated here rather than presented as
    # equally good candidates.
    print(f"\n   {'cost center':<26}{'matches outright':<20}{'near miss -- needs a rule'}")
    print("   " + "-" * 74)
    pairs, near, rejected = {}, {}, {}
    for centre in sorted(centres):
        key = normalize(centre)
        exact = [tag for tag in tag_counts if station_matches(tag, centre)]
        loose = [
            tag for tag in tag_counts
            if tag not in exact and key and normalize(tag) and (
                key in normalize(tag) or normalize(tag) in key
            )
        ]
        # A top-level tag is never a near miss for a station beneath it. This
        # caught `Indiana` being offered as a rule for `Indianapolis`, which
        # would have handed Salem's and Sellersburg's trucks to Indianapolis
        # -- the exact mistake the no-substring rule exists to prevent, made
        # by the thing suggesting the rules.
        suggestable = [t for t in loose if is_child_tag(t, parents)]
        parent_level = [t for t in loose if t not in suggestable]
        pairs[centre] = exact
        if suggestable:
            near[centre] = suggestable
        if parent_level:
            rejected[centre] = parent_level
        print(f"   {centre[:25]:<26}{(', '.join(exact) or '--')[:19]:<20}"
              f"{', '.join(loose)[:28]}")

    if near:
        print(f"\n   {len(near)} cost center(s) have only a near miss. Each needs a line")
        print("   in state/tag_cost_center_map.json -- a substring rule would make")
        print("   the same call for a pair that is genuinely two places:\n")
        for centre, tags in sorted(near.items()):
            for tag in tags:
                print(f'      "{tag}": "{centre}"')
    if rejected:
        print(f"\n   NOT offered as rules, though the names overlap: these are")
        print("   top-level tags, and a station sits beneath them rather than")
        print("   being them:")
        for centre, tags in sorted(rejected.items()):
            for tag in tags:
                children = sorted(
                    t for t, p in (parents or {}).items()
                    if p and p == (parents or {}).get(tag)
                ) or "its children"
                print(f"      {tag} -> {centre}   would also claim every other")
                print(f"         station under {tag}")
    out["near_misses"] = near
    out["rejected_near_misses"] = rejected

    # Tags a near-miss rule reaches are not unclaimed -- they just need the
    # line above pasting in. Listing them together read as though four more
    # stations had no tag at all.
    suggested = {t for tags in near.values() for t in tags}
    unclaimed = [
        t for t in tag_counts
        if not any(t in h for h in pairs.values()) and t not in suggested
    ]
    if suggested:
        print(f"\n   Reachable once the rules above are in place "
              f"({len(suggested)}): {', '.join(sorted(suggested))}")
    if unclaimed:
        print(f"\n   Tags naming no cost center at all ({len(unclaimed)}): "
              f"{', '.join(sorted(unclaimed))[:200]}")

    out["cost_centers"] = sorted(centres)
    out["tag_cost_center_candidates"] = pairs


def report_traumasoft_roster(ts_vehicles, out):
    """
    What the fleet sheets actually see, as opposed to what the API returns.

    Everything above this reads the raw vehicle list, which is not the roster
    any report is built from: the fleet sheets drop deleted and disabled
    records, drop the non-fleet statuses, and apply VehicleExclusions before
    counting anything. Reporting the raw list alongside findings about the
    reports invites the conclusion that the sheets are full of iPads, which
    they are not.

    So this runs the same filter chain the reports run -- by importing it
    rather than restating it, since a second copy of the rule would drift --
    and reports the funnel. What matters is the last row: the records that
    survive every filter and still collide, because those are the ones
    actually double-counted on a live sheet.
    """
    print("\n9. WHAT THE FLEET SHEETS ACTUALLY SEE")
    print("   " + "-" * 66)

    total = len(ts_vehicles)
    counts = Counter()
    for vehicle in ts_vehicles:
        for key, value in vehicle.items():
            if populated(value):
                counts[key] += 1
    print(f"   Traumasoft returns {total} record(s). Field population\n"
          f"   (presence, not truth -- a boolean reads 100% when it is\n"
          f"   always returned, whatever its value):\n")
    print(f"   {'field':<28}{'present':>10}")
    print("   " + "-" * 40)
    for key in sorted(counts, key=lambda k: (-counts[k], k)):
        print(f"   {key[:27]:<28}{pct(counts[key], total):>10}")

    try:
        import traumasoft_reports as R
    except Exception as exc:  # noqa: BLE001 -- probe must still be useful without it
        print(f"\n   traumasoft_reports unavailable ({type(exc).__name__}: {exc}).")
        print("   Cannot reproduce the report's filter chain; funnel skipped.")
        out["roster"] = {"field_population": dict(counts), "funnel": None}
        return

    flagged = lambda v: (  # noqa: E731
        R._is_truthy_flag(v.get("deleted")) or R._is_truthy_flag(v.get("disabled"))
    )
    exclusions = R.VehicleExclusions()

    # The list call asks the server to omit these. Whether it honours that is
    # worth knowing on its own: if deleted records arrive anyway, every caller
    # that trusts the parameter instead of re-checking the field is wrong.
    really_deleted = sum(1 for v in ts_vehicles if R._is_truthy_flag(v.get("deleted")))
    really_disabled = sum(1 for v in ts_vehicles if R._is_truthy_flag(v.get("disabled")))
    print(f"\n   deleted=true returned despite include_deleted=false:   {really_deleted}")
    print(f"   disabled=true returned despite include_disabled=false: {really_disabled}")
    if really_deleted or really_disabled:
        print("   The server did NOT honour the parameter. The reports re-check")
        print("   the field, so their sheets are unaffected -- but anything that")
        print("   trusts the parameter alone is counting deleted trucks.")

    survivors, dropped = [], defaultdict(list)
    for vehicle in ts_vehicles:
        name = str(vehicle.get("name") or "?")
        if flagged(vehicle):
            dropped["deleted or disabled"].append(name)
        elif vehicle.get("vehicle_status") in R.NON_FLEET_STATUSES:
            dropped[f"status {vehicle.get('vehicle_status')}"].append(name)
        else:
            reason = exclusions.excludes(vehicle)
            if reason:
                dropped[f"excluded: {reason}"].append(name)
            else:
                survivors.append(vehicle)

    print(f"\n   {'filter':<42}{'dropped':>9}")
    print("   " + "-" * 54)
    for reason in sorted(dropped, key=lambda r: -len(dropped[r])):
        print(f"   {reason[:41]:<42}{len(dropped[reason]):>9}")
    print("   " + "-" * 54)
    print(f"   {'REACHES THE FLEET SHEETS':<42}{len(survivors):>9}")

    # The exclusions file is local knowledge; say whether there is any, since
    # an empty one means only the built-in name patterns are doing the work.
    if not exclusions.ids and not exclusions.names:
        print(f"\n   No {R.VEHICLE_EXCLUSIONS_FILE} in effect -- built-in name")
        print(f"   patterns only: {', '.join(exclusions.patterns)}")

    # The part that still matters: collisions among the records that survive.
    by_vin, by_name = defaultdict(list), defaultdict(list)
    for vehicle in survivors:
        vin = normalize_vin(vehicle.get("vin"))
        if len(vin) == 17:
            by_vin[vin].append(vehicle)
        by_name[str(vehicle.get("name") or "").strip().lower()].append(vehicle)

    vin_dupes = {v: rows for v, rows in by_vin.items() if len(rows) > 1}
    name_dupes = {n: rows for n, rows in by_name.items() if len(rows) > 1}

    print(f"\n   Among those {len(survivors)}:")
    print(f"      same name twice:      {len(name_dupes)}")
    print(f"      same VIN, any name:   {len(vin_dupes)}")
    if vin_dupes:
        print("\n   One truck counted twice on a live sheet:")
        for vin, rows in sorted(vin_dupes.items()):
            names = ", ".join(sorted({str(r.get('name') or '?') for r in rows}))
            statuses = ", ".join(sorted({str(r.get('vehicle_status') or '?') for r in rows}))
            print(f"      {vin:<20}{names[:28]:<30}{statuses[:22]}")
        print("\n   These survive every filter, so each inflates the fleet count")
        print("   by one and splits its own history across two records. A name")
        print("   pattern cannot catch them -- the names are both legitimate.")

    out["roster"] = {
        "field_population": dict(counts),
        "returned": total,
        "reaches_sheets": len(survivors),
        "dropped": {k: sorted(v) for k, v in dropped.items()},
        "surviving_vin_duplicates": [
            {"vin": v,
             "names": sorted({str(r.get("name") or "?") for r in rows}),
             "statuses": sorted({str(r.get("vehicle_status") or "?") for r in rows})}
            for v, rows in sorted(vin_dupes.items())
        ],
    }


def normalize_vin(value):
    """
    A VIN reduced to what is comparable between the two systems.

    Case and stray punctuation differ between hand-entered and telematics-fed
    records; nothing else about a VIN is safe to alter, so this only upper-
    cases and drops anything that is not a letter or digit.
    """
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def report_vin_join(ts_vehicles, sam_vehicles, out):
    """
    The VIN join, measured against the name join it would replace.

    A VIN is the only identifier both systems hold independently -- Traumasoft
    from whoever typed it, Samsara from the gateway -- so it is the one key
    that cannot drift when somebody renames a unit. It also settles the
    duplicate-name records, which a name join cannot see past: two rows called
    A-101 are one truck if they carry one VIN and two trucks if they do not.

    Reported as a comparison rather than a number, because replacing a working
    join with a better-sounding one that quietly covers fewer trucks is the
    failure this is meant to prevent.
    """
    print("\n8. VIN JOIN")
    print("   " + "-" * 66)

    ts_total = len(ts_vehicles)
    ts_with = [v for v in ts_vehicles if normalize_vin(v.get("vin"))]
    sam_with = [v for v in sam_vehicles if normalize_vin(v.get("vin"))]
    print(f"   Traumasoft: {len(ts_with)} of {ts_total} record(s) carry a VIN "
          f"({pct(len(ts_with), ts_total).strip()})")
    print(f"   Samsara:    {len(sam_with)} of {len(sam_vehicles)} "
          f"({pct(len(sam_with), len(sam_vehicles)).strip()})")

    # A VIN is 17 characters. Anything else is a typo or a placeholder, and
    # joining on it would be joining on nothing.
    odd = [
        (str(v.get("name") or "?"), v.get("vin"))
        for v in ts_with if len(normalize_vin(v.get("vin"))) != 17
    ]
    if odd:
        print(f"\n   Traumasoft VINs that are not 17 characters ({len(odd)}):")
        for name, vin in odd[:20]:
            print(f"      {name[:26]:<28}{str(vin)[:22]:<24}"
                  f"{len(normalize_vin(vin))} chars")

    sam_by_vin = defaultdict(list)
    for vehicle in sam_with:
        sam_by_vin[normalize_vin(vehicle.get("vin"))].append(vehicle)
    ts_by_vin = defaultdict(list)
    for vehicle in ts_with:
        ts_by_vin[normalize_vin(vehicle.get("vin"))].append(vehicle)

    # The duplicate-name records are the reason for doing this at all, so say
    # plainly which ones one VIN resolves and which ones it does not.
    shared = [
        (vin, rows) for vin, rows in ts_by_vin.items() if len(rows) > 1
    ]
    if shared:
        print(f"\n   One VIN on several Traumasoft records ({len(shared)}) -- "
              f"these are one truck:")
        for vin, rows in shared:
            names = ", ".join(sorted({str(r.get('name') or '?') for r in rows}))
            statuses = ", ".join(sorted({str(r.get('vehicle_status') or '?') for r in rows}))
            print(f"      {vin[:19]:<21}{names[:26]:<28}{statuses[:24]}")

    matched = {v: rows for v, rows in ts_by_vin.items() if v in sam_by_vin}
    ts_only = sorted(set(ts_by_vin) - set(sam_by_vin))
    sam_only = sorted(set(sam_by_vin) - set(ts_by_vin))

    print(f"\n   VINs in both systems: {len(matched)}")
    print(f"   Traumasoft only:      {len(ts_only)}")
    print(f"   Samsara only:         {len(sam_only)}")

    joined_records = sum(len(rows) for rows in matched.values())
    print(f"   Traumasoft records reached by VIN: {joined_records} of {ts_total} "
          f"({pct(joined_records, ts_total).strip()})")

    if sam_only:
        print(f"\n   In Samsara with no Traumasoft VIN ({len(sam_only)}) -- these "
              f"vanish\n   if Samsara is the roster and the join is VIN:")
        for vin in sam_only[:25]:
            names = ", ".join(v.get("name", "?") for v in sam_by_vin[vin])
            print(f"      {vin[:19]:<21}{names[:36]}")

    # Which join covers more, and where they disagree. A unit reached by name
    # but not by VIN is a truck the switch would lose.
    overrides = SR.load_vehicle_overrides()
    ts_names = sorted({str(v.get("name") or "").strip() for v in ts_vehicles} - {""})
    by_name, _, _ = SR.match_vehicles(ts_names, sam_vehicles, overrides)
    name_reached = set(by_name)
    vin_reached = {
        str(r.get("name") or "").strip()
        for rows in matched.values() for r in rows
    }
    lost = sorted(name_reached - vin_reached)
    gained = sorted(vin_reached - name_reached)
    print(f"\n   Name join reached {len(name_reached)} unit(s); "
          f"VIN join reaches {len(vin_reached)}.")
    if gained:
        print(f"   VIN rescues ({len(gained)}): {', '.join(gained)[:180]}")
    if lost:
        print(f"   VIN loses ({len(lost)}): {', '.join(lost)[:180]}")
        print("   Those need a VIN in Traumasoft before the switch, or they")
        print("   drop off every sheet.")

    # externalIds is populated on every Samsara vehicle; if it carries a
    # Traumasoft id it beats both joins outright.
    keys = Counter()
    for vehicle in sam_vehicles:
        for key in (vehicle.get("externalIds") or {}):
            keys[key] += 1
    if keys:
        print(f"\n   Samsara externalIds keys: "
              f"{', '.join(f'{k} ({c})' for k, c in keys.most_common())}")

    out["vin"] = {
        "traumasoft_with_vin": len(ts_with),
        "traumasoft_total": ts_total,
        "samsara_with_vin": len(sam_with),
        "matched_vins": len(matched),
        "traumasoft_only": ts_only,
        "samsara_only": sam_only,
        "records_reached": joined_records,
        "malformed": [{"name": n, "vin": v} for n, v in odd],
        "duplicate_vin_records": [
            {"vin": v, "names": sorted({str(r.get("name") or "?") for r in rows})}
            for v, rows in shared
        ],
        "name_join_only": lost,
        "vin_join_only": gained,
        "external_id_keys": dict(keys),
    }


def report_vehicle_status(ts_vehicles, out):
    """
    The current status picture.

    Not what the worksheet asks for -- it wants days out of service across a
    past month, and this API carries no status history -- but it says how many
    vehicles are out right now, which is the number a current-state version of
    that report would be built on.
    """
    print("\n7. TRAUMASOFT VEHICLE STATUS (current state only)")
    print("   " + "-" * 66)

    statuses = Counter(
        str(v.get("vehicle_status") or "(none)") for v in ts_vehicles
    )
    for status, count in statuses.most_common():
        print(f"   {status[:40]:<42}{count:>5}")
    print("\n   This is a snapshot. vehicle_status carries no history and the")
    print("   API has no work-order endpoint, so days-out for a past month")
    print("   cannot be derived from it -- only accumulated going forward.")
    out["vehicle_status"] = dict(statuses)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--json", dest="json_path", help="Also write the findings as JSON.")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    try:
        api = TraumasoftAPI()
        ts_vehicles = api.list_vehicles()
    except TraumasoftAPIError as exc:
        log.error("Traumasoft API failed: %s", exc)
        return 2
    except RuntimeError as exc:
        log.error("%s", exc)
        return 2

    try:
        from samsara_api import SamsaraClient
        samsara = SamsaraClient(read_only=True)
        sam_vehicles = samsara.list_vehicles()
    except Exception as exc:  # noqa: BLE001 -- any failure means "no Samsara half"
        log.error("Samsara not reachable (%s: %s).", type(exc).__name__, exc)
        log.error("SAMSARA_API_TOKEN must be set in .env -- see .env.example.")
        return 2

    print("=" * 72)
    print("  Samsara tag probe -- can Samsara supply the cost center?")
    print(f"  Read-only. GET only, no writes.               build {build_id()}")
    print("=" * 72)

    out = {}
    report_fleet_shape(sam_vehicles, out)
    tag_counts = report_tag_vocabulary(sam_vehicles, out)
    report_attributes(sam_vehicles, out)
    matched = report_join(ts_vehicles, sam_vehicles, out)
    centres = fetch_cost_centers(api)
    report_tags_per_unit(ts_vehicles, matched, out, centres=centres)
    report_tags_vs_cost_centers(api, tag_counts or Counter(), out, centres=centres,
                                parents=out.get("tag_parents"))
    report_vehicle_status(ts_vehicles, out)
    report_vin_join(ts_vehicles, sam_vehicles, out)
    report_traumasoft_roster(ts_vehicles, out)

    print("\n" + "=" * 72)
    print("  Safe to paste. Operational values only, no patient data.")
    print("=" * 72 + "\n")

    if args.json_path:
        with open(args.json_path, "w", encoding="utf-8") as handle:
            json.dump(out, handle, indent=2, default=str)
        print(f"  Written to {args.json_path}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
