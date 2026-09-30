"""
Build a vehicle -> cost center map from Samsara's vehicle tags.

WHY THIS EXISTS. Traumasoft puts no cost center on a vehicle, and none on a
shift profile either -- `Lists/Schedule/ShiftProfiles` returns id and name and
nothing else. So every report that needs "whose truck is this" has had to infer
it from behaviour: the shift profile a vehicle ran legs under, then the crew
who staffed that profile, then those employees' `cost_center_name`. That chain
answers "where did this truck work" and is being read as "where does this truck
belong", and the two come apart exactly when it matters -- a unit covering a
neighbouring station for a week, a truck that ran nothing at all.

Samsara tags, IF they carry the station, are a statement of ownership rather
than an inference from behaviour, and they cover a vehicle that never moved.
That is a better source. Whether these tags do carry the station is not
something this script assumes: run `probe_samsara_tags.py` first -- its
sections 2, 5 and 6 report the tag vocabulary, the tags per unit, and which
tags look like cost center names. This script turns that into a file once the
answer is yes.

A TAG IS CURRENT, NOT HISTORICAL. Samsara returns today's tags. Using them for
July assumes no vehicle changed station in between, which is a far safer
assumption for ownership than for behaviour but is still an assumption, and
the written file records the date it was built so a reader can judge it.

NOTHING IS GUESSED. A vehicle whose tags map to two different cost centers is
left unset rather than assigned to one of them, and reported. A tag matching no
cost center name is reported. The output is a plain JSON file, so anything this
cannot work out can be typed in by hand and will not be overwritten unless
--overwrite is passed.

STRICTLY READ-ONLY against both APIs. Writes one local file.

Usage:
    python probe_samsara_tags.py                 # does this even work?
    python build_vehicle_cost_centers.py         # then build the map
    python build_vehicle_cost_centers.py --compare-days 30
"""

import argparse
import json
import logging
import os
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

import traumasoft_reports as R
import vehicle_assignment_by_day as A
import vehicle_productivity as VP
import vehicle_service_days as VS

log = logging.getLogger(__name__)

OUTPUT_FILE = os.path.join("state", "vehicle_cost_centers.json")
TAG_MAP_FILE = os.path.join("state", "tag_cost_center_map.json")


# =============================
# TAGS
# =============================
def tag_names(vehicle):
    """Every tag name on a Samsara vehicle, however the payload spells it."""
    names = []
    for tag in vehicle.get("tags") or []:
        if isinstance(tag, dict):
            name = str(tag.get("name") or tag.get("id") or "").strip()
        else:
            name = str(tag).strip()
        if name:
            names.append(name)
    return names


# Attribute names worth reading for a station. Some fleets put the station in
# an attribute rather than a tag, and only looking at tags would report an
# empty answer on a tenant that had the data all along.
#
# An ALLOWLIST rather than every attribute, because this tenant's attributes
# are `Asset Status` (Active / Out of Service), `Unit Type` (Secure Car /
# Wheelchair / BLS) and `Vehicle Status` (Retired). Feeding those values into
# a cost-center matcher is asking for the day somebody names a station
# something that collides with one of them.
STATION_ATTRIBUTES = ("station", "location", "base", "cost center",
                      "costcenter", "garage", "domicile", "market")


def attribute_values(vehicle, attribute_names=STATION_ATTRIBUTES):
    """Values from the attributes that plausibly name a station."""
    values = []
    for attr in vehicle.get("attributes") or []:
        if not isinstance(attr, dict):
            continue
        name = str(attr.get("name") or "").strip().lower()
        if not any(marker in name for marker in attribute_names):
            continue
        for value in attr.get("stringValues") or []:
            text = str(value).strip()
            if text:
                values.append(text)
    return values


def load_tag_map(path):
    """
    Hand-written tag -> cost center, for tags a name match cannot reach.

    Wins over the name match, being a decision rather than an observation --
    the same precedence the cost-center overrides use.
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, ValueError) as exc:
        log.warning("Could not read %s (%s).", path, exc)
        return {}
    # Either {"tags": {...}} or a bare {tag: cost center} dict. A bare file is
    # what somebody writes by hand, so it has to work; values that are not
    # strings are skipped rather than stringified, which is what keeps a
    # "note" list in a wrapped file from becoming a cost center.
    mapping = payload
    if isinstance(payload, dict) and isinstance(payload.get("tags"), dict):
        mapping = payload["tags"]
    if not isinstance(mapping, dict):
        return {}
    return {
        str(tag).strip().lower(): centre.strip()
        for tag, centre in mapping.items()
        if str(tag).strip() and isinstance(centre, str) and centre.strip()
    }


def resolve_label(label, centres, tag_map):
    """
    (cost center, how) for one tag or attribute value, or (None, reason).

    Exact name first, then the same legal-entity-wrapper normalisation the
    assignment report uses, so 'Boardman' and 'Lynx EMS LLC dba Lynx Boardman'
    are one station here too. No substring pass: a tag is a short label and
    substring matching short labels is how 'Columbus' claims two states.
    """
    key = str(label or "").strip()
    if not key:
        return None, "empty"
    if key.lower() in tag_map:
        return tag_map[key.lower()], "tag map"
    for centre in centres:
        if centre.strip().lower() == key.lower():
            return centre, "exact name"
    normalized = A.normalize_centre(key)
    if normalized:
        for centre in centres:
            if A.normalize_centre(centre) == normalized:
                return centre, "name match ignoring the legal entity wrapper"
    return None, "matches no cost center"


# =============================
# BUILD
# =============================
def resolve_vehicle(labels, centres, tag_map):
    """
    (cost center, how, note) for one vehicle's labels.

    A vehicle whose labels name two different cost centers is left unset. It
    is genuinely ambiguous, and picking one would produce a station roster
    that looks authoritative and is partly invented.
    """
    hits = {}
    for label in labels:
        centre, how = resolve_label(label, centres, tag_map)
        if centre:
            hits.setdefault(centre, (label, how))
    if not hits:
        return None, None, ("no label matched a cost center" if labels
                            else "no tags or attributes")
    if len(hits) > 1:
        return None, None, "labels name " + " and ".join(sorted(hits))
    centre, (label, how) = next(iter(hits.items()))
    return centre, f"{how}: {label}", None


def build_map(ts_vehicles, sam_vehicles, centres, tag_map):
    """
    normalized vehicle name -> {cost_center, source, ...}, plus what it could not do.

    Joined VIN first then name, reusing vehicle_productivity.build_samsara_index
    so this cannot disagree with the productivity report about which Samsara
    vehicle is which truck.
    """
    index, ambiguous_join = VP.build_samsara_index(ts_vehicles, sam_vehicles)
    sam_by_id = {str(v.get("id")): v for v in sam_vehicles}

    resolved, unresolved = {}, []
    for vehicle in ts_vehicles:
        ts_id = str(vehicle.get("id"))
        display = str(vehicle.get("name") or "").strip()
        key = VS.normalize_name(display)
        if not key:
            continue

        sam_id, how_joined = index.get(ts_id, (None, "unmatched"))
        if sam_id is None:
            unresolved.append((display, "no Samsara match", how_joined))
            continue
        sam = sam_by_id.get(sam_id) or {}
        labels = tag_names(sam) + attribute_values(sam)
        centre, source, note = resolve_vehicle(labels, centres, tag_map)
        if centre is None:
            unresolved.append((display, note, how_joined))
            continue
        resolved[key] = {
            "vehicle_name": display,
            "cost_center": centre,
            "source": source,
            "samsara_join": how_joined,
            "samsara_tags": labels,
        }
    return resolved, unresolved, ambiguous_join


def compare_with_crew_map(resolved, legs, cost_center_map):
    """
    Where the tag answer and the behaviour-derived answer disagree.

    The whole reason for doing this, and the only thing that says whether the
    switch changes any number. Agreement is reassurance; a disagreement is a
    truck that has been filed under the wrong station in every report built so
    far, or a tag nobody has kept up to date. Both are worth seeing by name.
    """
    from_legs, _ = VS.cost_centers_from_legs(legs, cost_center_map)
    agree, differ, tag_only, legs_only = [], [], [], []
    for key, entry in resolved.items():
        behaviour = from_legs.get(key)
        if behaviour is None:
            tag_only.append(entry["vehicle_name"])
        elif A.normalize_centre(behaviour) == A.normalize_centre(entry["cost_center"]):
            agree.append(entry["vehicle_name"])
        else:
            differ.append((entry["vehicle_name"], entry["cost_center"], behaviour))
    for key, centre in from_legs.items():
        if key not in resolved:
            legs_only.append(key)
    return agree, differ, tag_only, sorted(legs_only)


# =============================
# OUTPUT
# =============================
def write_map(path, resolved, built_from):
    payload = {
        "note": [
            "vehicle name -> cost center, built from Samsara vehicle tags.",
            "",
            "Traumasoft puts no cost center on a vehicle and none on a shift",
            "profile, so without this file every report infers a truck's",
            "station from the shift profile it ran legs under -- which answers",
            "where it worked, not where it belongs.",
            "",
            "THESE TAGS WERE CURRENT WHEN THIS WAS BUILT, not historical. Using",
            "it for a past month assumes no vehicle changed station since.",
            "",
            "Hand edits are safe: rebuild refuses to overwrite without",
            "--overwrite. Vehicles this could not resolve are absent rather",
            "than guessed, and callers fall back to their own sources.",
        ],
        "built": datetime.now().isoformat(timespec="seconds"),
        "built_from": built_from,
        "vehicles": {
            key: {
                "vehicle_name": entry["vehicle_name"],
                "cost_center": entry["cost_center"],
                "source": entry["source"],
            }
            for key, entry in sorted(resolved.items())
        },
    }
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=False)
    return payload


def load_vehicle_cost_centers(path):
    """
    normalized vehicle name -> cost center, for the reports that consume this.

    Missing is not an error: it is the normal state before this has been built,
    and the caller falls back to inferring from behaviour as before.
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, ValueError) as exc:
        log.warning("Could not read %s (%s).", path, exc)
        return {}
    vehicles = payload.get("vehicles") if isinstance(payload, dict) else payload
    resolved = {}
    for key, entry in (vehicles or {}).items():
        centre = entry.get("cost_center") if isinstance(entry, dict) else entry
        if centre:
            resolved[VS.normalize_name(key)] = str(centre).strip()
    return resolved


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("--out", default=OUTPUT_FILE)
    parser.add_argument("--tag-map", default=TAG_MAP_FILE,
                        help="Hand-written tag -> cost center, for tags a name "
                             "match cannot reach.")
    parser.add_argument("--compare-days", type=int, default=30,
                        help="Days of recent trips to compare the tag answer "
                             "against the behaviour-derived one (0 to skip).")
    parser.add_argument("--cost-center-map", default="state/shift_cost_center_map.json")
    parser.add_argument("--overwrite", action="store_true",
                        help="Replace an existing file, discarding hand edits.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Report what it would write, write nothing.")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    if os.path.exists(args.out) and not args.overwrite and not args.dry_run:
        log.error("%s already exists. It may carry hand edits, so it is not "
                  "replaced without --overwrite. Use --dry-run to see what "
                  "would change.", args.out)
        return 2

    print("=" * 72)
    print("  Vehicle -> cost center, from Samsara tags")
    print("  Read-only against both APIs. Writes one local file.")
    print("=" * 72)

    try:
        from traumasoft_api import TraumasoftAPI
        api = TraumasoftAPI()
        ts_vehicles = api.list_vehicles()
        centres = sorted({
            str(e.get("cost_center_name")).strip()
            for e in api.list_employees()
            if e.get("cost_center_name")
        })
    except Exception as exc:  # noqa: BLE001 -- reporting the failure IS the result
        log.error("Traumasoft unavailable (%s: %s).", type(exc).__name__, exc)
        return 2

    try:
        from samsara_api import SamsaraClient
        sam_vehicles = SamsaraClient(read_only=True).list_vehicles()
    except Exception as exc:  # noqa: BLE001
        log.error("Samsara unavailable (%s: %s). Nothing to build from.",
                  type(exc).__name__, exc)
        return 2

    tag_map = load_tag_map(args.tag_map)
    resolved, unresolved, ambiguous_join = build_map(
        ts_vehicles, sam_vehicles, centres, tag_map)

    print(f"\n1. WHAT THE TAGS COULD ANSWER")
    print("   " + "-" * 66)
    print(f"   Traumasoft vehicles:     {len(ts_vehicles)}")
    print(f"   Samsara vehicles:        {len(sam_vehicles)}")
    print(f"   Cost center names known: {len(centres)}")
    print(f"   Resolved to a station:   {len(resolved)}")
    print(f"   Not resolved:            {len(unresolved)}")
    if resolved:
        print(f"\n   {'cost center':<34}{'vehicles':>9}")
        print("   " + "-" * 43)
        for centre, count in Counter(
            e["cost_center"] for e in resolved.values()
        ).most_common():
            print(f"   {centre[:33]:<34}{count:>9}")

    print(f"\n2. WHAT IT COULD NOT")
    print("   " + "-" * 66)
    if not unresolved:
        print("   Every vehicle resolved.")
    for reason, count in Counter(note for _, note, _ in unresolved).most_common():
        print(f"      {count:>4}  {reason}")
    for display, note, _ in unresolved[:25]:
        print(f"      {display[:24]:<26}{note}")
    if len(unresolved) > 25:
        print(f"      ... and {len(unresolved) - 25} more")
    if ambiguous_join:
        print(f"\n   Ambiguous Samsara joins, left unmatched ({len(ambiguous_join)}):")
        for name, why in ambiguous_join[:10]:
            print(f"      {str(name)[:24]:<26}{why}")

    if args.compare_days > 0:
        print(f"\n3. AGAINST THE BEHAVIOUR-DERIVED ANSWER "
              f"(last {args.compare_days} days)")
        print("   " + "-" * 66)
        try:
            end = date.today() - timedelta(days=1)
            legs = VS.fetch_legs(api, end - timedelta(days=args.compare_days - 1), end)
            cost_center_map = R.CostCenterMap(path=args.cost_center_map)
            agree, differ, tag_only, legs_only = compare_with_crew_map(
                resolved, legs, cost_center_map)
            print(f"   Agree:                        {len(agree)}")
            print(f"   DISAGREE:                     {len(differ)}")
            print(f"   Tag has one, behaviour none:  {len(tag_only)}")
            print(f"   Behaviour has one, tag none:  {len(legs_only)}")
            if differ:
                print(f"\n   {'vehicle':<16}{'tag says':<24}{'its legs say'}")
                print("   " + "-" * 62)
                for name, tagged, behaviour in sorted(differ)[:25]:
                    print(f"   {str(name)[:15]:<16}{str(tagged)[:23]:<24}{behaviour}")
                print("\n   Each of these is either a truck filed under the wrong")
                print("   station in every report so far, or a tag nobody kept up")
                print("   to date. Both are worth settling before this file is used.")
            if tag_only:
                print(f"\n   {len(tag_only)} vehicle(s) the tags place and recent legs")
                print("   cannot -- they ran nothing. These are the whole point:")
                print("      " + ", ".join(tag_only[:20])[:220])
        except Exception as exc:  # noqa: BLE001
            log.warning("Comparison skipped (%s: %s).", type(exc).__name__, exc)

    print("\n" + "=" * 72)
    if args.dry_run:
        print(f"  Dry run. {len(resolved)} vehicle(s) would be written to {args.out}.")
    else:
        write_map(args.out, resolved, f"Samsara tags, {len(sam_vehicles)} vehicles")
        print(f"  Written to {args.out} -- {len(resolved)} vehicle(s).")
        print("  Hand-edit anything it could not resolve; rebuilding refuses")
        print("  to overwrite without --overwrite.")
    print("=" * 72 + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
