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
          f"({pct(len(matched), len(ts_names)).strip()} joined)\n")

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
        "traumasoft_units": len(ts_names),
        "samsara_vehicles": len(sam_vehicles),
        "matched": {k: v.get("name") for k, v in matched.items()},
        "unmatched": unmatched,
        "ambiguous": {k: [v.get("name") for v in hits] for k, hits in ambiguous.items()},
    }
    return matched


def report_tags_per_unit(ts_vehicles, matched, out):
    """
    The table the whole probe exists for: each unit, its status, its tags.

    Ordered by Traumasoft name so a naming convention -- if the tags follow
    one -- is visible down the column rather than having to be inferred.
    """
    print("\n5. TAGS PER TRAUMASOFT UNIT")
    print("   " + "-" * 66)
    print(f"   {'Traumasoft unit':<22}{'status':<20}{'Samsara tags'}")
    print("   " + "-" * 72)

    rows = []
    for vehicle in sorted(ts_vehicles, key=lambda v: str(v.get("name") or "")):
        name = str(vehicle.get("name") or "").strip()
        if not name:
            continue
        status = str(vehicle.get("vehicle_status") or "-")
        sam = matched.get(name)
        if sam is None:
            tags_text = "-- no Samsara match --"
            tags = None
        else:
            tags = [t for t, _ in tag_entries(sam)]
            tags_text = ", ".join(tags) if tags else "(none)"
        print(f"   {name[:21]:<22}{status[:19]:<20}{tags_text[:34]}")
        rows.append({
            "name": name,
            "vehicle_status": vehicle.get("vehicle_status"),
            "samsara_name": sam.get("name") if sam else None,
            "tags": tags,
        })
    out["units"] = rows


def report_tags_vs_cost_centers(api, tag_counts, out):
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

    print(f"\n   {'cost center':<34}{'candidate tag(s)'}")
    print("   " + "-" * 66)
    pairs = {}
    for centre in sorted(centres):
        key = normalize(centre)
        hits = [
            tag for tag in tag_counts
            if key and normalize(tag) and (
                key in normalize(tag) or normalize(tag) in key
            )
        ]
        pairs[centre] = hits
        print(f"   {centre[:33]:<34}{', '.join(hits)[:32] if hits else '-- none --'}")

    unclaimed = [t for t in tag_counts if not any(t in h for h in pairs.values())]
    if unclaimed:
        print(f"\n   Tags matching no cost center ({len(unclaimed)}): "
              f"{', '.join(sorted(unclaimed))[:200]}")

    out["cost_centers"] = sorted(centres)
    out["tag_cost_center_candidates"] = pairs


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
    report_tags_per_unit(ts_vehicles, matched, out)
    report_tags_vs_cost_centers(api, tag_counts or Counter(), out)
    report_vehicle_status(ts_vehicles, out)

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
