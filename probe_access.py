"""
Sweep every documented endpoint and report what this API key can actually reach.

Scope changes are made in Traumasoft's UI, and nothing in a response says which
scopes a key holds -- an endpoint outside the key's scope simply refuses. So the
only way to know what a change opened up is to ask each endpoint in turn and
record how it answers.

Also probes the ePCR surface. The spec names `ThirdParty/Data/Epcr/Huly` under
"not included -- private or non-partner integrations", so it is undocumented
rather than absent, and whether it reads at all is unknown. That question is
worth more than the rest of the sweep: OTP's arrival time and UHU's utilized
time both came from the ePCR originally, and both are degraded without it.

Strictly read-only. Every call is a GET, and the write-shaped rtype the spec
mentions for Huly is deliberately never sent.

Usage:
    python probe_access.py [--full]

    --full   also walk paths already known to work, rather than just
             confirming them with a single row
"""

import sys
import json
import logging
import argparse
from pathlib import Path
from datetime import date, datetime, timedelta

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("probe-access")

# Everything the spec documents, minus the /{id} reads (which need an id) and
# minus anything that is not a GET. Trip carries its own required rtype.
DOCUMENTED = [
    ("Data/Attachments", {"rtype": "GetTypes"}),
    ("Data/Billing/FeeSchedules", {}),
    ("Data/Billing/PayorCategories", {}),
    ("Data/Billing/Payors", {}),
    ("Data/Billing/Schedules", {}),
    ("Data/Cad/CallTypes", {}),
    ("Data/Cad/CancelReasons", {}),
    ("Data/Cad/Departments", {}),
    ("Data/Cad/Facilities", {}),
    ("Data/Cad/FacilityTypes", {}),
    ("Data/Cad/LevelOfService", {}),
    ("Data/Cad/Natures", {}),
    ("Data/Cad/Patients", {}),
    ("Data/Cad/Priorities", {}),
    ("Data/Cad/Subzones", {}),
    ("Data/Cad/Timestamps", {}),
    ("Data/Cad/Zones", {}),
    ("Data/Fleet/CustomStatus", {}),
    ("Data/Fleet/GpsGeofence", {}),
    ("Data/Fleet/Vehicles", {}),
    ("Data/Hr/EmployeeLevels", {}),
    ("Data/Organization", {}),
    ("Data/Schedule/EarningCodes", {}),
    ("Data/Schedule/PayTypes", {}),
    ("Data/Schedule/Shifts", {}),
    ("Data/User/Employees", {}),
    ("Data/User/Users", {}),
    ("Lists/Cad/CallTypes", {}),
    ("Lists/Cad/Companies", {}),
    ("Lists/Cad/CostCenters", {}),
    ("Lists/Cad/Facilities", {}),
    ("Lists/Cad/Mds", {}),
    ("Lists/Cad/Payors", {}),
    ("Lists/Cad/Posts", {}),
    ("Lists/Cad/Priorities", {}),
    ("Lists/Cad/Subzones", {}),
    ("Lists/Cad/Zones", {}),
    ("Lists/Fleet/Equipment", {}),
    ("Lists/Fleet/Vehicles", {}),
    ("Lists/Fleet/Vendors", {}),
    ("Lists/HumanResources/Certifications", {}),
    ("Lists/HumanResources/JobTitles", {}),
    ("Lists/Schedule/Radios", {}),
    ("Lists/Schedule/ShiftProfiles", {}),
    ("Lists/Schedule/Stations", {}),
    ("Lists/User/Employees", {}),
    ("Lists/User/LicenseLevels", {}),
    ("Lists/User/Supervisors", {}),
]

# Named in the scope UI but with no path in the spec. If a scope change made
# these real, they will answer.
UNDOCUMENTED = [
    ("Data/User/UserCertifications", {}),
    ("Data/User/UserCostCenters", {}),
    ("Lists/User/UserCertifications", {}),
    ("Lists/User/UserCostCenters", {}),
]

# The reason for the sweep. Bare paths and read-shaped rtypes only -- the
# write-shaped HulyUpdateTrip is never sent.
EPCR = [
    ("Data/Epcr/Huly", {}),
    ("Data/Epcr/Huly", {"rtype": "GetRuns"}),
    ("Data/Epcr/Huly", {"rtype": "GetTrips"}),
    ("Data/Epcr/Runs", {}),
    ("Data/Epcr", {}),
    ("Lists/Epcr/Huly", {}),
]


def describe(payload):
    """A one-line shape summary, plus the field names of the first row."""
    rows = TraumasoftAPI.extract_rows(payload)
    if isinstance(payload, dict):
        envelope = ", ".join(sorted(payload)[:6])
    elif isinstance(payload, list):
        envelope = "bare array"
    else:
        envelope = type(payload).__name__
    fields = sorted(rows[0])[:14] if rows and isinstance(rows[0], dict) else []
    return len(rows), envelope, fields


def probe(api, path, params, results, label):
    full = f"ThirdParty/{path}"
    try:
        payload = api.get(full, params=dict(params, rows=1) if params is not None else {"rows": 1})
    except TraumasoftAPIError as exc:
        results.append({
            "group": label, "path": path, "params": params,
            "ok": False, "status": exc.status_code, "detail": str(exc)[:200],
        })
        return False
    except Exception as exc:  # transport failure, not an API answer
        results.append({
            "group": label, "path": path, "params": params,
            "ok": False, "status": None, "detail": f"{type(exc).__name__}: {exc}"[:200],
        })
        return False

    count, envelope, fields = describe(payload)
    results.append({
        "group": label, "path": path, "params": params, "ok": True, "status": 200,
        "rows_on_page": count, "envelope": envelope, "fields": fields,
    })
    return True


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--full", action="store_true")
    args = parser.parse_args()

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("%s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Credentials rejected under every signing scheme.")
        return 1

    results = []

    log.info("Probing %s documented endpoints ...", len(DOCUMENTED))
    for path, params in DOCUMENTED:
        probe(api, path, params, results, "documented")

    log.info("Probing the Trip endpoint ...")
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    probe(api, "Data/Cad/Trip",
          {"rtype": "GetTrips", "trip_date": yesterday, "range_days": 1},
          results, "documented")

    log.info("Probing %s scope-UI names with no documented path ...", len(UNDOCUMENTED))
    for path, params in UNDOCUMENTED:
        probe(api, path, params, results, "undocumented")

    log.info("Probing the ePCR surface (%s candidates) ...", len(EPCR))
    for path, params in EPCR:
        probe(api, path, params, results, "epcr")

    reachable = [r for r in results if r["ok"]]
    refused = [r for r in results if not r["ok"]]

    print()
    print("=" * 78)
    print(f"ACCESS SWEEP   {len(reachable)} reachable, {len(refused)} refused")
    print("=" * 78)

    for group, title in (("documented", "DOCUMENTED ENDPOINTS"),
                         ("undocumented", "SCOPE-UI NAMES WITH NO DOCUMENTED PATH"),
                         ("epcr", "ePCR SURFACE")):
        rows = [r for r in results if r["group"] == group]
        if not rows:
            continue
        print(f"\n{title}")
        for r in rows:
            tag = f"?{r['params'].get('rtype')}" if r["params"].get("rtype") else ""
            name = f"{r['path']}{tag}"
            if r["ok"]:
                print(f"  OK    {name:<44} {r['rows_on_page']} row(s), envelope: {r['envelope']}")
                if r["fields"]:
                    print(f"        fields: {', '.join(r['fields'])}")
            else:
                print(f"  {str(r['status'] or 'ERR'):<5} {name:<44} {r['detail'][:70]}")

    epcr_ok = [r for r in results if r["group"] == "epcr" and r["ok"]]
    print()
    print("=" * 78)
    if epcr_ok:
        print("ePCR IS REACHABLE. Paths that answered:")
        for r in epcr_ok:
            tag = f"?{r['params'].get('rtype')}" if r["params"].get("rtype") else ""
            print(f"    {r['path']}{tag}  -- {r['rows_on_page']} row(s)")
            if r["fields"]:
                print(f"      fields: {', '.join(r['fields'])}")
        print("  Send me the JSON. If any of these carry per-run timestamps,")
        print("  OTP and UHU can both go back to clinician-recorded times.")
    else:
        print("ePCR is still not reachable with this key.")
        print("  Every candidate refused. That surface needs separate credentials")
        print("  from Traumasoft, not a scope change on this key.")
    print("=" * 78)

    out = Path("access_sweep.json")
    out.write_text(json.dumps({
        "generated": datetime.now().isoformat(timespec="seconds"),
        "base_url": api.base_url,
        "results": results,
    }, indent=2, default=str), encoding="utf-8")
    print(f"\nWritten to {out.resolve()} -- send me this file.\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
