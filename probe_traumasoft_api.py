"""
Traumasoft ThirdParty API probe.

Answers the questions the OpenAPI spec leaves open before any report code is
rewritten against the API. Read-only: it issues GETs only.

What it resolves:

  1. Do the HMAC credentials work at all?
  2. /Schedule/Shifts documents no date filter and no pagination. How many rows
     come back, what date span do they cover, and do any undocumented filter
     parameters actually work?
  3. Does a Shift row carry the fields the current SQL filters on
     (published, timeoff_type, schedule_type, cost center)?
  4. What do trip timestamp keys look like, and which one is the real
     "arrived" time now that ePCR is out of scope?
  5. Is the vehicle field allowlist really closed (no cost center, no
     status_reason, no work orders)?
  6. Can cost center be resolved for a trip at all, and by which path?

Usage:

    export TS_API_BASE_URL=https://your-tenant.traumasoft.com
    export TS_API_KEY=...
    export TS_API_SECRET=...
    python probe_traumasoft_api.py [YYYY-MM-DD] [--out DIR]

Writes a findings summary plus raw JSON samples to the output directory
(default: api_probe/).
"""

import os
import sys
import json
import logging
from pathlib import Path
from datetime import datetime, timedelta, date

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("probe")

# Candidate filter parameters that are NOT in the spec. If any of these change
# the Shifts result, the staffing report and the UHU denominator become
# server-filterable instead of full-pull-and-filter-locally.
CANDIDATE_SHIFT_FILTERS = [
    {"start_date": "{date}"},
    {"date": "{date}"},
    {"shift_date": "{date}"},
    {"from": "{date}", "to": "{date}"},
    {"start_time": "{date}"},
    {"begin_date": "{date}", "end_date": "{date}"},
    {"page": 1, "rows": 10},
]

# Fields the current SQL needs that the documented Shift schema omits.
SHIFT_FIELDS_WANTED = [
    "published",
    "timeoff_type",
    "schedule_type",
    "cost_center",
    "cost_center_name",
    "cost_center_id",
    "group",
    "group_id",
    "unit_id",
    "schedule_id",
    "user_id",
    "shift_name",
    "vehicle_name",
    "start_time",
    "end_time",
    "license_level",
    "division",
    "district",
]

# Vehicle columns the Daily Vehicle Overview needs. The spec's allowlist is
# id, name, vehicle_status, vin, odometer, disabled, deleted (+ enrichment),
# so these are expected to be absent -- this confirms it against the live API.
VEHICLE_FIELDS_WANTED = [
    "status_reason",
    "cost_center_id",
    "cost_center_name",
    "division_id",
    "district_id",
    "group_id",
    "oos_since",
    "odometer_type",
]


def _sample(rows, n=3):
    return rows[:n]


def _keys_union(rows):
    keys = set()
    for row in rows:
        if isinstance(row, dict):
            keys.update(row.keys())
    return sorted(keys)


def _write(out_dir, name, payload):
    path = Path(out_dir) / name
    path.write_text(json.dumps(payload, indent=2, default=str))
    return path


def probe_auth(api, findings):
    log.info("[1/6] Detecting the HMAC scheme against /Data/Organization ...")

    # The key-creation screen may issue only a key and no separate secret, so
    # which (formula, secret) pair the tenant accepts has to be established
    # rather than assumed.
    detected = api.detect_auth_mode()
    if detected is None:
        findings["auth"] = {
            "ok": False,
            "error": "no HMAC scheme accepted",
            "schemes_tried": [
                {"mode": m, "secret": label} for m, _, label in api.candidate_auth_schemes()
            ],
        }
        log.error("    Every signing scheme was rejected.")
        log.error("    Check, in order:")
        log.error("      - whether the key screen offers a secret you have not captured")
        log.error("      - this machine's clock (the timestamp is valid for 300 seconds)")
        log.error("      - that the key has read scopes enabled for Organization")
        return None

    mode, secret_label = detected
    log.info("    Accepted: formula=%s, secret=%s", mode, secret_label)

    try:
        org = api.get_organization()
    except TraumasoftAPIError as exc:
        findings["auth"] = {
            "ok": False,
            "error": str(exc),
            "status": exc.status_code,
            "auth_mode": mode,
        }
        log.error("    Organization fetch failed after auth succeeded: %s", exc)
        return None

    findings["auth"] = {
        "ok": True,
        "auth_mode": mode,
        "secret_source": secret_label,
        "organization_rows": len(org),
    }
    log.info("    OK - %s organization rows", len(org))
    return org


def probe_shifts(api, run_date, out_dir, findings):
    log.info("[2/6] Pulling /Data/Schedule/Shifts (no documented date filter) ...")
    result = {}
    try:
        shifts = api.list_shifts()
    except TraumasoftAPIError as exc:
        findings["shifts"] = {"ok": False, "error": str(exc), "status": exc.status_code}
        log.error("    Shifts failed: %s", exc)
        return None

    result["ok"] = True
    result["row_count"] = len(shifts)
    result["observed_keys"] = _keys_union(shifts)
    result["missing_fields_needed_by_current_sql"] = [
        f for f in SHIFT_FIELDS_WANTED if f not in result["observed_keys"]
    ]

    starts = sorted(
        str(s.get("start_time")) for s in shifts if isinstance(s, dict) and s.get("start_time")
    )
    result["earliest_start_time"] = starts[0] if starts else None
    result["latest_start_time"] = starts[-1] if starts else None
    result["distinct_start_dates"] = len({s[:10] for s in starts})
    result["covers_requested_date"] = any(s[:10] == run_date for s in starts)

    log.info(
        "    %s rows spanning %s distinct dates (%s .. %s)",
        result["row_count"],
        result["distinct_start_dates"],
        result["earliest_start_time"],
        result["latest_start_time"],
    )
    if result["missing_fields_needed_by_current_sql"]:
        log.warning(
            "    Shift rows lack: %s",
            ", ".join(result["missing_fields_needed_by_current_sql"]),
        )

    # Try undocumented filters: a filter "works" if it changes the row count.
    log.info("    Testing %s undocumented filter shapes ...", len(CANDIDATE_SHIFT_FILTERS))
    filter_results = []
    for candidate in CANDIDATE_SHIFT_FILTERS:
        params = {
            k: (v.format(date=run_date) if isinstance(v, str) else v)
            for k, v in candidate.items()
        }
        try:
            filtered = api.list_shifts(**params)
            changed = len(filtered) != len(shifts)
            filter_results.append(
                {
                    "params": params,
                    "row_count": len(filtered),
                    "changed_result": changed,
                    "likely_supported": changed,
                }
            )
            log.info(
                "      %s -> %s rows%s",
                params,
                len(filtered),
                "  <-- FILTER APPLIED" if changed else "",
            )
        except TraumasoftAPIError as exc:
            filter_results.append(
                {"params": params, "error": str(exc), "status": exc.status_code}
            )
            log.info("      %s -> %s", params, exc.status_code)

    result["undocumented_filters"] = filter_results
    result["server_side_date_filtering"] = any(
        f.get("likely_supported") for f in filter_results
    )

    _write(out_dir, "shifts_sample.json", _sample(shifts, 5))
    findings["shifts"] = result
    return shifts


def probe_trips(api, run_date, out_dir, findings):
    log.info("[3/6] Pulling /Data/Cad/Trip?rtype=GetTrips for %s ...", run_date)
    result = {}
    try:
        trips = api.get_trips(run_date, range_days=1)
    except TraumasoftAPIError as exc:
        findings["trips"] = {"ok": False, "error": str(exc), "status": exc.status_code}
        log.error("    GetTrips failed: %s", exc)
        return None

    result["ok"] = True
    result["row_count"] = len(trips)
    result["observed_keys"] = _keys_union(trips)

    # The timestamps array is "status name -> ISO time" maps. Collect every
    # status name seen so the OTP "arrived" timestamp can be chosen correctly.
    ts_names = set()
    with_ts = 0
    for trip in trips:
        stamps = trip.get("timestamps") if isinstance(trip, dict) else None
        if not stamps:
            continue
        with_ts += 1
        for entry in stamps:
            if isinstance(entry, dict):
                ts_names.update(entry.keys())

    result["trips_with_timestamps"] = with_ts
    result["observed_timestamp_names"] = sorted(ts_names)
    result["has_pickup_time"] = sum(1 for t in trips if t.get("pickup_time"))
    result["has_shift_name"] = sum(1 for t in trips if t.get("shift_name"))
    result["has_vehicle_id"] = sum(1 for t in trips if t.get("vehicle_id"))
    result["has_late_reasons"] = sum(1 for t in trips if t.get("late_reasons"))
    result["distinct_call_types"] = sorted(
        {str(t.get("call_type")) for t in trips if t.get("call_type")}
    )
    result["any_cost_center_field"] = [
        k for k in result["observed_keys"] if "cost" in k.lower()
    ]

    log.info("    %s trip legs; %s carry timestamps", result["row_count"], with_ts)
    log.info("    timestamp names seen: %s", ", ".join(result["observed_timestamp_names"]) or "none")
    if not result["any_cost_center_field"]:
        log.warning("    No cost-center field on trip legs (expected) - needs indirect resolution")

    _write(out_dir, "trips_sample.json", _sample(trips, 5))
    findings["trips"] = result
    return trips


def probe_vehicles(api, out_dir, findings):
    log.info("[4/6] Pulling /Data/Fleet/Vehicles and confirming the field allowlist ...")
    result = {}
    try:
        vehicles = api.list_vehicles()
    except TraumasoftAPIError as exc:
        findings["vehicles"] = {"ok": False, "error": str(exc), "status": exc.status_code}
        log.error("    Vehicles failed: %s", exc)
        return None

    result["ok"] = True
    result["row_count"] = len(vehicles)
    result["observed_keys"] = _keys_union(vehicles)
    result["distinct_vehicle_status"] = sorted(
        {str(v.get("vehicle_status")) for v in vehicles if v.get("vehicle_status")}
    )
    result["with_shift_name"] = sum(1 for v in vehicles if v.get("shift_name"))

    # Ask for the fields the vehicle report needs but the allowlist omits.
    # Unknown names are silently ignored, so anything that comes back is real.
    try:
        probed = api.list_vehicles(fields=["id", "name"] + VEHICLE_FIELDS_WANTED)
        returned = _keys_union(probed)
        result["unlisted_fields_that_returned"] = [
            f for f in VEHICLE_FIELDS_WANTED if f in returned
        ]
    except TraumasoftAPIError as exc:
        result["field_probe_error"] = str(exc)
        result["unlisted_fields_that_returned"] = []

    log.info("    %s vehicles; statuses: %s", result["row_count"], result["distinct_vehicle_status"])
    if result.get("unlisted_fields_that_returned"):
        log.info("    Bonus fields available: %s", result["unlisted_fields_that_returned"])
    else:
        log.warning(
            "    Allowlist confirmed closed - no status_reason / cost center / OOS history"
        )

    _write(out_dir, "vehicles_sample.json", _sample(vehicles, 5))
    findings["vehicles"] = result
    return vehicles


def probe_people(api, out_dir, findings):
    log.info("[5/6] Pulling /Data/User/Employees for cost-center coverage ...")
    result = {}
    try:
        employees = api.list_employees()
    except TraumasoftAPIError as exc:
        findings["employees"] = {"ok": False, "error": str(exc), "status": exc.status_code}
        log.error("    Employees failed: %s", exc)
        return None

    result["ok"] = True
    result["row_count"] = len(employees)
    result["observed_keys"] = _keys_union(employees)
    result["with_cost_center_name"] = sum(1 for e in employees if e.get("cost_center_name"))
    result["distinct_cost_centers"] = sorted(
        {str(e.get("cost_center_name")) for e in employees if e.get("cost_center_name")}
    )
    result["with_license_level"] = sum(1 for e in employees if e.get("license_level"))

    log.info(
        "    %s employees; %s carry cost_center_name (%s distinct)",
        result["row_count"],
        result["with_cost_center_name"],
        len(result["distinct_cost_centers"]),
    )

    _write(out_dir, "employees_sample.json", _sample(employees, 5))
    findings["employees"] = result
    return employees


def probe_cost_center_resolution(api, shifts, trips, employees, findings):
    """
    Cost center is absent from both trips and vehicles. Test whether it can be
    reached indirectly: trip.shift_name -> shift.user_id -> employee.cost_center_name.
    """
    log.info("[6/6] Testing indirect cost-center resolution for trips ...")
    result = {}

    if not (shifts and trips and employees):
        result["ok"] = False
        result["reason"] = "one or more prerequisite pulls failed"
        findings["cost_center_resolution"] = result
        return

    emp_cc = {
        e.get("user_id"): e.get("cost_center_name")
        for e in employees
        if e.get("user_id") and e.get("cost_center_name")
    }
    shift_cc = {}
    for shift in shifts:
        name = shift.get("shift_name")
        cc = emp_cc.get(shift.get("user_id"))
        if name and cc:
            shift_cc.setdefault(name, set()).add(cc)

    ambiguous = {k: sorted(v) for k, v in shift_cc.items() if len(v) > 1}
    resolvable = sum(1 for t in trips if t.get("shift_name") in shift_cc)

    result["ok"] = True
    result["shift_profiles_mapped"] = len(shift_cc)
    result["ambiguous_shift_profiles"] = len(ambiguous)
    result["ambiguous_examples"] = dict(list(ambiguous.items())[:5])
    result["trips_resolvable_via_shift_name"] = resolvable
    result["trips_total"] = len(trips)
    result["resolution_rate"] = round(resolvable / len(trips), 3) if trips else 0.0

    log.info(
        "    %s/%s trips resolve to a cost center via shift_name (%s ambiguous profiles)",
        resolvable,
        len(trips),
        len(ambiguous),
    )
    findings["cost_center_resolution"] = result


def write_summary(out_dir, findings, run_date):
    lines = [
        "# Traumasoft ThirdParty API probe",
        "",
        f"Probe date (trip/shift target): **{run_date}**",
        "",
    ]

    auth = findings.get("auth", {})
    lines += ["## Credentials", ""]
    lines.append("- OK" if auth.get("ok") else f"- FAILED: {auth.get('error')}")
    lines.append("")

    shifts = findings.get("shifts", {})
    lines += ["## Schedule/Shifts", ""]
    if shifts.get("ok"):
        lines += [
            f"- Rows returned: **{shifts['row_count']}**",
            f"- Distinct start dates in one pull: **{shifts['distinct_start_dates']}**",
            f"- Span: `{shifts['earliest_start_time']}` .. `{shifts['latest_start_time']}`",
            f"- Includes the requested date: **{shifts['covers_requested_date']}**",
            f"- Server-side date filtering available: **{shifts['server_side_date_filtering']}**",
            f"- Fields the current SQL needs but Shifts omits: "
            f"`{', '.join(shifts['missing_fields_needed_by_current_sql']) or 'none'}`",
            "",
        ]
    else:
        lines += [f"- FAILED: {shifts.get('error')}", ""]

    trips = findings.get("trips", {})
    lines += ["## Cad/Trip (GetTrips)", ""]
    if trips.get("ok"):
        lines += [
            f"- Trip legs for the day: **{trips['row_count']}**",
            f"- Legs carrying timestamps: **{trips['trips_with_timestamps']}**",
            f"- Timestamp names observed: `{', '.join(trips['observed_timestamp_names']) or 'none'}`",
            f"- Legs with `late_reasons`: **{trips['has_late_reasons']}**",
            f"- Cost-center fields present: `{', '.join(trips['any_cost_center_field']) or 'none'}`",
            "",
        ]
    else:
        lines += [f"- FAILED: {trips.get('error')}", ""]

    vehicles = findings.get("vehicles", {})
    lines += ["## Fleet/Vehicles", ""]
    if vehicles.get("ok"):
        lines += [
            f"- Vehicles: **{vehicles['row_count']}**",
            f"- Distinct statuses: `{', '.join(vehicles['distinct_vehicle_status'])}`",
            f"- Unlisted fields that actually returned: "
            f"`{', '.join(vehicles.get('unlisted_fields_that_returned') or []) or 'none'}`",
            "",
        ]
    else:
        lines += [f"- FAILED: {vehicles.get('error')}", ""]

    people = findings.get("employees", {})
    lines += ["## User/Employees", ""]
    if people.get("ok"):
        lines += [
            f"- Employees: **{people['row_count']}**",
            f"- With `cost_center_name`: **{people['with_cost_center_name']}** "
            f"({len(people['distinct_cost_centers'])} distinct)",
            "",
        ]
    else:
        lines += [f"- FAILED: {people.get('error')}", ""]

    ccr = findings.get("cost_center_resolution", {})
    lines += ["## Cost-center resolution for trips", ""]
    if ccr.get("ok"):
        lines += [
            f"- Trips resolvable via `shift_name`: "
            f"**{ccr['trips_resolvable_via_shift_name']}/{ccr['trips_total']}** "
            f"({ccr['resolution_rate']:.1%})",
            f"- Shift profiles mapped: **{ccr['shift_profiles_mapped']}**, "
            f"ambiguous: **{ccr['ambiguous_shift_profiles']}**",
            "",
        ]
    else:
        lines += [f"- Not evaluated: {ccr.get('reason', 'n/a')}", ""]

    path = Path(out_dir) / "PROBE_FINDINGS.md"
    path.write_text("\n".join(lines))
    return path


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    run_date = args[0] if args else (date.today() - timedelta(days=1)).isoformat()
    try:
        datetime.strptime(run_date, "%Y-%m-%d")
    except ValueError:
        log.error("Date must be YYYY-MM-DD, got %r", run_date)
        return 2

    out_dir = "api_probe"
    if "--out" in sys.argv:
        out_dir = sys.argv[sys.argv.index("--out") + 1]
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("%s", exc)
        log.error("Set TS_API_BASE_URL and TS_API_KEY (and TS_API_SECRET if you have one)")
        log.error("in a .env file next to this script, or in the environment.")
        return 2

    findings = {"probe_date": run_date, "base_url": api.base_url}

    probe_auth(api, findings)
    if not findings["auth"].get("ok"):
        _write(out_dir, "findings.json", findings)
        return 1

    shifts = probe_shifts(api, run_date, out_dir, findings)
    trips = probe_trips(api, run_date, out_dir, findings)
    vehicles = probe_vehicles(api, out_dir, findings)
    employees = probe_people(api, out_dir, findings)
    probe_cost_center_resolution(api, shifts, trips, employees, findings)

    _write(out_dir, "findings.json", findings)
    summary = write_summary(out_dir, findings, run_date)
    log.info("Done. Findings written to %s", summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())
