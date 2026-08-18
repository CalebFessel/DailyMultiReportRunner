# Migrating the Daily Multi-Report Runner to the Traumasoft ThirdParty API

The report bundle currently reads ~19 tables directly over ODBC. This document
maps each report onto the Traumasoft ThirdParty REST API (OpenAPI v1.0.0) and
records exactly what does and does not survive the move.

Two Traumasoft APIs exist and they are not interchangeable:

| API | Shape | Useful here? |
|---|---|---|
| **CAD2CAD** (`CAD2CAD API V1.18`) | MQTT/REST trip + AVL interchange for brokers | No — push-only, no schedule or fleet data, no history |
| **ThirdParty** (this spec) | Scoped, paginated REST with HMAC auth | Yes — this is the one to build against |

Base URL is `{tenant}/api`, e.g. `https://your-tenant.traumasoft.com/api`.

## Authentication

Four headers on every request:

| Header | Value |
|---|---|
| `X-TS-APIKEY` | issued API key |
| `X-TS-TIMESTAMP` | UNIX seconds, valid for 300s |
| `X-TS-ID` | fresh random number per request |
| `X-TS-AUTHORIZATION` | `hmac_sha256(body + timestamp + nonce, secret)` |

GPS Geofence uses a different formula: `hmac_sha256(api_key, timestamp + secret + nonce)`.

### When the key screen issues no secret

Both formulas take a `secret` "paired with that key", but the key-creation UI
may hand back only a single value. In that case the key doubles as its own
secret, so there are four plausible signing combinations rather than one.

Neither script guesses. `detect_auth_mode()` (and `Resolve-TsAuthMode` in the
PowerShell probe) tries each combination against `/ThirdParty/Data/Organization`
and keeps the first that returns a non-error status:

| Order | Formula | Signing secret |
|---|---|---|
| 1 | default | the supplied secret, or the key when none was supplied |
| 2 | default | the API key |
| 3 | legacy | the supplied secret |
| 4 | legacy | the API key |

Duplicates collapse when key and secret are the same value, so a key-only setup
tries two combinations, not four. The probe reports which one worked; pin it
afterwards with `TS_API_AUTH_MODE` to skip the detection round-trip.

If all combinations are rejected, the causes in likelihood order are: a secret
shown once at creation time that was not captured; a host clock more than five
minutes out (the timestamp is valid for 300 seconds — check `w32tm /query
/status`); or the key lacking read scope for Organization.

## Rate limits and paging

- 100 read requests/minute, 50 writes/minute; over the limit returns `429` with
  `Retry-After`. The client throttles to ~92/min and honours `Retry-After`.
- Page size is capped at **100** for Vehicles, Users, Payors, and Facilities.
- Two paginator dialects are in use and both must be handled:
  `{current_page, total_pages, total_records}` and jqGrid's `{page, total, records}`.
- Some endpoints return a bare array (Shifts, Organization, GetTrips) and small
  reference lists return a single named key (`custom_statuses`, `pay_types`, ...).

## Report-by-report mapping

### On-Time Performance

| Current source | Replacement |
|---|---|
| `cad_trip_legs` + `cad_trip_legs_rev` | `GET /ThirdParty/Data/Cad/Trip?rtype=GetTrips&trip_date=&range_days=` |
| `sched_unit_types` (call type join) | `call_type` string, already resolved on `TripLegSummary` |
| `rev.pickup_time` | `pickup_time` |
| `epcr_v2_values.field_value` (field 549) | **not available** — see below |

`range_days` is inclusive and capped at 31, so backfill works up to a month per call.

**ePCR is explicitly out of scope for this API** (the spec excludes
`ThirdParty/Data/Epcr/Huly` and directs partners to dedicated credentials). The
current OTP "actual" time comes from ePCR field 549, so it has to be replaced by
a CAD status time from the `timestamps` array on each leg. That array is a list
of `status name -> ISO time` maps; run `probe_traumasoft_api.py` to see which
names this tenant actually emits before picking one.

Expect OTP percentages to shift after the switch — the new series will not tie
to the historical one. `TripLegSummary` also exposes `late_reasons`, which the
current report has no equivalent for.

### Staffing

| Current source | Replacement |
|---|---|
| `sched_template_shift_assignments` | `GET /ThirdParty/Data/Schedule/Shifts` |
| `sched_units` (shift profile) | `shift_name` |
| `users` (crew names) | `GET /ThirdParty/Data/User/Employees` joined on `user_id` |
| `sched_unit_certification_templates` | `license_level` on the shift row |
| `cost_centers` | `cost_center_name` on the employee row |

Two problems, both worth confirming with the probe before writing report code:

1. **`/Schedule/Shifts` documents no date filter and no pagination** — only
   `include_deleted`. If the endpoint returns every shift, the report has to pull
   everything and filter client-side. The probe tests a set of undocumented
   filter parameters to find out whether server-side filtering exists.
2. **The `Shift` schema omits fields the current SQL filters on**: `published`,
   `timeoff_type` (the absent/noshow/lwop exclusion), and `schedule_type`. Without
   them, crew counts will include people the current report excludes.

### Unit-Hour Utilization

- **Loaded hours / run counts (numerator)** — derived from the trip `timestamps`
  array. This is arguably better than the current
  `est_trip_duration + shifted_amount` estimate, since the timestamps are actual.
- **Scheduled hours (denominator)** — `start_time`/`end_time` overlap from
  `/Schedule/Shifts`, so it inherits the date-filter problem above.
- **By shift profile** — `shift_name` appears on both `Shift` and
  `TripLegSummary`, which makes it a clean join key.
- **By cost center** — see below.

### Daily Vehicle Overview

| Sheet | Status |
|---|---|
| Summary | `GET /ThirdParty/Data/Fleet/Vehicles` + `/Fleet/CustomStatus` |
| All In Service | covered (`vehicle_status`) |
| In Use | covered (`shift_name` / `trip_status` enrichment, or the day's trips) |
| Unused In Service | covered (in-service roster minus used) |
| Out Of Service | **partial** |

The vehicle field allowlist is closed and explicit: `id, name, vehicle_status,
vin, odometer, disabled, deleted`, plus enrichment (`shift_name`, `trip_status`,
`trip_timestamp`, `post_id`, `post_status`, `post_timestamp`). Nothing else is
reachable, which drops four columns from the Out Of Service sheet:

- `status_reason` — no equivalent field
- `oos_since` / `total_days_out_of_service` — no vehicle status history endpoint
- `odometer_started` / `odometer_completed` — no work-order endpoint
- `work_order_station` — same

`oos_since` can be rebuilt locally: `Daily_Vehicle_Overview_APPEND.xlsx` already
snapshots this report daily with 730-day retention, so the first snapshot in
which a vehicle appears out of service becomes its OOS-since date. That is exact
from changeover forward and needs no API support. The work-order columns have no
workaround short of Traumasoft exposing a maintenance scope.

Note the current report already reads `sched_vehicles.status` as a *current*
value, so the API's current-status-only model is parity here, not a regression.

## Cost centers

Cost center is the weakest link in the whole migration. It appears on
**employees** (`cost_center_name`) and in the DDG hierarchy
(`GET /ThirdParty/Data/Organization` → `cost_center_ids[]`), but **not on trips
and not on vehicles**.

Available paths, best first:

1. **Employees** — direct. Staffing works with no extra machinery.
2. **Trip → `shift_name` → Shift → `user_id` → Employee `cost_center_name`.**
   Viable, and the probe measures the actual resolution rate and flags shift
   profiles that map to more than one cost center.
3. **Shift `division`/`district` → Organization → `cost_center_ids`.** Coarser —
   `Shift` carries division and district but not group, so a district spanning
   several cost centers is ambiguous.
4. **A maintained `shift_profile → cost_center` mapping**, seeded from
   `GET /ThirdParty/Lists/Schedule/ShiftProfiles`. Least elegant, most reliable.

## Scope UI vs. spec

The API-key scope screen lists `UserCertifications` and `UserCostCenters` as
readable entities, but this spec defines no `/Data/.../UserCertifications` or
`/Data/.../UserCostCenters` paths — only `/Lists/HumanResources/Certifications`
and `/Lists/Cad/CostCenters`, which are `id` + `name` lookups. Either the spec
lags the scope list or those scopes gate fields on other endpoints. Worth asking
Traumasoft, since dedicated certification and cost-center endpoints would remove
most of the awkwardness above.

## Confirmed against the live tenant

Two probes ran against `lynxems.traumasoft.com`. Auth resolved to the
documented default formula with a real paired secret. What they settled:

| Question | Answer |
|---|---|
| Do trips backfill? | **Yes** — 90 days back returned 665 legs. OTP and the UHU numerator keep full history and `--date` backfill still works. |
| Can `/Schedule/Shifts` be steered? | **No** — 751 rows, `today-1..today+2`, byte-identical for requests at −30, +0 and +30 days. |
| Is the arrival timestamp there? | **Yes** — `at_scene`, plus a distinct `at_scene: At Patient Bedside`. |
| Real cost-center coverage for OTP | **100%** — all 233 OTP-scorable legs carry `shift_name`. |
| Is the vehicle allowlist closed? | **Yes** — no unlisted field returned. |

### The 66.7% was a red herring

The first probe's headline cost-center coverage counted every leg. Broken out
by status, the unattributable legs are exactly the ones OTP never scores:

| Trip status | Legs | With `shift_name` | With `at_scene` |
|---|---|---|---|
| Clear | 374 | 374 | 374 |
| Canceled | 200 | 94 | 3 |
| Disregard | 108 | 12 | 0 |
| No Transport | 17 | 4 | 0 |
| unknown | 25 | 0 | 0 |

Of the 233 legs carrying both a `pickup_time` and an `at_scene` stamp — the
population OTP actually scores — **every one** has `shift_name`. Cost-center
attribution for OTP is complete.

Note that 374 legs reach `Clear` but only 233 have a scheduled `pickup_time`,
so a third of completed legs are unscoreable for lateness. The current SQL
filters on `rev.pickup_time` too, so it likely excludes the same legs, but that
has not been confirmed against a parallel run.

### What being present-only forces

Because shifts describe only the present, two things are accumulated to disk
rather than queried (`TS_STATE_DIR`, default `state/`):

- **`shift_cost_center_map.json`** — `shift_name` → cost center, learned from
  each day's visible shift window. Historical legs are attributed through it.
  Crew counts per cost center are kept rather than a single winner, so the
  dominant cost center wins and contested profiles stay visible. 11 of 128
  profiles map to more than one cost center, carrying 35 of 483 legs (7.2%);
  `West Virginia Admin` spans five, which is a float pool rather than a station.
- **`vehicle_oos_history.json`** — first date each vehicle was observed out of
  service, which is what `oos_since` and days-out are derived from. Cleared
  when a vehicle returns to service.

Both are cold-start empty: a vehicle already out of service on changeover day
shows a blank `oos_since` until it cycles, and a shift profile only becomes
attributable once it has been seen staffed.

### Vehicle status mapping

The status strings replace the status id + `status_reason LIKE` filtering:

| Status | Count | Treatment |
|---|---|---|
| In Service | 102 | in service |
| Out of Service | 36 | out of service |
| Out of Service - Collision | 4 | out of service |
| Retired | 36 | excluded from the fleet |
| New - Waiting for Delivery | 12 | excluded from the fleet |
| Waiting for Inspection | 1 | excluded from the fleet |

### Correction

An earlier draft listed `late_reasons` as a capability gained. It came back on
**0 legs**. Do not build on it.

## What is unaffected

Everything downstream of the data pull is source-agnostic and stays as-is:
Excel formatting, the append/dedupe/retention workbooks, email delivery, and the
optional `status_logger` integration. Only the SQL constants and the eight
`pd.read_sql_query(...)` calls need replacing.

## Running the probe

Credentials come from a `.env` file next to the scripts, so no shell exporting
is needed on any platform:

```
copy .env.example .env      # Windows;  cp .env.example .env on Linux/macOS
```

Then fill in `TS_API_BASE_URL`, `TS_API_KEY`, and `TS_API_SECRET`.

### Windows, nothing installed (no Python, no clone)

`probe_traumasoft_api.ps1` is standalone — pure PowerShell, no dependencies,
no repository. Save that one file anywhere and run it:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\probe_traumasoft_api.ps1 -Date 2026-08-17
```

It prompts for the base URL, key, and secret (press Enter at the secret prompt
if the key screen never issued one). Everything it needs — HMAC-SHA256, HTTPS,
JSON — is built into .NET.

### Windows, with the repository checked out

```powershell
cd C:\path\to\DailyMultiReportRunner
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass   # if the script is blocked
.\run_probe.ps1 -Date 2026-08-17
```

`run_probe.ps1` locates a real Python, builds `.venv`, and runs the Python
probe. Use it when Python is available; use the standalone probe above when it
is not.

`run_probe.ps1` locates a real Python 3.10+, creates `.venv`, installs
`requests` and `python-dotenv`, and runs the probe. It deliberately validates
each interpreter by asking for its version, because the Microsoft Store app
alias answers to `python` without being Python — that alias is what produces
"Python was not found; run without arguments to install from the Microsoft
Store". If no usable interpreter exists, the script prints where to look for the
one that already runs the daily report.

Note that PowerShell has no `export`. To set variables for a single session
instead of using `.env`:

```powershell
$env:TS_API_BASE_URL = "https://your-tenant.traumasoft.com"
$env:TS_API_KEY      = "..."
$env:TS_API_SECRET   = "..."
```

### Linux / macOS

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
python probe_traumasoft_api.py 2026-08-17
```

Either way the probe writes `api_probe/PROBE_FINDINGS.md`,
`api_probe/findings.json`, and raw JSON samples for each entity. It issues GETs
only, and `api_probe/` and `.env` are both gitignored.
