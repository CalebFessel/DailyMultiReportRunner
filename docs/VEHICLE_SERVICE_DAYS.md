# Days in service, days out: how July can be answered

Management asked, for July, by cost center, split AMB / MH / WC:

> - How many AMB, MH and WC vehicles were supposed to be on the road.
> - How many AMB, MH and WC vehicles were out of service from July 1 through
>   July 31.
>   **NOTE: A total count is misleading because they were not OOS the whole
>   month. For each vehicle, I seek something close to out of 31 days it was
>   running X days.**

That note is the right instinct and it is what makes the question answerable
or not.

## The API cannot answer this, and never could

Traumasoft's ThirdParty API returns **one current `vehicle_status` per
vehicle**. There is no status history endpoint and no work-order endpoint —
`docs/API_MIGRATION.md` records `oos_since` and `total_days_out_of_service`
as having no API equivalent, and `probe_traumasoft_api.py` confirmed
`cost_center_id`, `cost_center_name`, `status_reason` and `oos_since` absent
against the live API.

So asking the API today tells you about today. Nothing in it knows what any
vehicle's status was on 14 July.

## What can answer it: the daily report's own archive

`Daily_Vehicle_Overview_APPEND.xlsx` has been snapshotting the Out Of Service
sheet **every day, one row per out-of-service vehicle**, with 730-day
retention (`APPEND_RETENTION_DAYS`). July 2026 is well inside that window.

A vehicle in the 5 July snapshot was out of service on 5 July. Count the
snapshot days it appears on and the answer is **exact**, not inferred. That is
what `vehicle_service_days.py` does.

**Everything depends on having that file.** It lives in the daily run's
append directory (`Reports/Append/` by default). If it is on the reporting
machine, point `--append` at it and this is answerable today. If it only
existed on the Azure VM, it has to come back with the rest of the recovery
before July can be answered at all.

## The denominator is not 31

It is the number of days in July the daily report actually ran.

A day nobody snapshotted is a day nobody observed. Folding those into "in
service" would inflate every vehicle's uptime — in the direction that flatters
the fleet and misleads the reader. So every vehicle gets three counts that sum
to the calendar month:

| Count | Means |
| --- | --- |
| `days_out_of_service` | Seen out of service in a snapshot |
| `days_in_service` | Snapshotted, and not out of service |
| `days_not_observed` | No snapshot exists for that day |

Section 1 of the report states the real denominator and names the gaps as
date ranges. If the runner was down for a week in July, that shows as seven
unobserved days, not as seven days of perfect uptime.

## "In service" and "running" are different numbers

"In service" means not carrying an out-of-service label. It does not mean the
truck worked — a unit can be in service, staffed, and simply never dispatched.
That is the whole point of the Non-Productive category in
`docs/VEHICLE_PRODUCTIVITY.md`.

So `days_ran_a_call` is counted separately, from the period's own trip legs,
which backfill reliably. Both numbers are in the export. The gap between them
is real and worth reading.

## Cost center is not on a vehicle

It is not on trips either — only on employees. The route is
`trip → shift_name → shift → user_id → employee.cost_center_name`, accumulated
into `state/shift_cost_center_map.json` by the daily run, and this report uses
the repository's own `CostCenterMap` resolver so it cannot disagree with every
other sheet.

Attribution runs best-source-first and **labels which source answered**, in
`cost_center_source`:

1. **The period's own trip legs.** Period-accurate. Preferred.
2. **Trips before the period** (`--attribution-lookback-days`, default 90).
   Weaker — a unit can move between stations. This exists because a vehicle
   out of service for the whole of July ran nothing in July, and those are
   exactly the rows being asked about.
3. **The vehicle row's live `shift_name`.** Weakest: it is today's shift, not
   July's.
4. Otherwise `UNKNOWN`. Add it to `state/shift_cost_center_overrides.json`.

## AMB / MH / WC comes from the unit name, not from data

There is no vehicle class field anywhere in the API. Class is read off the
fleet number, which is a naming convention that can change.

The default rules are `AMB=a-;MH=m-;WC=wc-`, overridable with
`VEHICLE_CLASS_PATTERNS` or `--class-patterns`. Longest prefix wins.

**Confirm MH before sending anything out.** This repository documents `M-` as
**secure car**, not Medicar — see
`probe_uhu_sources.SINGLE_CREW_VEHICLE_PREFIXES`. Whether management's "MH" is
this tenant's `M-` is not something the data says. Section 5 of the report
prints every name prefix in the fleet, how many vehicles carry it, and what it
currently maps to, so the mapping can be corrected in one pass rather than
trusted.

## Two more things the report checks rather than assumes

**Vehicles retired since July are still counted.** A truck out of service
through July and retired in August is gone from today's roster. Dropping it
would understate exactly what is being measured, so the vehicle set is the
current roster *union* everything the archive saw, and `in_current_fleet`
marks the difference.

**The Summary sheet cross-checks the row counts.** The same workbook appends a
daily Out Of Service *count*. If that count and the number of Out Of Service
*rows* disagree on a day, one of them is wrong and the report says which days.

## Running it

```
python vehicle_service_days.py --month 2026-07 --xlsx July_Vehicle_Service_Days.xlsx
python vehicle_service_days.py --month 2026-07 --append "C:\Reports\Append"
python vehicle_service_days.py --month 2026-07 --no-trips   # archive only
```

Read-only: the append workbook is never written back. The export has three
sheets — per vehicle, by cost center × class, and by class.
