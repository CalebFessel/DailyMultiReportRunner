# DailyMultiReportRunner

This project contains a Python script that generates a bundle of operational Excel reports (on-time performance, staffing, vehicle utilization, and unit-hour utilization) for a midnight-to-midnight reporting window. [file:1]

The script is designed to be scheduled (for example, via Windows Task Scheduler or cron) and can email the generated reports to a distribution list. [file:1]

---

## Features

- **On-Time Performance (OTP)** by cost center and call type, with a reconciliation sheet to catch discrepancies. [file:1]  
- **Staffing report** showing active staffing as-of now and staffing for tomorrow’s schedule. [file:1]  
- **Daily vehicle overview** including total/in-service/out-of-service vehicles, usage for the day, unused in-service units, and out-of-service details. [file:1]  
- **Unit-Hour Utilization (UHU)**:
  - UHU by cost center.
  - UHU by shift profile. [file:1]  
- **Excel output** with formatted tables, frozen headers, and auto-fit columns. [file:1]  
- **Append workbooks** that maintain historical snapshots with configurable retention and de-duplication. [file:1]  
- **Email notifications** with the report bundle attached plus a status/health email. [file:1]  
- **Optional status dashboard integration** via a pluggable `status_logger` module. [file:1]  

---

## Requirements

### Python

- Python 3.10+ recommended (for `zoneinfo`). [file:1]  

### Python packages

Install dependencies with:

```bash
pip install -r requirements.txt
```

Suggested `requirements.txt`:

```text
pyodbc
pandas
openpyxl
python-dotenv
```

The standard library covers `logging`, `smtplib`, `email`, `ssl`, `pathlib`, etc. [file:1]

### Database / ODBC

- An ODBC DSN pointing to the reporting database (MySQL or compatible) with access to the required views/tables. [file:1]  
- The DSN name is provided via `DB_DSN` (or `TS_DSN` as a fallback). [file:1]  

---

## Configuration

All configuration is driven via environment variables (or a local `.env` file if using `python-dotenv`). [file:1]

Create a `.env` file like:

```env
# Database
DB_DSN=MyReportingDSN

# Output / retention
OUTPUT_DIR=Reports
RETENTION_DAYS=14
APPEND_RETENTION_DAYS=730

# Excel styling
EXCEL_TABLE_STYLE=TableStyleMedium9

# Email
SMTP_SERVER=smtp.example.com
SMTP_PORT=000
SMTP_USER=reports@example.com
SMTP_PASS=changeme
SMTP_FROM=reports@example.com
SMTP_EHLO_HOST=alerts.example.com

# Modes / recipients
TEST_MODE=true
TEST_MODE_RECIPIENT=reports-test@example.com
STATUS_EMAIL_RECIPIENT=reports-status@example.com

# Optional overrides used in some environments
REPORT_TEST_EMAIL=reports-test@example.com
REPORT_STATUS_EMAIL=reports-status@example.com

# Optional alternate DSN
TS_DSN=MyAlternateDSN
```

Key notes: [file:1]

- `OUTPUT_DIR` is the root folder where daily Excel files are written (default `Reports`). [file:1]  
- `APPEND_DIR` is automatically set to `OUTPUT_DIR/Append` to hold append/snapshot workbooks. [file:1]  
- `RETENTION_DAYS` controls how long individual daily `.xlsx` files are kept before cleanup. [file:1]  
- `APPEND_RETENTION_DAYS` controls how far back snapshot rows are kept inside append workbooks based on `snapshot_date` or `work_date`. [file:1]  
- `TEST_MODE=true` sends the bundle only to `TEST_MODE_RECIPIENT`; when `false`, it sends to the hard-coded `PROD_RECIPIENTS` list in the script. [file:1]  

You should edit `PROD_RECIPIENTS` in the script to match your environment or move them into environment variables. [file:1]

---

## Reports Generated

Each run produces the following Excel files in `OUTPUT_DIR` (names include the metrics date): [file:1]

- `CompanyWide_OTP_YYYY-MM-DD.xlsx`
  - “OTP by Call Type”
  - “OTP by Cost Center”
  - “OTP Reconciliation” (optional, can be disabled via `ENABLE_OTP_RECONCILIATION_SHEET` in the script). [file:1]  

- `Staffing_Report_YYYY-MM-DD.xlsx`
  - “Active Now” (current staffing as-of timestamp).  
  - “Tomorrow” (next day’s staffing). [file:1]  

- `Daily_Vehicle_Overview_YYYY-MM-DD.xlsx`
  - “Summary”
  - “In Use”
  - “Unused In Service”
  - “All In Service”
  - “Out Of Service” [file:1]  

- `Daily_UHU_By_Cost_Center_YYYY-MM-DD.xlsx` [file:1]  
- `Daily_UHU_By_Shift_Profile_YYYY-MM-DD.xlsx` [file:1]  

For each report category, an append workbook in `APPEND_DIR` accumulates snapshot data with de-duplication and retention logic, for example: [file:1]

- `CompanyWide_OTP_APPEND.xlsx`  
- `Staffing_Report_APPEND.xlsx`  
- `Daily_Vehicle_Overview_APPEND.xlsx`  
- `Daily_UHU_By_Cost_Center_APPEND.xlsx`  
- `Daily_UHU_By_Shift_Profile_APPEND.xlsx`  

---

## Where OTP Gets Its Two Timestamps

On-time performance is one subtraction: an arrival stamp minus a scheduled
pickup. Both halves are configurable, because both were lost when direct
database access went away.

| Half | Env var | Default | Source |
|---|---|---|---|
| Arrived | `TS_ARRIVAL_TIMESTAMP_KEYS` | `at_scene` | the leg's CAD `timestamps` map |
| Scheduled | `TS_PICKUP_TIME_KEYS` | `pickup_time` | the CAD grid's scheduled pickup |

Both accept a comma-separated preference chain; the first value present on a
leg wins. A leg missing either half is not scored.

**Arrival.** The old SQL scored against ePCR field 549, which the ThirdParty
API does not expose. `at_scene` stands in for it, and on the evidence it stands
in well: it published **78.2%** on 2026-09-15 against a historical series that
ran in the high 70s and low 80s. That match says field 549 was recording
arrival *at the scene*, which is what `at_scene` is. It is corroboration rather
than a tie-out — the old numbers did not survive the changeover, so there is
nothing to reconcile row by row — but the series does not step.
`at_scene: At Patient Bedside` led that chain until 2026-09-16. It was removed
on the belief that it is not captured here; the probe then found it on 229 of
385 completed legs, so that belief was wrong. It stays out for a better reason:
bedside lands *after* `at_scene` — you arrive, then you reach the patient — so a
chain that falls through scores some legs at the scene and others at the
patient and publishes both under one heading.

Pick one stamp and apply it to every leg. Measured on 2026-09-15:

| Arrival stamp | Scored | On time | Median delta |
|---|---|---|---|
| `at_scene` | 363 of 363 | **78.2%** | −3.7 min (early) |
| bedside | 228 of 363 | 58.3% | +5.8 min |
| bedside → `at_scene` (old chain) | 363 | 64.5% | +2.0 min |

Twenty points of OTP ride on that choice. `at_scene` wins on two counts. It
matches the historical range, and it scores every leg — bedside's 58.3% is
computed on the 63% of legs where a crew recorded a second status, and a crew
that skips it is plausibly on a busier call, so those legs are not a random
sample to drop. Scoring on bedside would make OTP partly a measure of
documentation discipline.

The 7.8-minute median gap between the two stamps is worth knowing on its own:
that is scene-to-patient time on the calls where it is recorded.

**Scheduled.** `pickup_time` is what the old SQL compared against, and on
2026-09-15 it was present on **100%** of completed legs — nothing to recover,
and no alternative field recovers anything. An August probe day showed only
233 of 374 carrying it, which is what the check below was built to explain; if
that gap ever returns, there are two very different reasons it could be true:

- the report is reading the wrong field — `appt_time` or
  `requested_pickup_time` might be populated where `pickup_time` is not; or
- those legs were never scheduled. An emergency call has no promised pickup,
  so there is nothing to be late against, and dropping it is correct.

Do not guess between them:

```powershell
python probe_otp_coverage.py 2026-09-15        # one busy weekday
python probe_otp_coverage.py 2026-09-01 --days 14
```

It reports, per candidate field, how many legs carry it, how many legs *only*
it can date, and what call types those legs are — then, for the arrival stamps,
**the on-time percentage each choice would actually publish**, and how far apart
`at_scene` and bedside land on the legs carrying both. Coverage says what a
stamp can date; only the second table says what the report would say. If the recovered legs are scheduled transfer work, change
`TS_PICKUP_TIME_KEYS`. If they are emergency call types, leave it alone: a
field that "recovers" them is inventing a deadline nobody gave the crew.

### Rebuilding OTP for days already past

Trips backfill roughly 90 days, so OTP — unlike UHU and staffing — can be
restated for a window that has already gone by.

```powershell
python rebuild_otp_history.py 2026-08-01                  # to yesterday
python rebuild_otp_history.py 2026-08-01 2026-09-15       # explicit end date
```

This matters because of what the arrival stamp used to default to. Until
2026-09-16 the chain was led by `at_scene: At Patient Bedside`, which meant legs
carrying it were scored at the patient and the rest at the scene — one column,
two definitions — and the published figure ran roughly fourteen points below the
pre-changeover series. A rebuild restates the window on one stamp.

It writes `CompanyWide_OTP_Rebuild_<start>_to_<end>.xlsx`: Summary, Daily,
By Cost Center, By Call Type, and Daily by Cost Center.

Three things the Summary sheet says that are worth knowing before you send the
workbook to anyone:

- **The window percentage is summed, not averaged.** `(early + on time) /
  scored` across every day. Averaging the daily percentages would weight a quiet
  Sunday like a full Monday.
- **A day that returned no legs is blank, not zero.** A 0% day and a day with no
  data are different claims and the sheet keeps them apart.
- **Cost centers come from today's map.** Cost center is not on a trip; it is
  resolved through `shift_name → crew → employee`, which the API answers only
  for the current shift window. A profile that changed cost centers during the
  window carries its current one across the whole rebuild.

**It does not write to the append workbooks.** Those are the daily runner's
record of what it published on the day, and a second writer backfilling them
would make the file's contents depend on which process ran last. If the rebuilt
series is the one you want to keep, it belongs beside the appends, not inside
them.

### Can the ePCR time be reached at all?

`probe_epcr_huly.py` asks that properly. The spec puts `Data/Epcr/Huly` out of
scope and a bare GET answered 501 — but these endpoints dispatch on `rtype`,
and `Cad/Trip` already behaves differently with and without one, so a single
unparameterised refusal proves little.

```powershell
python probe_epcr_huly.py 2026-09-15
```

It sweeps read-shaped actions (`GetRuns`, `GetTimestamps`, `GetFieldValues`, …)
and neighbouring paths, records what each answers, and flags any payload
containing something that looks like a scene arrival. It is **GET-only**; the
write-shaped `HulyUpdateTrip` is excluded, and `test_otp_sources.py` enforces
that by walking the probe's syntax tree rather than trusting a comment.

If everything refuses, the findings file in `api_probe/` is the evidence to
send Traumasoft support: it lists exactly what was tried and what came back.
Getting a read path opened is what would make OTP comparable to history again.

---

## Staffing Review

For the manager's question — **who is actually being put on a truck, and who
isn't** — use this rather than the hours report.

```powershell
python staffing_review.py                            # Cincinnati, last 60 days
python staffing_review.py --cost-center Toledo --days 30
python staffing_review.py --all-levels               # every position in the cost center
```

Output is `Staffing_Review_<cost center>_<start>_to_<end>.xlsx`: Summary,
Review, **Never Crewed**, Roster. Everyone appears; zero-day people sort to the
top, and the console prints them plus anyone not seen in 14+ days.

**It reads assignments, not punches**, and unlike hours this history already
exists. The daily run has been appending `Staffing_Report_APPEND.xlsx` since
the changeover, one row per unit per run, each carrying its crew's names and
ids — retention is 730 days, so the window goes back as far as the runner has
been going. Nothing needs to accrue first.

Reading assignments is also the better measure here. Crews are paid from
Paycor, not Traumasoft, so they have no reason to close a Traumasoft punch —
punch-outs are unreliable by construction and hours built on them vary with who
remembered. Whether someone was *put on a unit* comes from the schedule, not
from crew discipline.

It reads **both** sheets the daily run records, and they are not equivalent:

| Sheet | What it holds | Why it matters |
|---|---|---|
| **Tomorrow** | every unit whose shift *starts* that day | A whole day's schedule, independent of when the runner fired. Read first. |
| **Active Now** | units on shift at the *instant* the runner ran | A point-in-time sample. On its own it would put every crew whose shift missed that moment into the never-crewed list — a morning run would bury the night shift. |

Rows are filed under the day their **shift started**, not the day the report
ran. The Tomorrow sheet is written on one day and describes the next, so using
the run date would shift every one of those rows back a day.

Two things the Summary sheet states, both easy to misread:

- **A "day crewed" is a day assigned to a unit, not hours.** Someone rostered
  and sent home early counts the same as someone who worked the full shift.
- **Days with no record are counted and named — and they can manufacture false
  "never crewed" entries.** Someone who worked only on days the runner didn't
  cover looks identical to someone who didn't work. Check the covered-days
  figure before acting on that list. A missing append file and an empty one are
  reported differently: the first is a setup problem, the second means nobody
  was crewed.

The `hours_recorded` column fills in from `employee_hours_report.py`'s append
where it exists, and stays blank otherwise — blank meaning "not recorded",
never 0.00.

## Hours per Employee, Including the Zeros

```powershell
python employee_hours_report.py                       # Cincinnati, last 60 days
python employee_hours_report.py --cost-center Toledo --days 30
python employee_hours_report.py --list-levels         # what levels this tenant uses
python employee_hours_report.py --list-cost-centers   # what cost centers exist
```

Defaults to cost center **Cincinnati** and levels **OH EMT - Driver**,
**OH EMT - Non Driver**, **OH NEMT** — this tenant's real names, taken from
`--list-levels`. Levels are state-prefixed here (OH, WV, IN and MD each have
their own EMT and NEMT entries), so the unprefixed spellings match nobody.
Matching is case-insensitive against both `level` and `license_level` and
tolerant of spacing around the hyphen. Every matching
employee is listed **including those with no hours at all** — a zero is the
finding, and it sorts to the top of the sheet. Run `--list-levels` first if the
roster comes back empty; the names have to match what the tenant actually uses.

Output is `Employee_Hours_<cost center>_<start>_to_<end>.xlsx` with Summary,
Employees, Daily and Roster sheets.

### The hours are not backfillable — read this before asking for 60 days

The roster is complete and current. **The hours are not.** Hours come from
shift punches, and `/Schedule/Shifts` returns only `today-1..today+2` and
ignores every date filter — verified byte-identical for requests at −30, +0 and
+30 days. There is no historical punch endpoint.

So the report appends what it can see to `Employee_Hours_APPEND.xlsx` and the
window fills in **one day at a time from the first run**. Ask for 60 days today
and you get today's hours and a Summary sheet saying `Days of hours actually
recorded: 1` with 59 missing. Nothing can recover a day that was never
recorded.

Run it daily, alongside `daily_report_runner_api.py`, and the 60-day question
answers itself two months out. The Summary sheet states the covered window on
every run, so a partial answer is never mistaken for a full one.

`--no-append` gives a look without recording it — which means that day is lost.
The Summary sheet says so when you use it.

---

## How the Window Works

The script runs over a midnight-to-midnight reporting window. [file:1]

- **Default**: If no date is supplied on the command line, it runs for “yesterday” (local time) from 00:00 to today 00:00. [file:1]  
- **Backfill**: You can pass `YYYY-MM-DD` as an argument to backfill that day’s window. [file:1]  

The window is computed by:

- `window_midnight(report_end_date, now_dt)` → `(window_start, window_end)`. [file:1]  

Key derived dates: [file:1]

- `metrics_date` = `window_start.date()`; used in filenames and as the “work day” for OTP/UHU.  
- `staffing_asof` and `staffing_tomorrow_date` determine which records populate the staffing report.  
- `vehicle_date` drives the vehicle summary date filters.  

---

## Running the Script

### Local (ad hoc)

Activate your virtual environment and run:

```bash
python Daily_MultiReport_Runner.py
```

Optional CLI arguments: [file:1]

- `YYYY-MM-DD` – run the reports for that calendar day (backfill window).  
- `--no-email` – generate all files and log output but skip sending emails entirely.  

Examples:

```bash
# Normal run (yesterday's data, send emails)
python Daily_MultiReport_Runner.py

# Backfill for 2026-03-20 (send emails)
python Daily_MultiReport_Runner.py 2026-03-20

# Backfill without email
python Daily_MultiReport_Runner.py 2026-03-20 --no-email
```

### Scheduling

- On Windows, use **Task Scheduler** to run `python Daily_MultiReport_Runner.py` once per day after midnight.  
- On Linux, use **cron** with a line such as:

```cron
10 1 * * * /usr/bin/python /path/to/Daily_MultiReport_Runner.py >> /var/log/daily_reports.log 2>&1
```

Adjust paths and times for your environment. [file:1]

---

## Email Behavior

- Main email:
  - Subject: `Daily Reports Bundle - YYYY-MM-DD [TEST]` (suffix added when `TEST_MODE=true`). [file:1]  
  - Body: summary of each report (success/failure, row counts, file paths). [file:1]  
  - Attachments: all successfully generated `.xlsx` files for that run. [file:1]  

- Status email:
  - Sent to `STATUS_EMAIL_RECIPIENT`. [file:1]  
  - Includes the same summary, plus any error from the main email send (if applicable). [file:1]  

If `--no-email` is passed, both the main and status emails are skipped but logging and file output still occur. [file:1]

---

## Samsara Route Dispatch (Optional)

`push_samsara_routes.py` turns a day of Traumasoft trips into Samsara routes,
so a driver sees their scheduled transports as stops on the Samsara app
instead of only in CAD. It is independent of the reporting job: the daily run
neither needs nor touches it.

**Dry run is the default.** Without `--publish` the script reads both systems,
builds the exact payloads it would send, prints the plan, and stops. The
Samsara client is constructed read-only until `--publish` is passed, so a dry
run cannot write even if something else is wrong.

```bash
python push_samsara_routes.py                     # tomorrow, dry run
python push_samsara_routes.py 2026-09-10          # that day, dry run
python push_samsara_routes.py --json plan.json    # write the payloads out to read
python push_samsara_routes.py --publish           # actually create them
python push_samsara_routes.py --replace --publish # delete ours for that day first
python push_samsara_routes.py --vehicle M-12      # one unit only
```

### What it builds

One route per vehicle per day, named `[TS] M-12 — Thu Sep 10`. Every eligible
leg contributes two stops — pickup then drop-off — and Samsara orders the whole
list by scheduled arrival time, so the driver sees the day in sequence. Each
stop carries the run number, call type, level of service and priority in its
notes; the pickup also carries the patient name.

A stop uses a **registered Samsara address** when one matches the Traumasoft
facility name, because that brings the geofence somebody already drew around
the site. Otherwise it falls back to a single-use location built from the
leg's coordinates and postal address, which Samsara treats as a 300 m circle —
good enough for a house, coarse on a hospital campus. Registering your busiest
facilities in Samsara measurably improves arrival detection.

### How units are matched

Traumasoft and Samsara name the same unit differently — `M-12 Ford E450` and
`M12 - Medic 12` — but share the unit designator. Both names are collapsed to
`M12` and joined on that, so `M-12`, `M12`, `m 12` and `M-012` all agree while
`M12A` and `M12B` stay distinct.

Two cases are never guessed at, only reported:

- a unit with no Samsara vehicle at all
- a unit whose prefix matched several Samsara vehicles

Both are listed in the plan with the prefix that was tried. Fix them in
`state/samsara_vehicle_overrides.json` (copy the `.example.json`) rather than
renaming units in either system. An override always wins over the prefix rule.

### What is not pushed

Only legs a driver can actually be routed to. A leg is skipped when it has no
scheduled `pickup_time`, no location on one end, no vehicle assigned, or a
call type listed in `SAMSARA_EXCLUDED_CALL_TYPES`. 911 and on-demand work has
no schedule to route against and falls out here rather than landing in Samsara
with a guessed arrival time. Every skipped leg is counted by reason in the
plan, so a day that pushes fewer runs than expected explains itself.

### Time zones — read this before publishing

Traumasoft returns `pickup_time` as **local time with no UTC offset** on every
leg. Samsara wants an absolute instant. Something has to bridge that, and
getting it wrong is not subtle: a 07:30 pickup sent as `07:30Z` dispatches the
unit at 03:30 local.

Measured on 30 days of this tenant's data: `pickup_time` is naive on **all
18,026 legs**, and the per-leg `timezone` field is empty on **100%** of them.

The offset is resolved in this order:

1. `SAMSARA_TENANT_TIMEZONE` — an IANA zone (`America/New_York`), resolved
   against each trip's own date so daylight saving is handled
2. `SAMSARA_TENANT_UTC_OFFSET` — a fixed offset, wrong half the year
3. an offset on `pickup_time` itself, if the tenant ever sends one
4. the leg's `timezone` field
5. the status timestamps on the legs

**Set one of the first two.** Step 5 is the trap: it works fine for a day
already worked, which is why a 30-day probe reports a clean `-04:00` — but
tomorrow's trips have no status timestamps, because nothing has happened to
them yet, and tomorrow is exactly what this job pushes. If nothing in the
chain resolves, **the run refuses to publish** and says so. It will not guess.

### Drop-off times

Samsara requires a scheduled arrival on every stop, but Traumasoft only
schedules the pickup. The drop-off time is taken from `appt_time` — the hour
the patient is actually due — then `dropoff_eta`, and only then estimated at
pickup + `SAMSARA_DEFAULT_TRANSPORT_MINUTES` (45 by default). On real data only
about a third of legs carry an appointment time, so most drop-offs are
estimates — the plan counts them.

That matters because Samsara sorts **every** stop on a route by arrival time.
An uncapped 45-minute guess sorts past the unit's next pickup whenever two
legs are closer together than that, and the driver is told to collect the next
patient before delivering the one on board. So an estimated drop-off is pulled
back to just short of the next pickup; the plan counts those as "Capped". A
real appointment time is never moved — if it overlaps, the overlap is real.
A time landing *before* its own pickup is still pushed one minute past it.

On real data 18% of estimated drop-offs collided with the unit's next pickup
at a flat 45 minutes. Capping stops that misordering the driver, but it is
still a sign the estimate is too blunt.

There are three ways to make it, and the probe measures which one is actually
least wrong on your data rather than leaving it to taste:

| Model | What it does |
|---|---|
| `flat` (default) | `SAMSARA_DEFAULT_TRANSPORT_MINUTES` for everything |
| per level of service | `SAMSARA_TRANSPORT_MINUTES_BY_LOS`, e.g. a wheelchair van and a stretcher transport differ |
| `distance` | each leg from its own straight-line distance, `base + miles / mph` |

Section **5b** calibrates the flat numbers from the legs that do carry an
appointment time; section **5c** scores all three against that same ground
truth and prints the config for whichever wins. If none beats the flat
default it says so and tells you to leave it alone.

The distance model exists because the spread is mostly *within* a level of
service, not between: `Non Emergency` alone runs p25=30 to p90=60, and what
varies there is the trip, not the category. Its fitted speed comes out well
below any road speed — that is the straight line absorbing the difference from
the road, and it should not be "corrected" upward.

### Excluding by level of service

`SAMSARA_EXCLUDED_LOS` drops legs by `los`, case- and hyphen-insensitively
(real data carries `Non Emergency` and `Non-Emergency` as one thing typed two
ways). Empty by default. `Emergency` is the usual candidate — it is dispatched
in the moment, so a route built the night before describes nothing. Genuine
typos in the source data still have to be named individually.

### Before the first run: the readiness probe

`probe_samsara_readiness.py` is read-only and answers the questions the
mapping would otherwise guess at — run it before configuring anything:

```bash
python probe_samsara_readiness.py --days 30
```

It reports field coverage on real legs (how often coordinates and appointment
times are actually populated), whether trip stamps carry a UTC offset, your
real `call_type` and `los` vocabulary against the exclusion list, the
eligibility funnel with a reason for every dropped leg, how many drop-off
times would be estimated rather than scheduled, and the actual
Traumasoft↔Samsara vehicle match table. It works with or without a Samsara
token — without one you still get the Traumasoft-side unit prefixes.

**It never prints patient-identifying values.** Names, MRNs, phone numbers and
street addresses are counted, never shown, so its output is safe to paste into
an issue or a chat.

### Configuration

```env
SAMSARA_API_TOKEN=...                    # required
# SAMSARA_DEFAULT_TRANSPORT_MINUTES=45
# SAMSARA_EXCLUDED_CALL_TYPES=standby,cancel,no transport,dry run
```

The token needs read access to Vehicles and Addresses and write access to
Routes (Settings → Organization → API Tokens).

---

## Punch Quality and the Paycor Timecard Push (Optional)

Crews currently clock in twice — Traumasoft and Paycor — and only Paycor decides
what they are paid. Nothing rewards a complete Traumasoft punch, so punch-outs
there are unreliable by construction. That is a reporting problem, not just an
HR one: every unit-hour figure in this bundle rests on those punches.

`unit_punches_by_instance` bounds a punch with no end at the shift's end. For a
crew still on the road that is right. For a crew who simply never clocked out it
credits the whole scheduled shift — so `worked_hours` quietly becomes
`scheduled_hours` for that unit, and the gap the worked-hours denominator exists
to expose closes itself. Utilization then reads low by exactly that much.

### Measuring it first

```
python probe_punch_quality.py                      # the window as it stands
python probe_punch_quality.py --json punches.json  # append a daily baseline
```

Read-only, Traumasoft alone, no Paycor credentials needed. It reports punch-out
completion, who is not closing punches, where it concentrates, crew rostered who
never clocked in at all, and the headline number: **how many unit hours in the
UHU denominator are measured versus manufactured by the fallback.** That last
figure is computed by running the real worked-hours path twice, once over a feed
with open punches stripped — it is the fallback's contribution itself, not an
estimate of it.

**It is a four-day window, not a history.** `/Schedule/Shifts` returns
`today-1..today+2` and ignores every date filter, so punches cannot be
backfilled. A baseline accrues by running this daily and keeping the output;
a day not captured is gone. `--json` appends a row per run for that purpose.

### Pushing punches into Paycor

`push_paycor_timecards.py` sends Traumasoft punches to Paycor as timecard
punches. Three modes, meant to be used in order:

```
python push_paycor_timecards.py                       # 1. dry run
python push_paycor_timecards.py --reconcile           # 2. compare, write nothing
python push_paycor_timecards.py --reconcile --publish # 3. send them
```

**Mode 2 is the one that matters before going live.** It reads what Paycor
already holds for the same window and compares it punch by punch, so you can
prove the employee mapping resolves, the clocks agree, and the two systems are
describing the same shifts — while both sides are still read-only. A window
where every Traumasoft punch already matches Paycor is the result you want:
it means making Traumasoft authoritative changes who is trusted, not what
anyone is paid.

Rails that cannot be configured away:

- **An open punch is never sent.** A punch with no clock-out is not a payable
  record, and the shift-end fallback that makes it usable for reporting would
  here mean inventing the end of somebody's paid day.
- **An employee who cannot be mapped unambiguously is refused, not guessed.**
  Two people sharing a payroll number is exactly the case where a guess pays
  the wrong person. `state/paycor_employee_overrides.json` takes the decisions.
- **A punch Paycor already holds is skipped**, so a re-run does not double-pay.
  A publish without a reconciliation read is refused outright.
- **The first failure stops the run.** One bad punch is a fix; two hundred is
  an incident.
- **Production needs `PAYCOR_ENVIRONMENT=production` *and*
  `--i-understand-this-is-payroll`.** The client defaults to Paycor's sandbox,
  and anything but the exact string `production` resolves there, so a typo
  fails safe.

`--limit N` sends only the first N punches, for a first live test.

### What the write actually looks like

Built against Paycor's own OpenAPI spec, committed at
`docs/paycor-public-api-v1.json`. Three things in it are worth knowing before
you read the plan, because each one contradicts a reasonable assumption:

- **Punches are events, not intervals.** `POST /v1/legalentities/{id}/CreatePunches`
  takes an array where each object is a *single clock event* with one
  `punchDateTime` and a `punchStatusType` of In/Out/Auto/Transfer. So one
  Traumasoft punch row becomes **two** Paycor punches. The plan prints both.
- **Three guids are required and Paycor defaults none of them.** `employeeId`
  and `departmentId` come from the Paycor employee record; `activityTypeId` is
  a per-tenant choice (`PAYCOR_ACTIVITY_TYPE`, default `Work`). A punch missing
  any of them is refused rather than sent with a placeholder.
- **A 202 is not success.** CreatePunches validates asynchronously: it returns a
  tracking id, and whether the punches landed is only visible by reading
  `GET /v1/legalentities/{id}/punchErrorLog/{trackingId}` afterwards. Every
  publish here reads that log back before reporting anything as sent, and a
  batch whose log cannot be read is reported as **unverified** rather than
  counted either way.

Idempotency is by correlation id, not timestamps. Each punch carries a
`correlationId` derived deterministically from the Traumasoft punch id, and
Paycor returns it on read — so "have I sent this already?" is an exact lookup
that survives a re-run, a restart, and clocks that disagree by a minute.

A sandbox test is also reversible: `GET /v1/employees/{id}/employeePunches`
returns each punch's `punchId`, and `DELETE /v1/employees/{id}/DeletePunches`
takes those ids. A test you cannot undo is not a test to run against payroll.

### Credentials

Paycor wants both an OAuth 2.0 bearer token and an APIm subscription key on
every call. The bearer comes from an authorization-code grant, which needs a
human in a browser once; the refresh token that falls out is what goes in
`.env` as `PAYCOR_REFRESH_TOKEN`. `PAYCOR_ACCESS_TOKEN` accepts a token pasted
straight out of the developer portal, which is enough to try a read before
wiring the flow up properly. See the Paycor block in `.env.example`.

---

## Status Dashboard Integration (Optional)

The bottom of the script contains optional integration with a “status dashboard” via a separate `status_logger` module. [file:1]

The expected interface (if used):

- `get_cnxn()` – returns a DB connection for the status system.  
- `get_job_id(cnxn, job_name)` – resolves/creates a job ID.  
- `start_run(cnxn, job_id)` – inserts a “run started” record and returns a run ID.  
- `finish_run(cnxn, run_id, status, row_count, output_file, error_message)` – updates the run status.  
- `STATUS_SUCCESS` and `STATUS_FAILED` constants. [file:1]  

If `status_logger` cannot be imported, the script still runs normally and simply skips status dashboard calls. [file:1]

---

## Logging and Retention

- Log files are written to `OUTPUT_DIR/logs/DailyReports_YYYY-MM-DD.log`. [file:1]  
- Old daily report `.xlsx` files older than `RETENTION_DAYS` are automatically deleted from `OUTPUT_DIR`. [file:1]  
- Append workbooks are pruned based on `snapshot_date` or `work_date` using `APPEND_RETENTION_DAYS`. [file:1]  

---

## Customization

You can safely customize: [file:1]

- SQL filters (for example, excluded cost centers or certification IDs).  
- The `PROD_RECIPIENTS` list and email subjects.  
- The default DSN, output paths, and Excel table style.  

When modifying SQL, ensure parameter counts still match the number of `?` placeholders so `_expand_params_to_markers` and `pandas.read_sql_query` continue to work correctly. [file:1]
