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
requests
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

---

## TraumaSoft → Samsara Work Order Sync

`TraumaSoft_Samsara_WorkOrder_Sync.py` mirrors TraumaSoft vehicle maintenance work orders into **Samsara Maintenance → Work Orders**. Work orders can come from any of four sources (`TS_WO_SOURCE`, auto-detected when unset):

| Source | When it's used | Needs |
|---|---|---|
| `vehicles` (default with API credentials) | `TS_API_BASE_URL` / `TS_API_KEY` are set (the same `.env` the API-based reports use) | the project's `traumasoft_api.py` next to the script. TraumaSoft's third-party API has **no work-order endpoint** (its OpenAPI spec lists 94 paths; Fleet has only Vehicles, Equipment, CustomStatus and GpsGeofence), so this source watches each vehicle's `vehicle_status`: entering an out-of-service status opens a Samsara work order, returning to service completes it. See **Vehicle status episodes** below. |
| `api` | `TS_WO_SOURCE=api` | a TraumaSoft endpoint that returns work-order rows (`TS_WO_API_PATH`); only useful if TraumaSoft adds one. `--probe-api` checks candidate paths. |
| `db` | `DB_DSN` is set | an ODBC DSN to the TraumaSoft reporting DB; reads `vm_work_orders` joined to `sched_vehicles` / `cost_centers` |
| `csv` | `TS_WO_CSV_FILE` is set | an export of the work order list with at least an id, a vehicle name/number and a description column |

Rows from every source go through the same column resolution and mapping, so `--discover`, `TS_WO_COLUMN_MAP` and the status/priority maps work the same way for all of them.

- New TraumaSoft work orders are created in Samsara (`POST /maintenance/work-orders`).
- Changes to status, description, priority, odometer, due/completed dates are pushed as updates (`PATCH /maintenance/work-orders`).
- Every Samsara description starts with `[TS-WO:<id>]`; a ledger maps TraumaSoft ids to Samsara ids so reruns are idempotent. The ledger is either a JSON file (`Reports/State/samsara_workorder_ledger.json`, the default) or a SQL mapping table (see **Mapping table** below). If the ledger is lost, `--rebuild-ledger` reconstructs it from those tags.
- Vehicles are matched to Samsara assets by **VIN** first, then by name (case/punctuation-insensitive, so `M-12` matches `M 12`), then by `externalIds[SAMSARA_VEHICLE_EXTERNAL_ID_KEY]`. `SAMSARA_VEHICLE_MATCH` changes which is tried first. Work orders whose vehicle isn't found are skipped and listed in the summary email.
- TraumaSoft statuses/priorities/categories are mapped to Samsara's enums (`Open`, `Assigned`, `In Progress`, `On Hold`, `Pending Parts`, `Pending Approval`, `Completed`, `Closed`, `Cancelled`; `Low/Medium/High/Urgent`; `Annual/Corrective/Damage Repair/Preventive/Recall`). Common text values map automatically; numeric ids or custom labels can be mapped with `TS_WO_STATUS_MAP`, `TS_WO_PRIORITY_MAP`, `TS_WO_CATEGORY_MAP` (JSON).
- Odometers are converted from TraumaSoft miles (or km when `odometer_type` says so) to Samsara meters; timestamps are emitted as RFC 3339 in `TS_TIMEZONE` (default `America/New_York`).

### Vehicle status episodes (the `vehicles` source)

Every run fetches `GET /ThirdParty/Data/Fleet/Vehicles` (`id`, `name`, `vehicle_status`, `vin`, `odometer`) and keeps a small state file (`TS_VEHICLE_STATE_FILE`, default `Reports/State/ts_vehicle_status_state.json`):

- A vehicle whose status becomes out-of-service starts an **episode** with id `V<vehicle_id>-<yyyymmddHHMM>`. The episode is pushed to Samsara as an **Open** work order on that asset, category `Corrective`, odometer from TraumaSoft, description `[TS-WO:V56-202610140900] Vehicle M-13 marked 'Out of Service' in TraumaSoft since 2026-10-14 09:00`.
- While the vehicle stays out of service, further status changes (e.g. `Out of Service -> In Shop`) are appended to the description.
- When the vehicle returns to an in-service status, the episode is **Completed** in Samsara with `completedAtTime` set, and it keeps being checked for `TS_VEHICLE_CLOSED_RETENTION_DAYS` (30) so late edits still sync.
- Vehicles are matched to Samsara assets by **VIN first**, then name, then `externalIds` (`SAMSARA_VEHICLE_MATCH` picks which to try first; the others are fallbacks).

Which statuses count as out of service: by default any `vehicle_status` containing one of `TS_OOS_KEYWORDS` (`out of service`, `oos`, `shop`, `mechanical`, `maintenance`, `repair`, `down`, `broke`). To be exact, set `TS_OOS_STATUSES` to a JSON list of the status names, e.g. `TS_OOS_STATUSES=["Out of Service", "In Shop"]`. Run `--discover` first: it lists TraumaSoft's custom statuses and the current vehicle counts per status, starring the ones treated as out of service, and shows the episodes a real run would create. `--dry-run` does not write the state file, so the first real run starts the episodes.

This also restores the out-of-service history (`oos_since`, days out) the vehicle report lost when the reports moved from the database to the API, since the state file records every episode.

### Finding the TraumaSoft work-order endpoint

TraumaSoft's OpenAPI spec (v1.0.0, 94 paths) has no work-order or maintenance endpoint; its `CustomStatus.notify_group` even names a `work_orders` domain with no API. If a later spec adds one, probe for it with the existing client and your `.env`:

```bash
python TraumaSoft_Samsara_WorkOrder_Sync.py --probe-api
python TraumaSoft_Samsara_WorkOrder_Sync.py --probe-api --probe-path ThirdParty/Data/SomethingElse
```

It prints each candidate path with its HTTP status, and for any that answer, the row count and field names. Set `TS_WO_API_PATH` to the one that returns work orders (and `TS_WO_API_PARAMS` as JSON if it needs filters), then run `--discover`. If TraumaSoft confirms there is no work-order endpoint, use the `csv` source with a scheduled export, or the `db` source if Direct Database Access is available.

### Samsara API token

Create an API token in Samsara (Settings → API Tokens) with these scopes: **Read Vehicles**, and **Read Work Orders** + **Write Work Orders** (listed under the *Closed Beta* category — ask your Samsara rep to enable Work Orders API access if they are missing). Put it in `SAMSARA_API_TOKEN`. EU orgs set `SAMSARA_BASE_URL=https://api.eu.samsara.com`.

### First run

The exact `vm_work_orders` columns are not identical across TraumaSoft installs, so the sync discovers them at runtime. Start with:

On Windows, drop `TraumaSoft_Samsara_WorkOrder_Sync.py`, `samsara_workorders_api.py`, `Run-WorkOrderSync.ps1` and the `sql/` folder into the Reports project folder (the one that holds `traumasoft_api.py` and `.env`). If the project's `.venv` is broken (its `python.exe` points at a Python that was uninstalled), recreate it:

```powershell
winget install --id Python.Python.3.13 -e        # or the python.org installer with "Add to PATH"
cd "<Reports project folder>"
Remove-Item -Recurse -Force .venv
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`Run-WorkOrderSync.ps1` finds Python for you (`PYTHON_EXE`, `.venv`, the `py` launcher, PATH, or the usual install folders) and passes arguments through, e.g. `.\Run-WorkOrderSync.ps1 -Install` then `.\Run-WorkOrderSync.ps1 --discover`. Note that Windows PowerShell 5.1 does not accept `&&`; put commands on separate lines or join them with `;`.

```bash
# 1. See the real columns, sample rows and distinct status/priority values
python TraumaSoft_Samsara_WorkOrder_Sync.py --discover

# 2. (If needed) tell the sync which statuses / columns to use in .env:
#    TS_OOS_STATUSES=["Out of Service", "In Shop"]          # vehicles source
#    TS_WO_COLUMN_MAP={"description": ["complaint"], "status": ["wo_status"]}
#    TS_WO_STATUS_MAP={"1": "Open", "2": "In Progress", "3": "Completed"}

# 3. Dry run - logs every create/update it would do, writes nothing to Samsara
python TraumaSoft_Samsara_WorkOrder_Sync.py --dry-run --no-email

# 4. Real run on a handful of work orders
python TraumaSoft_Samsara_WorkOrder_Sync.py --limit 5
python TraumaSoft_Samsara_WorkOrder_Sync.py --ts-id 12345

# 5. Normal scheduled run (last TS_WO_LOOKBACK_DAYS days + any open ledger items)
python TraumaSoft_Samsara_WorkOrder_Sync.py
```

Other flags: `--source api|db|csv`, `--probe-api`, `--since YYYY-MM-DD`, `--lookback-days N`, `--rebuild-ledger`, `--init-ledger-table`, `--no-email`.

If the default query doesn't fit, point `TS_WO_SQL_FILE` at a `.sql` file. It must take one `?` parameter (the lookback start timestamp) and return the work order columns plus `vehicle_name`, `cost_center_name` and optionally `vehicle_odometer_type` / `vehicle_vin`.

### Endpoint table

Every call the sync makes, in the order it makes them:

| Step | Method | Samsara endpoint | Token scope | Samsara rate limit | What the sync does with it |
|---|---|---|---|---|---|
| 1 | `GET` | `/fleet/vehicles?limit=512` (paged via `pagination.endCursor`) | Read Vehicles | 25 req/sec | Builds the vehicle index used to turn a TraumaSoft vehicle into a Samsara `assetId` (by name, `externalIds`, or VIN) |
| 2 | `POST` | `/maintenance/work-orders` | Write Work Orders | 100 req/min | Creates one Samsara work order per new TraumaSoft work order |
| 3 | `PATCH` | `/maintenance/work-orders` (work order `id` in the body) | Write Work Orders | 100 req/min | Pushes status / description / priority / odometer / date changes for work orders already in the ledger |
| – | `GET` | `/maintenance/work-orders` (paged) | Read Work Orders | 5 req/sec | Only with `--rebuild-ledger`: scans descriptions for `[TS-WO:<id>]` to rebuild the ledger |

Base URL is `SAMSARA_BASE_URL` (default `https://api.samsara.com`). All requests send `Authorization: Bearer <SAMSARA_API_TOKEN>`.

TraumaSoft side: one `SELECT` against `TS_WO_TABLE` (default `vm_work_orders`) joined to `sched_vehicles` and `cost_centers`, filtered to rows changed in the last `TS_WO_LOOKBACK_DAYS`, plus a second `SELECT ... WHERE id IN (...)` for ledger work orders that are still open.

### Field mapping (TraumaSoft → Samsara)

| Logical field | TraumaSoft column candidates (first match wins) | Samsara field | Transform |
|---|---|---|---|
| id | `id`, `work_order_id`, `wo_id` | `description` prefix | `[TS-WO:<id>]` tag, also the ledger key |
| vehicle | `vehicle_id` → `sched_vehicles.name` | `assetId` | Name match (punctuation/case-insensitive), or `externalIds[SAMSARA_VEHICLE_EXTERNAL_ID_KEY]`, or VIN |
| description | `description`, `complaint`, `notes`, `details`, `problem`, `work_description`, `summary`, `title`, `name` | `description` | Prefixed with the tag; trailer adds TraumaSoft WO#, station, raw status |
| status | `status`, `status_name`, `state`, `wo_status`, else `status_id` | `status` | Text → Samsara enum via built-in map + `TS_WO_STATUS_MAP`; no status but a completed date → `Completed`; unknown → `Open` |
| priority | `priority`, `priority_name`, `priority_id` | `priority` | Text/number → `Low/Medium/High/Urgent`; default `SAMSARA_DEFAULT_PRIORITY` |
| category | `category`, `type`, `work_type`, `wo_type`, `service_type` | `category` | Text → `Annual/Corrective/Damage Repair/Preventive/Recall`; omitted if unknown |
| odometer | `odometer_started`, `odometer`, `mileage`, `odometer_completed` (+ `odometer_type`) | `odometerMeters` | miles × 1609.344, or km × 1000 |
| engine hours | `engine_hours`, `hours` | `engineHours` | integer |
| due | `due_date`, `date_due`, `due_at` | `dueAtTime` | RFC 3339 in `TS_TIMEZONE` |
| started | `date_started`, `started_at`, `start_date`, `date_start` | `startedAtTime` | create only; RFC 3339 |
| completed | `date_completed`, `completed_at`, `completed`, `date_closed`, `closed_at`, `close_date`, `date_finished` | `completedAtTime` | only when status is Completed/Closed/Cancelled |
| changed-since filter | `date_updated`, `updated_at`, `modified`, … then `date_created`, … | – | `COALESCE(updated, created, completed) >= lookback` |
| soft delete | `deleted`, `is_deleted` | – | rows with a non-zero value are excluded |
| cost center | `cost_center_id` → `cost_centers.name` | `description` trailer | `Station: <name>` |

Run `--discover` to see which real columns were picked; anything that resolved to `None` can be pointed at the right column with `TS_WO_COLUMN_MAP`.

### Mapping table (SQL ledger)

Set `SAMSARA_WO_LEDGER_DSN` to keep the TraumaSoft ↔ Samsara mapping in a database table instead of the JSON file. The table is `samsara_workorder_map` (override with `SAMSARA_WO_LEDGER_TABLE`), one row per synced work order:

| Column | Type | Meaning |
|---|---|---|
| `ts_wo_id` | VARCHAR(64), primary key | TraumaSoft `vm_work_orders.id`, or a vehicle episode id `V<vehicle_id>-<yyyymmddHHMM>` |
| `samsara_wo_id` | VARCHAR(64), unique | Samsara work order id |
| `samsara_asset_id` | VARCHAR(64) | Samsara vehicle/asset id the work order was created against |
| `ts_vehicle_id` | INT | TraumaSoft `sched_vehicles.id` |
| `vehicle_name` | VARCHAR(100) | Vehicle name at sync time |
| `samsara_status` | VARCHAR(32) | Last status pushed to Samsara |
| `fingerprint` | CHAR(64) | SHA-256 of the last PATCH payload; unchanged ⇒ no call |
| `last_action` | VARCHAR(16) | `create`, `update` or `rebuild` |
| `first_synced_at` / `last_synced_at` | DATETIME | Audit timestamps |

DDL is in `sql/samsara_workorder_map.mysql.sql` and `sql/samsara_workorder_map.mssql.sql`. Either run it yourself or let the script do it:

```bash
# .env
SAMSARA_WO_LEDGER_DSN=MyStatusDSN          # any ODBC DSN the job may write to
SAMSARA_WO_LEDGER_DIALECT=mysql            # mysql | mssql
# SAMSARA_WO_LEDGER_TABLE=samsara_workorder_map

python TraumaSoft_Samsara_WorkOrder_Sync.py --init-ledger-table   # creates the table if missing
```

The table is loaded into memory at start-up and every create/update is written through immediately (update-then-insert, committed per work order), so a crash mid-run never loses a Samsara id. Migrating from the JSON ledger: run `--rebuild-ledger` once with the DSN set and it will repopulate the table from the `[TS-WO:<id>]` tags in Samsara.

### Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `SAMSARA_API_TOKEN` | – | Bearer token (required for non-dry runs) |
| `SAMSARA_BASE_URL` | `https://api.samsara.com` | API host |
| `SAMSARA_VEHICLE_MATCH` | `vin` | Which match to try first: `vin`, `name` or `externalId` (the others are fallbacks) |
| `SAMSARA_VEHICLE_EXTERNAL_ID_KEY` | `traumasoftVehicleId` | externalIds key holding the TraumaSoft `sched_vehicles.id` |
| `SAMSARA_DEFAULT_PRIORITY` | `Medium` | Used when TraumaSoft has no priority |
| `SAMSARA_SEND_INITIAL_STATUS` | `true` | Send status/started/completed on create; falls back to create-then-PATCH if rejected |
| `TS_WO_SOURCE` | auto | `vehicles`, `api`, `db` or `csv` (see table above) |
| `TS_OOS_STATUSES` / `TS_OOS_KEYWORDS` | – / built-in list | Which TraumaSoft vehicle statuses open a work order (`vehicles` source) |
| `TS_VEHICLE_STATE_FILE` | `OUTPUT_DIR/State/ts_vehicle_status_state.json` | Episode state for the `vehicles` source |
| `TS_VEHICLE_ODOMETER_UNIT` / `TS_VEHICLE_CLOSED_RETENTION_DAYS` | `miles` / `30` | Odometer unit on TraumaSoft vehicles; how long completed episodes keep syncing |
| `TS_WO_API_PATH` / `TS_WO_API_PARAMS` | `ThirdParty/Data/WorkOrders` / – | TraumaSoft API endpoint and optional JSON query params for the `api` source |
| `TS_WO_CSV_FILE` | – | Export file for the `csv` source |
| `TS_WO_TABLE` | `vm_work_orders` | Source table for the `db` source |
| `TS_WO_LOOKBACK_DAYS` | `30` | How far back to look for new/changed work orders |
| `TS_WO_EXCLUDE_VEHICLE_IDS` | `1001,1002,1003` | Placeholder vehicles to ignore (same as the vehicle report) |
| `TS_WO_COLUMN_MAP` / `TS_WO_STATUS_MAP` / `TS_WO_PRIORITY_MAP` / `TS_WO_CATEGORY_MAP` | – | JSON overrides |
| `TS_WO_SQL_FILE` | – | Full source-query override |
| `SAMSARA_WO_LEDGER_FILE` | `OUTPUT_DIR/State/samsara_workorder_ledger.json` | Ledger location |
| `SAMSARA_WO_LEDGER_DSN` / `SAMSARA_WO_LEDGER_TABLE` / `SAMSARA_WO_LEDGER_DIALECT` | – / `samsara_workorder_map` / `mysql` | Use a SQL mapping table as the ledger instead of the JSON file |
| `SYNC_RECIPIENTS` | – | Comma-separated production recipients for the summary email |
| `TEST_MODE`, `TEST_MODE_RECIPIENT`, `STATUS_EMAIL_RECIPIENT`, `SMTP_*`, `DB_DSN`, `OUTPUT_DIR` | | Shared with the report runner |

The summary email subject gets an `(ATTENTION)` suffix when any work order failed or its vehicle couldn't be matched. Logs go to `OUTPUT_DIR/logs/SamsaraWorkOrderSync_YYYY-MM-DD.log`.

### Scheduling

Run it a few times a day alongside the report runner, e.g.

```cron
*/30 * * * * /usr/bin/python /path/to/TraumaSoft_Samsara_WorkOrder_Sync.py >> /var/log/samsara_wo_sync.log 2>&1
```

Samsara limits work order writes to 100/min; the client throttles itself and retries on 429/5xx.

### Tests

```bash
python -m unittest discover -s tests -v
```
