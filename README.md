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

`TraumaSoft_Samsara_WorkOrder_Sync.py` mirrors vehicle maintenance work orders from the TraumaSoft reporting database (`vm_work_orders`, joined to `sched_vehicles` and `cost_centers`) into **Samsara Maintenance → Work Orders**.

- New TraumaSoft work orders are created in Samsara (`POST /maintenance/work-orders`).
- Changes to status, description, priority, odometer, due/completed dates are pushed as updates (`PATCH /maintenance/work-orders`).
- Every Samsara description starts with `[TS-WO:<id>]`; a local JSON ledger (`Reports/State/samsara_workorder_ledger.json`) maps TraumaSoft ids to Samsara ids so reruns are idempotent. If the ledger is lost, `--rebuild-ledger` reconstructs it from those tags.
- Vehicles are matched to Samsara assets by **name** by default (case/punctuation-insensitive, so `M-12` matches `M 12`). Set `SAMSARA_VEHICLE_MATCH=externalId` (with `SAMSARA_VEHICLE_EXTERNAL_ID_KEY`) or `vin` to match differently. Work orders whose vehicle isn't found are skipped and listed in the summary email.
- TraumaSoft statuses/priorities/categories are mapped to Samsara's enums (`Open`, `Assigned`, `In Progress`, `On Hold`, `Pending Parts`, `Pending Approval`, `Completed`, `Closed`, `Cancelled`; `Low/Medium/High/Urgent`; `Annual/Corrective/Damage Repair/Preventive/Recall`). Common text values map automatically; numeric ids or custom labels can be mapped with `TS_WO_STATUS_MAP`, `TS_WO_PRIORITY_MAP`, `TS_WO_CATEGORY_MAP` (JSON).
- Odometers are converted from TraumaSoft miles (or km when `odometer_type` says so) to Samsara meters; timestamps are emitted as RFC 3339 in `TS_TIMEZONE` (default `America/New_York`).

### Samsara API token

Create an API token in Samsara (Settings → API Tokens) with these scopes: **Read Vehicles**, and **Read Work Orders** + **Write Work Orders** (listed under the *Closed Beta* category — ask your Samsara rep to enable Work Orders API access if they are missing). Put it in `SAMSARA_API_TOKEN`. EU orgs set `SAMSARA_BASE_URL=https://api.eu.samsara.com`.

### First run

The exact `vm_work_orders` columns are not identical across TraumaSoft installs, so the sync discovers them at runtime. Start with:

```bash
# 1. See the real columns, sample rows and distinct status/priority values
python TraumaSoft_Samsara_WorkOrder_Sync.py --discover

# 2. (If needed) tell the sync which columns to use and how to map values in .env:
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

Other flags: `--since YYYY-MM-DD`, `--lookback-days N`, `--rebuild-ledger`, `--no-email`.

If the default query doesn't fit, point `TS_WO_SQL_FILE` at a `.sql` file. It must take one `?` parameter (the lookback start timestamp) and return the work order columns plus `vehicle_name`, `cost_center_name` and optionally `vehicle_odometer_type` / `vehicle_vin`.

### Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `SAMSARA_API_TOKEN` | – | Bearer token (required for non-dry runs) |
| `SAMSARA_BASE_URL` | `https://api.samsara.com` | API host |
| `SAMSARA_VEHICLE_MATCH` | `name` | `name`, `externalId` or `vin` |
| `SAMSARA_VEHICLE_EXTERNAL_ID_KEY` | `traumasoftVehicleId` | externalIds key holding the TraumaSoft `sched_vehicles.id` |
| `SAMSARA_DEFAULT_PRIORITY` | `Medium` | Used when TraumaSoft has no priority |
| `SAMSARA_SEND_INITIAL_STATUS` | `true` | Send status/started/completed on create; falls back to create-then-PATCH if rejected |
| `TS_WO_TABLE` | `vm_work_orders` | Source table |
| `TS_WO_LOOKBACK_DAYS` | `30` | How far back to look for new/changed work orders |
| `TS_WO_EXCLUDE_VEHICLE_IDS` | `1001,1002,1003` | Placeholder vehicles to ignore (same as the vehicle report) |
| `TS_WO_COLUMN_MAP` / `TS_WO_STATUS_MAP` / `TS_WO_PRIORITY_MAP` / `TS_WO_CATEGORY_MAP` | – | JSON overrides |
| `TS_WO_SQL_FILE` | – | Full source-query override |
| `SAMSARA_WO_LEDGER_FILE` | `OUTPUT_DIR/State/samsara_workorder_ledger.json` | Ledger location |
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
