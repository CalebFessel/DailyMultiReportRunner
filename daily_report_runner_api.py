"""
Daily Multi-Report Runner, backed by the Traumasoft ThirdParty API.

Drop-in replacement for the ODBC runner. Produces the same five workbooks, the
same append/snapshot workbooks, and the same emails; only the data source
changed. Report content comes from traumasoft_reports, and everything
downstream -- Excel formatting, de-duplication, retention, delivery -- is the
original code, lifted unchanged into report_output.

    python daily_report_runner_api.py                 # yesterday
    python daily_report_runner_api.py 2026-08-17      # backfill that day
    python daily_report_runner_api.py --zip           # also bundle the day into one .zip
    python daily_report_runner_api.py --no-email      # never attempt delivery
    python daily_report_runner_api.py --dry-run       # write nothing, report what would happen

Email is optional. With no SMTP server and no recipients configured the run
writes its workbooks and stops there, which is the right behaviour when the
files are being sent by hand.

What backfill can and cannot do, per the live probes:

    OTP and the UHU numerator backfill properly -- GetTrips returned data 90
    days back. Staffing, and the scheduled-hours half of UHU, cannot: the
    shifts endpoint ignores every filter and always returns the same window
    around now. Backfilling an older date therefore produces a correct OTP
    report and a staffing report describing today, which is misleading unless
    it is called out. It is, loudly, in the run summary and the email body.
"""

import os
import sys
import zipfile
import logging
import traceback
from pathlib import Path
from datetime import datetime, timedelta, date

import pandas as pd

import report_output as OUT
import traumasoft_reports as R
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

# =============================
# CONFIG
# =============================
TEST_MODE = os.getenv("TEST_MODE", "false").strip().lower() in ("1", "true", "yes", "y")
TEST_MODE_RECIPIENT = (
    os.getenv("TEST_MODE_RECIPIENT")
    or os.getenv("REPORT_TEST_EMAIL")
    or "reports-test@example.com"
)
STATUS_EMAIL_RECIPIENT = (
    os.getenv("STATUS_EMAIL_RECIPIENT")
    or os.getenv("REPORT_STATUS_EMAIL")
    or "reports-status@example.com"
)
PROD_RECIPIENTS = [
    r.strip() for r in os.getenv("PROD_RECIPIENTS", "").split(",") if r.strip()
]

EMAIL_SUBJECT = "Daily Reports Bundle - {date}{test_suffix}"
JOB_NAME = "Daily Multi-Report Runner (API)"

# Reports whose data is only ever current, regardless of the date requested.
PRESENT_ONLY_REPORTS = ("Staffing", "UHU (scheduled hours)")


# =============================
# CLI
# =============================
def parse_cli_end_date(argv):
    for arg in argv[1:]:
        if arg.startswith("-"):
            continue
        try:
            return datetime.strptime(arg.strip(), "%Y-%m-%d").date()
        except ValueError:
            continue
    return None


def has_flag(argv, flag):
    flag = flag.lower()
    return any(a.lower() == flag for a in argv[1:])


def email_is_configured():
    """
    Whether delivery is even possible.

    With no SMTP credentials or no recipients there is nothing to attempt, and
    a run that produces its workbooks and stops is a success rather than a
    failure -- that is the normal state when the files are sent by hand.
    """
    has_smtp = bool(os.getenv("SMTP_USER")) and bool(
        os.getenv("SMTP_PASS") or os.getenv("SMTP_PASSWORD")
    )
    has_recipients = bool(PROD_RECIPIENTS) or (TEST_MODE and bool(TEST_MODE_RECIPIENT))
    return has_smtp and has_recipients


def bundle_zip(paths, output_dir, run_date_str):
    """Bundle the day's workbooks into one archive that is easy to attach."""
    zip_path = os.path.join(output_dir, f"Daily_Reports_{run_date_str}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            if os.path.exists(path):
                archive.write(path, arcname=os.path.basename(path))
    return zip_path


# =============================
# WORKBOOKS
# =============================
def _write_workbook(out_path, sheets):
    """Write {sheet name: DataFrame} to one formatted workbook."""
    Path(os.path.dirname(out_path) or ".").mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            OUT.write_df_sheet_with_table(
                writer,
                df if df is not None else pd.DataFrame(),
                sheet_name,
                f"{sheet_name}",
            )
    return out_path


def write_otp(reports, run_date_str, output_dir, append_dir):
    path = os.path.join(output_dir, f"CompanyWide_OTP_{run_date_str}.xlsx")
    _write_workbook(path, {
        "OTP by Call Type": reports["otp_by_call_type"],
        "OTP by Cost Center": reports["otp_by_cost_center"],
    })
    append_path = os.path.join(append_dir, "CompanyWide_OTP_APPEND.xlsx")
    OUT._append_to_workbook_xlsx(
        append_path, "OTP by Call Type", reports["otp_by_call_type"],
        dedupe_keys=["snapshot_date", "cost_center", "call_type"],
        snapshot_date_value=run_date_str,
    )
    OUT._append_to_workbook_xlsx(
        append_path, "OTP by Cost Center", reports["otp_by_cost_center"],
        dedupe_keys=["snapshot_date", "cost_center"],
        snapshot_date_value=run_date_str,
    )
    return (
        f"CallType={len(reports['otp_by_call_type'])}, "
        f"CostCenter={len(reports['otp_by_cost_center'])}",
        path,
    )


def write_staffing(reports, run_date_str, output_dir, append_dir):
    path = os.path.join(output_dir, f"Staffing_Report_{run_date_str}.xlsx")
    _write_workbook(path, {
        "Active Now": reports["staffing_active_now"],
        "Tomorrow": reports["staffing_tomorrow"],
    })
    append_path = os.path.join(append_dir, "Staffing_Report_APPEND.xlsx")
    for sheet, key in (("Active Now", "staffing_active_now"), ("Tomorrow", "staffing_tomorrow")):
        OUT._append_to_workbook_xlsx(
            append_path, sheet, reports[key],
            dedupe_keys=["snapshot_date", "cost_center", "shift_profile", "start_time", "end_time"],
            snapshot_date_value=run_date_str,
        )
    return (
        f"ActiveNow={len(reports['staffing_active_now'])}, "
        f"Tomorrow={len(reports['staffing_tomorrow'])}",
        path,
    )


def write_vehicles(reports, run_date_str, output_dir, append_dir):
    path = os.path.join(output_dir, f"Daily_Vehicle_Overview_{run_date_str}.xlsx")
    _write_workbook(path, {
        "Summary": reports["vehicle_summary"],
        "In Use": reports["vehicles_in_use"],
        "Unused In Service": reports["vehicles_unused_in_service"],
        "All In Service": reports["vehicles_all_in_service"],
        "Out Of Service": reports["vehicles_out_of_service"],
    })
    append_path = os.path.join(append_dir, "Daily_Vehicle_Overview_APPEND.xlsx")
    OUT._append_to_workbook_xlsx(
        append_path, "Summary", reports["vehicle_summary"],
        dedupe_keys=["snapshot_date", "metric"], snapshot_date_value=run_date_str,
    )
    OUT._append_to_workbook_xlsx(
        append_path, "Out Of Service", reports["vehicles_out_of_service"],
        # Keyed on the name, not the id: Excel reads a column of digit strings
        # back as integers, so "7" never matches the 7 that comes out of the
        # append file and a same-day re-run appends instead of replacing.
        dedupe_keys=["snapshot_date", "vehicle_name"], snapshot_date_value=run_date_str,
    )
    return f"Summary={len(reports['vehicle_summary'])}, OOS={len(reports['vehicles_out_of_service'])}", path


def write_runs(reports, run_date_str, output_dir, append_dir):
    path = os.path.join(output_dir, f"Daily_Run_Volume_{run_date_str}.xlsx")
    _write_workbook(path, {
        "Runs by Cost Center": reports["runs_by_cost_center"],
        "Runs by Vehicle": reports["runs_by_vehicle"],
    })

    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, "Daily_Run_Volume_By_Cost_Center_APPEND.xlsx"),
        "Runs by Cost Center", reports["runs_by_cost_center"],
        dedupe_keys=["snapshot_date", "cost_center_name"], snapshot_date_value=run_date_str,
    )
    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, "Daily_Run_Volume_By_Vehicle_APPEND.xlsx"),
        "Runs by Vehicle", reports["runs_by_vehicle"],
        # Keyed on the name for the same reason as the vehicle overview above.
        dedupe_keys=["snapshot_date", "vehicle_name"], snapshot_date_value=run_date_str,
    )

    cc = reports["runs_by_cost_center"]
    veh = reports["runs_by_vehicle"]
    total = int(cc["total_runs"].sum()) if not cc.empty else 0
    return (
        f"{total} runs across {len(cc)} cost center(s) and {len(veh)} vehicle(s)",
        [path],
    )


def write_uhu(reports, run_date_str, output_dir, append_dir):
    paths = []
    cc_path = os.path.join(output_dir, f"Daily_UHU_By_Cost_Center_{run_date_str}.xlsx")
    _write_workbook(cc_path, {"UHU by Cost Center": reports["uhu_by_cost_center"]})
    paths.append(cc_path)

    sp_path = os.path.join(output_dir, f"Daily_UHU_By_Shift_Profile_{run_date_str}.xlsx")
    _write_workbook(sp_path, {"UHU by Shift Profile": reports["uhu_by_shift_profile"]})
    paths.append(sp_path)

    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, "Daily_UHU_By_Cost_Center_APPEND.xlsx"),
        "UHU by Cost Center", reports["uhu_by_cost_center"],
        dedupe_keys=["snapshot_date", "cost_center_name"], snapshot_date_value=run_date_str,
    )
    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, "Daily_UHU_By_Shift_Profile_APPEND.xlsx"),
        "UHU by Shift Profile", reports["uhu_by_shift_profile"],
        dedupe_keys=["snapshot_date", "shift_profile_name"], snapshot_date_value=run_date_str,
    )
    return (
        f"ByCostCenter={len(reports['uhu_by_cost_center'])}, "
        f"ByShiftProfile={len(reports['uhu_by_shift_profile'])}",
        paths,
    )


# =============================
# MAIN
# =============================
def main():
    metrics_date = parse_cli_end_date(sys.argv) or (date.today() - timedelta(days=1))
    no_email = has_flag(sys.argv, "--no-email")
    dry_run = has_flag(sys.argv, "--dry-run")
    make_zip = has_flag(sys.argv, "--zip")
    run_date_str = metrics_date.isoformat()

    output_dir = OUT.OUTPUT_DIR
    append_dir = OUT.APPEND_DIR
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    Path(append_dir).mkdir(parents=True, exist_ok=True)

    log_path = OUT.setup_logging(output_dir, run_date_str)
    logging.info("=== %s ===", JOB_NAME)
    logging.info("Metrics date: %s   test_mode=%s   dry_run=%s", run_date_str, TEST_MODE, dry_run)
    logging.info("Log: %s", log_path)

    backfilling = metrics_date < (date.today() - timedelta(days=1))

    # --- connect ---
    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        logging.error("Configuration error: %s", exc)
        return 2
    if api.detect_auth_mode() is None:
        logging.error("Credentials rejected under every signing scheme; aborting.")
        return 1

    # --- fetch + build ---
    try:
        data = R.fetch_day(api, metrics_date)
        # One extra call: how recently each vehicle ran, so out-of-service
        # records that have not moved in weeks are visible as such.
        fleet_activity = R.fetch_fleet_activity(api, metrics_date)
        reports = R.build_all(data, fleet_activity=fleet_activity)
    except TraumasoftAPIError as exc:
        logging.error("API failure: %s", exc, exc_info=True)
        return 1

    results = []
    attachments = []

    writers = [
        ("Company-Wide OTP", write_otp),
        ("Staffing", write_staffing),
        ("Daily Vehicle Overview", write_vehicles),
        ("Unit-Hour Utilization", write_uhu),
        ("Run Volume", write_runs),
    ]
    for name, writer in writers:
        logging.info("--- Running: %s ---", name)
        if dry_run:
            logging.info("--- DRY RUN: %s not written ---", name)
            results.append((name, True, "dry run", None))
            continue
        try:
            rows_info, path = writer(reports, run_date_str, output_dir, append_dir)
            paths = path if isinstance(path, list) else [path]
            attachments.extend(paths)
            logging.info("--- OK: %s (%s)", name, rows_info)
            results.append((name, True, rows_info, ", ".join(paths)))
        except Exception as exc:  # one bad report must not lose the others
            logging.error("--- FAIL: %s: %s", name, exc, exc_info=True)
            results.append((name, False, "", str(exc)))

    # --- summary ---
    lines = [f"{JOB_NAME}", f"Metrics date: {run_date_str}", ""]
    for name, ok, info, detail in results:
        lines.append(f"{'OK  ' if ok else 'FAIL'}  {name}: {info or detail}")
    lines.append("")

    ambiguous = R.CostCenterMap().ambiguous()
    if ambiguous:
        lines.append(
            f"Note: {len(ambiguous)} shift profile(s) map to more than one cost center; "
            "the dominant one is used. See state/shift_cost_center_map.json."
        )
    if backfilling:
        lines.append("")
        lines.append(
            "WARNING: this is a backfill run. OTP and UHU run counts reflect "
            f"{run_date_str}, but {' and '.join(PRESENT_ONLY_REPORTS)} reflect TODAY -- the "
            "Traumasoft shifts endpoint always returns the current window and cannot be "
            "asked for a past date. Treat those sheets as current, not historical."
        )
    body = "\n".join(lines)
    logging.info("\n%s", body)

    if dry_run:
        logging.info("Dry run complete; nothing written or sent.")
        return 0

    if make_zip and attachments:
        try:
            zip_path = bundle_zip(attachments, output_dir, run_date_str)
            logging.info("Bundled %s files into %s", len(attachments), zip_path)
        except Exception as exc:
            logging.warning("Could not build the zip bundle: %s", exc)

    try:
        OUT.cleanup_old_files()
    except Exception as exc:
        logging.warning("Cleanup failed: %s", exc)

    # --- where the files are ---
    logging.info("")
    logging.info("Files for %s are in %s:", run_date_str, os.path.abspath(output_dir))
    for path in attachments:
        logging.info("  %s", os.path.basename(path))

    # --- email (optional) ---
    if no_email:
        logging.info("--no-email set; delivery skipped.")
        return 0

    if not email_is_configured():
        logging.info(
            "Email not configured (needs SMTP_USER, SMTP_PASS and PROD_RECIPIENTS); "
            "the workbooks above are ready to send by hand."
        )
        return 0

    recipients = [TEST_MODE_RECIPIENT] if TEST_MODE else PROD_RECIPIENTS
    subject = EMAIL_SUBJECT.format(date=run_date_str, test_suffix=" [TEST]" if TEST_MODE else "")
    send_error = None
    try:
        OUT.send_email(subject, body, recipients, attachments)
        logging.info("Bundle sent to %s", ", ".join(recipients))
    except Exception as exc:
        send_error = str(exc)
        logging.error("Failed to send bundle: %s", exc, exc_info=True)

    try:
        status_body = body if not send_error else f"{body}\n\nMAIN EMAIL FAILED: {send_error}"
        OUT.send_email(f"[STATUS] {subject}", status_body, [STATUS_EMAIL_RECIPIENT])
    except Exception as exc:
        logging.error("Failed to send status email: %s", exc)

    return 1 if (send_error or any(not ok for _, ok, _, _ in results)) else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        logging.error("Unhandled failure:\n%s", traceback.format_exc())
        sys.exit(1)
