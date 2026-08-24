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
    python daily_report_runner_api.py --region Indiana  # the same day, one region only

A regional run filters every sheet that names a cost center and writes its
workbooks under that region's name. It appends nothing: the company-wide run
already recorded the day for every cost center, that region's included. Which
cost centers make up a region is read from state/regions.json.

The Daily Vehicle Overview stays fleet-wide even under --region, because the
API puts no cost center on a vehicle. Run Volume by Vehicle is regional -- a
leg is attributable through its shift profile.

For a month-to-date regional bundle, see monthly_region_report.py.

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


def parse_region(argv):
    """
    `--region Indiana` or `--region="West Virginia"`.

    A regional run is a view of the same day, so it writes the day's workbooks
    filtered and leaves the append history alone -- the company-wide run owns
    that, and Indiana's rows are already in it. Two writers appending the same
    dates would only mean the file's contents depended on which run went last.
    """
    args = argv[1:]
    for index, arg in enumerate(args):
        lowered = arg.lower()
        if lowered.startswith("--region="):
            return arg.split("=", 1)[1].strip() or None
        if lowered == "--region":
            return args[index + 1].strip() if index + 1 < len(args) else None
    return None


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


def bundle_zip(paths, output_dir, run_date_str, tag=""):
    """Bundle the day's workbooks into one archive that is easy to attach."""
    zip_path = os.path.join(output_dir, f"Daily_Reports{tag}_{run_date_str}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            if os.path.exists(path):
                archive.write(path, arcname=os.path.basename(path))
    return zip_path


# =============================
# WORKBOOKS
# =============================
def _out_name(base, tag, run_date_str):
    """Reports/CompanyWide_OTP_Indiana_2026-08-23.xlsx, or without the region."""
    return f"{base}{tag}_{run_date_str}.xlsx"


def _append(append_dir, filename, sheet_name, df, dedupe_keys, snapshot_date_value):
    """
    Append a day's rows, unless this is a regional run.

    A regional bundle is a filtered view of a day the company-wide run already
    recorded, so it writes no history of its own.
    """
    if append_dir is None:
        return
    OUT._append_to_workbook_xlsx(
        os.path.join(append_dir, filename), sheet_name, df,
        dedupe_keys=dedupe_keys, snapshot_date_value=snapshot_date_value,
    )


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


def write_otp(reports, run_date_str, output_dir, append_dir, tag=""):
    path = os.path.join(output_dir, _out_name("CompanyWide_OTP", tag, run_date_str))
    _write_workbook(path, {
        "OTP by Call Type": reports["otp_by_call_type"],
        "OTP by Cost Center": reports["otp_by_cost_center"],
    })
    _append(
        append_dir, "CompanyWide_OTP_APPEND.xlsx",
        "OTP by Call Type", reports["otp_by_call_type"],
        ["snapshot_date", "cost_center", "call_type"], run_date_str,
    )
    _append(
        append_dir, "CompanyWide_OTP_APPEND.xlsx",
        "OTP by Cost Center", reports["otp_by_cost_center"],
        ["snapshot_date", "cost_center"], run_date_str,
    )
    return (
        f"CallType={len(reports['otp_by_call_type'])}, "
        f"CostCenter={len(reports['otp_by_cost_center'])}",
        path,
    )


def write_staffing(reports, run_date_str, output_dir, append_dir, tag=""):
    path = os.path.join(output_dir, _out_name("Staffing_Report", tag, run_date_str))
    _write_workbook(path, {
        "Active Now": reports["staffing_active_now"],
        "Tomorrow": reports["staffing_tomorrow"],
    })
    for sheet, key in (("Active Now", "staffing_active_now"), ("Tomorrow", "staffing_tomorrow")):
        _append(
            append_dir, "Staffing_Report_APPEND.xlsx", sheet, reports[key],
            ["snapshot_date", "cost_center", "shift_profile", "start_time", "end_time"],
            run_date_str,
        )
    # Rows are now every unit on shift, not only the crewed ones, so the count
    # alone would read as a jump in coverage. Say how many are short.
    def tally(key):
        df = reports[key]
        if df.empty or "staffing_status" not in df:
            return f"{len(df)}"
        short = int((df["staffing_status"] != "OK").sum())
        return f"{len(df)}" + (f" ({short} short)" if short else "")

    return (
        f"ActiveNow={tally('staffing_active_now')}, "
        f"Tomorrow={tally('staffing_tomorrow')}",
        path,
    )


def write_vehicles(reports, run_date_str, output_dir, append_dir, tag=""):
    path = os.path.join(output_dir, _out_name("Daily_Vehicle_Overview", tag, run_date_str))
    _write_workbook(path, {
        "Summary": reports["vehicle_summary"],
        "In Use": reports["vehicles_in_use"],
        "Unused In Service": reports["vehicles_unused_in_service"],
        "All In Service": reports["vehicles_all_in_service"],
        "Out Of Service": reports["vehicles_out_of_service"],
    })
    _append(
        append_dir, "Daily_Vehicle_Overview_APPEND.xlsx",
        "Summary", reports["vehicle_summary"],
        ["snapshot_date", "metric"], run_date_str,
    )
    _append(
        append_dir, "Daily_Vehicle_Overview_APPEND.xlsx",
        "Out Of Service", reports["vehicles_out_of_service"],
        # Keyed on the name, not the id: Excel reads a column of digit strings
        # back as integers, so "7" never matches the 7 that comes out of the
        # append file and a same-day re-run appends instead of replacing.
        ["snapshot_date", "vehicle_name"], run_date_str,
    )
    return f"Summary={len(reports['vehicle_summary'])}, OOS={len(reports['vehicles_out_of_service'])}", path


def write_runs(reports, run_date_str, output_dir, append_dir, tag=""):
    path = os.path.join(output_dir, _out_name("Daily_Run_Volume", tag, run_date_str))
    _write_workbook(path, {
        "Runs by Cost Center": reports["runs_by_cost_center"],
        "Runs by Vehicle": reports["runs_by_vehicle"],
    })

    _append(
        append_dir, "Daily_Run_Volume_By_Cost_Center_APPEND.xlsx",
        "Runs by Cost Center", reports["runs_by_cost_center"],
        ["snapshot_date", "cost_center_name"], run_date_str,
    )
    _append(
        append_dir, "Daily_Run_Volume_By_Vehicle_APPEND.xlsx",
        "Runs by Vehicle", reports["runs_by_vehicle"],
        # Keyed on the name for the same reason as the vehicle overview above.
        ["snapshot_date", "vehicle_name"], run_date_str,
    )

    cc = reports["runs_by_cost_center"]
    veh = reports["runs_by_vehicle"]
    total = int(cc["total_runs"].sum()) if not cc.empty else 0
    return (
        f"{total} runs across {len(cc)} cost center(s) and {len(veh)} vehicle(s)",
        [path],
    )


def write_uhu(reports, run_date_str, output_dir, append_dir, tag=""):
    paths = []
    cc_path = os.path.join(output_dir, _out_name("Daily_UHU_By_Cost_Center", tag, run_date_str))
    _write_workbook(cc_path, {"UHU by Cost Center": reports["uhu_by_cost_center"]})
    paths.append(cc_path)

    sp_path = os.path.join(output_dir, _out_name("Daily_UHU_By_Shift_Profile", tag, run_date_str))
    _write_workbook(sp_path, {"UHU by Shift Profile": reports["uhu_by_shift_profile"]})
    paths.append(sp_path)

    _append(
        append_dir, "Daily_UHU_By_Cost_Center_APPEND.xlsx",
        "UHU by Cost Center", reports["uhu_by_cost_center"],
        ["snapshot_date", "cost_center_name"], run_date_str,
    )
    _append(
        append_dir, "Daily_UHU_By_Shift_Profile_APPEND.xlsx",
        "UHU by Shift Profile", reports["uhu_by_shift_profile"],
        ["snapshot_date", "shift_profile_name"], run_date_str,
    )
    return (
        f"ByCostCenter={len(reports['uhu_by_cost_center'])}, "
        f"ByShiftProfile={len(reports['uhu_by_shift_profile'])}",
        paths,
    )


# =============================
# MODULE COMPATIBILITY
# =============================
# The runner and the data layer are separate files, and on a machine where they
# are copied in by hand it is easy to update one and not the other. Left
# unchecked that surfaces as a KeyError partway through a run, minutes and a few
# hundred API calls in, naming a report rather than the stale file. Check up
# front instead, and say which file to replace.
REQUIRED_BUILDERS = (
    "build_runs_by_cost_center",
    "build_runs_by_vehicle",
    "shift_instances",
    "resolve_shift_offset",
    "parse_shift_ts",
    "Regions",
    "cost_centers_in",
    "filter_reports_to_region",
)


def check_modules():
    """Fail immediately, and legibly, when the data layer is out of date."""
    missing = [name for name in REQUIRED_BUILDERS if not hasattr(R, name)]
    if not missing:
        return True
    logging.error("%s is out of date and does not provide: %s",
                  getattr(R, "__file__", "traumasoft_reports.py"), ", ".join(missing))
    logging.error(
        "Replace that file with the current version. If you saved it from a "
        "browser it may have landed alongside it as 'traumasoft_reports (1).py' "
        "rather than overwriting it."
    )
    return False


# =============================
# MAIN
# =============================
def main():
    metrics_date = parse_cli_end_date(sys.argv) or (date.today() - timedelta(days=1))
    no_email = has_flag(sys.argv, "--no-email")
    dry_run = has_flag(sys.argv, "--dry-run")
    make_zip = has_flag(sys.argv, "--zip")
    region = parse_region(sys.argv)
    run_date_str = metrics_date.isoformat()

    output_dir = OUT.OUTPUT_DIR
    append_dir = OUT.APPEND_DIR
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    Path(append_dir).mkdir(parents=True, exist_ok=True)

    log_path = OUT.setup_logging(output_dir, run_date_str)
    logging.info("=== %s ===", JOB_NAME)
    logging.info("Metrics date: %s   test_mode=%s   dry_run=%s", run_date_str, TEST_MODE, dry_run)
    logging.info("Log: %s", log_path)

    # Resolved before any API work: a mistyped region should cost a second, not
    # a full fetch.
    regions = R.Regions()
    tag = ""
    if region is not None:
        canonical = regions.canonical(region)
        if canonical is None:
            logging.error(
                "No region named %r in %s. Defined: %s. Copy "
                "state/regions.example.json if that file does not exist yet.",
                region, regions.path, ", ".join(regions.names()) or "(none)",
            )
            return 2
        region = canonical
        tag = "_" + region.replace(" ", "_")
        # The company-wide run owns the append history and has already recorded
        # this day for every cost center, Indiana's included.
        append_dir = None
        logging.info(
            "Region: %s -- workbooks are filtered to it, and the append history "
            "is left to the company-wide run.", region,
        )

    if not check_modules():
        return 2

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

    unmatched = []
    region_unattributed = 0
    if region:
        unmatched = regions.unmatched(R.cost_centers_in(reports))
        reports = R.filter_reports_to_region(reports, regions, region)
        # The sheets built by grouping legs are rebuilt from this region's legs
        # rather than filtered afterwards, or a vehicle that worked two regions
        # today lands wholly in one with the other's runs attached. Uses the
        # map build_all just refreshed and saved.
        rebuilt, _, region_unattributed = R.build_region_leg_reports(
            data["legs"], data["vehicles"], R.CostCenterMap(), regions, region
        )
        reports.update(rebuilt)

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
            rows_info, path = writer(reports, run_date_str, output_dir, append_dir, tag)
            paths = path if isinstance(path, list) else [path]
            attachments.extend(paths)
            logging.info("--- OK: %s (%s)", name, rows_info)
            results.append((name, True, rows_info, ", ".join(paths)))
        except Exception as exc:  # one bad report must not lose the others
            logging.error("--- FAIL: %s: %s", name, exc, exc_info=True)
            results.append((name, False, "", str(exc)))

    # --- summary ---
    lines = [f"{JOB_NAME}", f"Metrics date: {run_date_str}"]
    if region:
        lines.append(f"Region: {region}")
    lines.append("")
    for name, ok, info, detail in results:
        lines.append(f"{'OK  ' if ok else 'FAIL'}  {name}: {info or detail}")
    lines.append("")

    if region:
        lines.append(
            "The Daily Vehicle Overview is fleet-wide, not regional: the API puts "
            "no cost center on a vehicle, so its rosters cannot be scoped. Run "
            "Volume by Vehicle is regional, because a leg is attributable through "
            "its shift profile."
        )
        if region_unattributed:
            lines.append(
                f"{region_unattributed} leg(s) belong to no cost center and so to no "
                "region -- mostly calls cancelled before reaching a unit. This is why "
                "the regional totals do not sum to the company-wide sheet."
            )
        if unmatched:
            lines.append(
                f"Note: {len(unmatched)} cost center(s) belong to no region and are "
                f"in no regional bundle: {', '.join(unmatched)}. Add them to "
                f"{regions.path}."
            )
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
            zip_path = bundle_zip(attachments, output_dir, run_date_str, tag)
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
    if region:
        subject = f"{region} Reports Bundle - {run_date_str}" + (" [TEST]" if TEST_MODE else "")
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
