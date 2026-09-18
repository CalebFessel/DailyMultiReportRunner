"""
Hours worked per employee, for one cost center and a set of position levels.

    python employee_hours_report.py                          # Cincinnati, last 60 days
    python employee_hours_report.py --cost-center Toledo
    python employee_hours_report.py --days 30
    python employee_hours_report.py --list-levels            # what levels exist
    python employee_hours_report.py --list-cost-centers      # what cost centers exist

Every matching employee appears, including those with no hours at all. That is
the point of the report: a name with 0.00 hours is the finding, and a report
built by grouping punches would drop exactly those people.

READ THIS BEFORE TRUSTING A 60-DAY NUMBER
=========================================

Hours come from shift punches, and /Schedule/Shifts returns only
today-1..today+2 no matter what is asked of it -- verified byte-identical for
requests at -30, +0 and +30 days. There is no historical punch endpoint. So:

  * The ROSTER is complete and current. Every Cincinnati EMT-Driver, EMT-Non
    Driver and NEMT employee is listed today, zero-hour people included.

  * The HOURS only exist for days this report has already run. It appends what
    it can see to Employee_Hours_APPEND.xlsx, so the window fills in a day at a
    time from the first run. Ask for 60 days on day one and you get one day of
    hours and a Summary sheet saying so.

Nothing can backfill it. A day not recorded is gone, which is why this writes
its own append workbook rather than computing on demand.

Run it daily -- alongside daily_report_runner_api.py -- and the 60-day question
answers itself two months from now. The Summary sheet states the covered window
on every run so a partial answer is never mistaken for a full one.
"""

import os
import sys
import logging
from pathlib import Path
from collections import defaultdict
from datetime import datetime, timedelta, date

import pandas as pd

import report_output as OUT
import traumasoft_reports as R
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("employee-hours")

APPEND_FILE = "Employee_Hours_APPEND.xlsx"
DAILY_SHEET = "Employee Hours"

# Position levels to include, matched case-insensitively against the employee's
# `level` and `license_level`, ignoring spacing around a hyphen so
# "OH EMT - Driver", "OH EMT-Driver" and "OH EMT -  Driver" are one thing.
# Override with --levels.
#
# These are this tenant's actual names, taken from --list-levels rather than
# assumed. Levels are state-prefixed here -- OH, WV, IN, MD each have their own
# EMT and NEMT entries -- so the unprefixed "EMT - Driver" matches nobody, and
# a default that matched nothing would look like an empty roster rather than a
# configuration error.
#
# `level` holds the plain name and `license_level` carries a category prefix
# ("EMT - OH EMT - Driver"). Matching is on the whole value of either field, so
# the plain name in `level` is what lands, and "OH EMT - Driver" cannot sweep
# in a longer name that merely starts the same way.
DEFAULT_LEVELS = ["OH EMT - Driver", "OH EMT - Non Driver", "OH NEMT"]

DEFAULT_COST_CENTER = "Cincinnati"
DEFAULT_DAYS = 60


def normalize_level(value):
    """Collapse spacing and case so hyphenated level names compare equal."""
    text = " ".join(str(value or "").split()).lower()
    return text.replace(" - ", "-").replace(" -", "-").replace("- ", "-")


def parse_args(argv):
    args = {
        "cost_center": DEFAULT_COST_CENTER,
        "days": DEFAULT_DAYS,
        "levels": list(DEFAULT_LEVELS),
        "out": "Reports",
        "append_dir": os.path.join("Reports", "Append"),
        "list_levels": False,
        "list_cost_centers": False,
        "include_inactive": False,
        "no_append": False,
    }
    i = 0
    while i < len(argv):
        token = argv[i]
        lowered = token.lower()
        if lowered == "--cost-center":
            i += 1
            args["cost_center"] = argv[i]
        elif lowered == "--days":
            i += 1
            args["days"] = max(1, int(argv[i]))
        elif lowered == "--levels":
            i += 1
            args["levels"] = [part.strip() for part in argv[i].split(",") if part.strip()]
        elif lowered == "--out":
            i += 1
            args["out"] = argv[i]
        elif lowered == "--append-dir":
            i += 1
            args["append_dir"] = argv[i]
        elif lowered == "--list-levels":
            args["list_levels"] = True
        elif lowered == "--list-cost-centers":
            args["list_cost_centers"] = True
        elif lowered == "--include-inactive":
            args["include_inactive"] = True
        elif lowered == "--no-append":
            # For a look without recording it. Note this means the day is NOT
            # captured, and no later run can recover it.
            args["no_append"] = True
        else:
            raise SystemExit(f"Unknown option: {token}")
        i += 1
    return args


# =============================
# ROSTER
# =============================
def is_active(employee):
    """Employees currently on the books, by the several fields that say so."""
    if employee.get("disabled"):
        return False
    if employee.get("termination_date"):
        return False
    if str(employee.get("deactivated") or "").strip().lower() in ("1", "true", "yes"):
        return False
    status = str(employee.get("status") or "").strip().lower()
    return status not in ("inactive", "terminated", "disabled")


def employee_level(employee):
    """The position level, preferring `level` and falling back to license_level."""
    return employee.get("level") or employee.get("license_level") or ""


def matches_levels(employee, wanted):
    """
    True when either level field matches one of the requested positions.

    `wanted` of None means every level, which is how a caller asks for the
    whole cost center rather than a subset.
    """
    if wanted is None:
        return True
    normalized = {normalize_level(name) for name in wanted}
    return any(
        normalize_level(employee.get(field)) in normalized
        for field in ("level", "license_level")
    )


def matches_cost_center(employee, needle):
    return needle.strip().lower() in str(employee.get("cost_center_name") or "").lower()


def roster(employees, cost_center, levels, include_inactive):
    """
    Every employee the report is about, whether or not they worked.

    `levels` of None takes the whole cost center.
    """
    rows = []
    for employee in employees:
        if not matches_cost_center(employee, cost_center):
            continue
        if not matches_levels(employee, levels):
            continue
        active = is_active(employee)
        if not active and not include_inactive:
            continue
        rows.append({
            "user_id": employee.get("user_id"),
            "employee_num": employee.get("employee_num"),
            "last_name": employee.get("last_name"),
            "first_name": employee.get("first_name"),
            "level": employee_level(employee),
            "license_level": employee.get("license_level"),
            "cost_center_name": employee.get("cost_center_name"),
            "station": employee.get("station"),
            "hire_date": employee.get("hire_date"),
            "status": "Active" if active else "Inactive",
        })
    return pd.DataFrame(rows)


# =============================
# TODAY'S PUNCHES
# =============================
def hours_by_employee_day(shifts, shift_offset=None, now=None):
    """
    (user_id, date) -> hours on the clock, from the shift rows' punches.

    A shift row belongs to one crew member, so its punches are that person's.
    Punches are bucketed by the date the punch STARTED, which keeps an
    overnight shift's hours on the day the crew came on -- the same convention
    the unit-shift reports use.

    An open punch is bounded at the shift's end or at `now`, whichever comes
    first. Bounding only at the shift end credits hours nobody has worked yet:
    an overnight crew read at 06:00 would be paid to their 08:00 finish.
    """
    # parse_shift_ts returns tenant-local NAIVE time, so `now` has to be on
    # that same clock -- tenant_now(shift_offset) gives it. An aware datetime
    # would otherwise blow up mid-loop with a bare TypeError from a comparison
    # several frames down, which says nothing about what the caller got wrong.
    if now is not None and now.tzinfo is not None:
        raise ValueError(
            "now must be tenant-local naive time, as traumasoft_reports."
            "tenant_now(shift_offset) returns; got an aware datetime. Punch "
            "times are parsed naive, so an aware `now` cannot be compared."
        )

    totals = defaultdict(float)
    open_punches = defaultdict(int)
    still_running = defaultdict(int)

    for shift in shifts:
        if shift.get("deleted"):
            continue
        user_id = shift.get("user_id")
        if user_id is None:
            continue
        shift_end = R.parse_shift_ts(shift.get("end_time"), shift_offset)

        for punch in shift.get("punches") or []:
            if punch.get("deleted"):
                continue
            start = R.parse_shift_ts(punch.get("start_time"), shift_offset)
            end = R.parse_shift_ts(punch.get("end_time"), shift_offset)
            if not start:
                continue
            if not end:
                open_punches[user_id] += 1
                end = shift_end
                if end is None:
                    continue
                if now is not None and now < end:
                    end = now
                    still_running[user_id] += 1

            if end <= start:
                continue
            totals[(user_id, start.date())] += (end - start).total_seconds() / 3600.0

    return totals, open_punches, still_running


def todays_rows(roster_df, totals, open_punches, still_running):
    """One row per employee per date they were on the clock."""
    known = set(roster_df["user_id"]) if not roster_df.empty else set()
    by_user = roster_df.set_index("user_id").to_dict("index") if not roster_df.empty else {}

    rows = []
    for (user_id, work_date), hours in sorted(totals.items(), key=lambda kv: (str(kv[0][1]), kv[0][0])):
        if user_id not in known:
            continue
        person = by_user[user_id]
        rows.append({
            # ISO string, not a date object. Excel reads a written date back as
            # a Timestamp, which never equals a fresh date() -- so the de-dupe
            # key would miss and a re-run would append a second copy of the day
            # instead of replacing it. The rest of this project writes dates as
            # strings for the same reason.
            "work_date": work_date.isoformat(),
            "user_id": user_id,
            "employee_num": person.get("employee_num"),
            "last_name": person.get("last_name"),
            "first_name": person.get("first_name"),
            "level": person.get("level"),
            "cost_center_name": person.get("cost_center_name"),
            "hours_worked": round(hours, 2),
            "open_punches": open_punches.get(user_id, 0),
            "still_on_shift": still_running.get(user_id, 0),
        })
    return pd.DataFrame(rows)


# =============================
# WINDOW, FROM THE APPEND
# =============================
def window_rows(append_dir, start, end):
    """Recorded daily rows inside the window, or None if nothing is recorded."""
    df = OUT.read_append_sheet(os.path.join(append_dir, APPEND_FILE), DAILY_SHEET)
    if df is None or df.empty or "work_date" not in df.columns:
        return None
    dates = pd.to_datetime(df["work_date"], errors="coerce").dt.date
    keep = (dates >= start) & (dates <= end)
    out = df[keep].copy()
    out["work_date"] = dates[keep]
    return out


def summarize(roster_df, recorded, start, end):
    """
    One row per employee: hours over the window, zero included.

    The join is a left join onto the roster on purpose. Grouping the recorded
    rows alone would answer a different question -- who worked -- and silently
    omit every person this report exists to surface.
    """
    if roster_df.empty:
        return pd.DataFrame(), []

    if recorded is None or recorded.empty:
        totals = pd.DataFrame(columns=["user_id", "hours_worked", "days_worked"])
    else:
        totals = (
            recorded.groupby("user_id", dropna=False)
            .agg(
                hours_worked=("hours_worked", "sum"),
                days_worked=("work_date", "nunique"),
                last_worked=("work_date", "max"),
            )
            .reset_index()
        )

    merged = roster_df.merge(totals, on="user_id", how="left")
    merged["hours_worked"] = merged["hours_worked"].fillna(0.0).round(2)
    merged["days_worked"] = merged["days_worked"].fillna(0).astype(int)
    if "last_worked" not in merged.columns:
        merged["last_worked"] = pd.NaT
    merged["window_start"] = start
    merged["window_end"] = end

    columns = [
        "last_name", "first_name", "employee_num", "user_id", "level",
        "cost_center_name", "station", "status", "hire_date",
        "hours_worked", "days_worked", "last_worked", "window_start", "window_end",
    ]
    merged = merged[[c for c in columns if c in merged.columns]]
    merged = merged.sort_values(["hours_worked", "last_name", "first_name"],
                                ascending=[True, True, True])
    zero = merged[merged["hours_worked"] == 0]
    return merged, list(zip(zero["last_name"], zero["first_name"]))


def summary_sheet(args, roster_df, recorded, start, end, days_present, today_rows):
    """What the numbers cover, and what no run can ever recover."""
    asked = (end - start).days + 1
    missing = asked - len(days_present)

    def row(item, value, note=""):
        return {"item": item, "value": value, "note": note}

    total_hours = round(float(recorded["hours_worked"].sum()), 2) if recorded is not None and not recorded.empty else 0.0

    rows = [
        row("Cost center", args["cost_center"], "Matched against employee cost_center_name."),
        row("Levels included", ", ".join(args["levels"]),
            "Matched case-insensitively against `level` and `license_level`, "
            "ignoring spacing around the hyphen. --list-levels prints every "
            "level this tenant actually uses."),
        row("Employees listed", len(roster_df),
            "Everyone matching, including those with no hours. A zero is a "
            "finding, not a missing row."),
        row("Window asked for", f"{start} to {end}", f"{asked} day(s)"),
        row("Days of hours actually recorded", len(days_present),
            f"{missing} day(s) missing. " + (
                "THE WINDOW IS COMPLETE." if missing == 0 else
                "Hours come from shift punches, and /Schedule/Shifts serves only "
                "today-1..today+2 and ignores every date filter. A day exists in "
                "this report only because the report ran that day. Nothing can "
                "backfill the rest -- run it daily and the window fills in."
            )),
        row("Hours recorded in window", total_hours, "Sum across every listed employee."),
        row("Rows captured this run", len(today_rows) if today_rows is not None else 0,
            "Employee-days visible in the current shift window and written to "
            f"{APPEND_FILE}."),
        row("Append workbook", "written" if not args["no_append"] else "NOT written (--no-append)",
            os.path.join(args["append_dir"], APPEND_FILE) if not args["no_append"]
            else "This run recorded nothing; today's hours are not retrievable later."),
    ]
    if days_present:
        rows.append(row("Earliest day recorded", min(days_present), ""))
        rows.append(row("Latest day recorded", max(days_present), ""))
    return pd.DataFrame(rows)


# =============================
# DISCOVERY
# =============================
def print_levels(employees):
    counts = defaultdict(int)
    for employee in employees:
        for field in ("level", "license_level"):
            value = str(employee.get(field) or "").strip()
            if value:
                counts[(field, value)] += 1
    log.info("Levels in use on this tenant:")
    for (field, value), count in sorted(counts.items(), key=lambda kv: -kv[1]):
        log.info("  %5s  %-14s %s", count, field, value)


def print_cost_centers(employees):
    counts = defaultdict(int)
    for employee in employees:
        value = str(employee.get("cost_center_name") or "").strip() or "(none)"
        counts[value] += 1
    log.info("Cost centers on employee records:")
    for value, count in sorted(counts.items(), key=lambda kv: -kv[1]):
        log.info("  %5s  %s", count, value)


def main():
    args = parse_args(sys.argv[1:])

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("Configuration error: %s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Could not authenticate. Check TS_API_KEY / TS_API_SECRET.")
        return 1

    log.info("Fetching employees ...")
    try:
        employees = api.list_employees(include_disabled=True)
    except TraumasoftAPIError as exc:
        log.error("Employees failed: %s", exc)
        return 1
    log.info("  %s employee record(s)", len(employees))

    if args["list_levels"]:
        print_levels(employees)
        return 0
    if args["list_cost_centers"]:
        print_cost_centers(employees)
        return 0

    roster_df = roster(employees, args["cost_center"], args["levels"], args["include_inactive"])
    if roster_df.empty:
        log.error("No employees matched cost center '%s' at levels %s.",
                  args["cost_center"], ", ".join(args["levels"]))
        log.error("Run --list-cost-centers and --list-levels to see what this "
                  "tenant actually uses; the names must match.")
        return 1
    log.info("  %s employee(s) in %s at the requested levels",
             len(roster_df), args["cost_center"])

    log.info("Fetching shifts (rolling window -- today-1..today+2) ...")
    try:
        shifts = api.list_shifts()
    except TraumasoftAPIError as exc:
        log.error("Shifts failed: %s", exc)
        return 1

    # Shift times are UTC while trips are tenant-local; resolve the gap from a
    # day of trips so punches land on the right date and "now" is comparable.
    shift_offset = None
    try:
        legs = api.get_trips(date.today() - timedelta(days=1), range_days=1)
        shift_offset = R.resolve_shift_offset(legs)
    except TraumasoftAPIError as exc:
        log.warning("Could not read trips to resolve the shift clock: %s", exc)
    if shift_offset:
        log.info("  shift times shifted by %s to reach tenant-local time", shift_offset)
    now = R.tenant_now(shift_offset)

    totals, open_punches, still_running = hours_by_employee_day(shifts, shift_offset, now)
    today_rows = todays_rows(roster_df, totals, open_punches, still_running)
    log.info("  %s employee-day row(s) visible in the current window", len(today_rows))

    if not today_rows.empty and not args["no_append"]:
        OUT._append_to_workbook_xlsx(
            os.path.join(args["append_dir"], APPEND_FILE), DAILY_SHEET, today_rows,
            dedupe_keys=["work_date", "user_id"],
        )
        log.info("  recorded to %s", os.path.join(args["append_dir"], APPEND_FILE))
    elif args["no_append"]:
        log.warning("  --no-append: today's hours were NOT recorded and cannot "
                    "be recovered by a later run.")

    end = date.today() - timedelta(days=1)
    start = end - timedelta(days=args["days"] - 1)
    recorded = window_rows(args["append_dir"], start, end)
    days_present = sorted(set(recorded["work_date"])) if recorded is not None and not recorded.empty else []

    summary, zero_hour = summarize(roster_df, recorded, start, end)

    log.info("")
    log.info("--- %s, %s to %s ---", args["cost_center"], start, end)
    log.info("  employees listed : %s", len(summary))
    log.info("  days of hours    : %s of %s asked for", len(days_present), args["days"])
    log.info("  zero hours       : %s employee(s)", len(zero_hour))
    for last, first in zero_hour[:15]:
        log.info("      %s, %s", last, first)
    if len(zero_hour) > 15:
        log.info("      ... and %s more", len(zero_hour) - 15)

    if len(days_present) < args["days"]:
        log.warning("")
        log.warning("  The window is NOT complete. Hours exist only for days this")
        log.warning("  report has run: /Schedule/Shifts serves today-1..today+2 and")
        log.warning("  ignores date filters, so the rest cannot be backfilled by")
        log.warning("  anything. Run this daily and the window fills in.")

    sheets = {
        "Summary": summary_sheet(args, roster_df, recorded, start, end, days_present, today_rows),
        "Employees": summary,
        "Daily": recorded if recorded is not None else pd.DataFrame(),
        "Roster": roster_df,
    }

    out_dir = Path(args["out"])
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = "".join(ch for ch in args["cost_center"] if ch.isalnum() or ch in "-_") or "AllCostCenters"
    out_path = out_dir / f"Employee_Hours_{safe}_{start}_to_{end}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            OUT.write_df_sheet_with_table(
                writer, df if df is not None else pd.DataFrame(), name, name.replace(" ", "_")
            )

    log.info("")
    log.info("Wrote %s", out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
