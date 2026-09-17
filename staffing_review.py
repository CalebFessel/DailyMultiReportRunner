"""
Staffing review: who is actually being crewed, and who is not.

    python staffing_review.py                              # Cincinnati, last 60 days
    python staffing_review.py --cost-center Toledo --days 30
    python staffing_review.py --all-levels                 # every position, not just the three
    python staffing_review.py --list-levels

Every employee in the cost center at the requested levels appears, including
those who have not been on a unit once. A name with zero days is the finding.

WHY THIS READS ASSIGNMENTS AND NOT PUNCHES
==========================================

Two reasons, and the second is the important one.

1. Punches cannot be backfilled. /Schedule/Shifts serves today-1..today+2 and
   ignores every date filter, so hours only exist for days a report has already
   run. See employee_hours_report.py, which accrues them a day at a time.

2. Punches would be the wrong measure anyway. Crews are paid from Paycor, not
   Traumasoft, so they have no reason to close a Traumasoft punch -- punch-outs
   are unreliable by construction, and hours built on them read low and vary
   with who remembered. For "is this person being put on a truck", the
   assignment is the cleaner signal: it comes from the schedule, not from crew
   discipline.

So this reads the daily Staffing_Report_APPEND.xlsx, which has recorded one row
per unit per run since the changeover, each carrying its crew's names and ids.
Retention is 730 days, so the history is as long as the runner has been going.

WHAT A DAY MEANS HERE
=====================

The Active Now sheet is ONE LOOK AT THE BOARD per run, not a timesheet. A
person crewed at 07:45 and gone by 09:00 counts the same as one who worked the
whole shift, and a person who worked a day the run did not cover does not
appear at all. It measures how often someone was seen on a unit, which is what
a staffing review asks, and is not a substitute for hours. The Summary sheet
says so, and names the days it actually covers.
"""

import os
import re
import sys
import logging
from pathlib import Path
from collections import defaultdict
from datetime import datetime, timedelta, date

import pandas as pd

import report_output as OUT
import employee_hours_report as H
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("staffing-review")

STAFFING_APPEND = "Staffing_Report_APPEND.xlsx"

# Both sheets the daily run records, and they are NOT equivalent.
#
#   Tomorrow   -- every unit whose shift STARTS on the day described. A whole
#                 day's schedule, independent of when the runner happened to
#                 fire. This is the one that answers "was this person put on a
#                 truck", and it is read first.
#
#   Active Now -- units on shift at the INSTANT the runner ran (start <= now <=
#                 end). It is a point-in-time sample, and a biased one: a crew
#                 whose shift did not span the run time never appears, so a
#                 morning run makes night crews look like they never worked.
#                 Read only to fill gaps, never on its own.
#
# Reading Active Now alone would put every off-cycle crew in the never-crewed
# list, which is the exact opposite of what this report is for.
SHEETS = ("Tomorrow", "Active Now")
ACTIVE_SHEET = "Active Now"
TOMORROW_SHEET = "Tomorrow"

# crew_members is built as "First Last (ID 1234)", newline separated. The id is
# what matters -- names are not unique and change with marriages and typos.
CREW_ID = re.compile(r"\(ID\s*(\d+)\s*\)")


def parse_args(argv):
    args = {
        "cost_center": H.DEFAULT_COST_CENTER,
        "days": H.DEFAULT_DAYS,
        "levels": list(H.DEFAULT_LEVELS),
        "all_levels": False,
        "out": "Reports",
        "append_dir": os.path.join("Reports", "Append"),
        "include_inactive": False,
        "list_levels": False,
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
            args["levels"] = [p.strip() for p in argv[i].split(",") if p.strip()]
        elif lowered == "--all-levels":
            args["all_levels"] = True
        elif lowered == "--out":
            i += 1
            args["out"] = argv[i]
        elif lowered == "--append-dir":
            i += 1
            args["append_dir"] = argv[i]
        elif lowered == "--include-inactive":
            args["include_inactive"] = True
        elif lowered == "--list-levels":
            args["list_levels"] = True
        else:
            raise SystemExit(f"Unknown option: {token}")
        i += 1
    return args


def crew_ids(cell):
    """Every user id named in one crew_members cell."""
    return CREW_ID.findall(str(cell or ""))


def work_dates(df):
    """
    The day each row's shift actually started.

    Not snapshot_date. The Tomorrow sheet is written on one day and describes
    the next, so filing its rows under the run date would shift every one of
    them by a day -- and would put a Sunday-night shift under Saturday. The
    shift's own start_time is the day the crew worked; snapshot_date only says
    when the report looked.
    """
    if "start_time" in df.columns:
        started = pd.to_datetime(df["start_time"], errors="coerce").dt.date
        if started.notna().any():
            return started
    return pd.to_datetime(df.get("snapshot_date"), errors="coerce").dt.date


def assignments_in_window(append_dir, start, end):
    """
    (user_id, date) -> the units they were crewed on, from the daily records.

    Both sheets are read and unioned. Tomorrow carries a whole day's schedule;
    Active Now is a point-in-time sample that misses any shift not spanning the
    run. Together they cover more days than either alone, and a person counts
    for a date if either sheet put them on a unit that started it.

    Returns (None, None) when the workbook holds neither sheet -- the state
    before the daily runner has ever run. That is a different answer from an
    empty window, and the caller reports it differently: one is a setup problem
    with a remedy, the other means nobody was crewed.
    """
    path = os.path.join(append_dir, STAFFING_APPEND)
    frames = []
    for sheet in SHEETS:
        df = OUT.read_append_sheet(path, sheet)
        if df is not None and not df.empty:
            df = df.copy()
            df["_sheet"] = sheet
            frames.append(df)
    if not frames:
        return None, None

    combined = pd.concat(frames, ignore_index=True, sort=False)
    dates = work_dates(combined)
    keep = dates.notna() & (dates >= start) & (dates <= end)
    window = combined[keep].copy()
    if window.empty:
        return {}, []
    window["work_date"] = dates[keep]

    seen = defaultdict(lambda: defaultdict(set))
    for _, row in window.iterrows():
        day = row["work_date"]
        unit = row.get("shift_profile")
        for user_id in crew_ids(row.get("crew_members")):
            seen[str(user_id)][day].add(unit)
    return seen, sorted(set(window["work_date"]))


def hired_after(hire_date, when):
    """
    True when someone started after `when`, so far as the record says.

    A new hire has no assignment history because they were not employed yet,
    not because they are not being used. Left unmarked they sit in the
    never-crewed list next to people who genuinely are not working, which is
    the one error in this report that would embarrass whoever presents it.
    """
    parsed = pd.to_datetime(hire_date, errors="coerce")
    if pd.isna(parsed):
        return False
    return parsed.date() > when


def review(roster_df, seen, hours_by_user, start, end):
    """One row per employee: how often they were crewed, zero included."""
    if roster_df.empty:
        return pd.DataFrame(), []

    rows = []
    for _, person in roster_df.iterrows():
        user_id = str(person["user_id"])
        days = seen.get(user_id, {}) if seen else {}
        units = {unit for day_units in days.values() for unit in day_units if unit}
        day_list = sorted(days)
        last_seen = day_list[-1] if day_list else None
        rows.append({
            "last_name": person.get("last_name"),
            "first_name": person.get("first_name"),
            "employee_num": person.get("employee_num"),
            "user_id": person.get("user_id"),
            "level": person.get("level"),
            "cost_center_name": person.get("cost_center_name"),
            "station": person.get("station"),
            "status": person.get("status"),
            "hire_date": person.get("hire_date"),
            "days_crewed": len(day_list),
            "units_crewed": len(units),
            "first_seen": day_list[0] if day_list else None,
            "last_seen": last_seen,
            "days_since_last_seen": (end - last_seen).days if last_seen else None,
            # Marked rather than filtered out: a new hire with no assignments
            # is still worth a manager's eye, just for a different reason.
            "hired_during_window": hired_after(person.get("hire_date"), start),
            "hours_recorded": round(hours_by_user.get(user_id, 0.0), 2)
                              if hours_by_user else None,
        })

    df = pd.DataFrame(rows).sort_values(
        ["days_crewed", "last_name", "first_name"], ascending=[True, True, True]
    )
    never = df[(df["days_crewed"] == 0) & (~df["hired_during_window"])]
    return df, list(zip(never["last_name"], never["first_name"]))


def recorded_hours(append_dir, start, end):
    """
    Hours per user from the employee-hours append, when it exists.

    Supplementary only. It is absent until employee_hours_report.py has been
    run, and even then covers only the days it recorded, so the review does not
    depend on it -- it just fills a column when there is something to fill it
    with.
    """
    df = H.window_rows(append_dir, start, end)
    if df is None or df.empty or "hours_worked" not in df.columns:
        return {}
    totals = df.groupby("user_id")["hours_worked"].sum()
    return {str(user_id): float(value) for user_id, value in totals.items()}


def summary_sheet(args, roster_df, days_present, start, end, seen, never,
                  hours_by_user, result=None):
    asked = (end - start).days + 1
    missing = asked - len(days_present)

    def row(item, value, note=""):
        return {"item": item, "value": value, "note": note}

    crewed = len(roster_df) - len(never) if not roster_df.empty else 0
    new_hires = int(result["hired_during_window"].sum()) if result is not None and not result.empty else 0
    rows = [
        row("Cost center", args["cost_center"], "Matched against employee cost_center_name."),
        row("Levels included",
            "all" if args["all_levels"] else ", ".join(args["levels"]),
            "--list-levels prints every level this tenant uses."),
        row("Employees reviewed", len(roster_df),
            "Everyone matching, whether or not they were ever crewed."),
        row("Crewed at least once", crewed, ""),
        row("Never crewed in the window", len(never),
            "The finding this report exists for. Listed first in the Review "
            "sheet. Excludes anyone hired during the window, who has no "
            "history because they were not employed yet."),
        row("Hired during the window", new_hires,
            "Flagged by hired_during_window and kept out of the never-crewed "
            "count. Still worth a look, for a different reason."),
        row("Window asked for", f"{start} to {end}", f"{asked} day(s)"),
        row("Days actually recorded", len(days_present),
            ("THE WINDOW IS COMPLETE." if missing == 0 else
             f"{missing} day(s) have no record at all. Assignments come from "
             "the daily run's Staffing_Report_APPEND.xlsx, so a day exists "
             "only because the runner went that day. Retention is 730 days, so "
             "the history is as long as the runner has been going -- but a day "
             "it missed cannot be recovered, and someone who worked only on "
             "missing days will appear here as never crewed.")),
        row("Read from", " + ".join(SHEETS),
            "Tomorrow carries a whole day's schedule, so it is read first. "
            "Active Now is a point-in-time sample -- units on shift at the "
            "instant the runner fired -- which on its own would make every "
            "crew whose shift missed that moment look like they never worked. "
            "Rows are filed under the day their shift STARTED, not the day the "
            "report ran."),
        row("What a 'day crewed' means", "assigned to a unit that day",
            "It counts days a person was on a unit's crew, not hours. Someone "
            "rostered and sent home early counts the same as someone who "
            "worked the full shift."),
        row("Hours column", "populated" if hours_by_user else "empty",
            "Hours come from employee_hours_report.py's append and cover only "
            "the days that report has run. Supplementary; the review does not "
            "depend on it."),
    ]
    if days_present:
        rows.append(row("Earliest snapshot", min(days_present), ""))
        rows.append(row("Latest snapshot", max(days_present), ""))
    return pd.DataFrame(rows)


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

    if args["list_levels"]:
        H.print_levels(employees)
        return 0

    levels = None if args["all_levels"] else args["levels"]
    roster_df = H.roster(employees, args["cost_center"], levels, args["include_inactive"])

    if roster_df.empty:
        log.error("No employees matched cost center '%s'%s.", args["cost_center"],
                  "" if args["all_levels"] else f" at levels {', '.join(args['levels'])}")
        log.error("Run --list-levels to see what this tenant actually uses.")
        return 1
    log.info("  %s employee(s) to review", len(roster_df))

    end = date.today() - timedelta(days=1)
    start = end - timedelta(days=args["days"] - 1)

    seen, days_present = assignments_in_window(args["append_dir"], start, end)
    if seen is None:
        log.error("")
        log.error("No staffing history found at %s (neither %s sheet).",
                  os.path.join(args["append_dir"], STAFFING_APPEND),
                  " nor ".join(SHEETS))
        log.error("That file is written by daily_report_runner_api.py. Until it "
                  "has run at least once there is nothing to review -- and no "
                  "query can produce it after the fact, because /Schedule/Shifts "
                  "only ever answers for today-1..today+2.")
        return 1

    hours_by_user = recorded_hours(args["append_dir"], start, end)
    result, never = review(roster_df, seen, hours_by_user, start, end)

    log.info("")
    log.info("--- %s staffing review, %s to %s ---", args["cost_center"], start, end)
    log.info("  employees reviewed : %s", len(result))
    log.info("  days recorded      : %s of %s asked for", len(days_present), args["days"])
    new_hires = int(result["hired_during_window"].sum())
    log.info("  never crewed       : %s%s", len(never),
             f"  (plus {new_hires} hired during the window, listed apart)"
             if new_hires else "")
    for last, first in never[:20]:
        log.info("      %s, %s", last, first)
    if len(never) > 20:
        log.info("      ... and %s more", len(never) - 20)

    stale = result[(result["days_crewed"] > 0) & (result["days_since_last_seen"] >= 14)]
    if not stale.empty:
        log.info("  not seen in 14+ days: %s", len(stale))
        for _, person in stale.head(10).iterrows():
            log.info("      %s, %s -- last on %s (%s days)",
                     person["last_name"], person["first_name"],
                     person["last_seen"], int(person["days_since_last_seen"]))

    if len(days_present) < args["days"]:
        log.warning("")
        log.warning("  %s of %s days have no record. A person listed as never",
                    args["days"] - len(days_present), args["days"])
        log.warning("  crewed may simply have worked only on days the runner")
        log.warning("  did not cover -- read the never-crewed list with that in mind.")

    sheets = {
        "Summary": summary_sheet(args, roster_df, days_present, start, end,
                                 seen, never, hours_by_user, result),
        "Review": result,
        "Never Crewed": result[(result["days_crewed"] == 0)
                               & (~result["hired_during_window"])],
        "New Hires": result[result["hired_during_window"]],
        "Roster": roster_df,
    }

    out_dir = Path(args["out"])
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = "".join(ch for ch in args["cost_center"] if ch.isalnum() or ch in "-_") or "All"
    out_path = out_dir / f"Staffing_Review_{safe}_{start}_to_{end}.xlsx"
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
