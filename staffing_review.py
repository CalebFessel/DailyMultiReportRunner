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
ACTIVE_SHEET = "Active Now"

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


def assignments_in_window(append_dir, start, end):
    """
    (user_id, date) -> the units they were seen on, from the daily snapshots.

    Returns None when nothing has been recorded, which is the state before the
    daily runner has accumulated anything -- the caller says so rather than
    reporting a fleet nobody crewed.
    """
    df = OUT.read_append_sheet(os.path.join(append_dir, STAFFING_APPEND), ACTIVE_SHEET)
    if df is None or df.empty or "snapshot_date" not in df.columns:
        return None, None

    dates = pd.to_datetime(df["snapshot_date"], errors="coerce").dt.date
    keep = (dates >= start) & (dates <= end)
    window = df[keep].copy()
    if window.empty:
        return {}, []
    window["snapshot_date"] = dates[keep]

    seen = defaultdict(lambda: defaultdict(set))
    for _, row in window.iterrows():
        day = row["snapshot_date"]
        unit = row.get("shift_profile")
        for user_id in crew_ids(row.get("crew_members")):
            seen[str(user_id)][day].add(unit)
    return seen, sorted(set(window["snapshot_date"]))


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
            "hours_recorded": round(hours_by_user.get(user_id, 0.0), 2)
                              if hours_by_user else None,
        })

    df = pd.DataFrame(rows).sort_values(
        ["days_crewed", "last_name", "first_name"], ascending=[True, True, True]
    )
    never = df[df["days_crewed"] == 0]
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


def summary_sheet(args, roster_df, days_present, start, end, seen, never, hours_by_user):
    asked = (end - start).days + 1
    missing = asked - len(days_present)

    def row(item, value, note=""):
        return {"item": item, "value": value, "note": note}

    crewed = len(roster_df) - len(never) if not roster_df.empty else 0
    rows = [
        row("Cost center", args["cost_center"], "Matched against employee cost_center_name."),
        row("Levels included",
            "all" if args["all_levels"] else ", ".join(args["levels"]),
            "--list-levels prints every level this tenant uses."),
        row("Employees reviewed", len(roster_df),
            "Everyone matching, whether or not they were ever crewed."),
        row("Crewed at least once", crewed, ""),
        row("Never crewed in the window", len(never),
            "The finding this report exists for. Listed first in the Review sheet."),
        row("Window asked for", f"{start} to {end}", f"{asked} day(s)"),
        row("Days actually recorded", len(days_present),
            ("THE WINDOW IS COMPLETE." if missing == 0 else
             f"{missing} day(s) have no snapshot. Assignments come from the "
             "daily run's Staffing_Report_APPEND.xlsx, so a day exists only "
             "because the runner went that day. Retention is 730 days, so the "
             "history is as long as the runner has been going -- but a day it "
             "missed cannot be recovered.")),
        row("What a 'day crewed' means", "one look at the board",
            "The Active Now sheet is a single snapshot per run, not a "
            "timesheet. Someone crewed at 07:45 and gone by 09:00 counts the "
            "same as someone who worked the whole shift. This measures how "
            "often a person was seen on a unit, which is the staffing "
            "question -- it is not hours."),
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
        log.error("No staffing history found at %s.",
                  os.path.join(args["append_dir"], STAFFING_APPEND))
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
    log.info("  never crewed       : %s", len(never))
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
        log.warning("  %s of %s days have no snapshot. A person absent from the",
                    args["days"] - len(days_present), args["days"])
        log.warning("  report may simply have worked a day the runner did not cover.")

    sheets = {
        "Summary": summary_sheet(args, roster_df, days_present, start, end,
                                 seen, never, hours_by_user),
        "Review": result,
        "Never Crewed": result[result["days_crewed"] == 0],
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
