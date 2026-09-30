"""
Measure how good the Traumasoft punch record actually is.

Crews currently clock in and out in two places -- Traumasoft and Paycor -- and
only Paycor decides what they are paid. Nothing rewards a complete Traumasoft
punch, so punch-outs there are unreliable by construction. Every unit-hour
number this repo produces rests on those punches, which makes their quality a
reporting problem and not just an HR one.

`unit_punches_by_instance` bounds a punch with no end at the shift's end (or at
now, whichever is sooner). That was the right call for a crew still on the
road, but where a crew simply never clocked out it credits the whole scheduled
shift -- so `worked_hours` quietly collapses back into `scheduled_hours` for
that unit, and the gap the whole worked-hours design exists to expose closes
itself. This probe puts a number on that: how many unit hours in the UHU
denominator are measured, and how many are manufactured by the fallback.

It is also the before-and-after evidence for moving pay onto Traumasoft. The
same figures that argue for the change are the ones that show it worked.

**This is a four-day window, not a history.** /Schedule/Shifts returns
today-1..today+2 and ignores every date filter -- verified byte-identical for
requests at -30, +0 and +30 days -- so punches cannot be backfilled. A baseline
accrues by running this daily and keeping the output; a day not captured is
gone. --json writes a machine-readable row per run for exactly that purpose.

Read-only. Every call is a GET.

Usage:
    python probe_punch_quality.py                     # the window as it stands
    python probe_punch_quality.py --json punches.json # append a daily baseline
    python probe_punch_quality.py --employees 40      # longer offender list
"""

import os
import sys
import json
import hashlib
import logging
import argparse
from collections import Counter, defaultdict
from datetime import datetime, timedelta

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError
import traumasoft_reports as R

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("probe-punch-quality")

# A punch shorter than this is almost certainly a mis-click -- clock in, clock
# straight back out -- rather than work. Counted apart so it neither flatters
# the completion rate nor lands in the worked-hours total unremarked.
SUSPECT_PUNCH_MINUTES = float(os.getenv("PUNCH_SUSPECT_MINUTES", "5"))

# How long after a shift ends a still-open punch stops being "they are running
# late" and becomes "nobody is ever closing this".
STALE_PUNCH_HOURS = float(os.getenv("PUNCH_STALE_HOURS", "2"))


def build_id():
    """A short hash of the code actually running."""
    digest = hashlib.sha256()
    for path in (__file__, R.__file__):
        try:
            with open(path, "rb") as handle:
                digest.update(handle.read())
        except OSError:
            return "unknown"
    return digest.hexdigest()[:8]


def hours(delta):
    return delta.total_seconds() / 3600.0


def pct(part, whole):
    return (100.0 * part / whole) if whole else 0.0


def employee_index(employees):
    """user_id -> the employee record, for attributing punches to a person."""
    index = {}
    for emp in employees:
        uid = emp.get("user_id")
        if uid is not None:
            index[str(uid)] = emp
    return index


def employee_label(emp, user_id):
    if not emp:
        return f"user_id {user_id}"
    name = " ".join(
        part for part in (emp.get("first_name"), emp.get("last_name")) if part
    ).strip()
    num = emp.get("employee_num")
    if name and num:
        return f"{name} (#{num})"
    return name or (f"#{num}" if num else f"user_id {user_id}")


# =============================
# COLLECTION
# =============================
def collect_punches(shifts, offset, now):
    """
    One flat row per punch, with everything needed to judge it.

    Deliberately not routed through unit_punches_by_instance: that function
    clips punches to their unit-shift and merges crew rows, which is right for
    unit hours and wrong here. A punch is one person's clock event, and the
    question is whether that person closed it.
    """
    rows = []
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = R.profile_name(shift)
        if not name:
            continue
        shift_start = R.parse_shift_ts(shift.get("start_time"), offset)
        shift_end = R.parse_shift_ts(shift.get("end_time"), offset)
        user_id = shift.get("user_id")
        for punch in shift.get("punches") or []:
            if punch.get("deleted"):
                continue
            start = R.parse_shift_ts(punch.get("start_time"), offset)
            if not start:
                continue
            end = R.parse_shift_ts(punch.get("end_time"), offset)
            shift_over = bool(shift_end and now and now > shift_end)
            stale = bool(
                shift_end and now
                and now > shift_end + timedelta(hours=STALE_PUNCH_HOURS)
            )
            rows.append({
                "profile": name,
                "user_id": user_id,
                "vehicle_name": shift.get("vehicle_name"),
                "license_level": shift.get("license_level"),
                "shift_start": shift_start,
                "shift_end": shift_end,
                "start": start,
                "end": end,
                "open": end is None,
                # An open punch on a shift that has finished is a missed
                # punch-out. On a shift still running it is just a crew who
                # have not gone home yet, and means nothing at all.
                "missed_punch_out": end is None and shift_over,
                "stale": end is None and stale,
                "minutes": (end - start).total_seconds() / 60.0 if end else None,
            })
    return rows


def rostered_without_punches(shifts, offset):
    """
    Crew rows carrying no punch at all -- someone rostered who never clocked in.

    Distinct from a missed punch-out and worse: a missed punch-out still proves
    the person turned up. A row with no punches proves nothing either way, and
    the fallback cannot rescue it because there is no punch to bound.
    """
    seen = defaultdict(lambda: {"rows": 0, "with_punches": 0, "profiles": set()})
    for shift in shifts:
        if shift.get("deleted"):
            continue
        name = R.profile_name(shift)
        if not name:
            continue
        end = R.parse_shift_ts(shift.get("end_time"), offset)
        # Only judge shifts that have had their chance to be punched.
        entry = seen[shift.get("user_id")]
        entry["rows"] += 1
        entry["profiles"].add(name)
        if any(not p.get("deleted") for p in shift.get("punches") or []):
            entry["with_punches"] += 1
        entry.setdefault("last_end", end)
        if end and (entry.get("last_end") is None or end > entry["last_end"]):
            entry["last_end"] = end
    return seen


# =============================
# SECTIONS
# =============================
def report_window(shifts, offset, now):
    print("\n" + "=" * 78)
    print("1. WHAT THE FEED RETURNED")
    print("=" * 78)

    dates = Counter()
    for shift in shifts:
        if shift.get("deleted"):
            continue
        start = R.parse_shift_ts(shift.get("start_time"), offset)
        if start:
            dates[start.date()] += 1

    print(f"  Tenant now             : {now:%Y-%m-%d %H:%M} (local)")
    print(f"  Shift offset applied   : {hours(offset):+.1f}h from UTC")
    print(f"  Shift rows             : {len(shifts)}")
    live = [s for s in shifts if not s.get("deleted")]
    print(f"  Not deleted            : {len(live)}")
    with_punches = sum(
        1 for s in live if any(not p.get("deleted") for p in s.get("punches") or [])
    )
    print(f"  Carrying any punch     : {with_punches} ({pct(with_punches, len(live)):.1f}%)")
    print("\n  Dates present (the window cannot be steered):")
    for day in sorted(dates):
        print(f"    {day}  {dates[day]:>5} crew rows")
    if len(dates) <= 1:
        print("    -- only one day came back; the usual window is four.")
    return sorted(dates)


def report_completion(rows, now):
    print("\n" + "=" * 78)
    print("2. ARE PUNCHES BEING CLOSED?")
    print("=" * 78)

    total = len(rows)
    if not total:
        print("  No punches at all in the window. Nothing further can be measured.")
        return {}

    closed = [r for r in rows if not r["open"]]
    open_rows = [r for r in rows if r["open"]]
    missed = [r for r in open_rows if r["missed_punch_out"]]
    stale = [r for r in open_rows if r["stale"]]
    running = [r for r in open_rows if not r["missed_punch_out"]]
    suspect = [r for r in closed if (r["minutes"] or 0) < SUSPECT_PUNCH_MINUTES]

    print(f"  Punches in window      : {total}")
    print(f"  Closed                 : {len(closed)} ({pct(len(closed), total):.1f}%)")
    print(f"  Open                   : {len(open_rows)} ({pct(len(open_rows), total):.1f}%)")
    print(f"    shift still running  : {len(running)}   <- not a problem")
    print(f"    MISSED PUNCH-OUT     : {len(missed)}   <- shift is over, punch is not closed")
    print(f"      of those, stale    : {len(stale)}   (>{STALE_PUNCH_HOURS:g}h past shift end)")
    print(f"  Suspiciously short     : {len(suspect)} closed punches under "
          f"{SUSPECT_PUNCH_MINUTES:g} min")

    finished = len(closed) + len(missed)
    if finished:
        rate = pct(len(closed), finished)
        print(f"\n  Punch-out completion   : {rate:.1f}% "
              f"({len(closed)} of {finished} punches on finished shifts)")
        print("  This is the number to watch. If pay moves onto Traumasoft it")
        print("  should climb toward 100%; that is how you prove the change took.")
    else:
        print("\n  No finished shifts in the window, so completion cannot be scored yet.")

    return {
        "punches": total,
        "closed": len(closed),
        "open": len(open_rows),
        "still_running": len(running),
        "missed_punch_out": len(missed),
        "stale": len(stale),
        "suspect_short": len(suspect),
        "completion_pct": round(pct(len(closed), finished), 2) if finished else None,
    }


def day_is_finished(shifts, day, offset, now):
    """
    True when every shift starting on `day` has already ended.

    A day still in progress cannot be judged on punch discipline: its open
    punches are crews on the road, not missed punch-outs, and the fallback
    bounds them at `now` exactly as it should.
    """
    latest = None
    for shift in shifts:
        if shift.get("deleted"):
            continue
        start = R.parse_shift_ts(shift.get("start_time"), offset)
        end = R.parse_shift_ts(shift.get("end_time"), offset)
        if not start or not end or start.date() != day:
            continue
        if latest is None or end > latest:
            latest = end
    return latest is not None and latest <= now


def report_fabricated_hours(shifts, offset, now, dates):
    """
    How much of the UHU denominator is measured, and how much is invented.

    Computed by running the real worked-hours path twice: once as it ships, and
    once over a copy of the feed with every open punch stripped out. The
    difference is exactly what the shift-end fallback contributed -- not an
    estimate of it, the thing itself.

    **The headline counts finished days only.** An in-progress day is nearly
    all fallback by construction -- every crew currently on the road has an
    open punch, correctly bounded at now -- so pooling it with finished days
    produced a figure that was six times the real one and described nothing
    anybody publishes. The daily runner reports on yesterday, a finished day,
    so a finished day is what the number has to describe. In-progress days are
    still shown, marked, and left out of the total.
    """
    print("\n" + "=" * 78)
    print("3. HOW MUCH OF THE UHU DENOMINATOR IS REAL?")
    print("=" * 78)

    rules = R.UnitStaffingRules()

    def worked_total(feed, day):
        by_profile = R.unit_worked_hours(feed, day, offset, rules, now)
        return sum(by_profile.values())

    # A copy with open punches removed, so the fallback has nothing to bound.
    stripped = []
    for shift in shifts:
        clone = dict(shift)
        clone["punches"] = [
            p for p in shift.get("punches") or []
            if p.get("end_time")
        ]
        stripped.append(clone)

    print(f"  {'date':<12} {'as shipped':>12} {'punched only':>14} "
          f"{'fabricated':>12} {'share':>8}  state")
    print("  " + "-" * 72)

    totals = {"shipped": 0.0, "measured": 0.0}
    per_day = {}
    finished_days = []
    for day in dates:
        shipped = worked_total(shifts, day)
        measured = worked_total(stripped, day)
        gap = shipped - measured
        finished = day_is_finished(shifts, day, offset, now)
        if finished:
            totals["shipped"] += shipped
            totals["measured"] += measured
            finished_days.append(day)
        per_day[day.isoformat()] = {
            "worked_hours_as_shipped": round(shipped, 2),
            "worked_hours_measured": round(measured, 2),
            "fabricated_hours": round(gap, 2),
            "day_finished": finished,
        }
        print(f"  {day.isoformat():<12} {shipped:>12.2f} {measured:>14.2f} "
              f"{gap:>12.2f} {pct(gap, shipped):>7.1f}%  "
              f"{'finished' if finished else 'IN PROGRESS -- excluded'}")

    gap = totals["shipped"] - totals["measured"]
    print("  " + "-" * 72)
    print(f"  {'finished':<12} {totals['shipped']:>12.2f} {totals['measured']:>14.2f} "
          f"{gap:>12.2f} {pct(gap, totals['shipped']):>7.1f}%")

    if not finished_days:
        print("\n  No finished day in the window, so there is nothing to judge yet.")
        print("  Every open punch belongs to a crew still on the road. Re-run")
        print("  tomorrow, when today has ended.")
        return {"per_day": per_day, "finished_days": [], "fabricated_pct": None}

    print(f"\n  {pct(gap, totals['shipped']):.1f}% of the unit hours in the UHU")
    print("  denominator come from bounding an unclosed punch at the shift end,")
    print("  not from a crew clocking out. For those units worked_hours IS")
    print("  scheduled_hours, and utilization reads low by exactly that much.")
    print(f"\n  Finished day(s) only: {', '.join(d.isoformat() for d in finished_days)}.")
    print("  A day still running is nearly all fallback by construction -- every")
    print("  crew currently out has an open punch -- so counting it would say")
    print("  nothing about punch discipline. The daily runner reports on")
    print("  yesterday, which is why a finished day is the one that matters.")
    if gap > 0 and totals["shipped"]:
        implied = pct(totals["measured"], totals["shipped"])
        print(f"\n  Put the other way: {implied:.1f}% of the denominator is evidence.")
    return {"per_day": per_day,
            "finished_days": [d.isoformat() for d in finished_days],
            "worked_hours_as_shipped": round(totals["shipped"], 2),
            "worked_hours_measured": round(totals["measured"], 2),
            "fabricated_hours": round(gap, 2),
            "fabricated_pct": round(pct(gap, totals["shipped"]), 2)}


def report_by_employee(rows, employees, limit):
    print("\n" + "=" * 78)
    print("4. WHO IS NOT CLOSING PUNCHES")
    print("=" * 78)

    index = employee_index(employees)
    per_user = defaultdict(lambda: {"closed": 0, "missed": 0, "running": 0})
    for row in rows:
        entry = per_user[row["user_id"]]
        if not row["open"]:
            entry["closed"] += 1
        elif row["missed_punch_out"]:
            entry["missed"] += 1
        else:
            entry["running"] += 1

    scored = []
    for user_id, entry in per_user.items():
        finished = entry["closed"] + entry["missed"]
        if not finished:
            continue
        scored.append((entry["missed"], pct(entry["closed"], finished), user_id, entry))
    # Worst first: most missed, then lowest completion.
    scored.sort(key=lambda s: (-s[0], s[1]))

    offenders = [s for s in scored if s[0] > 0]
    if not offenders:
        print("  Every finished punch in the window was closed. Nothing to chase.")
        return []

    print(f"  {len(offenders)} of {len(scored)} crew with a finished shift left a "
          f"punch open.\n")
    print(f"  {'missed':>7} {'closed':>7} {'rate':>7}  employee")
    print("  " + "-" * 60)
    listed = []
    for missed, rate, user_id, entry in offenders[:limit]:
        emp = index.get(str(user_id))
        label = employee_label(emp, user_id)
        print(f"  {missed:>7} {entry['closed']:>7} {rate:>6.0f}%  {label}")
        listed.append({
            "user_id": user_id,
            "employee_num": (emp or {}).get("employee_num"),
            "cost_center": (emp or {}).get("cost_center_name"),
            "missed": missed,
            "closed": entry["closed"],
            "completion_pct": round(rate, 1),
        })
    if len(offenders) > limit:
        print(f"  ... and {len(offenders) - limit} more (--employees to see them)")

    print("\n  A short list means a training conversation. A long one means the")
    print("  system is the problem, not the people -- which is the argument for")
    print("  making Traumasoft the clock that pays.")
    return listed


def report_by_unit(rows):
    print("\n" + "=" * 78)
    print("5. WHERE IT CONCENTRATES")
    print("=" * 78)

    per_profile = defaultdict(lambda: {"closed": 0, "missed": 0})
    for row in rows:
        if row["open"] and not row["missed_punch_out"]:
            continue
        entry = per_profile[row["profile"]]
        entry["missed" if row["open"] else "closed"] += 1

    scored = []
    for name, entry in per_profile.items():
        finished = entry["closed"] + entry["missed"]
        if finished:
            scored.append((entry["missed"], pct(entry["closed"], finished), name, entry))
    scored.sort(key=lambda s: (-s[0], s[1]))

    worst = [s for s in scored if s[0] > 0][:15]
    if not worst:
        print("  No profile has a missed punch-out in this window.")
        return
    print(f"  {'missed':>7} {'closed':>7} {'rate':>7}  profile")
    print("  " + "-" * 60)
    for missed, rate, name, entry in worst:
        print(f"  {missed:>7} {entry['closed']:>7} {rate:>6.0f}%  {name}")


def report_never_punched(shifts, offset, employees, now):
    print("\n" + "=" * 78)
    print("6. ROSTERED BUT NEVER CLOCKED IN")
    print("=" * 78)

    index = employee_index(employees)
    seen = rostered_without_punches(shifts, offset)
    never = []
    for user_id, entry in seen.items():
        if entry["with_punches"]:
            continue
        last_end = entry.get("last_end")
        # Only count someone whose shift has actually finished; a crew rostered
        # for tonight has not failed to clock in yet.
        if last_end and now and now > last_end:
            never.append((user_id, entry))

    if not never:
        print("  Everyone with a finished shift in the window punched at least once.")
        return []

    print(f"  {len(never)} crew were rostered on a shift that has ended and")
    print("  recorded no punch at all. The fallback cannot help these -- there")
    print("  is no punch to bound, so they contribute zero worked hours and")
    print("  their unit reads as never crewed.\n")
    listed = []
    for user_id, entry in sorted(never, key=lambda n: -n[1]["rows"])[:20]:
        emp = index.get(str(user_id))
        profiles = ", ".join(sorted(entry["profiles"]))[:44]
        print(f"  {entry['rows']:>3} row(s)  {employee_label(emp, user_id):<34} {profiles}")
        listed.append({
            "user_id": user_id,
            "employee_num": (emp or {}).get("employee_num"),
            "rows": entry["rows"],
        })
    if len(never) > 20:
        print(f"  ... and {len(never) - 20} more")
    return listed


def report_paycor_readiness(shifts, employees):
    """
    Can a Traumasoft punch even be addressed to a Paycor employee?

    Whatever the push ends up looking like, it has to name the person on the
    Paycor side. `employee_num` is the only field in the Traumasoft employee
    record that plausibly carries a shared payroll identifier, so its coverage
    is a hard gate on the whole integration -- worth knowing now rather than
    after the credentials arrive.
    """
    print("\n" + "=" * 78)
    print("7. PAYCOR JOIN READINESS")
    print("=" * 78)

    punching_users = set()
    for shift in shifts:
        if shift.get("deleted"):
            continue
        if any(not p.get("deleted") for p in shift.get("punches") or []):
            punching_users.add(str(shift.get("user_id")))

    index = employee_index(employees)
    with_num, without_num, unknown = [], [], []
    for uid in punching_users:
        emp = index.get(uid)
        if emp is None:
            unknown.append(uid)
        elif str(emp.get("employee_num") or "").strip():
            with_num.append(uid)
        else:
            without_num.append(uid)

    total = len(punching_users)
    print(f"  Crew who punched in window : {total}")
    print(f"  Carrying employee_num      : {len(with_num)} ({pct(len(with_num), total):.1f}%)")
    print(f"  Missing employee_num       : {len(without_num)}")
    print(f"  Not in the employee roster : {len(unknown)}")

    dupes = Counter()
    for emp in employees:
        num = str(emp.get("employee_num") or "").strip()
        if num:
            dupes[num] += 1
    collisions = {num: n for num, n in dupes.items() if n > 1}
    print(f"  Duplicate employee_num     : {len(collisions)}")
    if collisions:
        print("    -- a duplicate cannot address a punch unambiguously; these")
        print("       need an override entry or a fix in Traumasoft:")
        for num, count in list(collisions.items())[:10]:
            print(f"       {num!r} used by {count} employees")

    if total and len(with_num) == total and not collisions:
        print("\n  Clean. Every punching crew member can be named on the Paycor side,")
        print("  ASSUMING Paycor keys on the same number -- which is unverified until")
        print("  you have credentials and can read a Paycor employee back.")
    else:
        print("\n  Not clean. The push cannot address the gaps above; they need")
        print("  state/paycor_employee_overrides.json or a fix in Traumasoft.")

    return {
        "punching_crew": total,
        "with_employee_num": len(with_num),
        "missing_employee_num": len(without_num),
        "not_in_roster": len(unknown),
        "duplicate_employee_num": len(collisions),
    }


def report_next_steps(summary):
    print("\n" + "=" * 78)
    print("8. WHAT THIS MEANS")
    print("=" * 78)
    fab = summary.get("hours", {}).get("fabricated_pct")
    comp = summary.get("completion", {}).get("completion_pct")
    if comp is not None:
        print(f"  * Punch-out completion is {comp:.1f}%. Record it every day --")
        print("    the window cannot be backfilled, so today's number only exists")
        print("    if today's run captured it.")
    if fab is not None:
        finished = summary.get("hours", {}).get("finished_days") or []
        print(f"  * {fab:.1f}% of the UHU denominator is the shift-end fallback")
        print("    rather than a measured punch, on the finished day(s) in the")
        print(f"    window ({', '.join(finished)}). That is the figure that")
        print("    describes what the daily runner publishes, because it reports")
        print("    on yesterday. A day still in progress is nearly all fallback")
        print("    by construction and says nothing about punch discipline.")
    else:
        print("  * No finished day in this window, so the UHU denominator cannot")
        print("    be judged yet. Re-run once today has ended.")
    print("  * Run this daily with --json to accumulate a baseline. Four days is")
    print("    all the API will ever show you at once.")
    print("  * worked_hours still rests on punches nobody is paid from, which is")
    print("    what the Paycor push is meant to change. Whether that matters is")
    print("    now an empirical question rather than an assumption -- read the")
    print("    punch-out completion figure in section 2 before arguing from it.")


# =============================
# MAIN
# =============================
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json", help="append a machine-readable summary row here")
    parser.add_argument("--employees", type=int, default=20,
                        help="how many crew to list in section 4 (default 20)")
    args = parser.parse_args()

    print("=" * 78)
    print("TRAUMASOFT PUNCH QUALITY")
    print(f"build {build_id()}   run {datetime.now():%Y-%m-%d %H:%M:%S}")
    print("=" * 78)
    print("Read-only. Every call below is a GET.")

    api = TraumasoftAPI()
    try:
        api.detect_auth_mode()
    except TraumasoftAPIError as exc:
        log.error("Could not authenticate: %s", exc)
        return 1

    log.info("Fetching shifts, employees and a day of trips...")
    shifts = api.list_shifts()
    employees = api.list_employees()
    # Trips are pulled only to recover the tenant's UTC offset, which the shift
    # feed does not carry and every comparison below depends on.
    legs = api.get_trips(datetime.now().date(), range_days=1)
    offset = R.resolve_shift_offset(legs)
    now = R.tenant_now(offset)

    dates = report_window(shifts, offset, now)
    rows = collect_punches(shifts, offset, now)
    completion = report_completion(rows, now)
    hours_summary = report_fabricated_hours(shifts, offset, now, dates) if dates else {}
    offenders = report_by_employee(rows, employees, args.employees)
    report_by_unit(rows)
    never = report_never_punched(shifts, offset, employees, now)
    readiness = report_paycor_readiness(shifts, employees)

    summary = {
        "build": build_id(),
        "captured_at": datetime.now().isoformat(timespec="seconds"),
        "tenant_now": now.isoformat(timespec="seconds"),
        "dates": [d.isoformat() for d in dates],
        "completion": completion,
        "hours": hours_summary,
        "paycor_readiness": readiness,
        "offenders": offenders,
        "never_punched": never,
    }
    report_next_steps(summary)

    if args.json:
        # Append rather than overwrite: the whole point is accumulating days
        # the API will not hand back a second time.
        existing = []
        if os.path.exists(args.json):
            try:
                with open(args.json, encoding="utf-8") as handle:
                    existing = json.load(handle)
            except (OSError, ValueError):
                log.warning("Could not read %s, starting a fresh baseline.", args.json)
            if not isinstance(existing, list):
                existing = [existing]
        existing.append(summary)
        with open(args.json, "w", encoding="utf-8") as handle:
            json.dump(existing, handle, indent=2, default=str)
        print(f"\nBaseline appended to {args.json} ({len(existing)} run(s) recorded).")

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
