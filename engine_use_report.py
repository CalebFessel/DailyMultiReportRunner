#!/usr/bin/env python3
"""Engine use report from Samsara engine-state history.

Reconstruction of the one-off script that produced
engine_use_<start>_<end>.xlsx (see that workbook's About sheet for the
methodology, reproduced in the About sheet this script writes).

Usage:
    SAMSARA_API_TOKEN=... python engine_use_report.py 2026-07-01 2026-07-31
    # optional: -o /path/to/output.xlsx

Token scopes: Read Vehicles, Read Vehicle Statistics, Read Tags.
Dependencies: pandas, openpyxl (same as Script).
"""

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import date, datetime, time, timedelta, timezone
from zoneinfo import ZoneInfo

import pandas as pd

API_BASE = "https://api.samsara.com"
LOCAL_TZ = ZoneInfo("America/New_York")

RUNNING = {"On", "Idle"}
LIGHT_USE_OFF_HOURS = 8.0          # off this long in a row => "light use"
RUNNING_ALL_DAY_HOURS = 23.5       # flag threshold
HIGH_MILES = 600.0                 # flag threshold
TYPE_TAGS = ("BLS", "Wheelchair", "Secure Car")


# ---------------------------------------------------------------------------
# Samsara REST helpers (stdlib only)
# ---------------------------------------------------------------------------

def _request(token, path, params=None):
    url = API_BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(
        url, headers={"Authorization": f"Bearer {token}", "Accept": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())


def _paginate(token, path, params=None):
    params = dict(params or {})
    while True:
        page = _request(token, path, params)
        yield page
        pagination = page.get("pagination") or {}
        if not pagination.get("hasNextPage"):
            return
        params["after"] = pagination.get("endCursor")


def _parse_time(value):
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


# ---------------------------------------------------------------------------
# Org data: vehicles with State / Base / Type from tags
# ---------------------------------------------------------------------------

def load_vehicles(token):
    vehicles = {}
    for page in _paginate(token, "/fleet/vehicles"):
        for v in page.get("data", []):
            created = v.get("createdAtTime")
            vehicles[str(v["id"])] = {
                "name": v.get("name") or str(v["id"]),
                "created": _parse_time(created) if created else None,
                "state": "",
                "base": "",
                "type": "",
            }
    # State = a root tag that has children; Base = the vehicle's tag whose
    # parent is that state; Type = membership in one of TYPE_TAGS.
    tags = []
    for page in _paginate(token, "/tags"):
        tags.extend(page.get("data", []))
    by_id = {str(t["id"]): t for t in tags}
    parents_with_children = {str(t.get("parentTagId")) for t in tags if t.get("parentTagId")}
    for t in tags:
        name = (t.get("name") or "").strip()
        parent = by_id.get(str(t.get("parentTagId") or ""))
        members = [str(m.get("id")) for m in (t.get("vehicles") or []) + (t.get("assets") or [])]
        for vid in members:
            v = vehicles.get(vid)
            if not v:
                continue
            if name in TYPE_TAGS and not v["type"]:
                v["type"] = name
            if parent is not None and str(t["id"]) not in parents_with_children:
                # leaf tag with a parent -> base under its parent state
                v["base"] = name
                v["state"] = (parent.get("name") or "").strip()
    return vehicles


# ---------------------------------------------------------------------------
# Stats history: engine states + odometers
# ---------------------------------------------------------------------------

def load_history(token, start_utc, end_utc):
    """Per-vehicle engine-state timeline and odometer readings."""
    engine = defaultdict(list)     # vid -> [(dt, value)]
    odo = defaultdict(list)        # vid -> [(dt, miles, source)]
    params = {
        "types": "engineStates,obdOdometerMeters,gpsOdometerMeters",
        "startTime": start_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "endTime": end_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    for page in _paginate(token, "/fleet/vehicles/stats/history", params):
        for v in page.get("data", []):
            vid = str(v["id"])
            for e in v.get("engineStates") or []:
                engine[vid].append((_parse_time(e["time"]), e["value"]))
            for key, src in (("obdOdometerMeters", "obd"), ("gpsOdometerMeters", "gps")):
                for e in v.get(key) or []:
                    odo[vid].append((_parse_time(e["time"]), e["value"] / 1609.344, src))
    for series in engine.values():
        series.sort(key=lambda x: x[0])
    for series in odo.values():
        series.sort(key=lambda x: x[0])
    return engine, odo


def day_bounds(d):
    start = datetime.combine(d, time.min, tzinfo=LOCAL_TZ)
    return start, start + timedelta(days=1)


def segments_for_day(series, day_start, day_end):
    """Yield (seg_start, seg_end, state) covering [day_start, day_end).

    state is an engine state, or None before the first reading (no signal).
    The last known state carries forward, matching how Samsara holds a
    state until the next event.
    """
    state, idx = None, 0
    for idx, (t, value) in enumerate(series):
        if t > day_start:
            idx -= 1
            break
        state = value
    else:
        idx = len(series) - 1
    cursor = day_start
    for t, value in series[idx + 1 :] if idx >= 0 else series:
        if t <= day_start:
            state = value
            continue
        if t >= day_end:
            break
        if t > cursor:
            yield cursor, t, state
            cursor = t
        state = value
    if cursor < day_end:
        yield cursor, day_end, state


def miles_for_day(readings, day_start, day_end):
    for source in ("obd", "gps"):
        in_day = [m for t, m, s in readings if s == source and day_start <= t < day_end]
        if len(in_day) >= 2:
            return in_day[-1] - in_day[0]
    return None


def analyze_day(series, odo_readings, d, created):
    day_start, day_end = day_bounds(d)
    if created and created > day_end:
        return {"Status": "not in Samsara yet"}

    hours = {"On": 0.0, "Idle": 0.0, "Off": 0.0, None: 0.0}
    starts, first_start, last_shutoff = 0, None, None
    longest_off, off_run_start = 0.0, None
    running_at_midnight = False
    prev_state, saw_any = None, False

    for seg_start, seg_end, state in segments_for_day(series, day_start, day_end):
        dur = (seg_end - seg_start).total_seconds() / 3600.0
        hours[state if state in ("On", "Idle", "Off") else None] += dur
        if seg_start == day_start and state in RUNNING:
            running_at_midnight = True
        if state is not None:
            saw_any = True
        if state in RUNNING and prev_state == "Off":
            starts += 1
            if first_start is None:
                first_start = seg_start
        if state == "Off" and prev_state in RUNNING:
            last_shutoff = seg_start
        if state == "Off":
            if off_run_start is None:
                off_run_start = seg_start
        elif off_run_start is not None:
            longest_off = max(longest_off, (seg_start - off_run_start).total_seconds() / 3600.0)
            off_run_start = None
        prev_state = state
    if off_run_start is not None:
        longest_off = max(longest_off, (day_end - off_run_start).total_seconds() / 3600.0)

    engine_hours = hours["On"] + hours["Idle"]
    if not saw_any:
        status = "no signal"
    elif engine_hours == 0:
        status = "not used"
    elif longest_off >= LIGHT_USE_OFF_HOURS:
        status = "light use"
    else:
        status = "used"

    miles = miles_for_day(odo_readings, day_start, day_end)
    flags = []
    if engine_hours >= RUNNING_ALL_DAY_HOURS:
        flags.append("running all day")
    if miles is not None and miles > HIGH_MILES:
        flags.append("over 600 miles")

    return {
        "Status": status,
        "Engine starts": starts,
        "Running at midnight": "yes" if running_at_midnight else "",
        "First start": first_start.astimezone(LOCAL_TZ).replace(tzinfo=None) if first_start else None,
        "Last shutoff": last_shutoff.astimezone(LOCAL_TZ).replace(tzinfo=None) if last_shutoff else None,
        "Engine hours": engine_hours,
        "Driving hours": hours["On"],
        "Idle hours": hours["Idle"],
        "Off hours": hours["Off"],
        "No-signal hours": hours[None],
        "Longest off (h)": longest_off,
        "Miles": miles,
        "Flags": "; ".join(flags) or None,
    }


# ---------------------------------------------------------------------------
# Report assembly
# ---------------------------------------------------------------------------

def mean_time(values):
    secs = [v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6 for v in values]
    if not secs:
        return None
    avg = sum(secs) / len(secs)
    return (datetime.min + timedelta(seconds=avg)).time()


def build_report(token, start_day, end_day, out_path):
    start_utc = day_bounds(start_day)[0].astimezone(timezone.utc)
    end_utc = day_bounds(end_day)[1].astimezone(timezone.utc)
    days = [start_day + timedelta(days=i) for i in range((end_day - start_day).days + 1)]

    print("Loading vehicles and tags...")
    vehicles = load_vehicles(token)
    print(f"{len(vehicles)} vehicles. Loading stats history (this is the slow part)...")
    engine, odo = load_history(token, start_utc, end_utc)

    daily_rows = []
    for vid, v in sorted(vehicles.items(), key=lambda kv: kv[1]["name"]):
        for d in days:
            row = {"Vehicle": v["name"], "Date": datetime.combine(d, time.min),
                   "State": v["state"], "Base": v["base"], "Type": v["type"]}
            row.update(analyze_day(engine.get(vid, []), odo.get(vid, []), d, v["created"]))
            daily_rows.append(row)
    daily = pd.DataFrame(daily_rows)

    # --- By Vehicle ---------------------------------------------------
    by_vehicle = []
    for (name,), grp in daily.groupby(["Vehicle"]):
        meta = grp.iloc[0]
        used = int((grp["Status"] == "used").sum())
        light = int((grp["Status"] == "light use").sum())
        signal_days = used + light + int((grp["Status"] == "not used").sum())
        used_days = used or None
        starts = int(grp["Engine starts"].fillna(0).sum())
        eng = grp["Engine hours"].fillna(0).sum()
        drv = grp["Driving hours"].fillna(0).sum()
        idl = grp["Idle hours"].fillna(0).sum()
        miles = grp["Miles"].dropna().sum()
        flags = sorted({f for cell in grp["Flags"].dropna() for f in cell.split("; ")})
        firsts = [t.time() for t in grp["First start"].dropna()]
        by_vehicle.append({
            "Vehicle": name, "State": meta["State"], "Base": meta["Base"],
            "Type": meta["Type"],
            "Days with signal": signal_days, "Days used": used, "Days light use": light,
            "Days not used": int((grp["Status"] == "not used").sum()),
            "Days no signal": int((grp["Status"] == "no signal").sum()),
            "Engine starts": starts,
            "Starts per used day": starts / used_days if used_days else None,
            "Engine hours": eng, "Driving hours": drv, "Idle hours": idl,
            "Idle %": idl / eng if eng else None,
            "Engine hrs per used day": eng / used_days if used_days else None,
            "Miles": miles or None,
            "Miles per used day": (miles / used_days) if (miles and used_days) else None,
            "Avg first start": mean_time(firsts),
            "Flags": "; ".join(flags) or None,
        })
    by_vehicle = pd.DataFrame(by_vehicle)

    # --- Fleet by Day --------------------------------------------------
    fleet_rows = []
    for d, grp in daily.groupby("Date"):
        used = int((grp["Status"] == "used").sum())
        light = int((grp["Status"] == "light use").sum())
        notused = int((grp["Status"] == "not used").sum())
        signal = used + light + notused
        fleet_rows.append({
            "Date": d, "Weekday": d.strftime("%a"),
            "Vehicles with signal": signal, "Used": used, "Light use": light,
            "Not used": notused, "No signal": int((grp["Status"] == "no signal").sum()),
            "% of signalled used": (used + light) / signal if signal else None,
            "Engine starts": int(grp["Engine starts"].fillna(0).sum()),
            "Engine hours": grp["Engine hours"].fillna(0).sum(),
            "Driving hours": grp["Driving hours"].fillna(0).sum(),
            "Idle hours": grp["Idle hours"].fillna(0).sum(),
            "Miles": grp["Miles"].dropna().sum(),
        })
    fleet = pd.DataFrame(fleet_rows)

    # --- By Base --------------------------------------------------------
    base_rows = []
    for (state, base), grp in daily.groupby(["State", "Base"]):
        used = int((grp["Status"] == "used").sum())
        signal = used + int((grp["Status"] == "light use").sum()) + int((grp["Status"] == "not used").sum())
        eng = grp["Engine hours"].fillna(0).sum()
        idl = grp["Idle hours"].fillna(0).sum()
        base_rows.append({
            "State": state, "Base": base, "Vehicles": grp["Vehicle"].nunique(),
            "Vehicle-days with signal": signal, "Vehicle-days used": used,
            "Utilisation %": used / signal if signal else None,
            "Engine hours": eng, "Idle %": idl / eng if eng else None,
            "Miles": grp["Miles"].dropna().sum(),
        })
    by_base = pd.DataFrame(base_rows).sort_values(["State", "Base"])

    # --- About ----------------------------------------------------------
    about = pd.DataFrame([
        ("Engine use report", f"{start_day:%b %d} - {end_day:%b %d, %Y}"),
        ("Source", "Samsara /fleet/vehicles/stats/history (engineStates, obdOdometerMeters / "
                   f"gpsOdometerMeters). Generated {datetime.now(LOCAL_TZ):%Y-%m-%d %H:%M}"),
        ("Days", "Local calendar days in America/New_York."),
        ("Vehicles", f"{len(vehicles)} vehicles currently in Samsara. State/Base/Type come from "
                     "today's Samsara tags, not the report period's."),
        ("", ""),
        ("Engine start", "Engine went from Off to On or Idle. A truck already running at midnight "
                         "is not counted as a start that day ('Running at midnight')."),
        ("Driving hours", "Samsara 'On': engine running and the vehicle moving."),
        ("Idle hours", "Samsara 'Idle': engine running, vehicle stationary. EMS units idle for "
                       "climate and equipment, so idle counts as in use."),
        ("Engine hours", "Driving + idle."),
        ("Miles", "Odometer change over the day (OBD odometer, else GPS odometer). Blank = fewer "
                  "than two readings, not zero."),
        ("", ""),
        ("Status: used", "Engine ran and was never off 8+ hours in a row."),
        ("Status: light use", "Engine ran, but sat off 8+ hours in a row."),
        ("Status: not used", "Signal present, engine off all day."),
        ("Status: no signal", "Gateway reported no engine state at all. Unknown, NOT unused."),
        ("Status: not in Samsara yet", "Vehicle was added to Samsara after that day."),
        ("Flag: running all day", "23.5+ engine hours. Samsara holds the last state until the next "
                                  "event, so a gateway that stopped reporting while running reads "
                                  "as running. Verify."),
        ("Flag: over 600 miles", "More than 600 miles in a day. Confirm with dispatch if it matters."),
    ])

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        by_vehicle.to_excel(writer, sheet_name="By Vehicle", index=False)
        daily.drop(columns=["Type"]).to_excel(writer, sheet_name="Daily Detail", index=False)
        fleet.to_excel(writer, sheet_name="Fleet by Day", index=False)
        by_base.to_excel(writer, sheet_name="By Base", index=False)
        about.to_excel(writer, sheet_name="About", index=False, header=False)
    print(f"Wrote {out_path}")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("start", type=date.fromisoformat)
    ap.add_argument("end", type=date.fromisoformat)
    ap.add_argument("-o", "--output")
    args = ap.parse_args()
    token = os.environ.get("SAMSARA_API_TOKEN")
    if not token:
        sys.exit("Set SAMSARA_API_TOKEN")
    out = args.output or f"engine_use_{args.start}_{args.end}.xlsx"
    build_report(token, args.start, args.end, out)


if __name__ == "__main__":
    main()
