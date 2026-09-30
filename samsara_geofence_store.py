#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Samsara geofence history store.

Pulls GPS samples from /fleet/vehicles/stats/history, keeps only those Samsara
tagged with an address (i.e. inside one of our geofences), and stores them in
SQLite so the history survives Samsara's ~450-day history window (edge measured 2026-08-27 between
2025-05-24 and 2025-06-13).

Two products:
  geofence_ping  raw matched samples - the durable record, cannot be re-fetched
                 once Samsara drops them
  visit          derived arrivals with entry/exit/dwell - rebuilt from pings at
                 any time, so the dwell rules stay changeable

Backfill runs OLDEST-FIRST. If the window is rolling retention the far end
expires daily, so it is fetched first; if the edge is instead the account start
date nothing is expiring and the order simply costs nothing. Re-run `horizon`
in a week - a moving edge means retention, a fixed one means account start.

Resumability is per vehicle-day. An interrupted multi-hour run picks up exactly
where it stopped rather than restarting.

Usage
-----
  # find the history edge (read-only)
  python samsara_geofence_store.py horizon

  # small test before committing to the full run
  python samsara_geofence_store.py backfill --start 2026-08-20 --end 2026-08-23

  # the real thing - oldest first, resumable, safe to re-run
  python samsara_geofence_store.py backfill --days 460

  # catch up since the last run (for the scheduler)
  python samsara_geofence_store.py incremental

  # (re)build visits from stored pings
  python samsara_geofence_store.py visits --gap-minutes 10

  python samsara_geofence_store.py stats
"""

import os
import sys
import time
import json
import sqlite3
import argparse
import datetime as dt
from urllib.parse import urlencode

import requests

BASE = os.environ.get("SAMSARA_BASE_URL", "https://api.samsara.com")
DEFAULT_DB = os.environ.get("SAMSARA_GEOFENCE_DB", r"C:\Reports\samsara_geofence.db")
MIN_INTERVAL = float(os.environ.get("SAMSARA_MIN_INTERVAL", "0.22"))


# --------------------------------------------------------------------------- auth

def utcnow():
    """Naive UTC now. datetime.utcnow() is deprecated on 3.12+."""
    return dt.datetime.now(dt.timezone.utc).replace(tzinfo=None)


def load_token():
    tok = os.environ.get("SAMSARA_API_TOKEN")
    if tok:
        return tok
    # fall back to a .env alongside the existing Reports integrations
    for folder in (r"C:\Reports", os.getcwd(), os.path.dirname(os.path.abspath(__file__))):
        for name in (".env", ".env.txt"):
            path = os.path.join(folder, name)
            if not os.path.isfile(path):
                continue
            try:
                with open(path, encoding="utf-8-sig") as fh:
                    for raw in fh:
                        line = raw.strip()
                        if line.startswith("export "):
                            line = line[7:].strip()
                        if "=" not in line or line.startswith("#"):
                            continue
                        k, _, v = line.partition("=")
                        if k.strip() == "SAMSARA_API_TOKEN":
                            v = v.strip().strip("'\"")
                            if v:
                                return v
            except OSError:
                pass
    raise SystemExit(
        "SAMSARA_API_TOKEN not set.\n"
        "  PowerShell: $env:SAMSARA_API_TOKEN = '<token>'\n"
        "  or add SAMSARA_API_TOKEN=<token> to C:\\Reports\\.env"
    )


class Api(object):
    def __init__(self, token):
        self.s = requests.Session()
        self.s.headers.update({"Authorization": "Bearer " + token})
        self._last = 0.0
        self.calls = 0

    def get(self, path, params):
        wait = MIN_INTERVAL - (time.time() - self._last)
        if wait > 0:
            time.sleep(wait)
        url = BASE + path + "?" + urlencode(params)
        for attempt in range(1, 7):
            self._last = time.time()
            self.calls += 1
            try:
                r = self.s.get(url, timeout=90)
            except requests.RequestException as exc:
                if attempt == 6:
                    raise
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 200:
                return r.json()
            if r.status_code == 401:
                raise SystemExit("401 - token missing the 'Read Vehicle Statistics' scope.")
            if r.status_code in (429,) or r.status_code >= 500:
                if attempt == 6:
                    r.raise_for_status()
                time.sleep(2 ** attempt)
                continue
            raise RuntimeError("HTTP %s on %s :: %s" % (r.status_code, path, r.text[:300]))
        raise RuntimeError("retries exhausted: " + path)


# --------------------------------------------------------------------------- db
SCHEMA = """
CREATE TABLE IF NOT EXISTS geofence_ping (
    vehicle_id   TEXT NOT NULL,
    ts           TEXT NOT NULL,
    vehicle_name TEXT,
    address_id   TEXT,
    address_name TEXT,
    lat          REAL,
    lon          REAL,
    speed_mph    REAL,
    heading      INTEGER,
    PRIMARY KEY (vehicle_id, ts)
);
CREATE INDEX IF NOT EXISTS ix_ping_addr_ts ON geofence_ping(address_id, ts);
CREATE INDEX IF NOT EXISTS ix_ping_veh_ts  ON geofence_ping(vehicle_id, ts);

CREATE TABLE IF NOT EXISTS ingest_log (
    vehicle_id  TEXT NOT NULL,
    day         TEXT NOT NULL,
    status      TEXT NOT NULL,
    samples     INTEGER,
    matched     INTEGER,
    fetched_at  TEXT,
    PRIMARY KEY (vehicle_id, day)
);

CREATE TABLE IF NOT EXISTS visit (
    visit_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    vehicle_id   TEXT,
    vehicle_name TEXT,
    address_id   TEXT,
    address_name TEXT,
    entry_ts     TEXT,
    exit_ts      TEXT,
    dwell_sec    INTEGER,
    ping_count   INTEGER,
    qualified    INTEGER  -- 1 = real visit, 0 = drive-through of the geofence
);
CREATE INDEX IF NOT EXISTS ix_visit_addr  ON visit(address_id, entry_ts);
CREATE INDEX IF NOT EXISTS ix_visit_veh   ON visit(vehicle_id, entry_ts);

CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
"""


def open_db(path):
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    con = sqlite3.connect(path)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.executescript(SCHEMA)
    cols = [r[1] for r in con.execute("PRAGMA table_info(visit)")]
    if "qualified" not in cols:
        con.execute("ALTER TABLE visit ADD COLUMN qualified INTEGER")
    # index depends on the column above, so it is created after the migration
    con.execute("CREATE INDEX IF NOT EXISTS ix_visit_qual ON visit(qualified, address_id)")
    con.commit()
    return con


# --------------------------------------------------------------------------- helpers
def iso(d):
    return d.strftime("%Y-%m-%dT%H:%M:%SZ")


def list_vehicles(api):
    out, after = [], None
    while True:
        p = {"limit": 512}
        if after:
            p["after"] = after
        j = api.get("/fleet/vehicles", p)
        out.extend(j.get("data") or [])
        pg = j.get("pagination") or {}
        if not pg.get("hasNextPage"):
            break
        after = pg.get("endCursor")
    return [(str(v["id"]), v.get("name") or str(v["id"])) for v in out]


def fetch_day(api, vehicle_ids, day):
    """One UTC day of gps samples for a batch of vehicles. Returns {vid: [samples]}."""
    start = iso(day)
    end = iso(day + dt.timedelta(days=1))
    got, after = {}, None
    while True:
        p = {"types": "gps", "startTime": start, "endTime": end,
             "vehicleIds": ",".join(vehicle_ids)}
        if after:
            p["after"] = after
        j = api.get("/fleet/vehicles/stats/history", p)
        for v in (j.get("data") or []):
            vid = str(v.get("id"))
            got.setdefault(vid, []).extend(v.get("gps") or [])
        pg = j.get("pagination") or {}
        if not pg.get("hasNextPage"):
            break
        after = pg.get("endCursor")
    return got


def store(con, vid, vname, samples):
    rows, matched = [], 0
    for g in samples:
        addr = g.get("address") or {}
        aid = addr.get("id")
        if not aid:
            continue  # outside every geofence - not geofence data, dropped
        matched += 1
        rows.append((vid, g.get("time"), vname, str(aid), addr.get("name"),
                     g.get("latitude"), g.get("longitude"),
                     g.get("speedMilesPerHour"), g.get("headingDegrees")))
    if rows:
        con.executemany(
            "INSERT OR IGNORE INTO geofence_ping "
            "(vehicle_id, ts, vehicle_name, address_id, address_name, lat, lon, speed_mph, heading) "
            "VALUES (?,?,?,?,?,?,?,?,?)", rows)
    return matched


# --------------------------------------------------------------------------- commands
def cmd_horizon(args):
    api = Api(load_token())
    veh = list_vehicles(api)
    if not veh:
        raise SystemExit("no vehicles returned")
    vid = veh[0][0]
    now = utcnow()
    print("probing retention edge with vehicle %s\n" % veh[0][1])
    lo, hi = 0, None
    for d in (330, 365, 380, 400, 420, 440, 460, 480, 500, 520, 545):
        a = now - dt.timedelta(days=d)
        j = api.get("/fleet/vehicles/stats/history",
                    {"types": "gps", "startTime": iso(a),
                     "endTime": iso(a + dt.timedelta(hours=6)), "vehicleIds": vid})
        n = sum(len(v.get("gps") or []) for v in (j.get("data") or []))
        print("  %4d days back (%s) : %5d samples" % (d, a.strftime("%Y-%m-%d"), n))
        if n > 0:
            lo = d
        elif hi is None:
            hi = d
    print("\nretention edge is between %d and %s days." % (lo, hi if hi else "?"))
    print("Backfill oldest-first: the far end expires daily.")


def cmd_backfill(args):
    api = Api(load_token())
    con = open_db(args.db)
    if getattr(args, "reset", False):
        # geofence matching is retroactive, so rows fetched under the old radii
        # are stale once geofences change. Clear and re-fetch.
        n = con.execute("SELECT COUNT(*) FROM geofence_ping").fetchone()[0]
        d = con.execute("SELECT COUNT(*) FROM ingest_log WHERE status='ok'").fetchone()[0]
        rng = con.execute("SELECT MIN(ts), MAX(ts) FROM geofence_ping").fetchone()
        # A --reset on a populated store throws away hours of irreplaceable
        # fetching. Never do it on a bare flag.
        if n > 0 and not getattr(args, "confirm_reset", False):
            print("REFUSING --reset: the store already holds data.")
            print("  pings        : {:,}".format(n))
            print("  vehicle-days : {:,}".format(d))
            print("  range        : %s .. %s" % (rng[0], rng[1]))
            print("")
            print("  Samsara's history window is finite - anything aged out cannot be")
            print("  re-fetched. If you really mean to discard the above, re-run with")
            print("  --confirm-reset. To add newer data instead, use 'incremental'.")
            raise SystemExit(2)
        if n > 0:
            con.execute("DELETE FROM geofence_ping")
            con.execute("DELETE FROM ingest_log")
            con.execute("DELETE FROM visit")
            con.commit()
            con.execute("VACUUM")
            print("--reset: cleared %d pings, ingest log and visits" % n)
    veh = list_vehicles(api)
    if args.vehicles:
        want = set(x.strip().upper() for x in args.vehicles.split(","))
        veh = [v for v in veh if v[1].upper() in want]
    print("vehicles: %d" % len(veh))

    today = utcnow().date()
    if args.start:
        d0 = dt.datetime.strptime(args.start, "%Y-%m-%d").date()
    else:
        d0 = today - dt.timedelta(days=args.days)
    d1 = dt.datetime.strptime(args.end, "%Y-%m-%d").date() if args.end else today

    days = []
    d = d0
    while d < d1:
        days.append(d)
        d += dt.timedelta(days=1)
    days.sort()  # OLDEST FIRST - the expiring end
    print("range %s .. %s  (%d days, oldest first)" % (d0, d1, len(days)))

    done = set()
    for r in con.execute("SELECT vehicle_id, day FROM ingest_log WHERE status='ok'"):
        done.add((r[0], r[1]))

    todo = [(v, day) for day in days for v in veh if (v[0], day.isoformat()) not in done]
    print("vehicle-days to fetch: %d  (already done: %d)\n" % (len(todo), len(done)))
    if not todo:
        print("nothing to do.")
        return

    batch = max(1, args.batch)
    t0 = time.time()
    total_matched = 0
    i = 0
    by_day = {}
    for v, day in todo:
        by_day.setdefault(day, []).append(v)

    for day in sorted(by_day.keys()):
        vlist = by_day[day]
        for k in range(0, len(vlist), batch):
            chunk = vlist[k:k + batch]
            ids = [c[0] for c in chunk]
            names = dict(chunk)
            try:
                got = fetch_day(api, ids, dt.datetime.combine(day, dt.time.min))
            except Exception as exc:
                for vid in ids:
                    con.execute("INSERT OR REPLACE INTO ingest_log VALUES (?,?,?,?,?,?)",
                                (vid, day.isoformat(), "error", 0, 0, utcnow().isoformat()))
                con.commit()
                print("  %s %s -> ERROR %s" % (day, ",".join(names[i2] for i2 in ids), exc))
                continue
            for vid in ids:
                samples = got.get(vid, [])
                matched = store(con, vid, names[vid], samples)
                total_matched += matched
                con.execute("INSERT OR REPLACE INTO ingest_log VALUES (?,?,?,?,?,?)",
                            (vid, day.isoformat(), "ok", len(samples), matched,
                             utcnow().isoformat()))
            con.commit()
            i += len(chunk)
            if i % (batch * 10) < batch:
                el = time.time() - t0
                rate = i / el if el else 0
                left = (len(todo) - i) / rate if rate else 0
                print("  %s  %d/%d vehicle-days  %d pings stored  %.0f/min  eta %.0f min"
                      % (day, i, len(todo), total_matched, rate * 60, left / 60))

    con.execute("INSERT OR REPLACE INTO meta VALUES ('last_backfill', ?)",
                (utcnow().isoformat(),))
    con.commit()
    print("\ndone. %d vehicle-days, %d pings stored, %d api calls, %.1f min"
          % (i, total_matched, api.calls, (time.time() - t0) / 60))


def cmd_incremental(args):
    con = open_db(args.db)
    row = con.execute("SELECT MAX(day) FROM ingest_log WHERE status='ok'").fetchone()
    last = row[0] if row and row[0] else None
    if not last:
        raise SystemExit("no prior ingest - run backfill first.")
    start = (dt.datetime.strptime(last, "%Y-%m-%d").date())
    args.start = start.isoformat()
    args.end = None
    print("incremental from %s" % args.start)
    cmd_backfill(args)


def cmd_visits(args):
    con = open_db(args.db)
    gap = args.gap_minutes * 60
    min_dwell = args.min_dwell_minutes * 60
    print("rebuilding visits (gap %d min, min dwell %d min)..." % (args.gap_minutes, args.min_dwell_minutes))
    con.execute("DELETE FROM visit")
    cur = con.execute(
        "SELECT vehicle_id, vehicle_name, address_id, address_name, ts "
        "FROM geofence_ping ORDER BY vehicle_id, ts")
    out = []
    cv = ca = None
    entry = last = None
    vname = aname = None
    count = 0

    def flush():
        if cv is None or entry is None:
            return
        e0 = dt.datetime.strptime(entry[:19], "%Y-%m-%dT%H:%M:%S")
        e1 = dt.datetime.strptime(last[:19], "%Y-%m-%dT%H:%M:%S")
        secs = int((e1 - e0).total_seconds())
        out.append((cv, vname, ca, aname, entry, last, secs, count,
                    1 if secs >= min_dwell else 0))

    for vid, vn, aid, an, ts in cur:
        if cv == vid and ca == aid and last is not None:
            t_prev = dt.datetime.strptime(last[:19], "%Y-%m-%dT%H:%M:%S")
            t_now = dt.datetime.strptime(ts[:19], "%Y-%m-%dT%H:%M:%S")
            if (t_now - t_prev).total_seconds() <= gap:
                last = ts
                count += 1
                continue
        flush()
        cv, ca, vname, aname = vid, aid, vn, an
        entry = last = ts
        count = 1
    flush()

    con.executemany(
        "INSERT INTO visit (vehicle_id, vehicle_name, address_id, address_name, "
        "entry_ts, exit_ts, dwell_sec, ping_count, qualified) "
        "VALUES (?,?,?,?,?,?,?,?,?)", out)
    con.commit()
    q = con.execute("SELECT COUNT(*) FROM visit WHERE qualified=1").fetchone()[0]
    d = con.execute("SELECT COUNT(*) FROM visit WHERE qualified=0").fetchone()[0]
    md = con.execute("SELECT AVG(dwell_sec)/60.0 FROM visit WHERE qualified=1").fetchone()[0]
    print("built %d segments: %d qualified visits, %d drive-throughs (kept, flagged)"
          % (len(out), q, d))
    print("  mean dwell on qualified visits: %.1f min" % (md or 0))


def cmd_stats(args):
    con = open_db(args.db)
    q = lambda s: con.execute(s).fetchone()
    print("db: %s  (%.1f MB)" % (args.db, os.path.getsize(args.db) / 1e6))
    print("pings          : %d" % q("SELECT COUNT(*) FROM geofence_ping")[0])
    print("visits         : %d" % q("SELECT COUNT(*) FROM visit")[0])
    r = q("SELECT MIN(ts), MAX(ts) FROM geofence_ping")
    print("ping range     : %s .. %s" % (r[0], r[1]))
    print("vehicles       : %d" % q("SELECT COUNT(DISTINCT vehicle_id) FROM geofence_ping")[0])
    print("addresses seen : %d" % q("SELECT COUNT(DISTINCT address_id) FROM geofence_ping")[0])
    ok = q("SELECT COUNT(*) FROM ingest_log WHERE status='ok'")[0]
    er = q("SELECT COUNT(*) FROM ingest_log WHERE status='error'")[0]
    print("ingest_log     : %d ok, %d error" % (ok, er))
    tot = q("SELECT COUNT(*) FROM visit")[0]
    qual = q("SELECT COUNT(*) FROM visit WHERE qualified=1")[0]
    print("qualified      : %d of %d segments (%d drive-throughs)" % (qual, tot, tot - qual))
    print("\ntop addresses by QUALIFIED visit count:")
    for r in con.execute("SELECT address_name, COUNT(*) c, AVG(dwell_sec)/60.0 d "
                         "FROM visit WHERE qualified=1 GROUP BY address_id "
                         "ORDER BY c DESC LIMIT 12"):
        print("  %-44s %5d visits  %6.1f min avg" % ((r[0] or "?")[:44], r[1], r[2] or 0))
    print("\nmost drive-throughs (geofence likely too wide):")
    for r in con.execute("SELECT address_name, COUNT(*) c FROM visit WHERE qualified=0 "
                         "GROUP BY address_id ORDER BY c DESC LIMIT 8"):
        print("  %-44s %5d" % ((r[0] or "?")[:44], r[1]))


ROLE_SCHEMA = """
CREATE TABLE IF NOT EXISTS site_role (
    address_id   TEXT NOT NULL,
    address_name TEXT,
    role         TEXT NOT NULL,   -- currently only 'post'
    start_month  TEXT NOT NULL,   -- YYYY-MM inclusive
    end_month    TEXT NOT NULL,   -- YYYY-MM inclusive
    source       TEXT NOT NULL,   -- 'auto' | 'manual'
    note         TEXT,
    PRIMARY KEY (address_id, start_month)
);
"""

POST_LONG_SEC = 7200      # a visit this long is posting, not a transport
POST_SHARE = 0.15         # month qualifies as a post month above this share
POST_SPLIT_SEC = 5400     # inside a post period, split posting from transport here
MIN_MONTH_VISITS = 20
MIN_PERIOD_MONTHS = 2   # a single flagged month is usually noise, not a post


def cmd_classify(args):
    """Detect which addresses served as crew posts, and WHEN.

    A static per-address flag is wrong: bases move. RiverVista was the Columbus
    post until May 2026 and a plain facility after, so the classification is
    scoped to month ranges. Rows with source='manual' are never overwritten.
    """
    con = open_db(args.db)
    con.executescript(ROLE_SCHEMA)
    con.commit()

    rows = con.execute("""
        SELECT address_id, address_name, substr(entry_ts,1,7) m,
               COUNT(*) n, SUM(CASE WHEN dwell_sec >= ? THEN 1 ELSE 0 END) longn
        FROM visit WHERE qualified=1
        GROUP BY address_id, m ORDER BY address_id, m""", (POST_LONG_SEC,)).fetchall()

    per = {}
    for aid, name, m, n, longn in rows:
        per.setdefault(aid, {"name": name, "months": []})["months"].append(
            (m, n, longn, (longn / float(n)) if n else 0.0))

    manual = set()
    for aid, s in con.execute("SELECT address_id, start_month FROM site_role WHERE source='manual'"):
        manual.add(aid)

    con.execute("DELETE FROM site_role WHERE source='auto'")
    made = 0
    detail = []
    for aid, d in per.items():
        if aid in manual:
            continue
        flags = []
        for m, n, longn, share in d["months"]:
            if n < MIN_MONTH_VISITS:
                flags.append((m, None))          # too thin to judge; inherit
            else:
                flags.append((m, share >= POST_SHARE))
        # carry the previous verdict through thin months
        last = False
        filled = []
        for m, f in flags:
            if f is None:
                f = last
            last = f
            filled.append((m, f))
        # collapse contiguous post months into ranges
        start = None
        for i, (m, f) in enumerate(filled):
            if f and start is None:
                start = m
            if start is not None and (not f or i == len(filled) - 1):
                end = m if (f and i == len(filled) - 1) else filled[i - 1][0]
                ay, am = map(int, start.split("-"))
                by, bm = map(int, end.split("-"))
                if (by - ay) * 12 + (bm - am) + 1 >= MIN_PERIOD_MONTHS:
                    con.execute("INSERT OR REPLACE INTO site_role VALUES (?,?,?,?,?,?,?)",
                                (aid, d["name"], "post", start, end, "auto", None))
                    detail.append((d["name"], start, end))
                    made += 1
                start = None
    con.commit()

    print("post periods detected: %d  (manual rows preserved: %d)" % (made, len(manual)))
    print("\n  %-42s %9s %9s" % ("site", "from", "to"))
    for nm, a, b in sorted(detail, key=lambda x: x[0])[:30]:
        print("  %-42s %9s %9s" % (nm[:42], a, b))
    print("\nOverride by hand where you know better, e.g.:")
    print("  INSERT INTO site_role VALUES('<address_id>','<name>','post','2025-09','2026-05','manual','moved');")


def cmd_mark(args):
    """Mark a site as a base/office (never a transport destination), or clear it.

    Some sites cannot be classified from movement data - a head office looks
    exactly like a fast facility, because crews really do arrive and leave.
    Only someone who knows the operation can say so.

      mark --name "COMMUNICARE CORPORATE" --as base
      mark --name "SOME SITE" --clear
      mark --list
    """
    con = open_db(args.db)
    con.executescript(ROLE_SCHEMA)

    if args.list or not args.name:
        print("sites marked as base/office (all visits excluded from reports):")
        rows = con.execute("SELECT address_id, address_name, note FROM site_role "
                           "WHERE role='base' ORDER BY address_name").fetchall()
        if not rows:
            print("  (none yet)")
        for aid, nm, note in rows:
            print("  %-46s %s  %s" % (nm[:46], aid, note or ""))
        if not args.name:
            print("\nCandidates worth reviewing - sites whose visits come from very few vehicles:")
            for nm, aid, v, n in con.execute("""
                SELECT address_name, address_id, COUNT(*) n, COUNT(DISTINCT vehicle_id) v
                FROM visit WHERE qualified=1 AND entry_ts>='2026-06'
                GROUP BY address_id HAVING n>=100 AND v<=6 ORDER BY n DESC LIMIT 12"""):
                print("  %-46s %5d visits from only %d vehicles" % (nm[:46], v, n))
        return

    hit = con.execute("SELECT address_id, address_name FROM visit WHERE address_name LIKE ? "
                      "GROUP BY address_id", ("%" + args.name + "%",)).fetchall()
    if not hit:
        raise SystemExit("no site matches %r" % args.name)
    if len(hit) > 1:
        print("ambiguous - matches %d sites:" % len(hit))
        for aid, nm in hit:
            print("   %s  %s" % (aid, nm))
        raise SystemExit("narrow the --name")
    aid, nm = hit[0]

    if args.clear:
        con.execute("DELETE FROM site_role WHERE address_id=? AND role='base'", (aid,))
        con.commit()
        print("cleared base marking on %s" % nm)
        return

    con.execute("INSERT OR REPLACE INTO site_role VALUES (?,?,?,?,?,?,?)",
                (aid, nm, "base", "1970-01", "2999-12", "manual", args.note or "not a transport destination"))
    con.commit()
    n = con.execute("SELECT COUNT(*), ROUND(SUM(dwell_sec)/3600.0) FROM visit "
                    "WHERE qualified=1 AND address_id=?", (aid,)).fetchone()
    print("marked %s as base/office." % nm)
    print("  %s visits / %s crew hours will now be excluded from every report." % (n[0], n[1]))


def _post_ranges(con):
    con.executescript(ROLE_SCHEMA)
    out = {}
    for aid, s, e in con.execute("SELECT address_id, start_month, end_month FROM site_role WHERE role='post'"):
        out.setdefault(aid, []).append((s, e))
    return out


def cmd_report(args):
    """Monthly facility turnaround, with posting excluded via site_role."""
    import csv as _csv
    con = open_db(args.db)
    posts = _post_ranges(con)
    if not posts:
        print("WARNING: no post periods recorded - run 'classify' first, or "
              "posting time will be counted as detention.\n")

    since = args.since or "2025-09"
    rows = con.execute("""
        SELECT address_id, address_name, substr(entry_ts,1,7) m, dwell_sec
        FROM visit WHERE qualified=1 AND substr(entry_ts,1,7) >= ?""", (since,)).fetchall()

    def is_posting(aid, m, dwell):
        for s, e in posts.get(aid, []):
            if s <= m <= e and dwell >= POST_SPLIT_SEC:
                return True
        return False

    # Sites marked role='base' are never transport destinations - head offices,
    # depots, crew bases. EVERY visit is excluded, not just the long ones. This
    # cannot be inferred from movement data: a crew coming and going from the
    # office looks exactly like a short handoff. It is asserted by someone who
    # knows the operation, via the 'mark' command.
    excluded = set(str(r[0]) for r in
                   con.execute("SELECT DISTINCT address_id FROM site_role WHERE role='base'"))

    def is_base_name(n):
        u = (n or "").upper()
        return (" BASE" in u or u.endswith("BASE") or "DEPOT" in u
                or "BASE LOCATION" in u or "INVENTORY" in u)

    agg = {}
    dropped = 0
    bases = 0
    for aid, name, m, dwell in rows:
        if str(aid) in excluded or is_base_name(name):
            bases += 1
            continue          # a facility turnaround report is not about bases
        if is_posting(aid, m, dwell):
            dropped += 1
            continue
        agg.setdefault((aid, m), {"name": name, "d": []})["d"].append(dwell / 60.0)

    def med(x):
        s = sorted(x); return s[len(s) // 2]
    def p90(x):
        s = sorted(x); return s[int(0.9 * (len(s) - 1))]

    months = sorted(set(m for _, m in agg.keys()))
    cur = months[-1]
    prev = months[-2] if len(months) > 1 else None
    print("transport visits since %s   (excluded: %s posting, %s at bases)"
          % (since, "{:,}".format(dropped), "{:,}".format(bases)))
    print("reporting month: %s   comparison: %s\n" % (cur, prev or "n/a"))

    cur_rows = [(v["name"], len(v["d"]), med(v["d"]), p90(v["d"]), sum(v["d"]) / 60.0, aid)
                for (aid, m), v in agg.items() if m == cur and len(v["d"]) >= args.min_visits]
    if not cur_rows:
        print("no facility met --min-visits %d in %s" % (args.min_visits, cur)); return
    fleet = med([r[2] for r in cur_rows])
    print("fleet median turnaround: %.0f min   (%d facilities >= %d visits)\n"
          % (fleet, len(cur_rows), args.min_visits))

    prev_med = {}
    if prev:
        for (aid, m), v in agg.items():
            if m == prev and len(v["d"]) >= args.min_visits:
                prev_med[aid] = med(v["d"])

    ranked = sorted(cur_rows, key=lambda r: -r[2])
    print("  %-38s %6s %7s %7s %8s %9s" % ("facility", "visits", "median", "p90", "hours", "vs prev"))
    for nm, n, md, p9, hrs, aid in ranked[:args.top]:
        delta = "%+.0f min" % (md - prev_med[aid]) if aid in prev_med else ""
        print("  %-38s %6d %7.0f %7.0f %8.0f %9s" % (nm[:38], n, md, p9, hrs, delta))
    if len(ranked) > args.top:
        print("  ... %d more in the CSV" % (len(ranked) - args.top))

    # the CSV carries EVERY facility, not just the printed ones
    out = []
    for rank, (nm, n, md, p9, hrs, aid) in enumerate(ranked, 1):
        out.append({"rank": rank, "facility": nm, "month": cur, "visits": n,
                    "median_min": round(md, 1), "p90_min": round(p9, 1),
                    "hours": round(hrs, 1),
                    "prev_median_min": round(prev_med[aid], 1) if aid in prev_med else "",
                    "delta_min": round(md - prev_med[aid], 1) if aid in prev_med else "",
                    "fleet_median_min": round(fleet, 1),
                    "vs_fleet_min": round(md - fleet, 1)})

    path = os.path.join(os.path.dirname(args.db) or ".", "facility_turnaround_%s.csv" % cur)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = _csv.DictWriter(fh, fieldnames=list(out[0].keys()))
        w.writeheader()
        for r in out:
            w.writerow(r)
    print("\nwritten to %s" % path)


def _haversine(a1, o1, a2, o2):
    import math
    R = 6371000.0
    p1, p2 = math.radians(a1), math.radians(a2)
    dp = p2 - p1
    dl = math.radians(o2 - o1)
    h = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(min(1, math.sqrt(h)))


def _pct(vals, p):
    if not vals:
        return 0.0
    s = sorted(vals)
    i = min(len(s) - 1, int(p / 100.0 * (len(s) - 1)))
    return s[i]


def cmd_resize(args):
    """Fit each geofence radius to its own measured footprint. RADIUS ONLY -
    centre, name, tags and externalIds are untouched (verified: PATCH is a true
    partial update). Dry-run unless --execute."""
    import csv as _csv
    con = open_db(args.db)
    api = Api(load_token())

    # measured footprint: pings that fall inside a QUALIFIED visit
    rows = con.execute(
        "SELECT p.address_id, p.address_name, p.lat, p.lon, v.qualified "
        "FROM geofence_ping p JOIN visit v "
        "  ON v.vehicle_id = p.vehicle_id AND p.ts BETWEEN v.entry_ts AND v.exit_ts"
    ).fetchall()
    if not rows:
        raise SystemExit("no joined pings - run 'visits' first.")

    agg = {}
    for aid, an, la, lo, q in rows:
        d = agg.setdefault(aid, {"name": an, "stop": [], "drive": 0})
        if q == 1:
            d["stop"].append((la, lo))
        else:
            d["drive"] += 1

    print("addresses with any data: %d" % len(agg))

    # current radii, straight from Samsara
    cur = {}
    after = None
    while True:
        p = {}
        if after:
            p["after"] = after
        j = api.get("/addresses", p)
        for a in (j.get("data") or []):
            gf = a.get("geofence") or {}
            circ = gf.get("circle")
            ext = a.get("externalIds") or {}
            cur[str(a["id"])] = {
                "name": a.get("name"),
                "radius": int(circ["radiusMeters"]) if circ else None,
                "polygon": bool(gf.get("polygon")),
                # addresses we created carry traumasoftFacilityId; anything else
                # pre-dates the import and may have been tuned by hand
                "source": "import" if ext.get("traumasoftFacilityId") else "curated",
            }
        pg = j.get("pagination") or {}
        if not pg.get("hasNextPage"):
            break
        after = pg.get("endCursor")
    print("addresses in Samsara : %d" % len(cur))

    plan, skipped = [], {"few_pings": 0, "polygon": 0, "unknown": 0, "small_change": 0, "curated": 0}
    for aid, d in agg.items():
        if len(d["stop"]) < args.min_pings:
            skipped["few_pings"] += 1
            continue
        info = cur.get(str(aid))
        if not info:
            skipped["unknown"] += 1
            continue
        if info["polygon"]:
            skipped["polygon"] += 1   # never convert a hand-drawn polygon to a circle
            continue
        la = sum(x[0] for x in d["stop"]) / len(d["stop"])
        lo = sum(x[1] for x in d["stop"]) / len(d["stop"])
        spread = [_haversine(la, lo, a, o) for a, o in d["stop"]]
        p95 = _pct(spread, 95)
        want = int(max(args.floor_m, min(args.ceiling_m, round(p95 * args.margin))))
        now = info["radius"]
        if now is None:
            skipped["unknown"] += 1
            continue
        if abs(want - now) < max(10, now * args.min_change):
            skipped["small_change"] += 1
            continue
        outside = sum(1 for s in spread if s > want)
        if info["source"] == "curated" and not args.include_curated:
            skipped["curated"] = skipped.get("curated", 0) + 1
            continue
        plan.append({
            "address_id": aid, "name": d["name"] or info["name"],
            "source": info["source"],
            "stopped_pings": len(d["stop"]), "drive_throughs": d["drive"],
            "p95_spread_m": int(p95), "current_m": now, "proposed_m": want,
            "direction": "shrink" if want < now else "widen",
            "stop_pings_outside_new": outside,
        })

    plan.sort(key=lambda r: -(r["drive_throughs"]))
    shrink = [r for r in plan if r["direction"] == "shrink"]
    widen = [r for r in plan if r["direction"] == "widen"]

    print("\n=== PLAN ===")
    print("  shrink : %d" % len(shrink))
    print("  widen  : %d" % len(widen))
    print("  skipped: %d too few stopped pings, %d polygon, %d small change,"
          % (skipped["few_pings"], skipped["polygon"], skipped["small_change"]))
    print("           %d hand-curated (use --include-curated to include them), %d not in Samsara"
          % (skipped["curated"], skipped["unknown"]))
    print("\n%-40s %7s %7s %7s %8s" % ("address", "cur", "new", "p95", "driveN"))
    for r in plan[:20]:
        print("%-40s %7d %7d %7d %8d"
              % (r["name"][:40], r["current_m"], r["proposed_m"], r["p95_spread_m"], r["drive_throughs"]))

    out = os.path.join(os.path.dirname(args.db) or ".", "geofence_resize_plan.csv")
    with open(out, "w", newline="", encoding="utf-8") as fh:
        w = _csv.DictWriter(fh, fieldnames=list(plan[0].keys()) if plan else ["address_id"])
        w.writeheader()
        for r in plan:
            w.writerow(r)
    print("\nplan written to %s" % out)

    if not args.execute:
        print("\nDRY RUN - nothing written. Re-run with --execute to apply.")
        return

    ok = fail = 0
    for i, r in enumerate(plan, 1):
        body = {"geofence": {"circle": {"radiusMeters": int(r["proposed_m"])}}}
        url = BASE + "/addresses/" + str(r["address_id"])
        try:
            wait = MIN_INTERVAL - (time.time() - api._last)
            if wait > 0:
                time.sleep(wait)
            api._last = time.time()
            resp = api.s.patch(url, json=body, timeout=60)
            if resp.status_code != 200:
                raise RuntimeError("HTTP %s :: %s" % (resp.status_code, resp.text[:200]))
            ok += 1
        except Exception as exc:
            fail += 1
            print("  FAILED %s: %s" % (r["name"][:40], exc))
        if i % 25 == 0:
            print("  %d/%d  ok=%d failed=%d" % (i, len(plan), ok, fail))
    print("\ndone. ok=%d failed=%d" % (ok, fail))
    print("Matching is retroactive - re-run backfill/visits to see the effect on history.")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["horizon", "backfill", "incremental", "visits",
                                        "stats", "resize", "classify", "report", "mark"])
    ap.add_argument("--db", default=DEFAULT_DB)
    ap.add_argument("--days", type=int, default=460)
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--vehicles", help="comma-separated vehicle names to limit to")
    ap.add_argument("--batch", type=int, default=8, help="vehicles per API call")
    ap.add_argument("--gap-minutes", type=int, default=10,
                    help="gap that splits one visit into two")
    ap.add_argument("--min-dwell-minutes", type=int, default=2,
                    help="dwell below this is flagged as a drive-through, not deleted "
                         "(2 min matches Samsara's own Time on Site definition)")
    ap.add_argument("--min-pings", type=int, default=30,
                    help="stopped pings required before an address is resized")
    ap.add_argument("--margin", type=float, default=1.15,
                    help="proposed radius = p95 of stopped-ping spread x this")
    ap.add_argument("--floor-m", type=int, default=75,
                    help="never propose tighter than this (GPS drift)")
    ap.add_argument("--ceiling-m", type=int, default=400)
    ap.add_argument("--min-change", type=float, default=0.20,
                    help="skip when the change is under this fraction of current radius")
    ap.add_argument("--reset", action="store_true",
                    help="backfill: wipe stored pings/visits first (needed after a "
                         "geofence resize, since matching is retroactive)")
    ap.add_argument("--name", help="mark: substring of the site name")
    ap.add_argument("--as", dest="as_role", default="base", help="mark: role to assign (base)")
    ap.add_argument("--note", help="mark: why")
    ap.add_argument("--clear", action="store_true", help="mark: remove the marking")
    ap.add_argument("--list", action="store_true", help="mark: list current markings")
    ap.add_argument("--since", help="report: earliest month, YYYY-MM (default 2025-09)")
    ap.add_argument("--min-visits", type=int, default=25,
                    help="report: minimum visits in the month for a facility to appear")
    ap.add_argument("--top", type=int, default=20, help="report: rows to print")
    ap.add_argument("--confirm-reset", action="store_true",
                    help="required alongside --reset when the store already holds data")
    ap.add_argument("--include-curated", action="store_true",
                    help="also resize the pre-existing hand-tuned addresses "
                         "(default: leave them alone)")
    ap.add_argument("--execute", action="store_true",
                    help="apply the resize; omit for a dry run")
    args = ap.parse_args()
    {"horizon": cmd_horizon, "backfill": cmd_backfill, "incremental": cmd_incremental,
     "visits": cmd_visits, "stats": cmd_stats, "resize": cmd_resize,
     "classify": cmd_classify, "report": cmd_report, "mark": cmd_mark}[args.command](args)


if __name__ == "__main__":
    main()
