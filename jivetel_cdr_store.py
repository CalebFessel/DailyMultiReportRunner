#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Jivetel (NetSapiens) call detail record store.

Pulls per-call CDR legs from /ns-api/v2/domains/<domain>/cdrs and keeps only
the fields that matter for call-center reporting, in SQLite. The API keeps
CDRs from 2025-01-01 (edge measured 2026-09-10), so this store is the durable
copy going forward.

Two products:
  cdr_leg   one row per CDR leg, trimmed to the reporting fields. A transfer
            produces a second leg with the same call_id.
  call      one row per phone call, grouped on call_id (call-orig-call-id):
            earliest start, summed talk/hold, first and last agent. Rebuilt
            from cdr_leg at any time.

API behaviour this script works around (all measured 2026-09-10):
  * limit > 100 silently switches the response to a 28-field summary that
    lacks call-orig-call-id, caller id and disconnect reason. Pages are
    therefore fixed at 100 and stepped with `start` (an offset; `offset` is
    ignored). One row in ~370k also arrives in that slim shape on its own
    (a leg that crosses UTC midnight); those are keyed on parent_cdr_id.
  * start_date / end_date are calendar days in UTC, inclusive. Timestamps
    come back as ISO-8601 with +00:00. Local (America/New_York) date and
    time are derived at ingest so reports never have to convert.
  * Dedupe on call_id, never on caller_id - a phone number that calls three
    times is three calls.
  * call-answer-datetime is set on every inbound leg because the queue or
    auto-attendant answers the SIP call. "Answered" here means a leg with
    talk_sec > 0; an inbound leg on a queue (term_user 9xxx) with talk_sec 0
    and 'Orig: Bye' is a caller who hung up while waiting.
  * call-direction: 0 = placed from an extension (outbound or internal),
    1 = inbound from outside, 2 = inbound legs that never connected
    (zero duration at a queue), 3 = rare, seen once.

Resumability is per day. A finished day is recorded in backfill_day and
skipped on re-run unless --force is given; today is never marked finished
because more legs will arrive.

Usage
-----
  # small test before the full run
  python jivetel_cdr_store.py backfill --start 2026-09-08 --end 2026-09-09

  # everything the API still has, oldest first, resumable
  python jivetel_cdr_store.py backfill --start 2025-01-01

  # catch up since the last run (for the scheduler): refetches yesterday
  # and today
  python jivetel_cdr_store.py incremental

  # (re)build the per-call table from stored legs
  python jivetel_cdr_store.py calls

  python jivetel_cdr_store.py stats
"""

import os
import sys
import time
import sqlite3
import argparse
import datetime as dt
from zoneinfo import ZoneInfo
from urllib.parse import urlencode

import requests
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

BASE = os.environ.get("JIVETEL_BASE_URL", "https://api.jivetel.net/ns-api/v2")
DOMAIN = os.environ.get("JIVETEL_DOMAIN", "LYNXEMSLLC")
TOKEN = os.environ.get("JIVETEL_API_TOKEN", "")
DEFAULT_DB = os.environ.get("JIVETEL_CDR_DB", r"C:\Reports\jivetel_cdr.db")
MIN_INTERVAL = float(os.environ.get("JIVETEL_MIN_INTERVAL", "0.2"))
TIMEOUT = int(os.environ.get("JIVETEL_API_TIMEOUT", "60"))
LOCAL_TZ = ZoneInfo("America/New_York")

PAGE = 100                      # anything larger drops to the summary schema
RETRY_STATUSES = {429, 500, 502, 503, 504}
CDR_EDGE = dt.date(2025, 1, 1)  # earliest day with data, measured 2026-09-10

# API field -> column. Everything not listed here is dropped on purpose:
# codec, RTP relay ports, packet counts, fax/video legs, SIP URIs, server
# MAC, sentiment placeholders.
FIELDS = {
    "id":                            "id",
    "call-orig-call-id":             "call_id",
    "call-parent-cdr-id":            "parent_cdr_id",
    "call-leg-ordinal-index":        "leg_index",
    "call-direction":                "direction",
    "call-start-datetime":           "start_utc",
    "call-ringing-datetime":         "ringing_utc",
    "call-answer-datetime":          "answer_utc",
    "call-disconnect-datetime":      "disconnect_utc",
    "call-total-duration-seconds":   "total_sec",
    "call-talking-duration-seconds": "talk_sec",
    "call-on-hold-duration-seconds": "hold_sec",
    "call-orig-caller-id":           "caller_id",
    "call-orig-from-name":           "caller_name",
    "call-orig-user":                "orig_user",
    "call-orig-request-user":        "dialed",
    "call-term-user":                "term_user",
    "call-term-caller-id":           "term_caller_id",
    "call-through-user":             "through_user",
    "call-disconnect-reason-text":   "disconnect_reason",
    "call-routing-class":            "routing_class",
    "call-tag":                      "tag",
    "call-disposition":              "disposition",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS cdr_leg (
    id                TEXT PRIMARY KEY,
    call_id           TEXT,
    parent_cdr_id     TEXT,
    leg_index         INTEGER,
    direction         INTEGER,
    start_utc         TEXT,
    ringing_utc       TEXT,
    answer_utc        TEXT,
    disconnect_utc    TEXT,
    start_local       TEXT,       -- America/New_York, ISO-8601
    local_date        TEXT,       -- YYYY-MM-DD in America/New_York
    total_sec         INTEGER,
    talk_sec          INTEGER,
    hold_sec          INTEGER,
    caller_id         TEXT,
    caller_name       TEXT,
    orig_user         TEXT,
    dialed            TEXT,
    term_user         TEXT,
    term_caller_id    TEXT,
    through_user      TEXT,
    disconnect_reason TEXT,
    routing_class     TEXT,
    tag               TEXT,
    disposition       TEXT,
    fetched_at        TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_leg_call   ON cdr_leg(call_id);
CREATE INDEX IF NOT EXISTS ix_leg_date   ON cdr_leg(local_date);
CREATE INDEX IF NOT EXISTS ix_leg_term   ON cdr_leg(term_user, local_date);

CREATE TABLE IF NOT EXISTS call (
    call_id           TEXT PRIMARY KEY,
    direction         INTEGER,
    start_utc         TEXT,
    start_local       TEXT,
    local_date        TEXT,
    answered          INTEGER,    -- 1 if any leg has talk_sec > 0. answer_utc is
                                   -- useless: the queue/IVR 'answers' every call
    legs              INTEGER,
    total_sec         INTEGER,
    talk_sec          INTEGER,
    hold_sec          INTEGER,
    caller_id         TEXT,
    caller_name       TEXT,
    orig_user         TEXT,
    dialed            TEXT,
    first_term_user   TEXT,
    last_term_user    TEXT,
    queues            TEXT,       -- comma-separated queue term_users touched
    disconnect_reason TEXT,       -- from the last leg
    built_at          TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_call_date ON call(local_date);

CREATE TABLE IF NOT EXISTS backfill_day (
    day        TEXT PRIMARY KEY,  -- UTC calendar day sent to the API
    legs       INTEGER NOT NULL,
    fetched_at TEXT NOT NULL
);
"""


# ----------------------------------------------------------------- HTTP ---

class Client:
    def __init__(self, token=TOKEN, base=BASE, domain=DOMAIN):
        if not token:
            sys.exit("JIVETEL_API_TOKEN is not set (put it in C:\\Reports\\.env)")
        self.base = f"{base}/domains/{domain}"
        self.s = requests.Session()
        self.s.headers.update({"Authorization": token, "Accept": "application/json"})
        self._last = 0.0
        self.calls = 0

    def get(self, path, params):
        url = f"{self.base}{path}?{urlencode(params)}"
        for attempt in range(6):
            wait = MIN_INTERVAL - (time.monotonic() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            self.calls += 1
            r = self.s.get(url, timeout=TIMEOUT)
            if r.status_code in RETRY_STATUSES:
                time.sleep(min(2 ** attempt, 30))
                continue
            r.raise_for_status()
            return r.json()
        raise RuntimeError(f"gave up on {url} after retries (last status {r.status_code})")

    def cdrs_for_day(self, day):
        """Every CDR leg for one UTC calendar day, full schema."""
        out, start = [], 0
        while True:
            page = self.get("/cdrs", {
                "start_date": day.isoformat(), "end_date": day.isoformat(),
                "limit": PAGE, "start": start,
            })
            if not isinstance(page, list):
                raise RuntimeError(f"unexpected response for {day}: {str(page)[:200]}")
            # A whole page in the 28-field shape means the page size tripped
            # the summary schema. A single slim row inside a full page is a
            # different thing: legs that straddle UTC midnight come back that
            # way on their own, and trim() handles them.
            if len(page) == PAGE and all("call-orig-call-id" not in r for r in page):
                raise RuntimeError("API returned the summary schema; PAGE must stay <= 100")
            out.extend(page)
            if len(page) < PAGE:
                return out
            start += PAGE


# -------------------------------------------------------------- storage ---

def open_db(path):
    con = sqlite3.connect(path)
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript(SCHEMA)
    return con


def _local(iso_utc):
    if not iso_utc:
        return None, None
    t = dt.datetime.fromisoformat(iso_utc).astimezone(LOCAL_TZ)
    return t.isoformat(), t.date().isoformat()


def _s(v):
    """Numbers and strings both arrive for user/number fields; store text."""
    return None if v is None else str(v)


def trim(rec, fetched_at):
    row = {col: rec.get(api) for api, col in FIELDS.items()}
    for k in ("caller_id", "orig_user", "dialed", "term_user", "term_caller_id", "through_user"):
        row[k] = _s(row[k])
    # Slim rows (legs spanning UTC midnight) carry no id or call id. Key them
    # on the parent CDR id so re-runs upsert instead of duplicating, and group
    # them under it too - for a single-leg call it is the call itself.
    if row["id"] is None and row["parent_cdr_id"]:
        row["id"] = "slim:" + row["parent_cdr_id"]
    if row["call_id"] is None:
        row["call_id"] = row["parent_cdr_id"]
    row["start_local"], row["local_date"] = _local(row["start_utc"])
    row["fetched_at"] = fetched_at
    return row


def upsert_legs(con, rows):
    if not rows:
        return
    cols = list(rows[0].keys())
    sql = (f"INSERT INTO cdr_leg ({','.join(cols)}) VALUES ({','.join('?' * len(cols))}) "
           f"ON CONFLICT(id) DO UPDATE SET " +
           ",".join(f"{c}=excluded.{c}" for c in cols if c != "id"))
    con.executemany(sql, [tuple(r[c] for c in cols) for r in rows])


def fetch_day(client, con, day, mark_done):
    now = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    recs = client.cdrs_for_day(day)
    rows = [trim(r, now) for r in recs]
    with con:
        upsert_legs(con, rows)
        if mark_done:
            con.execute("INSERT OR REPLACE INTO backfill_day VALUES (?,?,?)",
                        (day.isoformat(), len(rows), now))
    return len(rows)


def done_days(con):
    return {r[0] for r in con.execute("SELECT day FROM backfill_day")}


# ---------------------------------------------------------------- calls ---

def build_calls(con, since=None):
    """Collapse legs into one row per call_id. Idempotent."""
    where = "WHERE call_id IS NOT NULL"
    args = ()
    if since:
        where += " AND local_date >= ?"
        args = (since,)
    legs = con.execute(f"""
        SELECT call_id, direction, start_utc, start_local, local_date, answer_utc,
               total_sec, talk_sec, hold_sec, caller_id, caller_name, orig_user,
               dialed, term_user, disconnect_reason
        FROM cdr_leg {where}
        ORDER BY call_id, start_utc, leg_index""", args).fetchall()

    now = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    out, cur = [], None
    for leg in legs:
        (cid, direction, start_utc, start_local, local_date, answer, total, talk, hold,
         caller_id, caller_name, orig_user, dialed, term_user, reason) = leg
        if cur is None or cur["call_id"] != cid:
            if cur:
                out.append(cur)
            cur = dict(call_id=cid, direction=direction, start_utc=start_utc,
                       start_local=start_local, local_date=local_date, answered=0,
                       legs=0, total_sec=0, talk_sec=0, hold_sec=0,
                       caller_id=caller_id, caller_name=caller_name, orig_user=orig_user,
                       dialed=dialed, first_term_user=term_user, last_term_user=term_user,
                       queues=[], disconnect_reason=reason, built_at=now)
        cur["legs"] += 1
        cur["answered"] |= 1 if (talk or 0) > 0 else 0
        cur["total_sec"] += total or 0
        cur["talk_sec"] += talk or 0
        cur["hold_sec"] += hold or 0
        cur["last_term_user"] = term_user
        cur["disconnect_reason"] = reason
        if term_user and term_user.startswith("9") and len(term_user) == 4:
            if term_user not in cur["queues"]:
                cur["queues"].append(term_user)
    if cur:
        out.append(cur)

    for c in out:
        c["queues"] = ",".join(c["queues"]) or None
    if out:
        cols = list(out[0].keys())
        with con:
            con.executemany(
                f"INSERT OR REPLACE INTO call ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                [tuple(c[k] for k in cols) for c in out])
    return len(out)


# ------------------------------------------------------------- commands ---

def cmd_backfill(a):
    client, con = Client(), open_db(a.db)
    today = dt.datetime.now(dt.timezone.utc).date()
    start = dt.date.fromisoformat(a.start) if a.start else CDR_EDGE
    end = dt.date.fromisoformat(a.end) if a.end else today
    if start < CDR_EDGE:
        print(f"note: API has nothing before {CDR_EDGE}; starting there")
        start = CDR_EDGE
    skip = set() if a.force else done_days(con)
    days = [start + dt.timedelta(d) for d in range((end - start).days + 1)]
    todo = [d for d in days if d.isoformat() not in skip]
    print(f"{len(todo)} of {len(days)} days to fetch ({start} .. {end}), oldest first")
    total, t0 = 0, time.monotonic()
    for i, day in enumerate(todo, 1):
        n = fetch_day(client, con, day, mark_done=day < today)
        total += n
        print(f"  {day}  {n:5d} legs   ({i}/{len(todo)}, {client.calls} requests, "
              f"{time.monotonic() - t0:.0f}s)")
    n_calls = build_calls(con, since=(start - dt.timedelta(1)).isoformat())
    print(f"done: {total} legs stored, {n_calls} calls rebuilt from {start}")


def cmd_incremental(a):
    client, con = Client(), open_db(a.db)
    today = dt.datetime.now(dt.timezone.utc).date()
    # Refetch the last finished day too: legs that were still in progress at
    # UTC midnight land after the day was first fetched.
    for day in (today - dt.timedelta(1), today):
        n = fetch_day(client, con, day, mark_done=day < today)
        print(f"  {day}  {n:5d} legs")
    n_calls = build_calls(con, since=(today - dt.timedelta(2)).isoformat())
    print(f"done: {n_calls} calls rebuilt")


def cmd_calls(a):
    con = open_db(a.db)
    print(f"{build_calls(con, since=a.since)} calls built")


def cmd_stats(a):
    con = open_db(a.db)
    legs, lo, hi = con.execute("SELECT COUNT(*), MIN(local_date), MAX(local_date) FROM cdr_leg").fetchone()
    calls = con.execute("SELECT COUNT(*) FROM call").fetchone()[0]
    days = con.execute("SELECT COUNT(*) FROM backfill_day").fetchone()[0]
    print(f"legs {legs}  calls {calls}  days finished {days}  range {lo} .. {hi}")
    print("\ncalls by direction:")
    for d, n, ans in con.execute(
            "SELECT direction, COUNT(*), SUM(answered) FROM call GROUP BY direction"):
        print(f"  direction {d}: {n} calls, {ans} answered")
    print("\nlast 7 local days:")
    for row in con.execute("""
        SELECT local_date, COUNT(*), SUM(answered), SUM(legs > 1)
        FROM call GROUP BY local_date ORDER BY local_date DESC LIMIT 7"""):
        print("  %s  %5d calls  %5d answered  %3d multi-leg" % row)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--db", default=DEFAULT_DB)
    sub = p.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("backfill", help="fetch a date range, oldest first, resumable")
    b.add_argument("--start", help="YYYY-MM-DD (UTC day), default 2025-01-01")
    b.add_argument("--end", help="YYYY-MM-DD, default today")
    b.add_argument("--force", action="store_true", help="refetch days already marked done")
    b.set_defaults(fn=cmd_backfill)

    sub.add_parser("incremental", help="refetch yesterday and today").set_defaults(fn=cmd_incremental)

    c = sub.add_parser("calls", help="rebuild the per-call table from stored legs")
    c.add_argument("--since", help="only rebuild local dates >= this")
    c.set_defaults(fn=cmd_calls)

    sub.add_parser("stats").set_defaults(fn=cmd_stats)

    a = p.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
