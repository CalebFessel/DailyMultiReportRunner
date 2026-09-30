#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Dispatch 24-hour snapshot report.

An emailable, static version of the /dispatch dashboard: the same KPI strip,
call-distribution donut, needs-attention list, queue and agent tables, and
7-day trend lines, rendered as images and HTML so it survives in an inbox.

Numbers come from the same two Jivetel endpoints the live board uses, with
the same window (yesterday + today, which the board labels "last 24h") and
the same scope rules, so the report always agrees with the dashboard. One
extra graphic, inbound calls by hour for the trailing 24 hours, comes from
the local CDR store (jivetel_cdr.db) because the statistics endpoints only
resolve to whole days.

Not included: the turn-downs tile (needs the TraumaSoft database, which this
machine does not reach) and the live online/available bar (meaningless in a
report).

Usage
-----
  python dispatch_24h_report.py                       # main board -> HTML
  python dispatch_24h_report.py --scope wv
  python dispatch_24h_report.py --scope both
  python dispatch_24h_report.py --scope both --email a@x.com b@x.com

Output goes to C:\\Reports\\Reports\\dispatch_snapshot_<scope>_<stamp>.html.
Email needs SMTP_USER / SMTP_PASS (and optionally SMTP_FROM, SMTP_SERVER,
SMTP_PORT) in C:\\Reports\\.env, the same names the automation repo uses.
"""

import os
import io
import sys
import base64
import sqlite3
import smtplib
import argparse
import datetime as dt
from zoneinfo import ZoneInfo
from email.message import EmailMessage
from email.utils import make_msgid

import requests
from dotenv import load_dotenv
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

BASE   = os.environ.get("JIVETEL_BASE_URL", "https://api.jivetel.net/ns-api/v2")
DOMAIN = os.environ.get("JIVETEL_DOMAIN", "LYNXEMSLLC")
TOKEN  = os.environ.get("JIVETEL_API_TOKEN", "")
CDR_DB = os.environ.get("JIVETEL_CDR_DB", r"C:\Reports\jivetel_cdr.db")
OUT_DIR = os.environ.get("DISPATCH_REPORT_DIR", r"C:\Reports\Reports")
TZ = ZoneInfo("America/New_York")

# ── Mirrors Dashboard_JSON_Generator.py / dispatch.html ─────────────────
WV_QUEUE_NAMES = {"West Virginia Dispatch"}
WV_QUEUE_IDS   = {"9112"}
WV_EXT_MIN, WV_EXT_MAX = 400, 499
MIN_CALLS_FOR_ATTENTION = 5
TH = dict(sl_red=80, sl_yellow=95, aban_red=5, aban_yellow=2, wait_red=30, wait_yellow=15)
DIST_PALETTE = ['#2E86DE', '#2ECC71', '#FF9F1C', '#9B59B6', '#E63946', '#16A085', '#3C3489', '#888888']
C_VOLUME, C_ABANDON, C_WAIT, C_SL = '#185FA5', '#A32D2D', '#BA7517', '#0F6E56'
BOARD_TITLE = {"main": "Ohio Dispatch", "wv": "West Virginia Dispatch"}
ACCENT = {"main": "#3FB8FF", "wv": "#FFB347"}


# ── Jivetel ──────────────────────────────────────────────────────────────

def _get(path, params):
    r = requests.get(f"{BASE}/domains/{DOMAIN}{path}", params=params,
                     headers={"Authorization": TOKEN, "Accept": "application/json"}, timeout=30)
    r.raise_for_status()
    return r.json()


def _ext(agent_key):
    s = str(agent_key or "")
    if s.startswith("sip:"):
        s = s[4:]
    s = s.split("@", 1)[0]
    digits = ""
    for c in s:
        if c.isdigit():
            digits += c
        else:
            break
    return int(digits) if digits else None


def _is_wv_agent(key):
    e = _ext(key)
    return e is not None and WV_EXT_MIN <= e <= WV_EXT_MAX


def build_payload(qraw, araw, trend_inputs, scope):
    """Same filtering and arithmetic as _build_dispatch_payload in the generator."""
    def keep_q(name):
        return (name in WV_QUEUE_NAMES) if scope == "wv" else (name not in WV_QUEUE_NAMES)

    queues = []
    for qid, q in (qraw or {}).items():
        name = q.get("agent-name") or qid
        if not keep_q(name):
            continue
        queues.append(dict(
            id=qid, name=name,
            total=int(q.get("queue-calls-total-count", 0) or 0),
            answered=int(q.get("queue-calls-answered-count", 0) or 0),
            abandoned=int(q.get("queue-calls-abandoned-count", 0) or 0),
            avg_wait_sec=float(q.get("queue-calls-queueing-avg-duration-seconds", 0) or 0),
            avg_talk_min=float(q.get("queue-calls-talking-avg-duration-minutes", 0) or 0),
            service_level_pct=float(q.get("queue-calls-service-level-percent", 0) or 0)))
    queues.sort(key=lambda q: -q["total"])

    agents = []
    for aid, a in (araw or {}).items():
        if (_is_wv_agent(aid) if scope == "wv" else not _is_wv_agent(aid)):
            agents.append(dict(
                id=aid, name=a.get("agent-name") or aid,
                calls=int(a.get("queue-calls-answered-count", 0) or 0),
                talk_min=float(a.get("queue-calls-talking-avg-duration-minutes", 0) or 0),
                hold_sec=float(a.get("queue-calls-holding-avg-duration-seconds", 0) or 0),
                handle_min=float(a.get("queue-calls-handling-avg-duration-minutes", 0) or 0),
                total_talk_min=float(a.get("queue-calls-talking-duration-minutes", 0) or 0)))
    agents.sort(key=lambda a: -a["calls"])

    total = sum(q["total"] for q in queues)
    answered = sum(q["answered"] for q in queues)
    abandoned = sum(q["abandoned"] for q in queues)
    active = [q for q in queues if q["total"] > 0]
    summary = dict(
        total_calls=total, total_answered=answered, total_abandoned=abandoned,
        abandonment_pct=round(100 * abandoned / total, 1) if total else 0.0,
        weighted_avg_wait_sec=round(sum(q["avg_wait_sec"] * q["total"] for q in queues) / total, 1) if total else 0.0,
        avg_service_level_pct=round(sum(q["service_level_pct"] for q in active) / len(active), 1) if active else 0.0,
        active_queues=len(active))

    trend = dict(dates=[], total_calls=[], answered=[], abandoned=[], abandonment_pct=[], avg_wait_sec=[], service_level_pct=[])
    for ds, raw in trend_inputs:
        if raw is None:
            continue
        sc = [q for qid, q in raw.items() if keep_q(q.get("agent-name") or qid)]
        t = sum(int(q.get("queue-calls-total-count", 0) or 0) for q in sc)
        an = sum(int(q.get("queue-calls-answered-count", 0) or 0) for q in sc)
        ab = sum(int(q.get("queue-calls-abandoned-count", 0) or 0) for q in sc)
        ww = round(sum(float(q.get("queue-calls-queueing-avg-duration-seconds", 0) or 0)
                       * int(q.get("queue-calls-total-count", 0) or 0) for q in sc) / t, 1) if t else 0.0
        aq = [q for q in sc if int(q.get("queue-calls-total-count", 0) or 0) > 0]
        sl = round(sum(float(q.get("queue-calls-service-level-percent", 0) or 0) for q in aq) / len(aq), 1) if aq else 0.0
        trend["dates"].append(ds); trend["total_calls"].append(t); trend["answered"].append(an)
        trend["abandoned"].append(ab); trend["abandonment_pct"].append(round(100 * ab / t, 1) if t else 0.0)
        trend["avg_wait_sec"].append(ww); trend["service_level_pct"].append(sl)

    return dict(queues=queues, agents=agents, summary=summary, trend=trend)


def fetch_all():
    if not TOKEN:
        sys.exit("JIVETEL_API_TOKEN is not set in C:\\Reports\\.env")
    today = dt.datetime.now(TZ).date()
    yday = today - dt.timedelta(days=1)
    qraw = _get("/statistics/queue/per-queue", {"start_date": yday.isoformat(), "end_date": today.isoformat()})
    araw = _get("/statistics/agent", {"start_date": yday.isoformat(), "end_date": today.isoformat()})
    trend = []
    for off in range(6, -1, -1):
        d = (today - dt.timedelta(days=off)).isoformat()
        try:
            trend.append((d, _get("/statistics/queue/per-queue", {"start_date": d, "end_date": d})))
        except requests.RequestException:
            trend.append((d, None))
    return qraw, araw, trend, today, yday


# ── CDR store: trailing-24h hourly volume ────────────────────────────────

def hourly_from_cdr(scope, now):
    """Inbound calls per hour for the 24 hours ending now, from jivetel_cdr.db.
    Returns (labels, counts) or None when the store is missing or stale."""
    if not os.path.exists(CDR_DB):
        return None
    con = sqlite3.connect(CDR_DB)
    start = now - dt.timedelta(hours=24)
    rows = con.execute("""
        SELECT start_local, queues, last_term_user FROM call
        WHERE direction = 1 AND start_utc >= ? AND start_utc < ?""",
        (start.astimezone(dt.timezone.utc).isoformat(timespec="seconds"),
         now.astimezone(dt.timezone.utc).isoformat(timespec="seconds"))).fetchall()
    latest = con.execute("SELECT MAX(start_utc) FROM cdr_leg").fetchone()[0]
    con.close()
    if not rows:
        return None
    counts = {}
    for start_local, queues, term in rows:
        wv = (bool(queues) and any(q in WV_QUEUE_IDS for q in queues.split(","))) \
             or (term and term.isdigit() and WV_EXT_MIN <= int(term) <= WV_EXT_MAX)
        if (scope == "wv") != bool(wv):
            continue
        h = dt.datetime.fromisoformat(start_local).replace(minute=0, second=0, microsecond=0)
        counts[h] = counts.get(h, 0) + 1
    first = start.replace(minute=0, second=0, microsecond=0)
    hours = [first + dt.timedelta(hours=i) for i in range(25)]
    hours = [h for h in hours if h <= now]
    return ([h.strftime("%a %H:00") for h in hours], [counts.get(h, 0) for h in hours],
            dt.datetime.fromisoformat(latest).astimezone(TZ) if latest else None)


# ── Charts ───────────────────────────────────────────────────────────────

def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return buf.getvalue()


def chart_donut(queues):
    active = [q for q in queues if q["total"] > 0]
    fig, ax = plt.subplots(figsize=(5.2, 4.2))
    if not active:
        ax.text(0.5, 0.5, "No calls in the window", ha="center", va="center", color="#888")
        ax.axis("off")
        return _png(fig)
    top, rest = active[:8], active[8:]
    labels = [q["name"] for q in top] + ([f"Other ({len(rest)})"] if rest else [])
    values = [q["total"] for q in top] + ([sum(q["total"] for q in rest)] if rest else [])
    colors = DIST_PALETTE[:len(top)] + (["#cfcdc4"] if rest else [])
    ax.pie(values, colors=colors, startangle=90, counterclock=False,
           wedgeprops=dict(width=0.32, edgecolor="white", linewidth=2))
    ax.text(0, 0.06, f"{sum(values):,}", ha="center", va="center", fontsize=22, fontweight="bold", color="#222")
    ax.text(0, -0.14, "calls", ha="center", va="center", fontsize=10, color="#777")
    ax.legend(labels, loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=9)
    ax.set_aspect("equal")
    return _png(fig)


def chart_trends(t):
    fig, axes = plt.subplots(2, 2, figsize=(10, 5.6))
    labels = [d[5:] for d in t["dates"]]
    specs = [("Call volume", t["total_calls"], C_VOLUME, None),
             ("Abandonment %", t["abandonment_pct"], C_ABANDON, "%"),
             ("Avg wait (sec)", t["avg_wait_sec"], C_WAIT, "s"),
             ("Service level %", t["service_level_pct"], C_SL, "%")]
    for ax, (title, data, color, unit) in zip(axes.flat, specs):
        ax.plot(labels, data, color=color, linewidth=2.2, marker="o", markersize=4)
        ax.fill_between(labels, data, color=color, alpha=0.10)
        ax.set_title(title, fontsize=11, fontweight="bold", loc="left", color="#333")
        ax.grid(axis="y", color="#e6e6e6")
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        ax.tick_params(labelsize=8, colors="#555")
        if title.startswith("Service"):
            ax.set_ylim(0, 105)
        elif title.startswith("Abandon"):
            ax.set_ylim(0, max(5, max(data or [0]) * 1.25))
        else:
            ax.set_ylim(0, max(1, max(data or [0]) * 1.25))
        for x, y in zip(labels, data):
            ax.annotate(f"{y:g}{unit or ''}", (x, y), textcoords="offset points", xytext=(0, 6),
                        ha="center", fontsize=7, color="#444")
    fig.tight_layout(h_pad=1.6)
    return _png(fig)


def chart_hourly(labels, counts):
    fig, ax = plt.subplots(figsize=(10, 2.8))
    ax.bar(range(len(counts)), counts, color=C_VOLUME, width=0.7)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels([l if i % 3 == 0 else "" for i, l in enumerate(labels)], fontsize=8, color="#555")
    ax.grid(axis="y", color="#e6e6e6")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", labelsize=8, colors="#555")
    for i, c in enumerate(counts):
        if c:
            ax.annotate(str(c), (i, c), textcoords="offset points", xytext=(0, 3), ha="center", fontsize=7, color="#444")
    fig.tight_layout()
    return _png(fig)


# ── HTML ─────────────────────────────────────────────────────────────────

def _cls(val, red, yellow, bigger_is_better=False):
    if bigger_is_better:
        return "bad" if val < red else "mid" if val < yellow else "good"
    return "bad" if val > red else "mid" if val > yellow else "good"


def problem_queues(queues):
    out = []
    for q in queues:
        if q["total"] < MIN_CALLS_FOR_ATTENTION:
            continue
        issues = []
        if q["service_level_pct"] < TH["sl_red"]:
            issues.append(f"SLA {q['service_level_pct']:.1f}% (below {TH['sl_red']}%)")
        if q["abandoned"] and 100 * q["abandoned"] / q["total"] > TH["aban_red"]:
            issues.append(f"Abandonment {100 * q['abandoned'] / q['total']:.1f}% (over {TH['aban_red']}%)")
        if q["avg_wait_sec"] > TH["wait_red"]:
            issues.append(f"Avg wait {q['avg_wait_sec']:.1f}s (over {TH['wait_red']}s)")
        if issues:
            out.append((q, issues))
    return out


CSS = """
body{font-family:Segoe UI,Arial,sans-serif;color:#222;background:#f4f4f2;margin:0;padding:18px}
.wrap{max-width:960px;margin:0 auto;background:#fff;border:1px solid #e2e2de;border-radius:8px;overflow:hidden}
.hdr{background:#0f0f0d;color:#fff;padding:16px 22px;border-bottom:3px solid #C41E1E}
.hdr h1{margin:0;font-size:20px}.hdr .meta{color:#aaa;font-size:12px;margin-top:4px}
.sec{padding:16px 22px;border-top:1px solid #eee}.sec h2{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:#666;margin:0 0 10px}
.kpis{display:flex;flex-wrap:wrap;gap:10px}
.kpi{flex:1 1 120px;border:1px solid #e6e6e2;border-left:4px solid #bbb;border-radius:6px;padding:10px 12px;background:#fafaf8}
.kpi .l{font-size:11px;color:#777}.kpi .v{font-size:22px;font-weight:700;margin-top:2px}
.kpi.good{border-left-color:#0F6E56}.kpi.mid{border-left-color:#BA7517}.kpi.bad{border-left-color:#A32D2D}
table{border-collapse:collapse;width:100%;font-size:12.5px}th{background:#f1f1ee;text-align:left;padding:7px 8px;font-size:11px;color:#555;text-transform:uppercase;letter-spacing:.04em}
td{padding:6px 8px;border-top:1px solid #eee}.num{text-align:right;font-variant-numeric:tabular-nums}
td.good{color:#0F6E56;font-weight:600}td.mid{color:#BA7517;font-weight:600}td.bad{color:#A32D2D;font-weight:600}
.att{border:1px solid #f1c9c9;background:#fff6f6;border-radius:6px;padding:10px 12px;margin-bottom:8px}
.att b{color:#A32D2D}.att .i{font-size:12px;color:#555;margin-top:3px}
.ok{color:#0F6E56;font-size:13px}.note{font-size:11px;color:#888;margin-top:6px}
img{max-width:100%;height:auto;display:block}
.foot{padding:12px 22px;font-size:11px;color:#888;background:#fafaf8}
"""


def render_html(scope, p, imgs, today, yday, now, hourly_meta):
    s, q, a = p["summary"], p["queues"], p["agents"]
    active_q = [x for x in q if x["total"] > 0]
    hidden = len(q) - len(active_q)
    problems = problem_queues(q)
    kpi = lambda label, val, cls="": f'<div class="kpi {cls}"><div class="l">{label}</div><div class="v">{val}</div></div>'
    kpis = "".join([
        kpi("Total calls", f"{s['total_calls']:,}"),
        kpi("Answered", f"{s['total_answered']:,}", "good"),
        kpi("Abandoned", f"{s['total_abandoned']:,}", "bad" if s["total_abandoned"] else ""),
        kpi("Abandonment %", f"{s['abandonment_pct']:.1f}%", _cls(s["abandonment_pct"], TH["aban_red"], TH["aban_yellow"])),
        kpi("Avg wait", f"{s['weighted_avg_wait_sec']:.1f}s", _cls(s["weighted_avg_wait_sec"], TH["wait_red"], TH["wait_yellow"])),
        kpi("Service level", f"{s['avg_service_level_pct']:.1f}%", _cls(s["avg_service_level_pct"], TH["sl_red"], TH["sl_yellow"], True)),
    ])
    if problems:
        att = "".join(f'<div class="att"><b>{x["name"]}</b> &middot; {x["total"]} calls<div class="i">{" &nbsp;|&nbsp; ".join(iss)}</div></div>'
                      for x, iss in problems)
    else:
        att = '<div class="ok">No queue breached a threshold in this window.</div>'
    qrows = "".join(
        f'<tr><td>{x["name"]}</td><td class="num">{x["total"]:,}</td><td class="num">{x["answered"]:,}</td>'
        f'<td class="num">{x["abandoned"]:,}</td>'
        f'<td class="num {_cls(100 * x["abandoned"] / x["total"], TH["aban_red"], TH["aban_yellow"])}">{100 * x["abandoned"] / x["total"]:.1f}%</td>'
        f'<td class="num {_cls(x["avg_wait_sec"], TH["wait_red"], TH["wait_yellow"])}">{x["avg_wait_sec"]:.1f}</td>'
        f'<td class="num">{x["avg_talk_min"]:.2f}</td>'
        f'<td class="num {_cls(x["service_level_pct"], TH["sl_red"], TH["sl_yellow"], True)}">{x["service_level_pct"]:.1f}%</td></tr>'
        for x in active_q) or '<tr><td colspan="8" style="text-align:center;color:#888">No calls in any queue</td></tr>'
    arows = "".join(
        f'<tr><td>{x["name"]}</td><td class="num">{x["calls"]:,}</td><td class="num">{x["talk_min"]:.2f}</td>'
        f'<td class="num">{x["hold_sec"]:.1f}</td><td class="num">{x["handle_min"]:.2f}</td><td class="num">{x["total_talk_min"]:.1f}</td></tr>'
        for x in a if x["calls"] > 0) or '<tr><td colspan="6" style="text-align:center;color:#888">No agent took a queue call</td></tr>'
    hourly_block = ""
    if imgs.get("hourly"):
        stale = ""
        if hourly_meta and (now - hourly_meta) > dt.timedelta(hours=2):
            stale = f' Call store last updated {hourly_meta:%a %H:%M}; run jivetel_cdr_store.py incremental for a fuller picture.'
        hourly_block = f'''<div class="sec"><h2>Inbound calls by hour (trailing 24 hours)</h2>
      <img src="{imgs['hourly']}" alt="Calls by hour">
      <div class="note">From individual call records in the local CDR store, one count per inbound call. Queue statistics above count queue entries, so totals can differ slightly.{stale}</div></div>'''
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{BOARD_TITLE[scope]} snapshot {now:%Y-%m-%d %H:%M}</title><style>{CSS}</style></head>
<body><div class="wrap">
  <div class="hdr"><h1><span style="color:{ACCENT[scope]}">{BOARD_TITLE[scope].split(' Dispatch')[0]}</span> Dispatch &mdash; 24-hour snapshot</h1>
    <div class="meta">Window: {yday:%a %b %d} through {now:%a %b %d, %H:%M} ET &nbsp;&middot;&nbsp; Same window and sources as the live board &nbsp;&middot;&nbsp; Generated {now:%Y-%m-%d %H:%M}</div></div>
  <div class="sec"><div class="kpis">{kpis}</div></div>
  <div class="sec"><h2>Call distribution by queue</h2><img src="{imgs['donut']}" alt="Call distribution"></div>
  <div class="sec"><h2>Needs attention</h2>{att}
    <div class="note">Thresholds match the dashboard: service level below {TH['sl_red']}%, abandonment over {TH['aban_red']}%, average wait over {TH['wait_red']}s; queues under {MIN_CALLS_FOR_ATTENTION} calls are not flagged.</div></div>
  <div class="sec"><h2>7-day trends</h2><img src="{imgs['trends']}" alt="7-day trends"></div>
  {hourly_block}
  <div class="sec"><h2>Queues</h2>
    <table><tr><th>Queue</th><th class="num">Total</th><th class="num">Answered</th><th class="num">Abandoned</th><th class="num">Aban %</th><th class="num">Avg wait (s)</th><th class="num">Avg talk (min)</th><th class="num">Service level</th></tr>{qrows}</table>
    {f'<div class="note">{hidden} idle queue{"s" if hidden != 1 else ""} with no calls hidden.</div>' if hidden else ''}</div>
  <div class="sec"><h2>Agents who took calls</h2>
    <table><tr><th>Agent</th><th class="num">Calls</th><th class="num">Avg talk (min)</th><th class="num">Avg hold (s)</th><th class="num">Avg handle (min)</th><th class="num">Total talk (min)</th></tr>{arows}</table></div>
  <div class="foot">Source: Jivetel queue and agent statistics for {yday:%Y-%m-%d} and {today:%Y-%m-%d}, the same two-day range the dashboard labels "last 24h". Turn-downs and live agent status are not included in this report.</div>
</div></body></html>"""


# ── Email ────────────────────────────────────────────────────────────────

def send_email(subject, html_by_cid, pngs, recipients):
    user = os.environ.get("SMTP_USER"); pw = os.environ.get("SMTP_PASS") or os.environ.get("SMTP_PASSWORD")
    if not user or not pw:
        sys.exit("SMTP_USER / SMTP_PASS are not set in C:\\Reports\\.env; report was written but not emailed")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = os.environ.get("SMTP_FROM", user)
    msg["To"] = ", ".join(recipients)
    msg.set_content("This report is HTML. Open in an HTML-capable mail client.")
    msg.add_alternative(html_by_cid, subtype="html")
    for cid, png in pngs.items():
        msg.get_payload()[1].add_related(png, maintype="image", subtype="png", cid=cid)
    server = os.environ.get("SMTP_SERVER", "smtp.office365.com"); port = int(os.environ.get("SMTP_PORT", "587"))
    with smtplib.SMTP(server, port, timeout=30) as s:
        s.ehlo(); s.starttls(); s.ehlo(); s.login(user, pw); s.send_message(msg)


# ── Main ─────────────────────────────────────────────────────────────────

def build(scope, qraw, araw, trend, today, yday, now):
    p = build_payload(qraw, araw, trend, scope)
    pngs = {"donut": chart_donut(p["queues"]), "trends": chart_trends(p["trend"])}
    hourly = hourly_from_cdr(scope, now)
    hourly_meta = None
    if hourly:
        labels, counts, hourly_meta = hourly
        pngs["hourly"] = chart_hourly(labels, counts)
    b64 = {k: "data:image/png;base64," + base64.b64encode(v).decode() for k, v in pngs.items()}
    cids = {k: make_msgid(domain="lynx-dispatch") for k in pngs}
    html_file = render_html(scope, p, b64, today, yday, now, hourly_meta)
    html_mail = render_html(scope, p, {k: "cid:" + cids[k][1:-1] for k in pngs}, today, yday, now, hourly_meta)
    return p, html_file, html_mail, {cids[k][1:-1]: v for k, v in pngs.items()}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--scope", choices=["main", "wv", "both"], default="main")
    ap.add_argument("--email", nargs="+", metavar="ADDR", help="send the report to these addresses")
    ap.add_argument("--out", default=OUT_DIR)
    a = ap.parse_args()

    now = dt.datetime.now(TZ)
    qraw, araw, trend, today, yday = fetch_all()
    os.makedirs(a.out, exist_ok=True)
    for scope in (["main", "wv"] if a.scope == "both" else [a.scope]):
        p, html_file, html_mail, pngs = build(scope, qraw, araw, trend, today, yday, now)
        path = os.path.join(a.out, f"dispatch_snapshot_{scope}_{now:%Y%m%d_%H%M}.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html_file)
        s = p["summary"]
        print(f"{scope}: {path}  calls={s['total_calls']} answered={s['total_answered']} "
              f"abandoned={s['total_abandoned']} wait={s['weighted_avg_wait_sec']}s SL={s['avg_service_level_pct']}%")
        if a.email:
            send_email(f"{BOARD_TITLE[scope]} - 24-hour snapshot {now:%b %d %H:%M}", html_mail, pngs, a.email)
            print(f"   emailed to {', '.join(a.email)}")


if __name__ == "__main__":
    main()
