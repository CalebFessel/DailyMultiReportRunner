"""
Validate the API-backed OTP report against the historical append workbook.

Direct database access is gone, so a side-by-side run against the old SQL is
impossible. The append workbooks are the next best thing: they hold what the
SQL actually produced, day by day, for up to 730 days. Trips backfill roughly
90 days through the API, and that overlap is enough to answer the two questions
that decide whether the migrated report can be trusted.

  1. Which timestamp replaces ePCR field 549?
     The current OTP scores against an ePCR field this API does not expose.
     `at_scene` and `at_scene: At Patient Bedside` are both candidates. Rather
     than picking one and hoping, this recomputes OTP under each and reports
     which reproduces the historical numbers most closely.

  2. Does the API see the same population the SQL did?
     Only 233 of 374 completed legs on the probe day carried a scheduled
     pickup_time. If the historical run counts match the API's, the old SQL was
     filtering the same way. If history is consistently larger, it was scoring
     legs this rebuild silently drops -- which would quietly halve a
     denominator.

Read-only against both the API and the workbook.

Usage:
    python validate_against_history.py \
        --append "Reports/Append/CompanyWide_OTP_APPEND.xlsx" \
        --days 90
"""

import os
import sys
import json
import logging
import argparse
from pathlib import Path
from datetime import datetime, timedelta, date
from collections import defaultdict

import pandas as pd

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError
import traumasoft_reports as R

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("validate")

DEFAULT_APPEND = os.path.join("Reports", "Append", "CompanyWide_OTP_APPEND.xlsx")
COST_CENTER_SHEET = "OTP by Cost Center"

# Candidate arrival stamps, each tried on its own.
ARRIVAL_CANDIDATES = [
    ["at_scene: At Patient Bedside"],
    ["at_scene"],
    ["at_scene: At Patient Bedside", "at_scene"],
]

# GetTrips accepts an inclusive range capped at 31 days.
TRIP_CHUNK_DAYS = 31


def load_history(path, sheet=COST_CENTER_SHEET):
    """Read the append workbook's per-day, per-cost-center OTP rows."""
    if not Path(path).exists():
        raise FileNotFoundError(
            f"Append workbook not found: {path}\n"
            "Point --append at CompanyWide_OTP_APPEND.xlsx from the old runs."
        )
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    if "snapshot_date" not in df.columns:
        raise ValueError(f"'{sheet}' has no snapshot_date column; columns are {list(df.columns)}")
    df["snapshot_date"] = pd.to_datetime(df["snapshot_date"], errors="coerce").dt.date
    return df.dropna(subset=["snapshot_date"])


def fetch_legs_by_date(api, start_date, end_date):
    """
    Pull every leg between two dates and bucket it by scheduled pickup date.

    Chunked at the API's 31-day range cap, so 90 days costs three calls rather
    than ninety.
    """
    by_date = defaultdict(list)
    cursor = start_date
    while cursor <= end_date:
        span = min(TRIP_CHUNK_DAYS, (end_date - cursor).days + 1)
        log.info("Fetching trips %s +%s days ...", cursor, span)
        try:
            legs = api.get_trips(cursor, range_days=span)
        except TraumasoftAPIError as exc:
            log.error("  failed: %s", exc)
            legs = []
        for leg in legs:
            pickup = R.parse_ts(leg.get("pickup_time"))
            if pickup:
                by_date[pickup.date()].append(leg)
        log.info("  %s legs", len(legs))
        cursor += timedelta(days=span)
    return by_date


def compare_day(legs, history_rows, cost_center_map, arrival_keys):
    """Compare one day's API-derived OTP against the historical rows."""
    scored = R.scored_legs(legs, cost_center_map, arrival_keys=arrival_keys)
    api_df = R.build_otp_by_cost_center(scored)

    api_runs = int(api_df["total_runs"].sum()) if not api_df.empty else 0
    hist_runs = int(history_rows["total_runs"].sum()) if not history_rows.empty else 0

    # Company-wide on-time %, weighted by runs, is the headline number people
    # actually watch -- compare that rather than per-cost-center noise.
    api_pct = None
    if api_runs:
        on_time = int(api_df["on_time_runs"].sum() + api_df["early_runs"].sum())
        api_pct = round(100.0 * on_time / api_runs, 2)

    hist_pct = None
    if hist_runs and {"on_time_runs", "early_runs"} <= set(history_rows.columns):
        hist_on_time = int(history_rows["on_time_runs"].sum() + history_rows["early_runs"].sum())
        hist_pct = round(100.0 * hist_on_time / hist_runs, 2)

    return {
        "api_runs": api_runs,
        "hist_runs": hist_runs,
        "run_delta": api_runs - hist_runs,
        "api_on_time_pct": api_pct,
        "hist_on_time_pct": hist_pct,
        "pct_delta": (round(api_pct - hist_pct, 2) if api_pct is not None and hist_pct is not None else None),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--append", default=DEFAULT_APPEND, help="CompanyWide_OTP_APPEND.xlsx")
    parser.add_argument("--days", type=int, default=90, help="how far back to compare")
    parser.add_argument("--out", default="api_probe", help="output directory")
    args = parser.parse_args()

    Path(args.out).mkdir(parents=True, exist_ok=True)

    try:
        history = load_history(args.append)
    except (FileNotFoundError, ValueError) as exc:
        log.error("%s", exc)
        return 2

    end_date = date.today() - timedelta(days=1)
    start_date = end_date - timedelta(days=args.days)
    history = history[
        (history["snapshot_date"] >= start_date) & (history["snapshot_date"] <= end_date)
    ]
    overlap_dates = sorted(history["snapshot_date"].unique())
    if not overlap_dates:
        log.error(
            "No historical rows between %s and %s. Widen --days or check the workbook.",
            start_date, end_date,
        )
        return 2

    log.info("Comparing %s days of history (%s .. %s)",
             len(overlap_dates), overlap_dates[0], overlap_dates[-1])

    try:
        api = TraumasoftAPI()
    except ValueError as exc:
        log.error("%s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Credentials rejected under every signing scheme.")
        return 1

    # The cost-center map must be warm before historical legs can be attributed,
    # and it is only learnable from the present shift window.
    cost_center_map = R.CostCenterMap()
    log.info("Warming the cost-center map from the current shift window ...")
    cost_center_map.update(api.list_shifts(), api.list_employees())
    cost_center_map.save()
    log.info("  %s shift profiles known", len(cost_center_map.counts))

    legs_by_date = fetch_legs_by_date(api, overlap_dates[0], overlap_dates[-1])

    results = {}
    for candidate in ARRIVAL_CANDIDATES:
        label = " -> ".join(candidate)
        rows = []
        for day in overlap_dates:
            legs = legs_by_date.get(day, [])
            if not legs:
                continue
            comparison = compare_day(
                legs, history[history["snapshot_date"] == day], cost_center_map, candidate
            )
            comparison["date"] = day
            rows.append(comparison)
        results[label] = rows

    # --- summarise ---
    summary = []
    for label, rows in results.items():
        scored = [r for r in rows if r["pct_delta"] is not None]
        if not scored:
            summary.append({"arrival_keys": label, "days_compared": 0})
            continue
        mean_abs_pct = sum(abs(r["pct_delta"]) for r in scored) / len(scored)
        mean_run_delta = sum(r["run_delta"] for r in scored) / len(scored)
        mean_abs_run = sum(abs(r["run_delta"]) for r in scored) / len(scored)
        summary.append({
            "arrival_keys": label,
            "days_compared": len(scored),
            "mean_abs_on_time_pct_error": round(mean_abs_pct, 2),
            "mean_run_delta": round(mean_run_delta, 1),
            "mean_abs_run_delta": round(mean_abs_run, 1),
        })

    ranked = sorted(
        [s for s in summary if s.get("days_compared")],
        key=lambda s: s["mean_abs_on_time_pct_error"],
    )

    lines = ["# OTP validation against historical append workbook", ""]
    lines.append(f"Source: `{args.append}`")
    lines.append(f"Window: **{overlap_dates[0]} .. {overlap_dates[-1]}** ({len(overlap_dates)} days with history)")
    lines.append("")
    lines.append("## Which arrival timestamp reproduces history best?")
    lines.append("")
    lines.append("| Arrival stamp | Days | Mean abs. on-time % error | Mean run delta | Mean abs. run delta |")
    lines.append("|---|---|---|---|---|")
    for s in ranked:
        lines.append(
            f"| `{s['arrival_keys']}` | {s['days_compared']} | {s['mean_abs_on_time_pct_error']} | "
            f"{s['mean_run_delta']:+} | {s['mean_abs_run_delta']} |"
        )
    lines.append("")

    if ranked:
        best = ranked[0]
        lines.append(f"**Best fit: `{best['arrival_keys']}`** "
                     f"({best['mean_abs_on_time_pct_error']} percentage points mean absolute error).")
        lines.append("")
        lines.append("Set it with:")
        lines.append("")
        lines.append("```")
        lines.append(f"TS_ARRIVAL_TIMESTAMP_KEYS={best['arrival_keys'].replace(' -> ', ',')}")
        lines.append("```")
        lines.append("")
        if abs(best["mean_run_delta"]) > 1:
            direction = "fewer" if best["mean_run_delta"] < 0 else "more"
            lines.append(
                f"Note the API scores **{abs(best['mean_run_delta']):.1f} {direction} runs per day** on "
                "average than the SQL did. A consistent shortfall means the old query scored legs this "
                "rebuild drops -- most likely legs without a scheduled `pickup_time`. Investigate before "
                "trusting the percentages, since it moves the denominator."
            )
            lines.append("")

    lines.append("## Day by day (best-fitting stamp)")
    lines.append("")
    if ranked:
        lines.append("| Date | API runs | History runs | Delta | API on-time % | History on-time % | Delta |")
        lines.append("|---|---|---|---|---|---|---|")
        for row in results[ranked[0]["arrival_keys"]]:
            pct_delta = row["pct_delta"]
            pct_delta_text = "n/a" if pct_delta is None else f"{pct_delta:+}"
            lines.append(
                f"| {row['date']} | {row['api_runs']} | {row['hist_runs']} | {row['run_delta']:+} | "
                f"{row['api_on_time_pct']} | {row['hist_on_time_pct']} | {pct_delta_text} |"
            )
        lines.append("")

    report_path = Path(args.out) / "OTP_VALIDATION.md"
    report_path.write_text("\n".join(lines), encoding="utf-8")
    (Path(args.out) / "otp_validation.json").write_text(
        json.dumps({"summary": summary, "detail": {k: v for k, v in results.items()}},
                   indent=2, default=str),
        encoding="utf-8",
    )

    log.info("Wrote %s", report_path)
    for s in ranked:
        log.info("  %-45s  %.2f pp error, %+.1f runs/day",
                 s["arrival_keys"], s["mean_abs_on_time_pct_error"], s["mean_run_delta"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
