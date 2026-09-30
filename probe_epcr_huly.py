"""
What, if anything, can the ePCR (Huly) surface be read for?

OTP lost its original "arrived" value when direct database access went away:
the old SQL scored against ePCR field 549, and the ThirdParty spec puts ePCR
out of scope, listing `ThirdParty/Data/Epcr/Huly` and `Trip?rtype=HulyUpdateTrip`
under "Not included in this spec -- private or non-partner integrations". An
earlier read of Epcr/Huly answered 501.

501 from one unparameterised GET is weak evidence. These endpoints dispatch on
an `rtype` query parameter, and a surface that answers 501 to a bare call can
still answer 200 to a named action -- which is exactly the pattern Cad/Trip
follows (no rtype returns a single object, rtype=GetTrips returns the array).
So this enumerates named read actions and records what each one answers.

If any of them returns rows carrying a scene-arrival time, OTP can be scored
against the same value the historical numbers used, and the series becomes
comparable across the changeover instead of starting fresh.

STRICTLY READ-ONLY. It issues GET only, and the write-shaped `HulyUpdateTrip`
and `SetTimestamps` actions are deliberately excluded from every candidate
list below -- do not add them. A 501 or 404 here costs nothing; a stray write
against ePCR does not.

Usage:
    python probe_epcr_huly.py [YYYY-MM-DD] [--out DIR]

Writes a findings file (and any non-empty response bodies) to the output
directory, default api_probe/. Send that file to Traumasoft support if every
action refuses: it is the evidence for asking which credentials or endpoint
expose ePCR timestamps to a partner.
"""

import sys
import json
import logging
from pathlib import Path
from datetime import datetime, timedelta, date

from traumasoft_api import TraumasoftAPI, TraumasoftAPIError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("epcr-probe")

# Read-shaped action names to try against the ePCR surface. Every one is a
# guess: the spec documents none of them. Names follow the conventions the
# documented surfaces use (GetX / ListX / ReadX), since a dispatcher that
# accepts GetTrips is likely to accept GetRuns.
#
# Nothing here may create, update or delete. HulyUpdateTrip and SetTimestamps
# are excluded on purpose.
HULY_RTYPES = [
    None,                    # bare GET -- reproduces the original 501
    "GetRuns",
    "GetRun",
    "GetTrips",
    "GetTrip",
    "GetRunTimes",
    "GetTimestamps",
    "GetTimes",
    "GetFields",
    "GetFieldValues",
    "GetValues",
    "GetForm",
    "GetForms",
    "GetPcr",
    "GetPcrs",
    "GetReports",
    "GetIncidents",
    "GetPatients",
    "GetTypes",
    "GetConfigs",
    "List",
]

# Other paths worth one GET each: if ePCR data reaches partners at all, it may
# be hanging off a surface that is in the spec rather than off Huly.
NEIGHBOURING_PATHS = [
    ("ThirdParty/Data/Epcr", None),
    ("ThirdParty/Data/Epcr/Runs", None),
    ("ThirdParty/Data/Epcr/Huly", "GetRuns"),
    ("ThirdParty/Lists/Epcr/Huly", None),
    ("ThirdParty/Lists/Epcr/Fields", None),
    ("ThirdParty/Data/Cad/Trip", "GetTripDetails"),
    ("ThirdParty/Data/Cad/Trip", "GetTimestamps"),
    ("ThirdParty/Data/Cad/Timestamps", None),
]

# Field names that would mean the surface carries what OTP needs. Field 549 is
# the historical one; the rest are how it might be labelled in a JSON payload.
ARRIVAL_FIELD_HINTS = [
    "549", "on_scene", "onscene", "at_scene", "arrive", "arrival",
    "scene", "patient_contact", "bedside",
]


def parse_args(argv):
    args = {"day": None, "out": "api_probe"}
    rest = []
    i = 0
    while i < len(argv):
        token = argv[i]
        if token == "--out":
            i += 1
            args["out"] = argv[i]
        elif token.startswith("--"):
            raise SystemExit(f"Unknown option: {token}")
        else:
            rest.append(token)
        i += 1
    args["day"] = (
        datetime.strptime(rest[0], "%Y-%m-%d").date() if rest
        else date.today() - timedelta(days=1)
    )
    return args


def try_get(api, path, params):
    """One GET. Returns (outcome, status, payload_or_message)."""
    try:
        payload = api.get(path, params=params)
    except TraumasoftAPIError as exc:
        return "refused", exc.status_code, str(exc)[:400]
    except Exception as exc:  # noqa: BLE001 - a probe reports, it does not crash
        return "error", None, f"{type(exc).__name__}: {exc}"[:400]
    if payload is None:
        return "empty", 200, None
    return "ok", 200, payload


def looks_like_arrival(payload):
    """Does anything in this payload resemble a scene-arrival field?"""
    blob = json.dumps(payload, default=str).lower()
    return sorted({hint for hint in ARRIVAL_FIELD_HINTS if hint in blob})


def describe(payload):
    if isinstance(payload, list):
        first = payload[0] if payload else None
        return {
            "shape": "array",
            "rows": len(payload),
            "keys": sorted(first.keys()) if isinstance(first, dict) else None,
        }
    if isinstance(payload, dict):
        return {"shape": "object", "rows": 1, "keys": sorted(payload.keys())}
    return {"shape": type(payload).__name__, "rows": None, "keys": None}


def main():
    args = parse_args(sys.argv[1:])
    out_dir = Path(args["out"])
    out_dir.mkdir(parents=True, exist_ok=True)

    # One attempt per candidate: most of these are expected to refuse, and
    # retrying a 501 four times with backoff would make the sweep take minutes.
    try:
        api = TraumasoftAPI(max_retries=0)
    except ValueError as exc:
        log.error("Configuration error: %s", exc)
        return 2
    if api.detect_auth_mode() is None:
        log.error("Could not authenticate. Check TS_API_KEY / TS_API_SECRET.")
        return 1

    day = args["day"].isoformat()
    findings = {"date": day, "base_url": api.base_url, "attempts": []}
    promising = []

    def record(path, rtype, params):
        outcome, status, payload = try_get(api, path, params)
        label = f"{path}?rtype={rtype}" if rtype else path
        entry = {
            "path": path, "rtype": rtype, "params": params,
            "outcome": outcome, "status": status,
        }
        if outcome == "ok":
            entry.update(describe(payload))
            hits = looks_like_arrival(payload)
            entry["arrival_field_hints"] = hits
            log.info("  %-46s -> 200 %s", label, entry["shape"])
            if hits:
                log.info("      ^ contains %s", ", ".join(hits))
            promising.append((label, payload))
        elif outcome == "empty":
            log.info("  %-46s -> 200 (no body)", label)
        else:
            entry["message"] = payload
            log.info("  %-46s -> %s", label, status or "error")
        findings["attempts"].append(entry)

    log.info("Sweeping ThirdParty/Data/Epcr/Huly rtypes ...")
    for rtype in HULY_RTYPES:
        params = {"trip_date": day}
        if rtype:
            params["rtype"] = rtype
        record("ThirdParty/Data/Epcr/Huly", rtype, params)

    log.info("")
    log.info("Sweeping neighbouring paths ...")
    for path, rtype in NEIGHBOURING_PATHS:
        params = {"trip_date": day}
        if rtype:
            params["rtype"] = rtype
        record(path, rtype, params)

    statuses = {}
    for entry in findings["attempts"]:
        key = str(entry["status"] or entry["outcome"])
        statuses[key] = statuses.get(key, 0) + 1
    findings["status_summary"] = statuses

    log.info("")
    log.info("--- Summary ---")
    for status, count in sorted(statuses.items()):
        log.info("  %s: %s attempt(s)", status, count)

    if promising:
        log.info("")
        log.info("%s call(s) returned data. Samples written for inspection.", len(promising))
        for label, payload in promising:
            safe = label.replace("/", "_").replace("?", "_").replace("=", "-")
            sample = out_dir / f"epcr_{safe}.json"
            sample.write_text(json.dumps(payload, indent=2, default=str)[:2_000_000],
                              encoding="utf-8")
            log.info("  %s", sample)
    else:
        log.info("")
        log.info("Nothing readable. Every ePCR action refused, which is consistent")
        log.info("with the spec: this surface is not exposed to partner credentials.")
        log.info("The findings file is the evidence to send Traumasoft support when")
        log.info("asking what does expose an ePCR scene-arrival time.")

    path = out_dir / f"epcr_huly_probe_{day}.json"
    path.write_text(json.dumps(findings, indent=2, default=str), encoding="utf-8")
    log.info("")
    log.info("Wrote %s", path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
