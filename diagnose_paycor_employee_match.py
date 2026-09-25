"""
Explain, person by person, why a Traumasoft employee_num found no Paycor match.

The dry run can only say "matches no Paycor employee". That single sentence
covers several different problems -- a person genuinely absent from the Paycor
roster, a person present under a different number, a Traumasoft record holding
a username where a payroll number belongs -- and they need different fixes.
This separates them by falling back to the one identifier both systems
independently carry: the person's name.

Strictly read-only. It proposes overrides into a *candidate* file; it never
writes the one push_paycor_timecards.py actually loads. A wrong override pays
the wrong person, so the promotion from candidate to live stays a human act.

    python diagnose_paycor_employee_match.py
"""

import os
import re
import sys
import json
from collections import defaultdict
from datetime import datetime

# traumasoft_api loads .env as an import side effect, and paycor_api reads
# PAYCOR_ENVIRONMENT at import time to choose its base url. This order or a
# .env saying production quietly talks to the sandbox.
import traumasoft_api  # noqa: F401
import paycor_api
from traumasoft_api import TraumasoftAPI, TraumasoftAPIError
from paycor_api import PaycorClient, PaycorAPIError, PaycorAuthError

import push_paycor_timecards as P
import traumasoft_reports as R

CANDIDATE_FILE = os.path.join(P.STATE_DIR, "paycor_employee_overrides.candidate.json")

ID_FIELDS = ("employeeNumber", "alternateEmployeeNumber", "badgeNumber")


def norm_name(first, last):
    """Compare names without punctuation, case or spacing getting in the way."""
    raw = f"{last or ''} {first or ''}".lower()
    raw = re.sub(r"[^a-z\s]", " ", raw)
    return " ".join(raw.split())


def name_keys(first, last):
    """
    Progressively looser keys: full name, then last + first initial.

    The second catches Bob/Robert, which is the common reason a name match
    fails between a scheduling system and a payroll system.
    """
    full = norm_name(first, last)
    if not full:
        return []
    parts = full.split()
    keys = [full]
    if len(parts) >= 2:
        keys.append(f"{parts[0]} {parts[1][0]}")
    return keys


def main():
    print("=" * 78)
    print("PAYCOR EMPLOYEE MATCH DIAGNOSIS")
    print(f"environment: {paycor_api.ENVIRONMENT}")
    print("=" * 78)

    ts = TraumasoftAPI()
    try:
        ts.detect_auth_mode()
    except TraumasoftAPIError as exc:
        print(f"Could not authenticate to Traumasoft: {exc}")
        return 1

    print("Reading both rosters...")
    shifts = ts.list_shifts()
    ts_employees = ts.list_employees()
    legs = ts.get_trips(datetime.now().date(), range_days=1)
    offset = R.resolve_shift_offset(legs)
    now = R.tenant_now(offset)

    try:
        paycor = PaycorClient(read_only=True)
        paycor_roster = paycor.list_employees()
        activity_types = paycor.list_activity_types()
    except (PaycorAuthError, PaycorAPIError) as exc:
        print(f"Could not read from Paycor: {exc}")
        return 1

    print(f"Traumasoft employees: {len(ts_employees)}")
    print(f"Paycor employees:     {len(paycor_roster)}")

    # ---- What identifiers does the Paycor roster actually carry? ----
    print("\n" + "-" * 78)
    print("PAYCOR IDENTIFIER COVERAGE")
    print("-" * 78)
    for field in ID_FIELDS:
        present = [str(e.get(field)).strip() for e in paycor_roster
                   if str(e.get(field) or "").strip()]
        sample = ", ".join(present[:5])
        print(f"  {field:<26} {len(present):>5} of {len(paycor_roster)}"
              f"{'   e.g. ' + sample if sample else ''}")

    # A terminated-employee filter on the roster read would explain absences
    # far more simply than anything to do with the numbers themselves.
    status_fields = defaultdict(lambda: defaultdict(int))
    for emp in paycor_roster:
        for key in ("status", "employmentStatus", "employeeStatus", "isActive"):
            if key in emp:
                status_fields[key][str(emp.get(key))] += 1
    if status_fields:
        print("\n  Status fields on the roster:")
        for key, counts in status_fields.items():
            summary = ", ".join(f"{v}={c}" for v, c in
                                sorted(counts.items(), key=lambda b: -b[1])[:6])
            print(f"    {key:<24} {summary}")
    else:
        print("\n  The roster carries no status field, so whether it includes")
        print("  terminated employees cannot be told from here.")

    # ---- Reproduce the dry run's matching, then explain the failures ----
    index = P.paycor_employee_index(paycor_roster)
    overrides = P.load_overrides()
    activity_type_id, _ = P.resolve_activity_type(activity_types)
    sendable, refused = P.build_plan(
        shifts, ts_employees, offset, now, overrides, index, activity_type_id,
    )

    unmatched = {}
    for row in refused:
        reason = row.get("refused") or ""
        if "matches no Paycor employee" not in reason:
            continue
        num = str(row.get("employee_num") or "").strip()
        if num:
            unmatched.setdefault(num, row.get("employee_name"))

    if not unmatched:
        print("\nNo unmatched employee numbers in the current window.")
        return 0

    # ---- narrow the roster to the departments this operation actually uses ----
    #
    # The Paycor tenant is shared with the parent company, so most of this
    # roster never appears in Traumasoft. Matching a name against all of it
    # draws from thousands of unrelated people, and a surname plus an initial
    # is not rare across that many.
    #
    # Employees who matched on their payroll NUMBER are known to be ours, so
    # the departments they sit in describe this operation. Restricting the
    # name fallback to those departments cannot pay someone at the parent
    # company by accident.
    #
    # The failure mode is deliberately one-sided: one of ours in a department
    # no number-matched employee occupies is excluded and reported as absent,
    # which refuses a punch. The opposite error pays the wrong person.
    ours_departments = set()
    for emp in ts_employees:
        entry, how = P.resolve_employee(emp, overrides, index)
        if entry is not None and how == "employee_num":
            dept = entry.get("department_id")
            if dept:
                ours_departments.add(str(dept))

    def in_scope(emp):
        if not ours_departments:
            return True
        dept = (emp.get("department") or {}).get("id")
        return dept is not None and str(dept) in ours_departments

    in_scope_roster = [e for e in paycor_roster if in_scope(e)]
    print("\n" + "-" * 78)
    print("ROSTER SCOPE")
    print("-" * 78)
    print(f"  departments in use by employees matched on payroll number: "
          f"{len(ours_departments)}")
    print(f"  Paycor employees in those departments: {len(in_scope_roster)} "
          f"of {len(paycor_roster)}")
    if not ours_departments:
        print("\n  No employee matched on a payroll number, so the departments")
        print("  this operation uses cannot be inferred and the whole roster is")
        print("  in scope. Read every name match with that in mind.")
    else:
        print("\n  Name matching is restricted to these departments. The rest of")
        print("  the tenant belongs to the parent company, and a name match")
        print("  against it would put someone else's employee on this payroll.")

    # Name -> Paycor entries, for the fallback match.
    by_name = defaultdict(list)
    for emp in in_scope_roster:
        if not emp.get("id"):
            continue
        entry = {
            "id": str(emp["id"]),
            "name": " ".join(p for p in (emp.get("firstName"), emp.get("lastName")) if p),
            "number": next((str(emp.get(f)).strip() for f in ID_FIELDS
                            if str(emp.get(f) or "").strip()), None),
        }
        for key in name_keys(emp.get("firstName"), emp.get("lastName")):
            by_name[key].append(entry)

    ts_by_num = {}
    for emp in ts_employees:
        num = str(emp.get("employee_num") or "").strip()
        if num:
            ts_by_num.setdefault(num.lower(), emp)

    confident, ambiguous, absent, nameless = {}, [], [], []
    matched_detail = {}
    loose_matches = []

    for num, name in sorted(unmatched.items()):
        ts_emp = ts_by_num.get(num.lower()) or {}
        keys = name_keys(ts_emp.get("first_name"), ts_emp.get("last_name"))
        if not keys:
            nameless.append((num, name))
            continue

        hits = []
        loose = False
        for index, key in enumerate(keys):
            hits = by_name.get(key) or []
            if hits:
                # keys[0] is the full name; anything after it matched on last
                # name plus first initial, which is a different claim.
                loose = index > 0
                break

        unique = {h["id"]: h for h in hits}
        if len(unique) == 1:
            hit = next(iter(unique.values()))
            confident[num.lower()] = hit["id"]
            matched_detail[num.lower()] = dict(hit, loose=loose)
            if loose:
                loose_matches.append((num, ts_emp, hit))
        elif len(unique) > 1:
            ambiguous.append((num, name, [h["name"] for h in unique.values()]))
        else:
            absent.append((num, name))

    print("\n" + "-" * 78)
    print(f"UNMATCHED EMPLOYEE NUMBERS IN THIS WINDOW: {len(unmatched)}")
    print("-" * 78)

    if confident:
        print(f"\n  Matched by name instead ({len(confident)}) -- the person is in Paycor,")
        print("  under a number Traumasoft does not carry:\n")
        for num, guid in sorted(confident.items()):
            hit = matched_detail.get(num, {})
            ts_name = next((n for k, n in unmatched.items() if k.lower() == num), None)
            mark = "  <-- CHECK" if hit.get("loose") else ""
            print(f"    TS {num:<14} {ts_name or '?':<28}"
                  f" -> Paycor {hit.get('name', '?')} (#{hit.get('number')}){mark}")

    if loose_matches:
        print(f"\n  {len(loose_matches)} of those matched on LAST NAME + FIRST INITIAL,")
        print("  not on the full name. That rule is for Bob against Robert, and it")
        print("  also accepts Christina against Christiana. Read these before")
        print("  promoting them; the rest of the table is an exact name match:\n")
        for num, ts_emp, hit in loose_matches:
            ts_full = " ".join(p for p in (ts_emp.get("first_name"),
                                           ts_emp.get("last_name")) if p)
            print(f"    TS {num:<14} {ts_full:<28} -> {hit['name']} (#{hit['number']})")

    if ambiguous:
        print(f"\n  Name matches more than one Paycor employee ({len(ambiguous)}).")
        print("  These need a human; a guess here pays the wrong person:\n")
        for num, name, names in ambiguous:
            print(f"    TS {num:<14} {name or '?':<28} -> {', '.join(names)}")

    if absent:
        print(f"\n  Not in the Paycor roster under any number or name ({len(absent)}).")
        print("  Either not yet set up in payroll, or terminated and filtered out:\n")
        for num, name in absent:
            print(f"    TS {num:<14} {name or '?'}")

    if nameless:
        print(f"\n  No name on the Traumasoft record ({len(nameless)}), so nothing to")
        print("  match on. Fix the Traumasoft record:\n")
        for num, name in nameless:
            print(f"    TS {num:<14} {name or '?'}")

    if confident:
        loose_keys = sorted(n.lower() for n, _e, _h in loose_matches)
        payload = {
            "_verify_first": loose_keys,
            "_comment": (
                "CANDIDATES ONLY, produced by diagnose_paycor_employee_match.py. "
                "Each entry was matched by name, not by number. Verify every line "
                "against payroll before copying it into "
                "paycor_employee_overrides.json -- a wrong guid pays the wrong person. The keys under _verify_first matched on last name and first initial only, so check those against payroll before anything else."
            ),
            "overrides": confident,
        }
        os.makedirs(P.STATE_DIR, exist_ok=True)
        with open(CANDIDATE_FILE, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, sort_keys=True)
        print(f"\n  Wrote {len(confident)} candidate override(s) to:")
        print(f"    {CANDIDATE_FILE}")
        print("\n  Review every line, then copy the verified ones into")
        print(f"    {P.OVERRIDES_FILE}")
        print("  under an \"overrides\" key. The push only reads the second file.")

    print("\n" + "=" * 78)
    print(f"  currently sendable : {len(sendable)}")
    print(f"  recoverable by override : {len(confident)}")
    print(f"  needs a human decision  : {len(ambiguous) + len(absent) + len(nameless)}")
    print("=" * 78)
    return 0


if __name__ == "__main__":
    sys.exit(main())
