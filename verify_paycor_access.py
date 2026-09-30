"""
Check that the Paycor credential, scopes and legal entity id agree.

Run this once PAYCOR_LEGAL_ENTITY_ID is set -- the id is also Paycor's
"Client Id", so payroll usually knows it without an API call. This exercises
only the read endpoints the punch push depends on, so a clean run here means
the dry run will get as far as building a plan.

Read-only. Nothing here can write.

    python verify_paycor_access.py
"""

import os
import sys

# .env loads on this import, and paycor_api reads PAYCOR_ENVIRONMENT at import
# time, so the order matters. See find_paycor_legal_entity.py.
import traumasoft_api  # noqa: F401
import paycor_api
from paycor_api import PaycorClient, PaycorAPIError, PaycorAuthError


CHECKS = (
    ("employees", "list_employees", "roster -- joins employee_num to Paycor guids"),
    ("departments", "list_departments", "departmentId, required on every punch"),
    ("activity types", "list_activity_types", "activityTypeId, required on every punch"),
)


def main():
    entity = paycor_api.env_credential("LEGAL_ENTITY_ID", "").strip()
    print("=" * 74)
    print("PAYCOR ACCESS VERIFICATION")
    print(f"environment:     {paycor_api.ENVIRONMENT}")
    print(f"legal entity id: {entity or '(not set)'}")
    print("=" * 74)

    if not entity:
        print("\nPAYCOR_LEGAL_ENTITY_ID is not set. It is the same number Paycor")
        print("calls your Client Id -- payroll will know it.")
        return 2

    try:
        client = PaycorClient(read_only=True)
    except PaycorAuthError as exc:
        print(f"\nConfiguration error: {exc}")
        return 2

    failures = 0
    for label, method, why in CHECKS:
        try:
            rows = getattr(client, method)()
        except PaycorAPIError as exc:
            failures += 1
            print(f"\n  FAIL  {label}: {exc}")
            if exc.status_code in (401, 403):
                print(f"        {exc.status_code} is a scope problem, not the id.")
            elif exc.status_code == 404:
                print("        404 usually means the legal entity id is wrong.")
            continue
        except PaycorAuthError as exc:
            print(f"\n  FAIL  {label}: {exc}")
            return 1
        print(f"\n  OK    {label}: {len(rows)} record(s)")
        print(f"        {why}")
        if label == "activity types" and rows:
            wanted = os.getenv("PAYCOR_ACTIVITY_TYPE", "Work").strip().lower()
            names = [str(r.get("name", "")) for r in rows if isinstance(r, dict)]
            match = [n for n in names if n.strip().lower() == wanted]
            print(f"        available: {', '.join(names[:8]) or '(unnamed)'}")
            if match:
                print(f"        PAYCOR_ACTIVITY_TYPE={wanted!r} resolves.")
            else:
                print(f"        PAYCOR_ACTIVITY_TYPE={wanted!r} does NOT match any of these.")
                print("        Set PAYCOR_ACTIVITY_TYPE to one of the names above.")
                failures += 1

    print("\n" + "=" * 74)
    if failures:
        print(f"{failures} check(s) failed. Fix these before the dry run.")
        return 1
    print("All checks passed. The credential, scopes and legal entity id agree.")
    print("Next:  python push_paycor_timecards.py            (dry run, reads only)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
