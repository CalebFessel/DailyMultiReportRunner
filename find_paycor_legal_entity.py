"""
Print the legal entity IDs this Paycor credential can see.

PAYCOR_LEGAL_ENTITY_ID is required by every other call and is not shown
anywhere obvious in the developer portal. The API will tell you:
GET /v1/legalentities/ActivatedLegalEntityTenantList takes no parameters and
returns "the Legal Entities/Tenants for the user logged in".

Read-only. Needs only the subscription key and a token; it is the one call
here that does *not* need the legal entity id, which is the point.

    python find_paycor_legal_entity.py
"""

import os
import sys
import json

# traumasoft_api loads .env as a side effect of import, and paycor_api reads
# PAYCOR_ENVIRONMENT at import time to pick its base URL. Import in this order
# or a .env saying production silently talks to the sandbox.
import traumasoft_api  # noqa: F401  (imported for the .env load)
import paycor_api
from paycor_api import PaycorClient, PaycorAPIError, PaycorAuthError


def main():
    print("=" * 70)
    print("PAYCOR LEGAL ENTITY LOOKUP")
    print(f"environment: {paycor_api.ENVIRONMENT}")
    try:
        client = PaycorClient(read_only=True)
    except PaycorAuthError as exc:
        print(f"\nConfiguration error: {exc}")
        return 2
    print(f"base url:    {client.base_url}")
    print("=" * 70)

    try:
        payload = client.request("GET", "v1/legalentities/ActivatedLegalEntityTenantList")
    except PaycorAPIError as exc:
        print(f"\nAPI error: {exc}")
        print("\n401/403 here means the token or scopes are wrong, not the id.")
        return 1
    except PaycorAuthError as exc:
        print(f"\nAuth error: {exc}")
        return 1

    entities = []
    if isinstance(payload, dict):
        entities = payload.get("userLegalEntities") or []
    elif isinstance(payload, list):
        entities = payload

    if not entities:
        print("\nNo legal entities returned. Raw response:")
        print(json.dumps(payload, indent=2)[:2000])
        return 1

    print(f"\n{len(entities)} legal entit{'y' if len(entities) == 1 else 'ies'} visible:\n")
    print(f"  {'legalEntityId':>14}  {'tenantId':>10}")
    print(f"  {'-' * 14}  {'-' * 10}")
    for item in entities:
        if not isinstance(item, dict):
            continue
        print(f"  {str(item.get('legalEntityId', '?')):>14}  {str(item.get('tenantId', '?')):>10}")

    if len(entities) == 1:
        only = entities[0].get("legalEntityId")
        print(f"\nOne entity, so this is unambiguous:\n\n    PAYCOR_LEGAL_ENTITY_ID={only}\n")
    else:
        print(
            "\nMore than one. Pick the entity the crews are paid from -- the wrong\n"
            "one will read and write another company's punches. Confirm with\n"
            "payroll before setting it, then re-run the dry run.\n"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
