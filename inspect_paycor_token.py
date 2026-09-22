"""
Decode the access token Paycor issues and show what it actually grants.

A 403 on an endpoint whose scope the portal shows as ON has two possible
causes that look identical from outside: the scope did not make it into the
token, or the token is valid but not for the legal entity being asked about.
Guessing between them costs an activation cycle each time.

A Paycor access token is a JWT, so its claims can be read locally. This
decodes the payload -- without verifying the signature, because the point is
to see what was issued, not to trust it -- and prints the scopes, the client
it was issued to, and any tenant or legal entity it names.

Prints claims only, never the token itself.

    python inspect_paycor_token.py
"""

import sys
import json
import base64

# .env loads on this import; paycor_api picks its base url at import time.
import traumasoft_api  # noqa: F401
import paycor_api
from paycor_api import PaycorClient, PaycorAuthError

# Claims worth calling out by name rather than leaving in the dump.
SCOPE_CLAIMS = ("scope", "scp", "scopes", "permissions", "roles")
ENTITY_CLAIMS = ("legalEntityId", "legal_entity_id", "tenantId", "tenant_id",
                 "clientId", "client_id", "sub", "aud", "iss", "azp")


def decode_segment(segment):
    """One JWT segment to a dict. Padding is stripped in JWTs; put it back."""
    padded = segment + "=" * (-len(segment) % 4)
    return json.loads(base64.urlsafe_b64decode(padded.encode("ascii")))


def main():
    print("=" * 78)
    print("PAYCOR ACCESS TOKEN INSPECTION")
    print(f"environment: {paycor_api.ENVIRONMENT}")
    print("=" * 78)

    try:
        client = PaycorClient(read_only=True)
    except PaycorAuthError as exc:
        print(f"\nCould not build a client: {exc}")
        return 1

    print(f"base url:  {client.base_url}")
    print(f"entity in .env: {client.legal_entity_id or '(not set)'}")
    print(f"client id: {client.client_id[:8] + '...' if client.client_id else '(not set)'}")

    print("\nFetching an access token...")
    try:
        token = client._token()
    except PaycorAuthError as exc:
        print(f"\nCould not get an access token: {exc}")
        print("\nThat is the problem to fix first -- the 403s are downstream of it.")
        return 1

    parts = token.split(".")
    if len(parts) != 3:
        print(f"\nThe access token is not a JWT ({len(parts)} segment(s)), so its")
        print("contents cannot be read here. That is itself worth knowing.")
        return 1

    try:
        header = decode_segment(parts[0])
        claims = decode_segment(parts[1])
    except Exception as exc:
        print(f"\nCould not decode the token: {exc}")
        return 1

    print(f"\nalgorithm: {header.get('alg')}   type: {header.get('typ')}")

    # ---- scopes ----
    print("\n" + "-" * 78)
    print("SCOPES IN THE TOKEN")
    print("-" * 78)
    found_scopes = False
    for key in SCOPE_CLAIMS:
        if key not in claims:
            continue
        found_scopes = True
        value = claims[key]
        items = value.split() if isinstance(value, str) else list(value)
        print(f"\n  {key} ({len(items)}):")
        for item in sorted(str(i) for i in items):
            print(f"    {item}")
    if not found_scopes:
        print("\n  The token carries no scope claim under any of:")
        print(f"    {', '.join(SCOPE_CLAIMS)}")
        print("\n  If Paycor enforces scopes without naming them in the token,")
        print("  the 403 is decided server-side by the app registration and no")
        print("  amount of re-minting will show up here.")

    # ---- who and what it is for ----
    print("\n" + "-" * 78)
    print("IDENTITY AND AUDIENCE")
    print("-" * 78)
    for key in ENTITY_CLAIMS:
        if key in claims:
            print(f"  {key:<18} {claims[key]}")

    entity = str(client.legal_entity_id or "").strip()
    if entity:
        blob = json.dumps(claims)
        print(f"\n  legal entity {entity} appears in the token: "
              f"{'yes' if entity in blob else 'no'}")
        if entity not in blob:
            print("  A token that never names this entity is the likelier")
            print("  explanation for a 403 than a missing scope would be.")

    # ---- everything else, for the things worth seeing that are not guessed ----
    print("\n" + "-" * 78)
    print("ALL CLAIMS")
    print("-" * 78)
    for key in sorted(claims):
        value = claims[key]
        rendered = json.dumps(value) if isinstance(value, (dict, list)) else str(value)
        if len(rendered) > 300:
            rendered = rendered[:300] + f"... ({len(rendered)} chars)"
        print(f"  {key:<24} {rendered}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
