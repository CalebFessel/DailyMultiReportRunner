"""
Turn a Paycor authorization code into a refresh token.

Paycor's App Activation page runs a PKCE authorization-code flow: it shows a
Code Verifier, you click Initiate, and it redirects to
hcm.paycor.com/appactivation/clientredirect with ?code=... in the address bar.
That code is short-lived and single use; this exchanges it for the refresh
token every other call needs.

The Code Verifier shown on the activation page must be the one from the *same*
attempt that produced the code. Reloading the page issues a new verifier and
invalidates the old pairing.

    python paycor_exchange_code.py --code <code> --code-verifier <verifier>

Prints the refresh token so it can be pasted into .env. That value is a
payroll credential: keep it out of chat, tickets and shared drives.
"""

import os
import sys
import json
import argparse
import urllib.parse
import urllib.request
import urllib.error

# .env loads on this import; paycor_api reads PAYCOR_ENVIRONMENT at import time.
import traumasoft_api  # noqa: F401
import paycor_api

DEFAULT_REDIRECT = "https://hcm.paycor.com/appactivation/clientredirect"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--code", required=True, help="the ?code=... value from the redirect")
    ap.add_argument("--code-verifier", required=True, help="Code Verifier from the activation page")
    ap.add_argument("--redirect-uri", default=DEFAULT_REDIRECT)
    args = ap.parse_args()

    # Resolved per environment, so exchanging a sandbox code uses the sandbox
    # credentials and cannot accidentally mint a production token.
    subscription_key = paycor_api.env_credential("SUBSCRIPTION_KEY", "").strip()
    client_id = paycor_api.env_credential("CLIENT_ID", "").strip()
    client_secret = paycor_api.env_credential("CLIENT_SECRET", "").strip()

    if not subscription_key:
        print(f"No subscription key for environment {paycor_api.ENVIRONMENT!r}.")
        print("Set PAYCOR_SUBSCRIPTION_KEY, or the environment-specific")
        print("PAYCOR_SANDBOX_SUBSCRIPTION_KEY / PAYCOR_PRODUCTION_SUBSCRIPTION_KEY.")
        return 2

    base = paycor_api.PRODUCTION_BASE_URL if paycor_api.IS_PRODUCTION else paycor_api.SANDBOX_BASE_URL
    base = os.getenv("PAYCOR_API_BASE_URL", base).rstrip("/")

    print("=" * 74)
    print("PAYCOR AUTHORIZATION CODE EXCHANGE")
    print(f"environment: {paycor_api.ENVIRONMENT}")
    print(f"base url:    {base}")
    print("=" * 74)

    form = {
        "grant_type": "authorization_code",
        "code": args.code,
        "code_verifier": args.code_verifier,
        "redirect_uri": args.redirect_uri,
    }
    if client_id:
        form["client_id"] = client_id
    if client_secret:
        form["client_secret"] = client_secret

    url = (
        f"{base}/sts/v1/common/token"
        f"?subscription-key={urllib.parse.quote(subscription_key)}"
    )
    req = urllib.request.Request(
        url,
        data=urllib.parse.urlencode(form).encode("utf-8"),
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
            "Ocp-Apim-Subscription-Key": subscription_key,
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        print(f"\nExchange failed ({exc.code}): {detail[:500]}")
        if exc.code == 400:
            print(
                "\n400 here is usually one of: the code was already used, the code\n"
                "expired (they are short-lived -- redo the activation and exchange\n"
                "immediately), or the verifier is from a different page load than\n"
                "the code."
            )
        return 1

    refresh = payload.get("refresh_token")
    access = payload.get("access_token")

    scope = "PAYCOR_PRODUCTION_" if paycor_api.IS_PRODUCTION else "PAYCOR_SANDBOX_"

    if refresh:
        print("\nSUCCESS. Put this in .env, then delete it from your terminal history:\n")
        print(f"    {scope}REFRESH_TOKEN={refresh}\n")
        print("(Or PAYCOR_REFRESH_TOKEN if you keep only one environment's")
        print("credentials in this .env.)\n")
    elif access:
        print("\nNo refresh token came back, only an access token. It expires in")
        print(f"~{int(payload.get('expires_in') or 3600) // 60} minutes, so this is a test credential, not a scheduled one:\n")
        print(f"    {scope}ACCESS_TOKEN={access}\n")
    else:
        print(f"\nNeither token present. Response keys: {sorted(payload)}")
        return 1

    print("Then:  python verify_paycor_access.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
