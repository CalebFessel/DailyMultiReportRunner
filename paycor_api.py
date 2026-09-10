"""
Paycor Public API client, scoped to what a timecard push needs.

Two things make this client different from the other two in this repo, and
both are deliberate.

**It defaults to the sandbox.** `PAYCOR_ENVIRONMENT` must say `production` in
so many words before a single call leaves for the real tenant. Samsara's client
defaults to the real API because a bad route wastes a driver's morning; a bad
punch changes what somebody is paid, and wage-hour errors belong to the
employer, not the vendor.

**Its write shape is unverified.** developers.paycor.com is unreachable from
the environment this was written in, so the read endpoints below are
corroborated only from secondary sources and the write endpoint is a best
reading of them. Every part of the write that could differ -- the path, the
HTTP method, the field names, the timestamp format -- is an environment
variable rather than a literal, so correcting it against the real docs is a
config change and not a rewrite. `describe_write_contract()` prints exactly
what this client believes, which is the first thing to check against Paycor's
own reference once you have portal access.

Auth is OAuth 2.0 plus an APIm subscription key; both are required on every
call. Paycor's flow is an authorization-code grant against secure.paycor.com,
which needs a human in a browser once. The refresh token that falls out of it
is what belongs in .env -- this client exchanges it for a short-lived access
token and re-exchanges when that expires.

Sources, all secondary:
  https://developers.paycor.com/explore
  https://www.merge.dev/blog/paycor-api
  https://rollout.com/integration-guides/paycor/
"""

import json
import logging
import os
import time
import urllib.error
import urllib.parse
import urllib.request

log = logging.getLogger(__name__)

PRODUCTION_BASE_URL = "https://apis.paycor.com"
SANDBOX_BASE_URL = "https://apis-sandbox.paycor.com"

# The environment gate. Anything other than the exact string "production"
# resolves to the sandbox, so a typo fails safe rather than going live.
ENVIRONMENT = os.getenv("PAYCOR_ENVIRONMENT", "sandbox").strip().lower()
IS_PRODUCTION = ENVIRONMENT == "production"

DEFAULT_TIMEOUT = int(os.getenv("PAYCOR_API_TIMEOUT", "60"))
DEFAULT_MIN_INTERVAL = float(os.getenv("PAYCOR_MIN_INTERVAL", "0.3"))

RETRY_STATUSES = {429, 500, 502, 503, 504}
MAX_RETRIES = 4

# Refresh a little before the token actually dies, so a long run does not fail
# on the one call that straddles the boundary.
TOKEN_EXPIRY_MARGIN_SECONDS = 60

# ---------------------------------------------------------------------------
# The unverified write contract. Every one of these is a guess until checked
# against Paycor's own reference; all are overridable so that checking them
# costs an .env edit rather than a code change.
# ---------------------------------------------------------------------------
PUNCH_WRITE_PATH = os.getenv(
    "PAYCOR_PUNCH_WRITE_PATH", "v1/legalentities/{legal_entity_id}/timecardpunches"
)
PUNCH_WRITE_METHOD = os.getenv("PAYCOR_PUNCH_WRITE_METHOD", "POST").upper()
PUNCH_FIELD_EMPLOYEE = os.getenv("PAYCOR_PUNCH_FIELD_EMPLOYEE", "employeeId")
PUNCH_FIELD_IN = os.getenv("PAYCOR_PUNCH_FIELD_IN", "punchInTime")
PUNCH_FIELD_OUT = os.getenv("PAYCOR_PUNCH_FIELD_OUT", "punchOutTime")
# Paycor's timecard times are wall-clock in the employee's own zone in the UI.
# Whether the API wants a naive local string or an offset-bearing one is
# exactly the kind of thing that silently shifts every punch by hours, so it
# is a setting with a loud default rather than an assumption.
PUNCH_TIME_FORMAT = os.getenv("PAYCOR_PUNCH_TIME_FORMAT", "%Y-%m-%dT%H:%M:%S")


class PaycorAPIError(RuntimeError):
    """A non-retryable error response from Paycor."""

    def __init__(self, status_code, message, path=None, body=None):
        self.status_code = status_code
        self.path = path
        self.body = body
        super().__init__(f"Paycor API {status_code} on {path}: {message}")


class PaycorReadOnlyError(RuntimeError):
    """Raised when a write is attempted on a read-only client."""


class PaycorAuthError(RuntimeError):
    """Raised when no usable credential could be assembled."""


def describe_write_contract():
    """What this client currently believes a punch write looks like."""
    return {
        "environment": "production" if IS_PRODUCTION else "sandbox",
        "base_url": PRODUCTION_BASE_URL if IS_PRODUCTION else SANDBOX_BASE_URL,
        "method": PUNCH_WRITE_METHOD,
        "path": PUNCH_WRITE_PATH,
        "body_fields": {
            "employee": PUNCH_FIELD_EMPLOYEE,
            "punch_in": PUNCH_FIELD_IN,
            "punch_out": PUNCH_FIELD_OUT,
        },
        "time_format": PUNCH_TIME_FORMAT,
        "verified": False,
    }


class PaycorClient:
    def __init__(
        self,
        subscription_key=None,
        refresh_token=None,
        access_token=None,
        client_id=None,
        client_secret=None,
        legal_entity_id=None,
        base_url=None,
        timeout=DEFAULT_TIMEOUT,
        min_interval=DEFAULT_MIN_INTERVAL,
        read_only=True,
    ):
        self.subscription_key = subscription_key or os.getenv("PAYCOR_SUBSCRIPTION_KEY", "")
        self.refresh_token = refresh_token or os.getenv("PAYCOR_REFRESH_TOKEN", "")
        self.client_id = client_id or os.getenv("PAYCOR_CLIENT_ID", "")
        self.client_secret = client_secret or os.getenv("PAYCOR_CLIENT_SECRET", "")
        self.legal_entity_id = legal_entity_id or os.getenv("PAYCOR_LEGAL_ENTITY_ID", "")
        self.timeout = timeout
        self.min_interval = min_interval
        self.read_only = read_only

        default_base = PRODUCTION_BASE_URL if IS_PRODUCTION else SANDBOX_BASE_URL
        self.base_url = (base_url or os.getenv("PAYCOR_API_BASE_URL", default_base)).rstrip("/")

        # A directly-supplied access token short-circuits the refresh dance,
        # which is how you test with a token pasted out of the developer portal
        # before wiring the OAuth flow up properly.
        self._access_token = access_token or os.getenv("PAYCOR_ACCESS_TOKEN", "") or None
        self._token_expires_at = float("inf") if self._access_token else 0.0
        self._last_call = 0.0

        if not self.subscription_key:
            raise PaycorAuthError(
                "PAYCOR_SUBSCRIPTION_KEY is not set. Every Paycor call needs it "
                "alongside the bearer token -- see .env.example."
            )
        if not self._access_token and not self.refresh_token:
            raise PaycorAuthError(
                "Set PAYCOR_REFRESH_TOKEN (preferred) or PAYCOR_ACCESS_TOKEN. "
                "The refresh token comes from running Paycor's authorization-code "
                "flow once in a browser; see the Paycor section of README.md."
            )
        if IS_PRODUCTION:
            log.warning(
                "Paycor client is pointed at PRODUCTION (%s). Writes from here "
                "change what people are paid.", self.base_url
            )
        else:
            log.info("Paycor client is pointed at the sandbox (%s).", self.base_url)

    # =============================
    # AUTH
    # =============================
    def _fetch_access_token(self):
        """
        Exchange the refresh token for an access token.

        Paycor puts the subscription key in the query string on this one call
        and in a header on every other, which is unusual enough to be worth
        stating rather than looking like a mistake.
        """
        url = (
            f"{self.base_url}/sts/v1/common/token"
            f"?subscription-key={urllib.parse.quote(self.subscription_key)}"
        )
        form = {
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
        }
        if self.client_id:
            form["client_id"] = self.client_id
        if self.client_secret:
            form["client_secret"] = self.client_secret

        body = urllib.parse.urlencode(form).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=body,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
                "Ocp-Apim-Subscription-Key": self.subscription_key,
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise PaycorAuthError(
                f"Token exchange failed ({exc.code}): {detail[:400]}"
            ) from exc

        token = payload.get("access_token")
        if not token:
            raise PaycorAuthError(
                f"Token response carried no access_token: {list(payload)}"
            )
        # Paycor rotates the refresh token on use in some configurations. Keep
        # whichever one came back so a long run does not authenticate once and
        # then fail on a stale credential.
        if payload.get("refresh_token"):
            self.refresh_token = payload["refresh_token"]
        self._access_token = token
        expires_in = float(payload.get("expires_in") or 3600)
        self._token_expires_at = time.time() + expires_in - TOKEN_EXPIRY_MARGIN_SECONDS
        log.info("Paycor access token acquired, valid ~%.0f minutes.", expires_in / 60)
        return token

    def _token(self):
        if not self._access_token or time.time() >= self._token_expires_at:
            return self._fetch_access_token()
        return self._access_token

    # =============================
    # TRANSPORT
    # =============================
    def _throttle(self):
        elapsed = time.monotonic() - self._last_call
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self._last_call = time.monotonic()

    def request(self, method, path, params=None, json_body=None, write=False):
        """
        One API call, with pacing, token refresh and retries.

        `write` is stated by the caller rather than inferred from the method, so
        a read-only client cannot be talked into a write by a helper that
        happens to POST. This is the same guard as the Samsara client and it
        matters more here.
        """
        if write and self.read_only:
            raise PaycorReadOnlyError(
                f"{method} {path} is a write and this client is read-only. "
                "Construct PaycorClient(read_only=False) to send it."
            )

        url = f"{self.base_url}/{path.lstrip('/')}"
        if params:
            clean = {k: v for k, v in params.items() if v is not None}
            if clean:
                url = f"{url}?{urllib.parse.urlencode(clean)}"

        body_bytes = None
        headers = {
            "Authorization": f"Bearer {self._token()}",
            "Ocp-Apim-Subscription-Key": self.subscription_key,
            "Accept": "application/json",
        }
        if json_body is not None:
            body_bytes = json.dumps(json_body).encode("utf-8")
            headers["Content-Type"] = "application/json"

        last_error = None
        for attempt in range(MAX_RETRIES):
            self._throttle()
            req = urllib.request.Request(url, data=body_bytes, headers=headers, method=method)
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    raw = resp.read().decode("utf-8")
                    return json.loads(raw) if raw.strip() else {}
            except urllib.error.HTTPError as exc:
                detail = exc.read().decode("utf-8", errors="replace")
                # A 401 mid-run usually means the token aged out under us.
                # Refresh once and try again before treating it as fatal.
                if exc.code == 401 and attempt == 0 and self.refresh_token:
                    log.info("Paycor 401 on %s, refreshing the token and retrying.", path)
                    self._access_token = None
                    headers["Authorization"] = f"Bearer {self._token()}"
                    continue
                if exc.code in RETRY_STATUSES and attempt < MAX_RETRIES - 1:
                    backoff = 2 ** attempt
                    log.warning("Paycor %s on %s, retrying in %ss", exc.code, path, backoff)
                    time.sleep(backoff)
                    last_error = exc
                    continue
                raise PaycorAPIError(exc.code, exc.reason, path=path, body=detail) from exc
            except urllib.error.URLError as exc:
                if attempt < MAX_RETRIES - 1:
                    backoff = 2 ** attempt
                    log.warning("Paycor network error on %s (%s), retrying in %ss",
                                path, exc.reason, backoff)
                    time.sleep(backoff)
                    last_error = exc
                    continue
                raise
        raise PaycorAPIError(0, f"exhausted retries ({last_error})", path=path)

    def paginate(self, path, params=None, max_pages=1000):
        """
        Walk Paycor's continuation-token pagination.

        Responses carry {"records": [...], "hasMoreResults": bool,
        "continuationToken": "..."}. Key names vary a little across their
        endpoints, so several spellings are accepted rather than assuming one.
        """
        params = dict(params or {})
        pages = 0
        while pages < max_pages:
            payload = self.request("GET", path, params=params)
            rows = (
                payload.get("records")
                or payload.get("results")
                or payload.get("data")
                or []
            )
            for row in rows:
                yield row
            token = payload.get("continuationToken") or payload.get("continuation_token")
            more = payload.get("hasMoreResults")
            if more is None:
                more = bool(token)
            if not more or not token:
                return
            params["continuationToken"] = token
            pages += 1

    # =============================
    # READS
    # =============================
    def _entity(self, legal_entity_id=None):
        entity = legal_entity_id or self.legal_entity_id
        if not entity:
            raise PaycorAPIError(
                0, "no legal entity id", path="(caller)",
                body="Set PAYCOR_LEGAL_ENTITY_ID or pass legal_entity_id.",
            )
        return entity

    def list_employees(self, legal_entity_id=None, include_terminated=False):
        """GET /v1/legalentities/{id}/employees -- the Paycor-side roster."""
        entity = self._entity(legal_entity_id)
        return list(self.paginate(
            f"v1/legalentities/{entity}/employees",
            params={"includeTerminated": "true" if include_terminated else None},
        ))

    def get_timecard_punches(self, start_date, end_date, legal_entity_id=None):
        """
        GET /v1/legalentities/{id}/timecardpunches over a date window.

        This is the reconciliation read: what Paycor already holds, to compare
        against what Traumasoft says before anything is written.
        """
        entity = self._entity(legal_entity_id)
        return list(self.paginate(
            f"v1/legalentities/{entity}/timecardpunches",
            params={"startDate": start_date, "endDate": end_date},
        ))

    def get_employee_punches(self, employee_id, start_date=None, end_date=None):
        """GET /v1/employees/{id}/timecardpunches -- one person's punches."""
        return list(self.paginate(
            f"v1/employees/{employee_id}/timecardpunches",
            params={"startDate": start_date, "endDate": end_date},
        ))

    # =============================
    # WRITE
    # =============================
    def build_punch_body(self, employee_id, punch_in, punch_out):
        """
        The request body, assembled from the configurable field names.

        Separate from `create_punch` so a dry run can show precisely what would
        be sent without a client that is able to send it.
        """
        body = {
            PUNCH_FIELD_EMPLOYEE: employee_id,
            PUNCH_FIELD_IN: punch_in.strftime(PUNCH_TIME_FORMAT),
        }
        if punch_out is not None:
            body[PUNCH_FIELD_OUT] = punch_out.strftime(PUNCH_TIME_FORMAT)
        return body

    def create_punch(self, employee_id, punch_in, punch_out, legal_entity_id=None):
        """
        Write one punch. Requires read_only=False.

        UNVERIFIED against Paycor's own documentation -- see the module
        docstring. Check `describe_write_contract()` against their reference
        before pointing this at production.
        """
        entity = self._entity(legal_entity_id)
        path = PUNCH_WRITE_PATH.format(legal_entity_id=entity)
        body = self.build_punch_body(employee_id, punch_in, punch_out)
        return self.request(PUNCH_WRITE_METHOD, path, json_body=body, write=True)
