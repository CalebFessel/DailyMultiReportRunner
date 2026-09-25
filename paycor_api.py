"""
Paycor Public API v1 client, scoped to what a timecard push needs.

Built against Paycor's own OpenAPI 3.0 spec, kept at
docs/paycor-public-api-v1.json. Endpoint paths, required fields and response
shapes below are taken from it verbatim rather than inferred -- including its
inconsistent casing, which is reproduced exactly because URL paths are
case-sensitive: the punch *read* is `/v1/legalEntities/...` with a capital E
while every other endpoint here is `/v1/legalentities/...`.

Two things make this client different from the Samsara one, and both are
deliberate.

**It defaults to the sandbox.** `PAYCOR_ENVIRONMENT` must say `production` in
so many words before a call leaves for the real tenant. A bad route wastes a
driver's morning; a bad punch changes what somebody is paid.

**A 202 is not success.** `CreatePunches` is asynchronous: it accepts the
batch and returns a tracking id, and whether the punches actually landed is
only visible by reading the punch error log for that id afterwards. Treating
the 202 as "sent" would report success for punches Paycor rejected, so
`create_punches` returns the tracking id and `punch_errors` is the other half
of the operation. Never call one without the other.

Auth is two headers, both required on every call:
    Authorization: Bearer <access token>
    Ocp-Apim-Subscription-Key: <APIm subscription key>

The bearer comes from Paycor's authorization-code grant, which needs a human
in a browser once; the refresh token that falls out of it is what belongs in
.env. See the Paycor section of README.md.
"""

import json
import logging
import os
import time
import urllib.error
import urllib.parse
import re
import sys
import tempfile
import urllib.request
import uuid

log = logging.getLogger(__name__)

# The spec lists only the production host. The sandbox host is Paycor's
# documented counterpart to it and is what this client points at by default.
PRODUCTION_BASE_URL = "https://apis.paycor.com"
SANDBOX_BASE_URL = "https://apis-sandbox.paycor.com"

# The environment gate. Anything but the exact string "production" resolves to
# the sandbox, so a typo fails safe rather than going live.
ENVIRONMENT = os.getenv("PAYCOR_ENVIRONMENT", "sandbox").strip().lower()
IS_PRODUCTION = ENVIRONMENT == "production"


def credential_var_name(name):
    """
    The environment variable that actually supplied this credential, or None.

    env_credential() resolves a value; this resolves where it came from, which
    is what a write-back needs. A rotated refresh token has to replace the name
    that produced it -- writing PAYCOR_REFRESH_TOKEN when the sandbox-scoped
    name is in use would leave the stale value winning on the next run.
    """
    prefix = "PAYCOR_PRODUCTION_" if IS_PRODUCTION else "PAYCOR_SANDBOX_"
    scoped = f"{prefix}{name}"
    if os.getenv(scoped, "").strip():
        return scoped
    if os.getenv(f"PAYCOR_{name}", "").strip():
        return f"PAYCOR_{name}"
    return None


def dotenv_path():
    """
    The .env the credentials were loaded from, if one was.

    traumasoft_api loads it and records where from. Read that through
    sys.modules rather than importing it, so this module keeps no dependency
    on the other -- and fall back to the same search when it was never loaded,
    so a Paycor-only script still finds the file.
    """
    override = os.getenv("PAYCOR_ENV_FILE", "").strip()
    if override:
        return override

    loader = sys.modules.get("traumasoft_api")
    loaded_from = getattr(loader, "_DOTENV_LOADED_FROM", None) if loader else None
    if loaded_from:
        return loaded_from

    here = os.path.dirname(os.path.abspath(__file__))
    for directory in (here, os.getcwd()):
        for name in (".env", ".env.txt"):
            candidate = os.path.join(directory, name)
            if os.path.isfile(candidate):
                return candidate
    return None


def write_env_value(path, key, value):
    """
    Replace one KEY=VALUE in an env file, leaving every other byte alone.

    Written to a temporary file in the same directory and moved into place, so
    an interruption cannot leave a half-written .env -- which on this file
    would mean losing every other credential in it.
    """
    try:
        with open(path, "r", encoding="utf-8-sig") as handle:
            lines = handle.read().splitlines()
    except OSError as exc:
        raise OSError(f"could not read {path}: {exc}") from exc

    pattern = re.compile(rf"^\s*(?:export\s+)?{re.escape(key)}\s*=")
    replaced = False
    for index, line in enumerate(lines):
        if pattern.match(line):
            lines[index] = f"{key}={value}"
            replaced = True
            break
    if not replaced:
        lines.append(f"{key}={value}")

    directory = os.path.dirname(os.path.abspath(path)) or "."
    handle = tempfile.NamedTemporaryFile(
        "w", encoding="utf-8", dir=directory, delete=False, newline="\n"
    )
    try:
        handle.write("\n".join(lines) + "\n")
        handle.close()
        os.replace(handle.name, path)
    except Exception:
        try:
            os.unlink(handle.name)
        except OSError:
            pass
        raise
    return replaced


def env_credential(name, default=""):
    """
    Read a credential for whichever environment is selected.

    Sandbox and production are separate Paycor tenants with separate keys,
    tokens and legal entity ids. Holding one set of PAYCOR_* names means
    switching environments is a hand edit of the .env -- which is exactly how
    production credentials end up pointed at the sandbox host, or sandbox
    credentials at production, with nothing in the output saying so.

    So an environment-specific name wins when it is set:

        PAYCOR_SANDBOX_SUBSCRIPTION_KEY     used when environment=sandbox
        PAYCOR_PRODUCTION_SUBSCRIPTION_KEY  used when environment=production
        PAYCOR_SUBSCRIPTION_KEY             used when neither is

    Both sets live in the .env at once and PAYCOR_ENVIRONMENT selects between
    them coherently. The bare name still works on its own, so an existing .env
    keeps behaving exactly as it did.
    """
    prefix = "PAYCOR_PRODUCTION_" if IS_PRODUCTION else "PAYCOR_SANDBOX_"
    scoped = os.getenv(f"{prefix}{name}", "").strip()
    if scoped:
        return scoped
    return os.getenv(f"PAYCOR_{name}", default)

DEFAULT_TIMEOUT = int(os.getenv("PAYCOR_API_TIMEOUT", "60"))
DEFAULT_MIN_INTERVAL = float(os.getenv("PAYCOR_MIN_INTERVAL", "0.3"))

RETRY_STATUSES = {429, 500, 502, 503, 504}
MAX_RETRIES = 4

# Refresh a little before the token actually dies, so a long run does not fail
# on the one call that straddles the boundary.
TOKEN_EXPIRY_MARGIN_SECONDS = 60

# Paycor documents punchDateTime as "local time of the associated employee",
# formatted YYYY-MM-DDTHH:MM:SS. Traumasoft punches are already tenant-local
# once parse_shift_ts has applied the offset, so they go out unconverted -- but
# only because both are the same wall clock, which is worth stating.
PUNCH_TIME_FORMAT = "%Y-%m-%dT%H:%M:%S"

# CreatePunches takes an array. Paycor documents no explicit ceiling, so this
# is a self-imposed one: a failed batch has to be diagnosed through a single
# error log, and a smaller batch makes that log readable.
MAX_PUNCH_BATCH = int(os.getenv("PAYCOR_MAX_PUNCH_BATCH", "100"))

# employeePunches caps its window at 31 days per the spec.
MAX_EMPLOYEE_PUNCH_DAYS = 31

# A fixed namespace, so the correlation id derived from a Traumasoft punch is
# the same every run. That is what makes a re-run recognise its own writes
# instead of duplicating them.
CORRELATION_NAMESPACE = uuid.UUID("6f1a4d2e-0c27-4c5a-9b3e-7d0f2a8c15b4")

PUNCH_STATUS_IN = "In"
PUNCH_STATUS_OUT = "Out"
PUNCH_STATUS_TYPES = ("Auto", "In", "Out", "Transfer")


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


def correlation_id(traumasoft_punch_id, half):
    """
    A deterministic correlation id for one half of a Traumasoft punch.

    Paycor stores `correlationId` on the punch and returns it when the punch is
    read back, which turns "have I already sent this?" into an exact lookup
    rather than a timestamp comparison with a tolerance. Deriving it from the
    Traumasoft punch id means the answer survives a re-run, a restart, and a
    clock that disagrees by a minute.
    """
    return str(uuid.uuid5(CORRELATION_NAMESPACE, f"ts-punch:{traumasoft_punch_id}:{half}"))


def describe_write_contract():
    """What this client sends, and where each part of it comes from."""
    return {
        "environment": "production" if IS_PRODUCTION else "sandbox",
        "base_url": PRODUCTION_BASE_URL if IS_PRODUCTION else SANDBOX_BASE_URL,
        "request": "POST /v1/legalentities/{legalEntityId}/CreatePunches",
        "body": "array of EmployeePunch",
        "required_fields": [
            "employeeId", "departmentId", "punchDateTime",
            "punchStatusType", "activityTypeId", "isTransfer",
        ],
        "punch_model": "one object per clock EVENT (In / Out), not per interval",
        "time_format": PUNCH_TIME_FORMAT,
        "accepted_response": "202 + resourceUrl.id (tracking id); errors only "
                             "visible via GET punchErrorLog/{trackingId}",
        "source": "docs/paycor-public-api-v1.json (Paycor Public API v1, OpenAPI 3.0)",
        "verified": True,
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
        self.subscription_key = subscription_key or env_credential("SUBSCRIPTION_KEY")
        self.refresh_token = refresh_token or env_credential("REFRESH_TOKEN")
        self.client_id = client_id or env_credential("CLIENT_ID")
        self.client_secret = client_secret or env_credential("CLIENT_SECRET")
        self.legal_entity_id = legal_entity_id or env_credential("LEGAL_ENTITY_ID")
        self.timeout = timeout
        self.min_interval = min_interval
        self.read_only = read_only

        default_base = PRODUCTION_BASE_URL if IS_PRODUCTION else SANDBOX_BASE_URL
        self.base_url = (base_url or os.getenv("PAYCOR_API_BASE_URL", default_base)).rstrip("/")

        # A directly-supplied access token short-circuits the refresh dance,
        # which is how you test with a token pasted out of the developer portal
        # before wiring the OAuth flow up properly.
        self._access_token = access_token or env_credential("ACCESS_TOKEN") or None
        self._token_expires_at = float("inf") if self._access_token else 0.0
        self._last_call = 0.0

        if not self.subscription_key:
            raise PaycorAuthError(
                "PAYCOR_SUBSCRIPTION_KEY is not set. Paycor wants it on every "
                "call alongside the bearer token -- see .env.example."
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
        form = {"grant_type": "refresh_token", "refresh_token": self.refresh_token}
        if self.client_id:
            form["client_id"] = self.client_id
        if self.client_secret:
            form["client_secret"] = self.client_secret

        req = urllib.request.Request(
            url,
            data=urllib.parse.urlencode(form).encode("utf-8"),
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
        # whichever came back, so a long run does not authenticate once and then
        # fail on a stale credential -- and write it back, or the next run
        # starts from a refresh token Paycor has already retired. That failure
        # arrives as an ordinary auth error on a morning nobody changed
        # anything, which is a bad way to find out.
        rotated = payload.get("refresh_token")
        if rotated and rotated != self.refresh_token:
            self.refresh_token = rotated
            self._persist_refresh_token(rotated)
        self._access_token = token
        expires_in = float(payload.get("expires_in") or 3600)
        self._token_expires_at = time.time() + expires_in - TOKEN_EXPIRY_MARGIN_SECONDS
        log.info("Paycor access token acquired, valid ~%.0f minutes.", expires_in / 60)
        return token

    def _persist_refresh_token(self, token):
        """
        Write a rotated refresh token back where it came from.

        Best effort by design: a failure here means the next run authenticates
        with a retired credential, which is worth a loud warning, but it must
        not take down a run that has already authenticated and may be part way
        through writing payroll.

        Set PAYCOR_PERSIST_REFRESH_TOKEN=false on a host with a read-only
        checkout or an injected secret, where the write would fail every time
        and the platform owns the credential anyway.
        """
        if os.getenv("PAYCOR_PERSIST_REFRESH_TOKEN", "true").strip().lower() in (
            "false", "0", "no"
        ):
            log.info("Paycor rotated the refresh token; persistence is disabled, "
                     "so update the stored credential yourself.")
            return False

        key = credential_var_name("REFRESH_TOKEN")
        if not key:
            # The token was passed in code, not read from the environment, so
            # there is nothing to write back to.
            log.warning(
                "Paycor rotated the refresh token, but the old one did not come "
                "from an environment variable, so it cannot be written back. "
                "The next run will use a retired credential."
            )
            return False

        path = dotenv_path()
        if not path:
            log.warning(
                "Paycor rotated the refresh token and no .env was found to write "
                "%s back to. Store the new value before the next run, or it will "
                "authenticate with a retired credential.", key
            )
            return False

        try:
            write_env_value(path, key, token)
        except Exception as exc:
            log.warning(
                "Paycor rotated the refresh token but %s could not be updated in "
                "%s (%s). Store the new value before the next run.", key, path, exc
            )
            return False

        # Keep the process's own environment in step, so anything reading it
        # later in this run sees the live credential rather than the retired one.
        os.environ[key] = token
        log.info("Paycor rotated the refresh token; %s updated in %s.", key, path)
        return True

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
                # Refresh once and retry before treating it as fatal.
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

        Paged endpoints answer with
        {"records": [...], "hasMoreResults": bool, "continuationToken": "..."}.
        A few endpoints (employeePunches among them) return a bare array with
        no envelope at all, so that shape is handled rather than assumed away.
        """
        params = dict(params or {})
        pages = 0
        while pages < max_pages:
            payload = self.request("GET", path, params=params)
            if isinstance(payload, list):
                for row in payload:
                    yield row
                return
            for row in payload.get("records") or []:
                yield row
            token = payload.get("continuationToken")
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

    def list_employees(self, legal_entity_id=None):
        """
        GET /v1/legalentities/{id}/employees

        Each record carries `id` (the employee guid every punch is addressed
        to), `employeeNumber` (the string that joins to Traumasoft's
        employee_num), `badgeNumber`, and `department` as a ResourceReference
        whose id is the department guid CreatePunches requires.
        """
        entity = self._entity(legal_entity_id)
        return list(self.paginate(f"v1/legalentities/{entity}/employees"))

    def list_departments(self, legal_entity_id=None):
        """GET /v1/legalentities/{id}/departments -- id, code, description."""
        entity = self._entity(legal_entity_id)
        return list(self.paginate(f"v1/legalentities/{entity}/departments"))

    def list_activity_types(self, legal_entity_id=None):
        """
        GET /v1/legalentities/{id}/activitytypes -- id, name, type.

        CreatePunches requires an activityTypeId and Paycor supplies no
        default, so one of these has to be chosen before anything can be sent.
        """
        entity = self._entity(legal_entity_id)
        return list(self.paginate(f"v1/legalentities/{entity}/activitytypes"))

    def get_timecard_punches(self, start_date, end_date, legal_entity_id=None):
        """
        GET /v1/legalEntities/{id}/punches -- note the capital E, per the spec.

        Returns TimeCardV3 records, which are punch *pairs*: punchIn, punchOut,
        hourAmount, employeeId, employeeNumber, departmentCode. Both dates are
        required. This is the shape that matches Traumasoft's own intervals,
        so it is what the reconciliation reads.
        """
        entity = self._entity(legal_entity_id)
        return list(self.paginate(
            f"v1/legalEntities/{entity}/punches",
            params={"startDate": start_date, "endDate": end_date},
        ))

    def get_employee_punches(self, employee_id, start_date, end_date):
        """
        GET /v1/employees/{id}/employeePunches -- individual punch events.

        Unlike the legal-entity read above this returns one record per clock
        event, carrying `punchId` and, crucially, `correlationId` -- which is
        what lets a re-run recognise punches it sent itself. The window is
        capped at 31 days by the spec.
        """
        return list(self.paginate(
            f"v1/employees/{employee_id}/employeePunches",
            params={"startDate": start_date, "endDate": end_date},
        ))

    def punch_errors(self, tracking_id, legal_entity_id=None):
        """
        GET /v1/legalentities/{id}/punchErrorLog/{trackingId}

        The other half of create_punches. An empty list here is the only
        evidence that an accepted batch actually landed.
        """
        entity = self._entity(legal_entity_id)
        return list(self.paginate(
            f"v1/legalentities/{entity}/punchErrorLog/{tracking_id}"
        ))

    # =============================
    # WRITES
    # =============================
    @staticmethod
    def build_punch(employee_id, department_id, activity_type_id, punch_datetime,
                    status, note=None, work_location_id=None, correlation=None):
        """
        One EmployeePunch object -- a single clock event, not an interval.

        This is the shape mistake worth naming: Paycor's punch model is
        event-based, so a shift that Traumasoft stores as one row with a start
        and an end becomes two objects here, an In and an Out.

        Separate from `create_punches` so a dry run can show exactly what would
        be sent without a client that is able to send it.
        """
        if status not in PUNCH_STATUS_TYPES:
            raise ValueError(
                f"punchStatusType must be one of {PUNCH_STATUS_TYPES}, got {status!r}"
            )
        body = {
            "employeeId": employee_id,
            "departmentId": department_id,
            "punchDateTime": punch_datetime.strftime(PUNCH_TIME_FORMAT),
            "punchStatusType": status,
            "activityTypeId": activity_type_id,
            # Every punch here is a plain clock event on one department. A
            # transfer punch means moving between departments or activities
            # mid-shift, which nothing in the Traumasoft feed describes.
            "isTransfer": False,
        }
        if note:
            # Paycor rejects a note outside 1..300 characters.
            body["note"] = str(note)[:300]
        if work_location_id:
            body["workLocationId"] = work_location_id
        if correlation:
            body["correlationId"] = correlation
        return body

    def create_punches(self, punches, legal_entity_id=None):
        """
        POST /v1/legalentities/{id}/CreatePunches. Requires read_only=False.

        Returns the tracking id from the 202 response. **That is an
        acknowledgement, not a result** -- Paycor validates asynchronously, so
        call `punch_errors(tracking_id)` afterwards to find out what actually
        landed. A caller that treats this return value as success will report
        punches as sent that Paycor threw away.
        """
        entity = self._entity(legal_entity_id)
        if not punches:
            return None
        if len(punches) > MAX_PUNCH_BATCH:
            raise ValueError(
                f"{len(punches)} punches exceeds the {MAX_PUNCH_BATCH} batch "
                "ceiling; send them in chunks so one error log stays readable."
            )
        payload = self.request(
            "POST", f"v1/legalentities/{entity}/CreatePunches",
            json_body=punches, write=True,
        )
        resource = (payload or {}).get("resourceUrl") or {}
        tracking_id = resource.get("id")
        if not tracking_id:
            log.warning(
                "CreatePunches returned no tracking id (%s); errors for this "
                "batch cannot be checked.", json.dumps(payload)[:200]
            )
        return tracking_id

    def delete_punches(self, employee_id, punches):
        """
        DELETE /v1/employees/{id}/DeletePunches. Requires read_only=False.

        Returns the tracking id from the 202, exactly like `create_punches`:
        the delete is applied asynchronously and a rejected one is visible
        only through `punch_errors(tracking_id)`. A caller that treats the
        return of this method as success will report punches as deleted that
        are still there -- which is how a sandbox probe came to announce a
        clean cleanup over punches Paycor had kept.

        Takes either punch guids or the punch records from
        `get_employee_punches`. Prefer the records: the spec requires that
        when a punch carries a punchRefId, the delete must name it too, and
        only the record knows whether it has one.

        This is what makes a test reversible, and a test you cannot undo is
        not one to run against payroll.
        """
        body = []
        for punch in punches:
            if isinstance(punch, dict):
                punch_id = punch.get("punchId") or punch.get("id")
                if not punch_id:
                    continue
                entry = {"punchId": str(punch_id)}
                # Required by the spec whenever it is present on the punch.
                ref_id = punch.get("punchRefId")
                if ref_id:
                    entry["punchRefId"] = str(ref_id)
                body.append(entry)
            elif punch:
                body.append({"punchId": str(punch)})

        if not body:
            return None
        if len(body) > MAX_PUNCH_BATCH:
            raise ValueError(
                f"{len(body)} punches exceeds the {MAX_PUNCH_BATCH} batch "
                "ceiling; delete them in chunks so one error log stays readable."
            )

        payload = self.request(
            "DELETE", f"v1/employees/{employee_id}/DeletePunches",
            json_body=body, write=True,
        )
        resource = (payload or {}).get("resourceUrl") or {}
        tracking_id = resource.get("id")
        if not tracking_id:
            log.warning(
                "DeletePunches returned no tracking id (%s); whether the "
                "delete was applied cannot be checked.", json.dumps(payload)[:200]
            )
        return tracking_id
