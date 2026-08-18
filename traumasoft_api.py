"""
Traumasoft ThirdParty API client.

Thin HMAC-authenticated REST client for the Traumasoft ThirdParty Data and
Lists endpoints, built to back the Daily Multi-Report Runner once direct
ODBC/database access is no longer available.

Auth, pagination shapes, rate limits, and endpoint paths follow the
Traumasoft ThirdParty OpenAPI spec (v1.0.0). See docs/API_MIGRATION.md for
how each report maps onto these endpoints.

Configuration is provided via environment variables:

    TS_API_BASE_URL   e.g. https://your-tenant.traumasoft.com
    TS_API_KEY        API key issued by Traumasoft (X-TS-APIKEY)
    TS_API_SECRET     Secret paired with that key (used for the HMAC)
    TS_API_TIMEOUT    Per-request timeout in seconds (default 60)
    TS_API_MIN_INTERVAL  Seconds between requests (default 0.65 -> ~92/min)
"""

import os
import time
import hmac
import json
import random
import hashlib
import logging
import urllib.parse
from datetime import date

import requests

# Optional dotenv, so credentials can live in a .env file next to the script
# instead of being exported by hand in every shell. Matches the pattern the
# report runner already uses.
try:
    from dotenv import load_dotenv

    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass

log = logging.getLogger(__name__)

# =============================
# CONFIG
# =============================
DEFAULT_TIMEOUT = float(os.getenv("TS_API_TIMEOUT", "60"))

# The documented read limit is 100 requests/minute. Stay just under it so a
# long backfill never trips a 429 in the first place.
DEFAULT_MIN_INTERVAL = float(os.getenv("TS_API_MIN_INTERVAL", "0.65"))

MAX_RETRIES = 4
RETRY_BACKOFF = (2, 4, 8, 16)  # seconds, matches the project's push retry policy
MAX_RETRY_AFTER = 120  # cap an over-long Retry-After so a run cannot stall

# Response envelopes that wrap their payload in a single named list.
_KNOWN_ROW_KEYS = (
    "rows",
    "users",
    "custom_statuses",
    "pay_types",
    "employee_levels",
    "fee_schedules",
    "schedules",
    "payor_categories",
    "attachment_types",
)


class TraumasoftAPIError(RuntimeError):
    """Raised when the API returns a non-retryable error response."""

    def __init__(self, status_code, message, path=None, body=None):
        self.status_code = status_code
        self.path = path
        self.body = body
        super().__init__(message)


# =============================
# CLIENT
# =============================
class TraumasoftAPI:
    """HMAC-signed client for /api/ThirdParty/* endpoints."""

    def __init__(
        self,
        base_url=None,
        api_key=None,
        api_secret=None,
        timeout=DEFAULT_TIMEOUT,
        min_interval=DEFAULT_MIN_INTERVAL,
        session=None,
    ):
        base_url = base_url or os.getenv("TS_API_BASE_URL", "")
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key or os.getenv("TS_API_KEY", "")
        self.api_secret = api_secret or os.getenv("TS_API_SECRET", "")
        self.timeout = timeout
        self.min_interval = min_interval
        self.session = session or requests.Session()
        self._last_request_at = 0.0

        if not self.base_url:
            raise ValueError("TS_API_BASE_URL is required (e.g. https://tenant.traumasoft.com)")
        if not self.api_key or not self.api_secret:
            raise ValueError("TS_API_KEY and TS_API_SECRET are required")

    # ---------- signing ----------
    def _auth_headers(self, body_str="", legacy_hmac=False):
        """
        Build the four required auth headers.

        Default scheme (all Data/Lists endpoints):
            hmac_sha256(body + timestamp + nonce, secret)

        Legacy scheme (GPS Geofence / older Postman collection):
            hmac_sha256(api_key, timestamp + secret + nonce)
        """
        timestamp = str(int(time.time()))
        nonce = str(random.randrange(10**9, 10**10))

        if legacy_hmac:
            message = self.api_key
            key = timestamp + self.api_secret + nonce
        else:
            message = (body_str or "") + timestamp + nonce
            key = self.api_secret

        digest = hmac.new(
            key.encode("utf-8"), message.encode("utf-8"), hashlib.sha256
        ).hexdigest()

        return {
            "X-TS-APIKEY": self.api_key,
            "X-TS-TIMESTAMP": timestamp,
            "X-TS-ID": nonce,
            "X-TS-AUTHORIZATION": digest,
            "Accept": "application/json",
        }

    # ---------- transport ----------
    def _throttle(self):
        elapsed = time.monotonic() - self._last_request_at
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        self._last_request_at = time.monotonic()

    def request(self, method, path, params=None, json_body=None, legacy_hmac=False):
        """Issue one signed request, retrying on 429 and transient failures."""
        url = f"{self.base_url}/api/{path.lstrip('/')}"
        body_str = json.dumps(json_body, separators=(",", ":")) if json_body is not None else ""

        last_error = None
        for attempt in range(MAX_RETRIES + 1):
            self._throttle()

            headers = self._auth_headers(body_str, legacy_hmac=legacy_hmac)
            if json_body is not None:
                headers["Content-Type"] = "application/json"

            try:
                resp = self.session.request(
                    method,
                    url,
                    params=params,
                    data=body_str.encode("utf-8") if json_body is not None else None,
                    headers=headers,
                    timeout=self.timeout,
                )
            except requests.RequestException as exc:
                last_error = exc
                if attempt >= MAX_RETRIES:
                    raise TraumasoftAPIError(None, f"{method} {path} failed: {exc}", path=path)
                wait = RETRY_BACKOFF[min(attempt, len(RETRY_BACKOFF) - 1)]
                log.warning("%s %s network error (%s); retrying in %ss", method, path, exc, wait)
                time.sleep(wait)
                continue

            if resp.status_code == 429:
                if attempt >= MAX_RETRIES:
                    raise TraumasoftAPIError(429, f"{method} {path} rate limited", path=path)
                wait = min(int(resp.headers.get("Retry-After", "60") or 60), MAX_RETRY_AFTER)
                log.warning("%s %s rate limited; waiting %ss", method, path, wait)
                time.sleep(wait)
                continue

            if resp.status_code >= 500:
                last_error = resp.text[:500]
                if attempt >= MAX_RETRIES:
                    raise TraumasoftAPIError(
                        resp.status_code, f"{method} {path} server error", path=path, body=last_error
                    )
                wait = RETRY_BACKOFF[min(attempt, len(RETRY_BACKOFF) - 1)]
                log.warning("%s %s -> %s; retrying in %ss", method, path, resp.status_code, wait)
                time.sleep(wait)
                continue

            if resp.status_code >= 400:
                raise TraumasoftAPIError(
                    resp.status_code,
                    f"{method} {path} -> {resp.status_code}: {resp.text[:500]}",
                    path=path,
                    body=resp.text[:2000],
                )

            if not resp.content:
                return None
            try:
                return resp.json()
            except ValueError:
                raise TraumasoftAPIError(
                    resp.status_code,
                    f"{method} {path} returned non-JSON body",
                    path=path,
                    body=resp.text[:2000],
                )

        raise TraumasoftAPIError(None, f"{method} {path} exhausted retries: {last_error}", path=path)

    def get(self, path, params=None, legacy_hmac=False):
        return self.request("GET", path, params=params, legacy_hmac=legacy_hmac)

    # ---------- response shapes ----------
    @staticmethod
    def extract_rows(payload):
        """
        Normalize the several envelope shapes this API uses into a list of dicts.

        Handles bare arrays (Shifts, Organization, GetTrips), paginated
        envelopes keyed on 'rows'/'users', and small reference lists keyed on
        their own plural name (custom_statuses, pay_types, ...).
        """
        if payload is None:
            return []
        if isinstance(payload, list):
            return payload
        if not isinstance(payload, dict):
            return []

        for key in _KNOWN_ROW_KEYS:
            value = payload.get(key)
            if isinstance(value, list):
                return value

        # Unknown single-list envelope: fall back to the only list present.
        lists = [v for v in payload.values() if isinstance(v, list)]
        if len(lists) == 1:
            return lists[0]
        return []

    @staticmethod
    def page_info(payload):
        """
        Return (current_page, total_pages) for whichever paginator shape came back.

        Standard/Users/Lists envelopes use current_page + total_pages; the
        jqGrid envelope (Vehicles, Payors, Facilities, Employees) uses
        page + total.
        """
        if not isinstance(payload, dict):
            return None, None
        if "current_page" in payload or "total_pages" in payload:
            return payload.get("current_page"), payload.get("total_pages")
        if "page" in payload or "total" in payload:
            return payload.get("page"), payload.get("total")
        return None, None

    def paginate(self, path, params=None, page_size=100, max_pages=1000):
        """Yield every row across all pages of a paginated endpoint."""
        params = dict(params or {})
        if page_size:
            params.setdefault("rows", page_size)

        page = 1
        while page <= max_pages:
            params["page"] = page
            payload = self.get(path, params=params)
            rows = self.extract_rows(payload)
            for row in rows:
                yield row

            current, total = self.page_info(payload)
            if total is None:
                # Unpaginated endpoint: one pass is the whole result.
                return
            try:
                total = int(total)
            except (TypeError, ValueError):
                return
            if page >= total or not rows:
                return
            page += 1

    # =============================
    # ENDPOINT HELPERS
    # =============================
    # -- Dispatch / CAD --
    def get_trips(self, trip_date, range_days=1, **filters):
        """
        GET /ThirdParty/Data/Cad/Trip?rtype=GetTrips

        Returns a list of TripLegSummary objects. `range_days` is inclusive
        from trip_date and is capped at 31 by the API.
        """
        if isinstance(trip_date, date):
            trip_date = trip_date.isoformat()
        if range_days is not None and range_days > 31:
            raise ValueError("range_days is capped at 31 by the Traumasoft API")

        params = {"rtype": "GetTrips", "trip_date": trip_date}
        if range_days is not None:
            params["range_days"] = range_days
        params.update({k: v for k, v in filters.items() if v is not None})
        return self.extract_rows(self.get("ThirdParty/Data/Cad/Trip", params=params))

    def list_call_types(self, include_deleted=False):
        return list(
            self.paginate(
                "ThirdParty/Data/Cad/CallTypes",
                params={"include_deleted": str(bool(include_deleted)).lower()},
            )
        )

    def list_zones(self, include_deleted=False):
        return list(
            self.paginate(
                "ThirdParty/Data/Cad/Zones",
                params={"include_deleted": str(bool(include_deleted)).lower()},
            )
        )

    def list_timestamp_types(self, include_deleted=False):
        return list(
            self.paginate(
                "ThirdParty/Data/Cad/Timestamps",
                params={"include_deleted": str(bool(include_deleted)).lower()},
            )
        )

    # -- Schedule --
    def list_shifts(self, include_deleted=False, **extra_params):
        """
        GET /ThirdParty/Data/Schedule/Shifts

        The spec documents only `include_deleted` here: no date filter and no
        pagination. `extra_params` exists so undocumented filters can be passed
        through once probe_traumasoft_api.py identifies any that work.
        """
        params = {"include_deleted": str(bool(include_deleted)).lower()}
        params.update({k: v for k, v in extra_params.items() if v is not None})
        return self.extract_rows(self.get("ThirdParty/Data/Schedule/Shifts", params=params))

    def list_shift_profiles(self):
        return list(self.paginate("ThirdParty/Lists/Schedule/ShiftProfiles"))

    # -- People --
    def list_employees(self, include_disabled=False):
        return list(
            self.paginate(
                "ThirdParty/Data/User/Employees",
                params={"include_disabled": str(bool(include_disabled)).lower()},
            )
        )

    def list_users(self, include_disabled=False, fields=None):
        params = {"include_disabled": str(bool(include_disabled)).lower()}
        if fields:
            params["fields[users]"] = ",".join(fields) if not isinstance(fields, str) else fields
        return list(self.paginate("ThirdParty/Data/User/Users", params=params))

    # -- Fleet --
    def list_vehicles(self, include_deleted=False, include_disabled=False, fields=None):
        params = {
            "include_deleted": str(bool(include_deleted)).lower(),
            "include_disabled": str(bool(include_disabled)).lower(),
        }
        if fields:
            params["fields[vehicles]"] = (
                ",".join(fields) if not isinstance(fields, str) else fields
            )
        return list(self.paginate("ThirdParty/Data/Fleet/Vehicles", params=params))

    def list_custom_statuses(self, include_deleted=False):
        return self.extract_rows(
            self.get(
                "ThirdParty/Data/Fleet/CustomStatus",
                params={"include_deleted": str(bool(include_deleted)).lower()},
            )
        )

    # -- Organization --
    def get_organization(self):
        """DDG hierarchy with the cost center ids attached to each group."""
        return self.extract_rows(self.get("ThirdParty/Data/Organization"))

    def list_cost_centers(self):
        return list(self.paginate("ThirdParty/Lists/Cad/CostCenters"))


def build_query_string(params):
    """Render params the way the API expects (used for logging/debugging)."""
    return urllib.parse.urlencode(params or {}, doseq=True)
