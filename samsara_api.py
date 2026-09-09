"""
Samsara Fleet API client, scoped to what route dispatch needs.

Deliberately small: vehicles (to resolve a Traumasoft unit to a Samsara
vehicle id), addresses (to reuse registered geofences where they exist), and
routes (to create, list and delete them).

Unlike the Traumasoft client in traumasoft_api.py, this one can write. Every
write goes through `request()` with `write=True`, and a client constructed
with `read_only=True` -- the default -- refuses those calls outright. The
caller has to opt into publishing; nothing here can create a route by
accident.

Auth is a plain bearer token, not the HMAC dance Traumasoft uses:
    Authorization: Bearer <SAMSARA_API_TOKEN>

Docs: https://developers.samsara.com/reference/createroute
"""

import json
import logging
import os
import time
import urllib.error
import urllib.parse
import urllib.request

log = logging.getLogger(__name__)

DEFAULT_BASE_URL = "https://api.samsara.com"

# Samsara publishes per-endpoint limits rather than one global number. The
# route endpoints sit at the low end, so pace every call to the slowest thing
# we touch instead of tracking a budget per path.
DEFAULT_MIN_INTERVAL = float(os.getenv("SAMSARA_MIN_INTERVAL", "0.4"))
DEFAULT_TIMEOUT = int(os.getenv("SAMSARA_API_TIMEOUT", "60"))

# 429 and 5xx are worth retrying; 4xx of any other kind is a bad request and
# retrying just repeats it.
RETRY_STATUSES = {429, 500, 502, 503, 504}
MAX_RETRIES = 4


class SamsaraAPIError(RuntimeError):
    """A non-retryable error response from Samsara."""

    def __init__(self, status_code, message, path=None, body=None):
        self.status_code = status_code
        self.path = path
        self.body = body
        super().__init__(f"Samsara API {status_code} on {path}: {message}")


class SamsaraReadOnlyError(RuntimeError):
    """Raised when a write is attempted on a read-only client."""


class SamsaraClient:
    def __init__(
        self,
        token=None,
        base_url=None,
        timeout=DEFAULT_TIMEOUT,
        min_interval=DEFAULT_MIN_INTERVAL,
        read_only=True,
    ):
        self.token = token or os.getenv("SAMSARA_API_TOKEN", "")
        self.base_url = (base_url or os.getenv("SAMSARA_API_BASE_URL", DEFAULT_BASE_URL)).rstrip("/")
        self.timeout = timeout
        self.min_interval = min_interval
        self.read_only = read_only
        self._last_call = 0.0
        if not self.token:
            raise RuntimeError(
                "SAMSARA_API_TOKEN is not set. Add it to .env -- see .env.example."
            )

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
        One API call, with pacing and retries.

        `write` is not inferred from the HTTP method -- the caller states it --
        so a read-only client cannot be talked into a write by a helper that
        happens to use POST.
        """
        if write and self.read_only:
            raise SamsaraReadOnlyError(
                f"{method} {path} is a write and this client is read-only. "
                "Construct SamsaraClient(read_only=False) to publish."
            )

        url = f"{self.base_url}/{path.lstrip('/')}"
        if params:
            clean = {k: v for k, v in params.items() if v is not None}
            if clean:
                url = f"{url}?{urllib.parse.urlencode(clean)}"

        body_bytes = None
        headers = {
            "Authorization": f"Bearer {self.token}",
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
                if exc.code in RETRY_STATUSES and attempt < MAX_RETRIES - 1:
                    backoff = 2 ** attempt
                    log.warning(
                        "Samsara %s on %s, retrying in %ss", exc.code, path, backoff
                    )
                    time.sleep(backoff)
                    last_error = exc
                    continue
                raise SamsaraAPIError(exc.code, exc.reason, path=path, body=detail) from exc
            except urllib.error.URLError as exc:
                if attempt < MAX_RETRIES - 1:
                    backoff = 2 ** attempt
                    log.warning("Samsara network error on %s (%s), retrying in %ss",
                                path, exc.reason, backoff)
                    time.sleep(backoff)
                    last_error = exc
                    continue
                raise
        raise SamsaraAPIError(0, f"exhausted retries ({last_error})", path=path)

    def paginate(self, path, params=None, max_pages=1000):
        """
        Walk Samsara's cursor pagination.

        Responses carry {"data": [...], "pagination": {"endCursor", "hasNextPage"}}.
        """
        params = dict(params or {})
        pages = 0
        while pages < max_pages:
            payload = self.request("GET", path, params=params)
            for row in payload.get("data") or []:
                yield row
            page = payload.get("pagination") or {}
            if not page.get("hasNextPage") or not page.get("endCursor"):
                return
            params["after"] = page["endCursor"]
            pages += 1

    # =============================
    # ENDPOINTS
    # =============================
    def list_vehicles(self, limit=512):
        """GET /fleet/vehicles -- id, name and tags for every vehicle."""
        return list(self.paginate("fleet/vehicles", params={"limit": limit}))

    def list_addresses(self, limit=512):
        """
        GET /addresses -- the registered geofences.

        Worth pulling because a stop that names an existing addressId inherits
        that geofence's shape, while a single-use location is always a 300m
        circle. Facilities that already exist in Samsara route better.
        """
        return list(self.paginate("addresses", params={"limit": limit}))

    def list_routes(self, start_time, end_time, limit=512):
        """GET /fleet/routes over an RFC3339 window."""
        return list(
            self.paginate(
                "fleet/routes",
                params={"startTime": start_time, "endTime": end_time, "limit": limit},
            )
        )

    def create_route(self, route):
        """POST /fleet/routes. Requires read_only=False."""
        return self.request("POST", "fleet/routes", json_body=route, write=True)

    def delete_route(self, route_id):
        """DELETE /fleet/routes/{id}. Requires read_only=False."""
        return self.request("DELETE", f"fleet/routes/{route_id}", write=True)
