"""
Minimal Samsara REST client used by the TraumaSoft -> Samsara work order sync.

Covers only what the sync needs:
  - GET  /fleet/vehicles              (asset lookup, paginated)
  - GET  /maintenance/work-orders     (ledger rebuild, paginated)
  - POST /maintenance/work-orders     (create)
  - PATCH /maintenance/work-orders    (update; work order id goes in the body)

Auth is a bearer token.  The token must have the "Read Work Orders" and
"Write Work Orders" scopes (Closed Beta category in the token editor) plus
"Read Vehicles".

Rate limits documented by Samsara:
  - GET  /maintenance/work-orders   5 req/sec
  - POST/PATCH work orders          100 req/min
  - GET  /fleet/vehicles            25 req/sec
The client throttles writes to stay under 100/min and retries 429/5xx with
backoff, honouring Retry-After when present.
"""

from __future__ import annotations

import logging
import time
from typing import Any, Dict, Iterator, List, Optional

import requests

DEFAULT_BASE_URL = "https://api.samsara.com"
EU_BASE_URL = "https://api.eu.samsara.com"

WORK_ORDER_STATUSES = (
    "Assigned",
    "Cancelled",
    "Closed",
    "Completed",
    "In Progress",
    "On Hold",
    "Open",
    "Pending Approval",
    "Pending Parts",
)
WORK_ORDER_PRIORITIES = ("Low", "Medium", "High", "Urgent")
WORK_ORDER_CATEGORIES = (
    "Annual",
    "Corrective",
    "Damage Repair",
    "Preventive",
    "Recall",
    "Unspecified",
)


class SamsaraApiError(RuntimeError):
    def __init__(self, status_code: int, message: str, body: Any = None):
        super().__init__(f"Samsara API {status_code}: {message}")
        self.status_code = status_code
        self.body = body


class SamsaraClient:
    def __init__(
        self,
        token: str,
        base_url: str = DEFAULT_BASE_URL,
        timeout: float = 30.0,
        max_retries: int = 5,
        write_min_interval_s: float = 0.65,  # ~92 writes/min, under the 100/min cap
        session: Optional[requests.Session] = None,
    ):
        if not token:
            raise ValueError("Samsara API token is required (SAMSARA_API_TOKEN).")
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.max_retries = max_retries
        self.write_min_interval_s = write_min_interval_s
        self._last_write_ts = 0.0
        self._session = session or requests.Session()
        self._session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )

    # ------------------------------------------------------------------ core
    def _request(
        self,
        method: str,
        path: str,
        params: Optional[Dict[str, Any]] = None,
        json: Optional[Dict[str, Any]] = None,
    ) -> Any:
        url = f"{self.base_url}/{path.lstrip('/')}"
        if params:
            params = {k: v for k, v in params.items() if v is not None}
        if json:
            json = {k: v for k, v in json.items() if v is not None}

        if method in ("POST", "PATCH", "DELETE"):
            self._throttle_write()

        attempt = 0
        while True:
            attempt += 1
            try:
                resp = self._session.request(
                    method, url, params=params, json=json, timeout=self.timeout
                )
            except requests.RequestException as e:
                if attempt > self.max_retries:
                    raise
                wait = min(2 ** attempt, 30)
                logging.warning("Samsara %s %s network error (%s); retry in %ss", method, path, e, wait)
                time.sleep(wait)
                continue

            if resp.status_code in (429, 500, 502, 503, 504) and attempt <= self.max_retries:
                retry_after = resp.headers.get("Retry-After")
                try:
                    wait = float(retry_after) if retry_after else min(2 ** attempt, 30)
                except ValueError:
                    wait = min(2 ** attempt, 30)
                logging.warning(
                    "Samsara %s %s -> %s; retry %s/%s in %.1fs",
                    method, path, resp.status_code, attempt, self.max_retries, wait,
                )
                time.sleep(wait)
                continue

            if resp.status_code >= 400:
                try:
                    body = resp.json()
                except ValueError:
                    body = resp.text
                raise SamsaraApiError(resp.status_code, str(body)[:1000], body)

            if resp.status_code == 204 or not resp.content:
                return None
            return resp.json()

    def _throttle_write(self) -> None:
        now = time.monotonic()
        gap = now - self._last_write_ts
        if gap < self.write_min_interval_s:
            time.sleep(self.write_min_interval_s - gap)
        self._last_write_ts = time.monotonic()

    def _paginate(self, path: str, params: Optional[Dict[str, Any]] = None) -> Iterator[Dict[str, Any]]:
        params = dict(params or {})
        after: Optional[str] = None
        while True:
            if after:
                params["after"] = after
            page = self._request("GET", path, params=params) or {}
            for item in page.get("data", []) or []:
                yield item
            pagination = page.get("pagination") or {}
            if not pagination.get("hasNextPage"):
                return
            after = pagination.get("endCursor")
            if not after:
                return

    # -------------------------------------------------------------- vehicles
    def list_vehicles(self) -> List[Dict[str, Any]]:
        """All vehicles in the org. Fields used by the sync: id, name, vin, externalIds."""
        return list(self._paginate("fleet/vehicles", {"limit": 512}))

    # ----------------------------------------------------------- work orders
    def list_work_orders(self, ids: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        params: Dict[str, Any] = {}
        if ids:
            params["ids"] = ",".join(ids)
        return list(self._paginate("maintenance/work-orders", params))

    def create_work_order(self, body: Dict[str, Any]) -> Dict[str, Any]:
        """POST /maintenance/work-orders. `assetId` is required. Returns the work order object."""
        if not body.get("assetId"):
            raise ValueError("assetId is required to create a Samsara work order")
        resp = self._request("POST", "maintenance/work-orders", json=body) or {}
        return resp.get("data", resp)

    def update_work_order(self, work_order_id: str, body: Dict[str, Any]) -> Dict[str, Any]:
        """PATCH /maintenance/work-orders. The work order id is sent in the body as `id`."""
        payload = dict(body)
        payload["id"] = work_order_id
        resp = self._request("PATCH", "maintenance/work-orders", json=payload) or {}
        return resp.get("data", resp)
