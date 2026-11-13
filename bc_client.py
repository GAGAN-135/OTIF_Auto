"""Business Central OData v4 client.

Two things this does that the Power Query version did not:

1. Follows ``@odata.nextLink`` to exhaustion. The M code read only
   ``Json.Document(...)[value]``, which is the *first page*. That works today
   only because the service happens to return everything in one response; it
   would silently truncate the report the day someone lowers the server's page
   limit.
2. Projects with ``$select`` so only the consumed columns cross the wire.
"""

from __future__ import annotations

import logging
from typing import Any

import requests
from requests.adapters import HTTPAdapter
from requests.auth import HTTPBasicAuth
from urllib3.util.retry import Retry

import config

log = logging.getLogger(__name__)


class BusinessCentralError(RuntimeError):
    """Raised for configuration, authentication, and transport failures."""


class BCClient:
    """Thin authenticated wrapper over the BC OData v4 endpoint."""

    def __init__(
        self,
        username: str | None = None,
        access_key: str | None = None,
        timeout: int = config.REQUEST_TIMEOUT,
    ) -> None:
        username = username if username is not None else config.BC_USERNAME
        access_key = access_key if access_key is not None else config.BC_ACCESS_KEY
        if not username or not access_key:
            raise BusinessCentralError(
                "Business Central credentials are missing. Copy .env.example to "
                ".env and set BC_USERNAME and BC_ACCESS_KEY (the Web Service "
                "Access Key from the BC user card, not your Windows password)."
            )

        self.timeout = timeout
        self.session = requests.Session()
        self.session.auth = HTTPBasicAuth(username, access_key)
        self.session.headers.update({
            "Accept": "application/json",
            # Ask BC not to bother computing formatted values we never read.
            "Prefer": "odata.include-annotations=\"-*\"",
        })

        # BC behind a load balancer occasionally returns 502/503 under load;
        # retry with backoff rather than failing the whole nightly run.
        retry = Retry(
            total=4,
            backoff_factor=2.0,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_maxsize=8)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

    def close(self) -> None:
        self.session.close()

    def __enter__(self) -> "BCClient":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def fetch(
        self,
        entity: str,
        select: list[str] | None = None,
        odata_filter: str | None = None,
    ) -> list[dict[str, Any]]:
        """Return every record of ``entity``, following server-side paging."""
        url = f"{config.company_url()}/{entity}"
        params: dict[str, str] | None = {}
        if select:
            params["$select"] = ",".join(select)
        if odata_filter:
            params["$filter"] = odata_filter
        params = params or None

        records: list[dict[str, Any]] = []
        page = 0
        while url:
            page += 1
            try:
                resp = self.session.get(url, params=params, timeout=self.timeout)
            except requests.RequestException as exc:
                raise BusinessCentralError(
                    f"{entity}: request failed on page {page} -- {exc}"
                ) from exc

            # The nextLink is absolute and already carries the query string, so
            # params must only be applied to the very first request.
            params = None

            if resp.status_code == 401:
                raise BusinessCentralError(
                    f"{entity}: 401 Unauthorized. Check BC_USERNAME and that "
                    "BC_ACCESS_KEY is a current Web Service Access Key, and that "
                    "the user has permission on this entity."
                )
            if resp.status_code == 404:
                raise BusinessCentralError(
                    f"{entity}: 404 Not Found. Verify the entity is published as "
                    f"a web service and that BC_COMPANY ({config.BC_COMPANY!r}) "
                    "is exactly right."
                )
            if not resp.ok:
                raise BusinessCentralError(
                    f"{entity}: HTTP {resp.status_code} on page {page} -- "
                    f"{resp.text[:400]}"
                )

            try:
                payload = resp.json()
            except ValueError as exc:
                raise BusinessCentralError(
                    f"{entity}: response on page {page} was not JSON -- "
                    f"{resp.text[:200]}"
                ) from exc

            batch = payload.get("value")
            if batch is None:
                raise BusinessCentralError(
                    f"{entity}: response on page {page} has no 'value' array. "
                    f"Keys present: {sorted(payload)[:10]}"
                )
            records.extend(batch)
            url = payload.get("@odata.nextLink")
            log.debug("%s page %d: %d rows (running total %d)",
                      entity, page, len(batch), len(records))

        log.info("%-30s %7d rows in %d page(s)", entity, len(records), page)
        return records
