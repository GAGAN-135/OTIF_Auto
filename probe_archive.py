"""Characterise the slow ArchiveSalesOrdeLinesBi endpoint.

Tries progressively different request shapes to find one the server can serve,
and to tell a server-side query problem apart from a payload-size problem.

    python probe_archive.py
"""

from __future__ import annotations

import sys
import time

import requests
from requests.auth import HTTPBasicAuth

import config

ENTITY = config.ENTITY_ARCHIVE_ORDER_LINES
FULL_SELECT = ",".join(config.SELECT_ARCHIVE_ORDER_LINES)

# (label, params, timeout seconds)
VARIANTS: list[tuple[str, dict | None, int]] = [
    ("$top=1, no $select",
     {"$top": "1"}, 240),
    ("$top=1, one field only",
     {"$top": "1", "$select": "Document_No"}, 240),
    ("$top=1, $filter Version_No eq 1",
     {"$top": "1", "$select": FULL_SELECT, "$filter": "Version_No eq 1"}, 240),
    ("$count only",
     None, 240),
    ("$top=100, full $select",
     {"$top": "100", "$select": FULL_SELECT}, 300),
    ("full pull, full $select (the real request)",
     {"$select": FULL_SELECT}, 900),
]


def main() -> int:
    problems = config.validate()
    if problems:
        for p in problems:
            print(f"config: {p}")
        return 1

    session = requests.Session()
    session.auth = HTTPBasicAuth(config.BC_USERNAME, config.BC_ACCESS_KEY)
    session.headers.update({"Accept": "application/json"})

    base = f"{config.company_url()}/{ENTITY}"
    print(f"entity: {ENTITY}\n")

    for label, params, timeout in VARIANTS:
        url = base + ("/$count" if params is None else "")
        print(f"-> {label}  (timeout {timeout}s)", flush=True)
        started = time.perf_counter()
        try:
            resp = session.get(url, params=params, timeout=timeout)
            elapsed = time.perf_counter() - started
        except requests.exceptions.ReadTimeout:
            print(f"   FAIL  read timeout after {timeout}s\n", flush=True)
            continue
        except requests.RequestException as exc:
            elapsed = time.perf_counter() - started
            print(f"   FAIL  {elapsed:6.1f}s  {type(exc).__name__}: {exc}\n", flush=True)
            continue

        if not resp.ok:
            print(f"   FAIL  {elapsed:6.1f}s  HTTP {resp.status_code}: "
                  f"{resp.text[:200]}\n", flush=True)
            continue

        size_mb = len(resp.content) / 1048576
        if params is None:
            print(f"   ok    {elapsed:6.1f}s  count = {resp.text.strip()[:40]}\n", flush=True)
            continue
        try:
            payload = resp.json()
            rows = len(payload.get("value", []))
            more = " (more pages)" if "@odata.nextLink" in payload else ""
            print(f"   ok    {elapsed:6.1f}s  {rows} row(s), "
                  f"{size_mb:.1f} MB{more}\n", flush=True)
        except ValueError:
            print(f"   ok    {elapsed:6.1f}s  {size_mb:.1f} MB non-JSON\n", flush=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
