"""Connection diagnostic for the Business Central OData endpoint.

Answers, in order: do the credentials work, how slow is each entity, and how
many rows does it hold? Run this before blaming the pipeline.

    python probe.py
    python probe.py --timeout 600
"""

from __future__ import annotations

import argparse
import sys
import time

import requests
from requests.auth import HTTPBasicAuth

import config

ENTITIES = [
    (config.ENTITY_ARCHIVE_ORDER_LINES, config.SELECT_ARCHIVE_ORDER_LINES),
    (config.ENTITY_OPEN_ORDER_LINES, config.SELECT_OPEN_ORDER_LINES),
    (config.ENTITY_SHIPMENT_LINES, config.SELECT_SHIPMENT_LINES),
    (config.ENTITY_SHIPMENT_HEADERS, config.SELECT_SHIPMENT_HEADERS),
    (config.ENTITY_ORDER_HEADERS_LIVE, config.SELECT_ORDER_HEADERS),
    (config.ENTITY_ORDER_HEADERS_ARCHIVE, config.SELECT_ORDER_HEADERS),
]


def probe(session: requests.Session, entity: str, params: dict, timeout: int,
          label: str) -> tuple[bool, float, str]:
    url = f"{config.company_url()}/{entity}"
    started = time.perf_counter()
    try:
        resp = session.get(url, params=params, timeout=timeout)
    except requests.exceptions.ReadTimeout:
        return False, time.perf_counter() - started, f"read timeout after {timeout}s"
    except requests.exceptions.ConnectTimeout:
        return False, time.perf_counter() - started, f"connect timeout after {timeout}s"
    except requests.RequestException as exc:
        return False, time.perf_counter() - started, f"{type(exc).__name__}: {exc}"

    elapsed = time.perf_counter() - started
    if resp.status_code == 401:
        return False, elapsed, "401 Unauthorized - check username / Web Service Access Key"
    if not resp.ok:
        return False, elapsed, f"HTTP {resp.status_code}: {resp.text[:160]}"

    try:
        payload = resp.json()
    except ValueError:
        return False, elapsed, f"non-JSON response: {resp.text[:120]}"

    if isinstance(payload, dict) and "value" in payload:
        n = len(payload["value"])
        has_next = "@odata.nextLink" in payload
        detail = f"{n} row(s){', more pages' if has_next else ''}"
        if n:
            detail += f", {len(payload['value'][0])} field(s)"
        return True, elapsed, detail
    return True, elapsed, f"{len(resp.content)} bytes"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--timeout", type=int, default=120,
                        help="per-request timeout in seconds (default 120)")
    parser.add_argument("--count", action="store_true",
                        help="also ask each entity for $count (can be slow)")
    args = parser.parse_args(argv)

    problems = config.validate()
    if problems:
        for p in problems:
            print(f"config: {p}")
        return 1

    print(f"endpoint : {config.BC_BASE_URL}")
    print(f"company  : {config.BC_COMPANY}")
    print(f"user     : {config.BC_USERNAME}")
    print(f"key      : {len(config.BC_ACCESS_KEY)} characters")
    print(f"timeout  : {args.timeout}s per request")
    print()

    session = requests.Session()
    session.auth = HTTPBasicAuth(config.BC_USERNAME, config.BC_ACCESS_KEY)
    session.headers.update({"Accept": "application/json"})

    failures = 0
    print(f"{'entity':<32}{'single row':<14}{'result'}")
    print("-" * 92)
    for entity, select in ENTITIES:
        ok, elapsed, detail = probe(
            session, entity,
            {"$top": "1", "$select": ",".join(select)},
            args.timeout, "top1",
        )
        flag = "ok" if ok else "FAIL"
        print(f"{entity:<32}{elapsed:>7.1f}s  {flag:<6}{detail}")
        if not ok:
            failures += 1

    if args.count:
        print()
        print(f"{'entity':<32}{'$count':<14}{'rows'}")
        print("-" * 92)
        for entity, _select in ENTITIES:
            url = f"{config.company_url()}/{entity}/$count"
            started = time.perf_counter()
            try:
                resp = session.get(url, timeout=args.timeout)
                elapsed = time.perf_counter() - started
                body = resp.text.strip()[:40] if resp.ok else f"HTTP {resp.status_code}"
                print(f"{entity:<32}{elapsed:>7.1f}s  {body}")
            except requests.RequestException as exc:
                elapsed = time.perf_counter() - started
                print(f"{entity:<32}{elapsed:>7.1f}s  {type(exc).__name__}")

    print()
    if failures:
        print(f"{failures} entity/entities failed.")
        return 2
    print("All entities reachable.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
