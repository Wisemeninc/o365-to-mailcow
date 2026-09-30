"""Minimal Microsoft Graph client: paging, throttling, retries, concurrency cap.

Only ever talks to graph.microsoft.com. Every request carries a timeout and the tool's
User-Agent. Concurrency per client is capped with a semaphore (Graph allows four
in-flight requests per mailbox); create one client per mailbox to get that behaviour.
"""

from __future__ import annotations

import email.utils
import logging
import random
import threading
import time
from collections.abc import Callable, Iterator
from urllib.parse import urlsplit

import requests

from . import USER_AGENT
from .auth import TokenProvider
from .config import GRAPH_HOST

log = logging.getLogger(__name__)

GRAPH_BASE = f"https://{GRAPH_HOST}/v1.0"
RETRY_STATUSES = {429, 503, 504}


class GraphError(Exception):
    def __init__(self, status: int, message: str, path: str) -> None:
        super().__init__(f"HTTP {status} for {path}: {message}")
        self.status = status
        self.path = path


class GraphClient:
    def __init__(
        self,
        tokens: TokenProvider,
        session: requests.Session | None = None,
        max_inflight: int = 4,
        timeout: tuple[float, float] = (15.0, 180.0),
        max_retries: int = 6,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self._tokens = tokens
        self._session = session or requests.Session()
        self._sem = threading.BoundedSemaphore(max_inflight)
        self._timeout = timeout
        self._max_retries = max_retries
        self._sleep = sleep

    # -- public ------------------------------------------------------------------------

    def get(self, path: str, params: dict | None = None, headers: dict | None = None) -> dict:
        resp = self._request(path, params=params, headers=headers, stream=False)
        return resp.json()

    def iter_pages(
        self, path: str, params: dict | None = None, headers: dict | None = None
    ) -> Iterator[dict]:
        """Yield every item of a collection, following @odata.nextLink."""
        url: str | None = path
        first = True
        while url:
            page = self.get(url, params=params if first else None, headers=headers)
            first = False
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")

    def get_delta(
        self, path: str, params: dict | None = None, headers: dict | None = None
    ) -> tuple[list[dict], str | None]:
        """Follow a delta query to the end. Returns (items, deltaLink)."""
        items: list[dict] = []
        url: str | None = path
        first = True
        delta_link: str | None = None
        while url:
            page = self.get(url, params=params if first else None, headers=headers)
            first = False
            items.extend(page.get("value", []))
            delta_link = page.get("@odata.deltaLink", delta_link)
            url = page.get("@odata.nextLink")
        return items, delta_link

    def get_bytes(self, path: str, headers: dict | None = None) -> bytes:
        resp = self._request(path, params=None, headers=headers, stream=True)
        try:
            return resp.content
        finally:
            resp.close()

    # -- internals ---------------------------------------------------------------------

    @staticmethod
    def _resolve(path: str) -> str:
        if path.startswith("https://"):
            host = urlsplit(path).netloc.lower()
            if host != GRAPH_HOST:
                raise GraphError(0, f"refusing request to non-Graph host {host}", path)
            return path
        return f"{GRAPH_BASE}/{path.lstrip('/')}"

    @staticmethod
    def _retry_after(resp: requests.Response, attempt: int) -> float:
        raw = resp.headers.get("Retry-After")
        if raw:
            try:
                return float(raw)
            except ValueError:
                when = email.utils.parsedate_to_datetime(raw)
                return max(0.0, when.timestamp() - time.time())
        return min(60.0, (2**attempt) + random.uniform(0, 1))  # noqa: S311 - jitter only

    def _request(
        self, path: str, params: dict | None, headers: dict | None, stream: bool
    ) -> requests.Response:
        url = self._resolve(path)
        base_headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
        if headers:
            base_headers.update(headers)
        refreshed = False
        for attempt in range(self._max_retries + 1):
            base_headers["Authorization"] = f"Bearer {self._tokens.get_token()}"
            with self._sem:
                try:
                    resp = self._session.get(
                        url, params=params, headers=base_headers,
                        timeout=self._timeout, stream=stream,
                    )
                except requests.RequestException as exc:
                    if attempt >= self._max_retries:
                        name = exc.__class__.__name__
                        raise GraphError(0, f"network error: {name}", path) from exc
                    self._sleep(self._retry_after_network(attempt))
                    continue
                if resp.status_code < 400:
                    return resp
                if resp.status_code == 401 and not refreshed:
                    refreshed = True
                    self._tokens.invalidate()
                    resp.close()
                    continue
                if resp.status_code in RETRY_STATUSES and attempt < self._max_retries:
                    delay = self._retry_after(resp, attempt)
                    log.info("Graph %s on %s; retrying in %.0fs", resp.status_code, path, delay)
                    resp.close()
                    self._sleep(delay)
                    continue
                message = self._error_message(resp)
                resp.close()
                raise GraphError(resp.status_code, message, path)
        raise GraphError(0, "retries exhausted", path)

    @staticmethod
    def _retry_after_network(attempt: int) -> float:
        return min(30.0, (2**attempt) + random.uniform(0, 1))  # noqa: S311

    @staticmethod
    def _error_message(resp: requests.Response) -> str:
        try:
            err = resp.json().get("error", {})
            return f"{err.get('code', '?')}: {err.get('message', '')}"[:200]
        except ValueError:
            return resp.text[:200]
