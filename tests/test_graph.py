"""GraphClient over a mocked transport (ISC-28..35)."""

from __future__ import annotations

import io
import threading
import time
from urllib.parse import urlsplit

import pytest
import requests
import responses

from o365_to_mailcow import USER_AGENT
from o365_to_mailcow.graph import GRAPH_BASE, GraphClient, GraphError

URL = f"{GRAPH_BASE}/users/a@x.com/mailFolders"


class Tokens:
    def __init__(self) -> None:
        self.n = 0
        self.invalidated = 0

    def get_token(self) -> str:
        self.n += 1
        return f"tok-{self.n}"

    def invalidate(self) -> None:
        self.invalidated += 1


def client(**kw) -> tuple[GraphClient, list[float], Tokens]:
    sleeps: list[float] = []
    tokens = Tokens()
    return GraphClient(tokens, sleep=sleeps.append, **kw), sleeps, tokens


@responses.activate
def test_iter_pages_follows_next_link_isc_28():
    responses.get(URL, json={"value": [1, 2], "@odata.nextLink": URL + "?$skip=2"})
    responses.get(URL + "?$skip=2", json={"value": [3]})
    g, _, _ = client()
    assert list(g.iter_pages("/users/a@x.com/mailFolders", params={"$top": 100})) == [1, 2, 3]
    assert "%24top=100" in responses.calls[0].request.url
    assert "top" not in responses.calls[1].request.url  # nextLink carries its own query


@responses.activate
def test_get_delta_returns_items_and_link():
    responses.get(URL + "/delta", json={"value": [{"id": "a"}], "@odata.nextLink": URL + "/d2"})
    responses.get(URL + "/d2", json={"value": [{"id": "b"}], "@odata.deltaLink": URL + "/dl"})
    g, _, _ = client()
    items, link = g.get_delta("/users/a@x.com/mailFolders/delta")
    assert [i["id"] for i in items] == ["a", "b"] and link == URL + "/dl"


@responses.activate
def test_429_sleeps_retry_after_isc_29():
    responses.get(URL, status=429, headers={"Retry-After": "7"})
    responses.get(URL, json={"ok": True})
    g, sleeps, _ = client()
    assert g.get("/users/a@x.com/mailFolders") == {"ok": True}
    assert sleeps == [7.0]


@responses.activate
def test_503_exponential_backoff_then_raise_isc_30():
    for _ in range(4):
        responses.get(URL, status=503)
    g, sleeps, _ = client(max_retries=3)
    with pytest.raises(GraphError) as exc:
        g.get("/users/a@x.com/mailFolders")
    assert exc.value.status == 503
    assert len(sleeps) == 3 and sleeps[0] < sleeps[1] < sleeps[2]
    assert 1 <= sleeps[0] < 2 and 4 <= sleeps[2] < 5


@responses.activate
def test_a_request_without_retries_raises_at_once_when_throttled():
    responses.get(URL, status=429, headers={"Retry-After": "300"})
    responses.get(URL, status=429, headers={"Retry-After": "7"})
    responses.get(URL, json={"ok": True})
    g, sleeps, _ = client()
    with pytest.raises(GraphError) as exc:
        g.get("/users/a@x.com/mailFolders", retries=0)
    assert exc.value.status == 429 and sleeps == [] and len(responses.calls) == 1
    # the override is for that one request: the next one retries as configured
    assert g.get("/users/a@x.com/mailFolders") == {"ok": True}
    assert sleeps == [7.0]


@responses.activate
def test_504_retried():
    responses.get(URL, status=504)
    responses.get(URL, json={})
    g, sleeps, _ = client()
    assert g.get(URL) == {} and len(sleeps) == 1


@responses.activate
def test_401_refreshes_once_isc_31():
    responses.get(URL, status=401)
    responses.get(URL, json={"v": 1})
    g, _, tokens = client()
    assert g.get(URL) == {"v": 1}
    assert tokens.invalidated == 1
    assert responses.calls[1].request.headers["Authorization"] == "Bearer tok-2"


@responses.activate
def test_second_401_raises_isc_31():
    responses.get(URL, status=401, json={"error": {"code": "InvalidAuthenticationToken"}})
    responses.get(URL, status=401, json={"error": {"code": "InvalidAuthenticationToken"}})
    g, _, tokens = client()
    with pytest.raises(GraphError) as exc:
        g.get(URL)
    assert exc.value.status == 401 and tokens.invalidated == 1


@responses.activate
def test_404_raises_immediately_with_code():
    responses.get(URL, status=404, json={"error": {"code": "ErrorItemNotFound", "message": "x"}})
    g, sleeps, _ = client()
    with pytest.raises(GraphError, match="ErrorItemNotFound") as exc:
        g.get(URL)
    assert exc.value.status == 404 and sleeps == []


def test_non_graph_host_refused_without_request_isc_35():
    class Boom(requests.Session):
        def get(self, *a, **kw):
            raise AssertionError("must not send")

    g = GraphClient(Tokens(), session=Boom())
    for url in ("https://evil.example.com/v1.0/me", "https://graph.microsoft.com.evil.io/x",
                "https://login.microsoftonline.com/x"):
        with pytest.raises(GraphError, match="non-Graph host"):
            g.get(url)


@responses.activate
def test_user_agent_and_auth_headers_isc_33():
    responses.get(URL, json={})
    g, _, _ = client()
    g.get(URL, headers={"Prefer": 'IdType="ImmutableId"'})
    h = responses.calls[0].request.headers
    assert h["User-Agent"] == USER_AGENT and USER_AGENT.startswith("o365-to-mailcow/")
    assert h["Authorization"] == "Bearer tok-1" and h["Prefer"] == 'IdType="ImmutableId"'


class Recorder(requests.Session):
    """Records kwargs and concurrency of every request."""

    def __init__(self, delay: float = 0.0) -> None:
        super().__init__()
        self.kwargs: list[dict] = []
        self.urls: list[str] = []
        self.active = 0
        self.max_active = 0
        self._lock = threading.Lock()
        self._delay = delay

    def get(self, url, **kw):
        with self._lock:
            self.kwargs.append(kw)
            self.urls.append(url)
            self.active += 1
            self.max_active = max(self.max_active, self.active)
        time.sleep(self._delay)
        with self._lock:
            self.active -= 1
        resp = requests.Response()
        resp.status_code = 200
        resp._content = b'{"value": []}'
        resp.raw = io.BytesIO(b'{"value": []}')  # get_bytes streams from raw
        return resp


def test_every_request_has_timeout_isc_34():
    rec = Recorder()
    g = GraphClient(Tokens(), session=rec)
    g.get("/me")
    g.get_bytes("/users/a/messages/1/$value")
    list(g.iter_pages("/users/a/mailFolders"))
    for kw in rec.kwargs:
        connect, read = kw["timeout"]
        assert connect > 0 and read > 0


def test_concurrency_capped_at_four_isc_32():
    rec = Recorder(delay=0.05)
    g = GraphClient(Tokens(), session=rec)
    threads = [threading.Thread(target=g.get, args=(f"/users/a/messages/{i}",))
               for i in range(12)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert len(rec.kwargs) == 12
    assert rec.max_active == 4


def test_only_graph_host_contacted_isc_35():
    rec = Recorder()
    g = GraphClient(Tokens(), session=rec)
    g.get("/users/a/mailFolders")
    g.get(f"{GRAPH_BASE}/users/a/contacts")
    assert {urlsplit(u).netloc for u in rec.urls} == {"graph.microsoft.com"}


def test_network_errors_retried_then_wrapped():
    class Flaky(requests.Session):
        def get(self, *a, **kw):
            raise requests.ConnectionError("down")

    sleeps: list[float] = []
    g = GraphClient(Tokens(), session=Flaky(), max_retries=2, sleep=sleeps.append)
    with pytest.raises(GraphError, match="network error"):
        g.get("/me")
    assert len(sleeps) == 2
