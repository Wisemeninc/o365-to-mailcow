"""Contacts migrator with a fake Graph, fake DAV and a monkeypatched converter."""

from __future__ import annotations

import pytest

from fakes_o365 import MAPPING, FakeDav, FakeGraph, make_config
from o365_to_mailcow import contacts_sync
from o365_to_mailcow.contacts_sync import ContactsMigrator
from o365_to_mailcow.graph import GraphError
from o365_to_mailcow.state import STATUS_DONE, State

U = "/users/alice@contoso.com"


def contact(cid: str, lm: str = "2024-01-01T00:00:00Z", **extra) -> dict:
    return {"id": cid, "displayName": cid.upper(), "lastModifiedDateTime": lm, **extra}


def routes() -> dict:
    return {
        f"{U}/contactFolders": [{"id": "cf1", "displayName": "Suppliers"}],
        f"{U}/contactFolders/cf1/childFolders": [{"id": "cf2", "displayName": "Local"}],
        f"{U}/contactFolders/cf2/childFolders": [],
        f"{U}/contacts": [contact("c1"), contact("c2")],
        f"{U}/contactFolders/cf1/contacts": [contact("c3", categories=["VIP"])],
        f"{U}/contactFolders/cf2/contacts": [contact("c4")],
        f"{U}/contacts/c1/photo/$value": b"\xff\xd8JPEG",
        f"{U}/contacts/c3/photo/$value": GraphError(500, "photo service down", "x"),
    }


class Converter:
    def __init__(self) -> None:
        self.calls: list[tuple[dict, bytes | None]] = []

    def __call__(self, c, photo=None):
        self.calls.append((c, photo))
        mod = contacts_sync.contacts_conv
        return mod.ConvertedContact(f"uid-{c['id']}", b"VCARD", c.get("lastModifiedDateTime"))


@pytest.fixture
def env(tmp_path, monkeypatch):
    conv = Converter()
    monkeypatch.setattr(contacts_sync.contacts_conv, "convert_contact", conv)
    cfg = make_config(tmp_path, contacts_photos=True)
    state = State(tmp_path / "state.db")
    yield cfg, state, FakeGraph(routes()), FakeDav(), conv
    state.close()


def mig(cfg, state, graph, dav, dry_run=False):
    return ContactsMigrator(cfg, graph, state, dav, MAPPING, dry_run)


def test_plan_recurses_child_folders_isc_127(env):
    cfg, state, graph, _, _ = env
    plan = mig(cfg, state, graph, None, dry_run=True).plan()
    assert [(c.name, c.slug, c.count) for c in plan.collections] == [
        ("Contacts", "personal", 2), ("Suppliers", "suppliers", 1), ("Local", "local", 1)]


def test_migrate_default_to_personal_others_mkcol_isc_84_90(env):
    cfg, state, graph, dav, _ = env
    res = mig(cfg, state, graph, dav).migrate()
    assert ("MKCOL", "suppliers") in dav.writes and ("MKCOL", "local") in dav.writes
    assert ("MKCOL", "personal") not in dav.writes
    assert set(dav.books["personal"]) == {"uid-c1", "uid-c2"}
    assert set(dav.books["suppliers"]) == {"uid-c3"}
    assert set(dav.books["local"]) == {"uid-c4"}
    assert res.failed == 0 and not res.fallbacks
    assert state.contact_counts(MAPPING.source)["personal"][STATUS_DONE] == 2


def test_photos_fetched_404_is_no_photo_isc_88_111(env):
    cfg, state, graph, dav, conv = env
    res = mig(cfg, state, graph, dav).migrate()
    photos = {c["id"]: p for c, p in conv.calls}
    assert photos == {"c1": b"\xff\xd8JPEG", "c2": None, "c3": None, "c4": None}
    assert res.failed == 0
    assert res.warnings == ["1 contact photo(s) could not be fetched and were left out"]


def test_photos_not_fetched_when_disabled(env, tmp_path):
    _, state, graph, dav, conv = env
    cfg = make_config(tmp_path, contacts_photos=False)
    mig(cfg, state, graph, dav).migrate()
    assert not [p for p in graph.paths() if p.endswith("/photo/$value")]


def test_second_run_idempotent_on_last_modified(env):
    cfg, state, graph, dav, conv = env
    mig(cfg, state, graph, dav).migrate()
    conv.calls.clear()
    res = mig(cfg, state, graph, dav).migrate()
    # c3's photo fetch failed (not a 404), so it alone is retried on the next run
    assert [c["id"] for c, _ in conv.calls] == ["c3"]
    assert sum(c.unchanged for c in res.collections) == 3
    conv.calls.clear()
    graph.routes[f"{U}/contacts"][1]["lastModifiedDateTime"] = "2025-01-01T00:00:00Z"
    mig(cfg, state, graph, dav).migrate()
    assert sorted(c["id"] for c, _ in conv.calls) == ["c2", "c3"]


def test_mkcol_refused_falls_back_to_personal_with_category_isc_84(env):
    cfg, state, graph, _, conv = env
    dav = FakeDav(refuse_mkcol=True)
    res = mig(cfg, state, graph, dav).migrate()
    assert set(dav.books["personal"]) == {"uid-c1", "uid-c2", "uid-c3", "uid-c4"}
    cats = {c["id"]: c.get("categories") for c, _ in conv.calls}
    assert cats["c3"] == ["VIP", "Suppliers"] and cats["c4"] == ["Local"]
    assert cats["c1"] is None
    assert len(res.fallbacks) == 2 and "Suppliers" in res.fallbacks[0]
    v = mig(cfg, state, graph, dav).verify()
    assert [(c.slug, c.graph_count, c.dav_count, c.mismatch) for c in v.collections] == [
        ("personal", 4, 4, False)]
    assert len(v.fallbacks) == 2


def test_put_failure_recorded(env):
    from o365_to_mailcow.dav import DavError

    cfg, state, graph, dav, _ = env
    dav.put_errors["uid-c2"] = DavError(400, "bad vcard")
    res = mig(cfg, state, graph, dav).migrate()
    assert res.failed == 1 and sum(c.put for c in res.collections) == 3


def test_missing_addressbook_home_isc_107(env):
    cfg, state, graph, _, _ = env
    dav = FakeDav(addressbook_home=False)
    res = mig(cfg, state, graph, dav).migrate()
    assert dav.writes == [] and "not found" in res.errors[0]


def test_dry_run_counts_without_dav(env):
    cfg, state, graph, _, conv = env
    res = mig(cfg, state, graph, None, dry_run=True).migrate()
    assert sum(c.would_put for c in res.collections) == 4 and conv.calls == []


def test_verify_per_book_isc_91(env):
    cfg, state, graph, dav, _ = env
    mig(cfg, state, graph, dav).migrate()
    v = mig(cfg, state, graph, dav).verify()
    assert [(c.slug, c.graph_count, c.dav_count, c.mismatch) for c in v.collections] == [
        ("personal", 2, 2, False), ("suppliers", 1, 1, False), ("local", 1, 1, False)]
