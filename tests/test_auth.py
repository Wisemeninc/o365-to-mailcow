"""TokenProvider with MSAL mocked (ISC-22..27)."""

from __future__ import annotations

import io
import stat
import types
from pathlib import Path

import msal
import pytest

from fakes_o365 import make_config
from o365_to_mailcow import auth
from o365_to_mailcow.auth import AuthError, TokenProvider
from o365_to_mailcow.config import DELEGATED_SCOPES


class FakeCache:
    def __init__(self) -> None:
        self.has_state_changed = True
        self.removed: list = []
        self.loaded: str | None = None

    def deserialize(self, text: str) -> None:
        self.loaded = text

    def serialize(self) -> str:
        return '{"AccessToken": {}}'

    def find(self, kind):
        return [{"secret": "at-1"}, {"secret": "at-2"}]

    def remove_at(self, item) -> None:
        self.removed.append(item)


class FakeApp:
    instances: list[FakeApp] = []

    def __init__(self, client_id, **kwargs) -> None:
        self.client_id = client_id
        self.kwargs = kwargs
        self.accounts: list[dict] = []
        self.calls: list[tuple[str, object]] = []
        self.result = {"access_token": "tok"}
        FakeApp.instances.append(self)

    def acquire_token_for_client(self, scopes):
        self.calls.append(("client", scopes))
        return self.result

    def get_accounts(self):
        return self.accounts

    def acquire_token_silent(self, scopes, account):
        self.calls.append(("silent", scopes))
        return {"access_token": "silent-tok"}

    def initiate_device_flow(self, scopes):
        self.calls.append(("device", scopes))
        return {"user_code": "ABCD", "message": "Go to https://microsoft.com/devicelogin "
                                                "and enter ABCD"}

    def acquire_token_by_device_flow(self, flow):
        self.calls.append(("device_token", flow["user_code"]))
        return self.result


@pytest.fixture(autouse=True)
def fake_msal(monkeypatch):
    FakeApp.instances.clear()
    monkeypatch.setattr(auth.msal, "ConfidentialClientApplication", FakeApp)
    monkeypatch.setattr(auth.msal, "PublicClientApplication", FakeApp)
    monkeypatch.setattr(auth.msal, "SerializableTokenCache", FakeCache)


def test_app_mode_client_credentials_isc_22(tmp_path):
    tp = TokenProvider(make_config(tmp_path))
    assert tp.get_token() == "tok"
    app = FakeApp.instances[0]
    assert app.client_id == "client"
    assert app.kwargs["authority"] == "https://login.microsoftonline.com/tenant"
    assert app.kwargs["client_credential"] == "client-secret-value-123"
    assert app.calls == [("client", ["https://graph.microsoft.com/.default"])]
    assert not (tmp_path / "state" / "msal_cache.bin").exists()  # nothing persisted


def test_delegated_device_code_prints_code_and_url_isc_23_24(tmp_path):
    out = io.StringIO()
    cfg = make_config(tmp_path, auth_mode="delegated", client_secret=None)
    tp = TokenProvider(cfg, out=out)
    assert tp.get_token() == "tok"
    app = FakeApp.instances[0]
    assert app.calls[0] == ("device", DELEGATED_SCOPES)
    assert "https://microsoft.com/devicelogin" in out.getvalue() and "ABCD" in out.getvalue()
    assert set(DELEGATED_SCOPES) == {"Mail.Read.Shared", "Calendars.Read.Shared",
                                     "Contacts.Read.Shared", "User.Read"}


def test_msal_adds_offline_access_itself_isc_24():
    """ISC-24 wants offline_access; MSAL rejects it as input and always adds it."""
    fake_self = types.SimpleNamespace(_exclude_scopes=frozenset())
    decorated = msal.ClientApplication._decorate_scope(fake_self, list(DELEGATED_SCOPES))
    assert "offline_access" in decorated
    with pytest.raises(ValueError):
        msal.ClientApplication._decorate_scope(fake_self, ["offline_access"])


def test_delegated_cache_persisted_mode_0600_isc_25(tmp_path):
    cfg = make_config(tmp_path, auth_mode="delegated", client_secret=None)
    TokenProvider(cfg, out=io.StringIO()).get_token()
    path = tmp_path / "state" / "msal_cache.bin"
    assert path.is_file()
    assert stat.S_IMODE(path.stat().st_mode) == 0o600
    tp2 = TokenProvider(cfg, out=io.StringIO())
    assert tp2._cache.loaded == '{"AccessToken": {}}'


def test_delegated_silent_when_account_cached(tmp_path):
    cfg = make_config(tmp_path, auth_mode="delegated", client_secret=None)
    tp = TokenProvider(cfg, out=io.StringIO())
    FakeApp.instances[0].accounts = [{"username": "admin"}]
    assert tp.get_token() == "silent-tok"
    assert [c[0] for c in FakeApp.instances[0].calls] == ["silent"]


def test_token_asked_from_msal_on_every_call_isc_26(tmp_path):
    """No token caching in our layer: MSAL refreshes before expiry, so a long run keeps
    getting fresh tokens. Simulated with a fake clock that expires tokens after 60 min."""
    clock = {"t": 0.0}
    tp = TokenProvider(make_config(tmp_path))
    app = FakeApp.instances[0]

    def acquire(scopes):
        return {"access_token": f"tok-{int(clock['t'] // 3600)}"}

    app.acquire_token_for_client = acquire
    assert tp.get_token() == "tok-0"
    clock["t"] = 3599
    assert tp.get_token() == "tok-0"
    clock["t"] = 3601  # past the 60-minute mark
    assert tp.get_token() == "tok-1"


def test_error_result_raises_auth_error(tmp_path):
    tp = TokenProvider(make_config(tmp_path))
    FakeApp.instances[0].result = {"error": "invalid_client",
                                   "error_description": "AADSTS7000215: Invalid secret"}
    with pytest.raises(AuthError, match="invalid_client: AADSTS7000215"):
        tp.get_token()


def test_invalidate_removes_access_tokens(tmp_path):
    tp = TokenProvider(make_config(tmp_path))
    tp.invalidate()
    assert tp._cache.removed == [{"secret": "at-1"}, {"secret": "at-2"}]


def test_no_write_scope_anywhere_in_src_isc_27():
    src = Path(auth.__file__).parent
    for path in src.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        for bad in ("Mail.ReadWrite", "Calendars.ReadWrite", "Contacts.ReadWrite",
                    "MailboxSettings", "Mail.Send"):
            assert bad not in text, f"{bad} in {path.name}"
