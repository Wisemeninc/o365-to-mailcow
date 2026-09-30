"""ImapDestination against an autospec'd IMAPClient: allowlist, reconnect, APPENDUID."""

from __future__ import annotations

from datetime import UTC, datetime
from unittest.mock import create_autospec

import pytest
from imapclient import IMAPClient
from imapclient.exceptions import IMAPClientAbortError, IMAPClientError, LoginError
from imapclient.imap_utf7 import decode, encode

from o365_to_mailcow.imap_dest import ImapConnectionError, ImapDestination, ImapError

# ISC-54: the only IMAPClient methods the tool may call, with the command each issues.
ALLOWED = {
    "login": "LOGIN", "list_folders": "LIST", "create_folder": "CREATE",
    "subscribe_folder": "SUBSCRIBE", "folder_status": "STATUS",
    "select_folder": "SELECT/EXAMINE", "search": "SEARCH", "append": "APPEND",
    "fetch": "FETCH", "logout": "LOGOUT",
}
FORBIDDEN = {"delete_messages", "expunge", "uid_expunge", "delete_folder", "add_flags",
             "set_flags", "remove_flags", "move", "copy", "rename_folder", "namespace"}


class Factory:
    """Hands out autospec'd IMAPClient mocks and remembers them."""

    def __init__(self, n: int = 3) -> None:
        self.clients = [create_autospec(IMAPClient, instance=True) for _ in range(n)]
        for c in self.clients:
            c.list_folders.return_value = [((b"\\HasNoChildren",), b"/", "INBOX")]
            c.folder_status.return_value = {b"MESSAGES": 3, b"UIDVALIDITY": 42}
            c.append.return_value = b"[APPENDUID 42 17] Append completed."
            c.search.return_value = [5]
            c.fetch.return_value = {5: {b"BODY[]": b"raw", b"SEQ": 1}}
        self.created = 0
        self.kwargs: list[dict] = []

    def __call__(self, host, **kwargs):
        self.kwargs.append({"host": host, **kwargs})
        client = self.clients[self.created]
        self.created += 1
        return client


def used_methods(factory: Factory) -> set[str]:
    return {call[0].split(".")[0] for c in factory.clients[:factory.created]
            for call in c.method_calls}


def make(factory: Factory) -> ImapDestination:
    return ImapDestination("mail.example.net", 993, "alice@example.net", "pw-secret",
                           client_factory=factory)


def test_full_session_uses_only_allowed_commands_isc_54_55():
    f = Factory()
    d = make(f)
    d.connect()
    assert d.delimiter == "/"
    f.clients[0].list_folders.return_value = []
    d.ensure_folder("Projekte/Bestätigungen")
    assert d.folder_status("INBOX") == (3, 42)
    assert d.has_message_id("INBOX", "<a@b>")
    assert d.append("INBOX", b"x", ["\\Seen"], datetime(2024, 1, 1, tzinfo=UTC)) == 17
    assert d.fetch_message("INBOX", 5) == b"raw"
    d.close()
    used = used_methods(f)
    assert used <= set(ALLOWED), used - set(ALLOWED)
    assert not used & FORBIDDEN


def test_connect_uses_tls_with_verification_and_timeouts():
    f = Factory()
    make(f).connect()
    kw = f.kwargs[0]
    assert kw["ssl"] is True and kw["port"] == 993
    assert kw["ssl_context"].verify_mode.name == "CERT_REQUIRED"
    assert kw["ssl_context"].check_hostname is True
    assert kw["timeout"].connect and kw["timeout"].read
    f.clients[0].login.assert_called_once_with("alice@example.net", "pw-secret")


def test_ensure_folder_is_idempotent_and_never_creates_inbox():
    f = Factory()
    d = make(f)
    c = f.clients[0]
    c.list_folders.return_value = [((), b"/", "Archive")]
    d.ensure_folder("Archive")
    c.create_folder.assert_not_called()
    c.subscribe_folder.assert_called_with("Archive")
    d.ensure_folder("INBOX")
    c.create_folder.assert_not_called()
    c.list_folders.return_value = [((), b"/", "Archive2")]  # wildcard-ish partial match
    d.ensure_folder("Archive")
    c.create_folder.assert_called_once_with("Archive")


def test_search_is_read_only_select():
    f = Factory()
    d = make(f)
    d.has_message_id("INBOX", "<a@b>")
    f.clients[0].select_folder.assert_called_with("INBOX", readonly=True)
    f.clients[0].search.assert_called_with(["HEADER", "Message-ID", "<a@b>"])


def test_append_without_appenduid_returns_none():
    f = Factory()
    f.clients[0].append.return_value = b"Append completed."
    assert make(f).append("INBOX", b"x", [], datetime.now(UTC)) is None


def test_append_passes_flags_and_date():
    f = Factory()
    when = datetime(2024, 1, 2, 3, 4, 5, tzinfo=UTC)
    make(f).append("Sent", b"mime", ["\\Seen", "Work"], when)
    f.clients[0].append.assert_called_once_with("Sent", b"mime", flags=("\\Seen", "Work"),
                                                 msg_time=when)


def test_dropped_connection_reconnects_once_isc_100():
    f = Factory()
    f.clients[0].append.side_effect = IMAPClientAbortError("socket closed")
    d = make(f)
    assert d.append("INBOX", b"x", [], datetime.now(UTC)) == 17
    assert f.created == 2
    f.clients[1].append.assert_called_once()


def test_second_connection_failure_raises_isc_100():
    f = Factory()
    f.clients[0].append.side_effect = OSError("reset")
    f.clients[1].append.side_effect = OSError("reset again")
    with pytest.raises(ImapConnectionError):
        make(f).append("INBOX", b"x", [], datetime.now(UTC))
    assert f.created == 2


def test_no_response_is_imap_error_without_retry_isc_99():
    f = Factory()
    f.clients[0].append.side_effect = IMAPClientError("append failed: [OVERQUOTA] Quota")
    with pytest.raises(ImapError, match="OVERQUOTA"):
        make(f).append("INBOX", b"x", [], datetime.now(UTC))
    assert f.created == 1


def test_login_failure_does_not_leak_password():
    f = Factory()
    f.clients[0].login.side_effect = LoginError("AUTHENTICATIONFAILED pw-secret?")
    with pytest.raises(ImapError) as exc:
        make(f).connect()
    assert "pw-secret" not in str(exc.value)


def test_modified_utf7_round_trip_isc_40():
    for name in ("Bestätigungen", "Projekte & Co", "日本語/メール", "~peter"):
        encoded = encode(name)
        assert encoded.isascii()
        assert decode(encoded) == name


def test_class_has_no_destructive_methods_isc_55():
    names = {n.lower() for n in dir(ImapDestination)}
    for bad in ("expunge", "delete", "store", "move", "rename"):
        assert not any(bad in n for n in names), bad
