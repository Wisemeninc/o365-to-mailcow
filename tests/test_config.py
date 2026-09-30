"""Config loading: file and env sources, secrets, permissions, mailbox lists (ISC-15..20)."""

from __future__ import annotations

import logging
import os
from pathlib import Path

import pytest

from o365_to_mailcow.config import (
    APP_SCOPES,
    ConfigError,
    MailboxMapping,
    load_config,
)

FULL = """
[microsoft]
tenant_id = "t-1"
client_id = "c-1"
auth_mode = "app"

[mailcow]
host = "mail.example.net"

[run]
state_dir = "{state}"
mailboxes = [
    "Alice@Contoso.com",
    {{ source = "bob@contoso.com", destination = "robert@example.net" }},
]
"""
ENV = {"O365MIG_CLIENT_SECRET": "s3cret-value", "O365MIG_MAILCOW_API_KEY": "key-value"}


def write(tmp_path: Path, text: str, mode: int = 0o600, name: str = "config.toml") -> Path:
    p = tmp_path / name
    p.write_text(text.format(state=tmp_path / "state"), encoding="utf-8")
    os.chmod(p, mode)
    return p


def test_loads_from_path_with_env_secrets_isc_15_16(tmp_path):
    cfg = load_config(write(tmp_path, FULL), env=ENV)
    assert cfg.tenant_id == "t-1" and cfg.client_id == "c-1" and cfg.auth_mode == "app"
    assert cfg.client_secret == "s3cret-value" and cfg.mailcow_api_key == "key-value"
    assert cfg.mailboxes == (MailboxMapping("alice@contoso.com", "alice@contoso.com"),
                             MailboxMapping("bob@contoso.com", "robert@example.net"))
    assert cfg.scopes == APP_SCOPES
    assert cfg.authority == "https://login.microsoftonline.com/t-1"
    assert cfg.max_message_bytes == 150 * 1024 * 1024
    assert cfg.calendar_attendees == "keep"


def test_loads_from_env_var_path_isc_15(tmp_path):
    path = write(tmp_path, FULL)
    cfg = load_config(None, env={**ENV, "O365MIG_CONFIG": str(path)})
    assert cfg.mailcow_host == "mail.example.net"


def test_no_config_path_is_error(tmp_path):
    with pytest.raises(ConfigError, match="O365MIG_CONFIG"):
        load_config(None, env={})
    with pytest.raises(ConfigError, match="not found"):
        load_config(tmp_path / "nope.toml", env={})


def test_permissive_mode_warns_naming_file_isc_17(tmp_path, caplog):
    path = write(tmp_path, FULL, mode=0o644)
    with caplog.at_level(logging.WARNING):
        load_config(path, env=ENV)
    assert str(path) in caplog.text and "chmod 600" in caplog.text


def test_private_mode_does_not_warn(tmp_path, caplog):
    with caplog.at_level(logging.WARNING):
        load_config(write(tmp_path, FULL), env=ENV)
    assert "readable by other users" not in caplog.text


def test_missing_keys_listed_in_one_error_isc_18(tmp_path):
    path = write(tmp_path, "[microsoft]\nauth_mode = 'app'\n")
    with pytest.raises(ConfigError) as exc:
        load_config(path, env={})
    text = str(exc.value)
    for key in ("microsoft.tenant_id", "microsoft.client_id", "mailcow.host",
                "mailcow.api_key", "microsoft.client_secret", "run.mailboxes"):
        assert key in text


def test_delegated_mode_needs_no_secret(tmp_path):
    text = FULL.replace('auth_mode = "app"', 'auth_mode = "delegated"')
    cfg = load_config(write(tmp_path, text), env={"O365MIG_MAILCOW_API_KEY": "k"})
    assert cfg.client_secret is None
    assert "Mail.Read.Shared" in cfg.scopes


def test_mailboxes_from_csv_with_mapping_isc_19_20(tmp_path):
    csv = tmp_path / "boxes.csv"
    csv.write_text("# source,destination\nCarol@contoso.com,carol@example.net\n"
                   "dave@contoso.com\n\n", encoding="utf-8")
    text = FULL.split("mailboxes = [")[0]  # no mailboxes in the file, only the CSV
    cfg = load_config(write(tmp_path, text), mailboxes_csv=str(csv), env=ENV)
    assert cfg.mailboxes == (MailboxMapping("carol@contoso.com", "carol@example.net"),
                             MailboxMapping("dave@contoso.com", "dave@contoso.com"))


def test_invalid_csv_row(tmp_path):
    csv = tmp_path / "boxes.csv"
    csv.write_text("not-an-address\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="invalid mailbox row"):
        load_config(write(tmp_path, FULL), mailboxes_csv=str(csv), env=ENV)


@pytest.mark.parametrize("old,new,match", [
    ('auth_mode = "app"', 'auth_mode = "magic"', "auth_mode"),
    ('host = "mail.example.net"', 'host = "https://mail.example.net/"', "bare hostname"),
    ('[run]', '[run]\ncalendar_attendees = "maybe"', "calendar_attendees"),
])
def test_invalid_values(tmp_path, old, new, match):
    with pytest.raises(ConfigError, match=match):
        load_config(write(tmp_path, FULL.replace(old, new)), env=ENV)


def test_env_secret_overrides_file(tmp_path):
    text = FULL.replace('auth_mode = "app"', 'auth_mode = "app"\nclient_secret = "from-file"')
    assert load_config(write(tmp_path, text), env=ENV).client_secret == "s3cret-value"
    env = {"O365MIG_MAILCOW_API_KEY": "k"}
    assert load_config(write(tmp_path, text), env=env).client_secret == "from-file"


def test_example_config_documents_every_key_isc_8():
    example = Path(__file__).resolve().parents[1] / "config.example.toml"
    text = example.read_text(encoding="utf-8")
    keys = ["tenant_id", "client_id", "auth_mode", "client_secret", "host", "api_key",
            "state_dir", "mailboxes", "mailboxes_csv", "parallel_mailboxes",
            "max_message_bytes", "calendar_exceptions_from_days",
            "calendar_exceptions_to_days", "contacts_photos", "calendar_attendees",
            "imap_port", "log_level", "source_folder_skip"]
    lines = text.splitlines()
    for key in keys:
        idx = next(i for i, line in enumerate(lines)
                   if line.lstrip("# ").startswith(f"{key} ="))
        assert lines[idx - 1].lstrip().startswith("#"), f"{key} has no comment above it"


def test_example_config_parses_and_loads(tmp_path):
    import shutil

    example = Path(__file__).resolve().parents[1] / "config.example.toml"
    target = tmp_path / "config.toml"
    shutil.copy(example, target)
    os.chmod(target, 0o600)
    cfg = load_config(str(target), env={"O365MIG_CLIENT_SECRET": "s3cret-value",
                                        "O365MIG_MAILCOW_API_KEY": "api-key-value"})
    assert cfg.mailboxes and any(m.aliases for m in cfg.mailboxes)
