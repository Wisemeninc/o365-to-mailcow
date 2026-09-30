"""Microsoft sign-in via MSAL.

Two modes:

* ``app``       - client credentials. No user involvement. Needs application permissions
                  (Mail.Read, Calendars.Read, Contacts.Read) with admin consent.
* ``delegated`` - device-code flow for an administrator who has Full Access on the target
                  mailboxes. Uses the ``.Shared`` delegated scopes. The token cache is kept
                  on disk (mode 0600) so a long run and later re-runs do not prompt again.

MSAL refreshes tokens itself; callers simply ask for a token before every request.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

import msal

from .config import Config

log = logging.getLogger(__name__)


class AuthError(Exception):
    pass


class TokenProvider:
    """Returns a valid bearer token for Microsoft Graph, refreshing as needed."""

    def __init__(self, cfg: Config, cache_path: Path | None = None, out=None) -> None:
        self._cfg = cfg
        self._scopes = cfg.scopes
        self._out = out or sys.stdout
        self._cache_path = cache_path or (cfg.state_dir / "msal_cache.bin")
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        if cfg.auth_mode == "app":
            self._app: msal.ClientApplication = msal.ConfidentialClientApplication(
                cfg.client_id,
                authority=cfg.authority,
                client_credential=cfg.client_secret,
                token_cache=self._cache,
            )
        else:
            self._app = msal.PublicClientApplication(
                cfg.client_id, authority=cfg.authority, token_cache=self._cache
            )

    # -- cache -------------------------------------------------------------------------

    def _load_cache(self) -> None:
        if self._cfg.auth_mode != "delegated":
            return
        try:
            if self._cache_path.is_file():
                self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:  # corrupt cache: start fresh, never crash
            log.warning("ignoring unreadable token cache %s: %s", self._cache_path, exc)

    def _save_cache(self) -> None:
        if self._cfg.auth_mode != "delegated" or not self._cache.has_state_changed:
            return
        self._cache_path.parent.mkdir(parents=True, exist_ok=True)
        flags = os.O_WRONLY | os.O_CREAT | os.O_TRUNC | getattr(os, "O_NOFOLLOW", 0)
        fd = os.open(self._cache_path, flags, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(self._cache.serialize())
        os.chmod(self._cache_path, 0o600)

    # -- tokens ------------------------------------------------------------------------

    def get_token(self) -> str:
        if self._cfg.auth_mode == "app":
            result = self._app.acquire_token_for_client(scopes=self._scopes)
        else:
            result = self._acquire_delegated()
        if "access_token" not in result:
            raise AuthError(
                f"{result.get('error', 'unknown_error')}: "
                f"{result.get('error_description', 'no description')}"
            )
        self._save_cache()
        return result["access_token"]

    def _acquire_delegated(self) -> dict:
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(self._scopes, account=accounts[0])
            if result and "access_token" in result:
                return result
        flow = self._app.initiate_device_flow(scopes=self._scopes)
        if "user_code" not in flow:
            raise AuthError(f"device flow failed: {flow.get('error_description', flow)}")
        print(flow["message"], file=self._out, flush=True)
        return self._app.acquire_token_by_device_flow(flow)

    def invalidate(self) -> None:
        """Drop cached access tokens so the next call fetches a fresh one (after a 401)."""
        for token in self._cache.find(msal.TokenCache.CredentialType.ACCESS_TOKEN):
            self._cache.remove_at(token)
