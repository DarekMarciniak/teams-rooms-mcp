"""MSAL public-client auth with a persistent token cache.

Device code flow runs in a background thread so the MCP tool can return the
code/URL to the user immediately (stdout belongs to the MCP protocol, so we
can't just print it).
"""

from __future__ import annotations

import logging
import os
import threading
from typing import Any

import msal
from mcp.server.mcpserver.exceptions import ToolError

from .config import Settings

log = logging.getLogger(__name__)


class AuthError(ToolError):
    pass


class Auth:
    def __init__(self, settings: Settings):
        self._s = settings
        self._lock = threading.Lock()
        self._cache = msal.SerializableTokenCache()
        if settings.token_cache_path.exists():
            self._cache.deserialize(settings.token_cache_path.read_text(encoding="utf-8"))
        self._app: msal.PublicClientApplication | None = None
        self._pending: dict[str, Any] | None = None

    @property
    def app(self) -> msal.PublicClientApplication:
        if self._app is None:
            if not self._s.client_id:
                raise AuthError(
                    "MS_CLIENT_ID is not configured. Register an app in Microsoft Entra ID "
                    "and set MS_CLIENT_ID / MS_TENANT_ID in the MCP server config (see README)."
                )
            self._app = msal.PublicClientApplication(
                self._s.client_id, authority=self._s.authority, token_cache=self._cache
            )
        return self._app

    def _persist(self) -> None:
        with self._lock:
            if not self._cache.has_state_changed:
                return
            path = self._s.token_cache_path
            path.parent.mkdir(parents=True, exist_ok=True)
            tmp = path.with_suffix(".tmp")
            tmp.write_text(self._cache.serialize(), encoding="utf-8")
            if os.name == "posix":
                os.chmod(tmp, 0o600)
            os.replace(tmp, path)
            self._cache.has_state_changed = False

    def account(self) -> dict | None:
        accounts = self.app.get_accounts()
        return accounts[0] if accounts else None

    def token(self) -> str | None:
        """Return a valid access token from cache (refreshing if needed), or None."""
        acct = self.account()
        if not acct:
            return None
        result = self.app.acquire_token_silent(self._s.scopes, account=acct)
        self._persist()
        if result and "access_token" in result:
            return result["access_token"]
        return None

    def start_device_flow(self) -> dict:
        with self._lock:
            pending = self._pending
            if pending and pending["thread"].is_alive():
                return _flow_info(pending["flow"])
            flow = self.app.initiate_device_flow(scopes=self._s.scopes)
            if "user_code" not in flow:
                raise AuthError(flow.get("error_description") or str(flow))
            pending = {"flow": flow, "result": None}

            def run() -> None:
                try:
                    pending["result"] = self.app.acquire_token_by_device_flow(flow)
                except Exception as e:  # noqa: BLE001 - surface any failure via auth_status
                    pending["result"] = {"error": type(e).__name__, "error_description": str(e)}
                self._persist()

            pending["thread"] = threading.Thread(target=run, daemon=True)
            self._pending = pending
            pending["thread"].start()
        return _flow_info(flow)

    def interactive_sign_in(self) -> dict:
        result = self.app.acquire_token_interactive(self._s.scopes, prompt="select_account")
        self._persist()
        if "access_token" not in result:
            raise AuthError(_explain(result.get("error_description") or str(result)))
        return {"status": "signed_in", "account": result.get("id_token_claims", {}).get("preferred_username")}

    def status(self) -> dict:
        acct = self.account()
        if acct and self.token():
            return {"status": "signed_in", "account": acct.get("username")}
        pending = self._pending
        if pending:
            if pending["thread"].is_alive():
                return {"status": "waiting_for_user", **_flow_info(pending["flow"])}
            result = pending["result"] or {}
            if "error" in result:
                return {"status": "failed", "error": _explain(result.get("error_description") or result["error"])}
        return {"status": "signed_out"}

    def sign_out(self) -> None:
        for acct in self.app.get_accounts():
            self.app.remove_account(acct)
        self._persist()
        self._s.token_cache_path.unlink(missing_ok=True)


def _explain(error: str) -> str:
    if any(code in error for code in ("AADSTS65001", "AADSTS90094", "AADSTS90008")):
        return (
            f"{error}\n\nSome requested permissions need admin consent in your tenant. Either ask an "
            "admin to approve them, or set MS_GRAPH_SCOPES to the no-admin set listed in the README."
        )
    if "AADSTS53003" in error or "AADSTS50105" in error:
        return f"{error}\n\nYour tenant blocks this client app. Ask IT, or use your own app registration (MS_CLIENT_ID)."
    return error


def _flow_info(flow: dict) -> dict:
    return {
        "status": "waiting_for_user",
        "verification_uri": flow["verification_uri"],
        "user_code": flow["user_code"],
        "message": flow["message"],
        "next_step": "Open the URL, enter the code and sign in, then call auth_status.",
    }
