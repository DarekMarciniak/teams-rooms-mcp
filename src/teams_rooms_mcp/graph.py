"""Thin async Microsoft Graph client: auth header, paging, throttling retries."""

from __future__ import annotations

import asyncio
from typing import Any

import httpx
from mcp.server.mcpserver.exceptions import ToolError

from .auth import Auth

GRAPH_V1 = "https://graph.microsoft.com/v1.0"
EVENTUAL = {"ConsistencyLevel": "eventual"}  # required for $search / $count advanced queries


class GraphError(ToolError):
    def __init__(self, status: int, body: str):
        self.status = status
        message = body
        try:
            import json

            err = json.loads(body).get("error", {})
            message = f"{err.get('code')}: {err.get('message')}"
        except Exception:  # noqa: BLE001
            pass
        hint = {
            401: " (token rejected - try auth_sign_out then auth_sign_in)",
            403: " (missing permission/admin consent, or you are not the meeting organizer)",
        }.get(status, "")
        super().__init__(f"Graph {status}: {message}{hint}")


class GraphClient:
    def __init__(self, auth: Auth):
        self._auth = auth
        self._http = httpx.AsyncClient(timeout=60)

    async def _token(self) -> str:
        token = await asyncio.to_thread(self._auth.token)
        if not token:
            raise GraphError(401, "Not signed in. Call the auth_sign_in tool first.")
        return token

    async def request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, str] | None = None,
        headers: dict[str, str] | None = None,
        json: Any = None,
    ) -> dict:
        url = path if path.startswith("https://") else GRAPH_V1 + path
        for attempt in range(5):
            token = await self._token()
            r = await self._http.request(
                method,
                url,
                params=params,
                json=json,
                headers={"Authorization": f"Bearer {token}", **(headers or {})},
            )
            if r.status_code in (429, 503, 504) and attempt < 4:
                delay = float(r.headers.get("Retry-After", 2**attempt))
                await asyncio.sleep(min(delay, 30))
                continue
            if r.status_code >= 400:
                raise GraphError(r.status_code, r.text)
            return r.json() if r.content else {}
        raise GraphError(429, "Throttled too many times")

    async def get(self, path: str, params: dict[str, str] | None = None, headers: dict[str, str] | None = None) -> dict:
        return await self.request("GET", path, params=params, headers=headers)

    async def get_all(
        self,
        path: str,
        params: dict[str, str] | None = None,
        headers: dict[str, str] | None = None,
        limit: int | None = None,
    ) -> list[dict]:
        """GET a collection, following @odata.nextLink until exhausted or `limit` reached."""
        items: list[dict] = []
        url: str | None = path
        while url:
            data = await self.request("GET", url, params=params, headers=headers)
            items.extend(data.get("value", []))
            if limit is not None and len(items) >= limit:
                return items[:limit]
            url = data.get("@odata.nextLink")
            params = None  # nextLink already carries the query
        return items
