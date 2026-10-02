"""Configuration, read from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

# Microsoft's own "Microsoft Graph Command Line Tools" public client (used by Graph PowerShell),
# so no app registration is needed. Override with MS_CLIENT_ID to use your own registration.
GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

# Delegated Microsoft Graph permissions. Most of these need admin consent in the tenant.
DEFAULT_SCOPES = [
    "User.Read",
    "User.Read.All",                   # directory (AD) lookups: department, job title, etc.
    "Group.Read.All",                  # groups, group calendars
    "GroupMember.Read.All",            # group rosters
    "Team.ReadBasic.All",              # joined teams
    "Channel.ReadBasic.All",           # team channels
    "ChannelMessage.Read.All",         # channel messages
    "Chat.Read",                       # meeting chats (incl. breakout-room chats)
    "Calendars.Read",                  # calendar events / join URLs
    "OnlineMeetings.Read",             # resolve online meetings
    "OnlineMeetingArtifact.Read.All",  # attendance reports
]


@dataclass(frozen=True)
class Settings:
    client_id: str | None
    tenant: str
    scopes: list[str] = field(default_factory=lambda: list(DEFAULT_SCOPES))
    data_dir: Path = Path.home() / ".teams-rooms-mcp"
    auth_mode: str = "device"  # "device" or "interactive"

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant}"

    @property
    def token_cache_path(self) -> Path:
        return self.data_dir / "token_cache.json"

    @property
    def db_path(self) -> Path:
        return self.data_dir / "tracking.db"


def load_settings() -> Settings:
    scopes_env = os.environ.get("MS_GRAPH_SCOPES", "").split()
    data_dir = os.environ.get("TEAMS_MCP_DATA_DIR")
    auth_mode = os.environ.get("MS_AUTH_MODE", "device").strip().lower()
    if auth_mode not in ("device", "interactive"):
        raise RuntimeError("MS_AUTH_MODE must be 'device' or 'interactive'")
    return Settings(
        client_id=os.environ.get("MS_CLIENT_ID", "").strip() or GRAPH_CLI_CLIENT_ID,
        tenant=os.environ.get("MS_TENANT_ID", "organizations").strip() or "organizations",
        scopes=scopes_env or list(DEFAULT_SCOPES),
        data_dir=Path(data_dir).expanduser() if data_dir else Path.home() / ".teams-rooms-mcp",
        auth_mode=auth_mode,
    )
