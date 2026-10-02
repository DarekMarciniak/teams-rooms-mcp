# teams-rooms-mcp

A local MCP server for Microsoft Teams. Through Microsoft Graph it reads:

- teams
- groups
- channel and meeting messages
- calendars
- attendance reports
- users in the directory (Microsoft Entra ID, formerly Azure AD)

It also keeps a **local tracking database**: which student was in which breakout room of which meeting, and which group (year/subject) they belong to.

It works with Claude Desktop (including Cowork), Claude Code, or any MCP client that runs stdio servers.

## 1. Sign-in: no app registration needed

You sign in with your own Microsoft 365 account through standard OAuth: device code, or a browser window. Everything runs with **your** permissions (delegated access), so the server can only see what you can see in Teams.

Every OAuth sign-in has to name a client app. By default this server uses Microsoft's own public client, **Microsoft Graph Command Line Tools** (`14d82eec-204b-4c2f-b7e8-296a70dab67e`). It is the same client Graph PowerShell (`Connect-MgGraph`) uses, so you don't need to register anything. No client secret is involved.

On first sign-in Microsoft asks you to consent to the permissions. Some of them need **admin consent** in the tenant:

| Needs admin consent | What breaks without it |
|---|---|
| `OnlineMeetingArtifact.Read.All` | attendance reports (`get_meeting_attendance`, `sync_meeting_attendance`) |
| `Group.Read.All`, `GroupMember.Read.All` | listing groups and rosters (`list_groups`, `get_group_members`, `sync_group_roster`) |
| `User.Read.All` | full directory profiles (department, job title, manager) |
| `ChannelMessage.Read.All` | channel messages |

Many school tenants have already approved Graph PowerShell for these. If yours hasn't, sign-in fails with "Need admin approval". You then have two options:

- **Ask an admin to approve them** (one click in the consent prompt, or Entra → Enterprise apps → *Microsoft Graph Command Line Tools* → Permissions).
- **Run with only the permissions a normal user can grant**:
  `MS_GRAPH_SCOPES="User.Read User.ReadBasic.All Team.ReadBasic.All Channel.ReadBasic.All Chat.Read Calendars.Read OnlineMeetings.Read"`
  This still gives you teams, channels, meeting and breakout-room chats, calendars and basic user lookup. Attendance then has to come from downloaded files (`import_attendance_file`) or manual assignment.

**Optional: your own app registration.** Use one if your tenant blocks the Graph CLI client, or if you want a dedicated, named app:

1. In Entra, create a new app registration (single tenant).
2. Under **Authentication**, add a *Mobile and desktop* platform with redirect `http://localhost`, and set **Allow public client flows = Yes**.
3. Add the delegated Graph permissions listed above.
4. Set `MS_CLIENT_ID` and `MS_TENANT_ID`.

## 2. Install

You don't need to clone anything. `uvx` downloads and runs the server straight from GitHub.

1. Install **uv** and **Git** (once). In PowerShell:
   ```powershell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   winget install Git.Git
   ```
2. Close and reopen PowerShell, then check it works. The first run downloads everything; the server then waits for input, so press `Ctrl+C`:
   ```powershell
   uvx --from git+https://github.com/DarekMarciniak/teams-rooms-mcp@v0.1.0 teams-rooms-mcp
   ```

## 3. Connect to Claude

**Claude Desktop / Cowork.** Open `%APPDATA%\Claude\claude_desktop_config.json`. In Claude Desktop that's *Settings → Developer → Edit Config*. Add the server below, then fully quit Claude (from the tray icon) and start it again:

```json
{
  "mcpServers": {
    "teams-rooms": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/DarekMarciniak/teams-rooms-mcp@v0.1.0", "teams-rooms-mcp"]
    }
  }
}
```

If the file already has an `mcpServers` section, add only the `"teams-rooms": {...}` entry inside it. If Claude can't find `uvx`, put its full path in `command`, e.g. `C:\Users\<name>\.local\bin\uvx.exe`.

**Claude Code:**

```powershell
claude mcp add teams-rooms -- uvx --from git+https://github.com/DarekMarciniak/teams-rooms-mcp@v0.1.0 teams-rooms-mcp
```

**Updating:** when a new version is tagged, change `@v0.1.0` to the new tag in the config and restart Claude.

**Development (from a clone):**

```powershell
git clone https://github.com/DarekMarciniak/teams-rooms-mcp
cd teams-rooms-mcp
uv sync
uv run teams-rooms-mcp
```

### Environment variables (all optional)

| Variable | Default | Meaning |
|---|---|---|
| `MS_CLIENT_ID` | Graph CLI client | Client ID; set only if you use your own app registration |
| `MS_TENANT_ID` | `organizations` | Your tenant ID or domain (e.g. `school.edu`). Recommended if you have accounts in several tenants. |
| `MS_AUTH_MODE` | `device` | `device`: Claude shows a code to enter at microsoft.com/devicelogin. `interactive`: a browser login window opens on this machine. |
| `MS_GRAPH_SCOPES` | full set above | Space-separated override, e.g. the no-admin set |
| `TEAMS_MCP_DATA_DIR` | `~/.teams-rooms-mcp` | Where the token cache and `tracking.db` are kept |

## 4. Sign in

Ask Claude to *"sign in to Teams"*. It calls `auth_sign_in` and shows you a URL and a code. Sign in there, then tell Claude you're done (it will check `auth_status`). Tokens are cached and refreshed automatically.

## Tools

| Area | Tools |
|---|---|
| Auth | `auth_sign_in`, `auth_status`, `auth_sign_out` |
| Directory | `whoami`, `search_users`, `get_user`, `get_user_groups` |
| Groups/Teams | `list_groups`, `get_group_members`, `list_my_teams`, `list_team_channels`, `list_channel_messages` |
| Meetings | `list_calendar_events`, `find_online_meeting`, `get_meeting_attendance`, `list_chats`, `list_chat_messages` |
| Tracking | `sync_group_roster`, `sync_meeting_attendance`, `import_attendance_file`, `record_room_assignment` |
| Reports | `tracking_overview`, `meeting_room_report`, `group_attendance`, `student_history`, `query_tracking_db`, `delete_tracked_meeting` |

## Breakout-room tracking: how it works

Graph has **no single API that says "student X was in breakout room Y"**. In Teams, each breakout room is its own small meeting with its own chat. So the server records presence per `(meeting_label, room, person)` and accepts it from three sources:

1. **Graph attendance reports** (`sync_meeting_attendance`). Run it once for the main meeting (`room="Main"`) and once per room meeting. To find each room's join URL, use `list_chats(topic_contains="Room")`. **Use the same `meeting_label` for all of them.** Whether Graph returns attendance reports for breakout-room meetings depends on the tenant. If it returns nothing, use source 2 or 3.
2. **Attendance files** downloaded from Teams (`import_attendance_file`). Supports the Teams UTF-16 export, or any CSV/TSV with Name and Email columns. If the file has a `Room` or `Breakout room` column, it is used per row.
3. **Manual assignment** (`record_room_assignment`). For example, from the breakout-room assignment screen.

Run `sync_group_roster` for each year/subject group. Every report then shows each student's groups, and `group_attendance` lists who in a class was present (and in which rooms), who was absent, and who attended without being in the group.

Example prompts:

- *"Sync rosters for all groups whose name starts with 'Y2'."*
- *"Find yesterday's Y2 Math meeting, sync main-meeting attendance as 'Math Y2 2026-10-01' linked to the Y2 Math group, then do the same for each breakout room chat created yesterday."*
- *"Who in Y2 Math missed today, and which room was Anna in?"*
- *"Across all meetings this month, which students were never put in the same room as each other?"* (Claude can answer this with `query_tracking_db`.)

### Limitations

- Attendance reports only exist for meetings **you organized**, or ones covered by an admin application access policy. They appear a few minutes after the meeting ends.
- Some permissions need admin consent (see section 1).
- The tracking DB holds personal data about students. It stays local in `TEAMS_MCP_DATA_DIR`, so treat that folder under your school's data-protection rules.
