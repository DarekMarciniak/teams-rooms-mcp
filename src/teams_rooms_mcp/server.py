"""MCP server: Microsoft Teams / Graph access plus breakout-room attendance tracking."""

from __future__ import annotations

import asyncio
import functools
import html
import logging
import re
import sqlite3
import sys
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

from .attendance_file import parse_attendance_file
from .auth import Auth
from .config import load_settings
from .graph import EVENTUAL, GraphClient
from .store import Store

logging.basicConfig(stream=sys.stderr, level=logging.INFO)  # stdout is the MCP channel

settings = load_settings()
auth = Auth(settings)
graph = GraphClient(auth)
store = Store(settings.db_path)

READ_ONLY = ToolAnnotations(read_only_hint=True)
DESTRUCTIVE = ToolAnnotations(destructive_hint=True)

mcp = MCPServer(
    "teams-rooms",
    instructions="""Microsoft Teams / Graph access for managing student groups and breakout rooms.

Typical workflow:
1. auth_status; if signed out call auth_sign_in and show the user the URL + code.
2. Rosters: list_groups / list_my_teams, then sync_group_roster for each year/subject group.
3. Find the class meeting: list_calendar_events (gives joinUrl) or find_online_meeting.
4. Record who was where, using a stable meeting_label (e.g. 'Math Y2 2026-10-01'):
   - main meeting: sync_meeting_attendance(room='Main', group_id=<class group>)
   - each breakout room: find its join URL with list_chats(topic_contains='Room') and call
     sync_meeting_attendance(room='Room 1', ...) with the SAME meeting_label
   - or import_attendance_file for files downloaded from Teams, or record_room_assignment manually.
5. Report: meeting_room_report, group_attendance, student_history, query_tracking_db.
Attendance reports are only available for meetings the signed-in user organized.""",
)


def tool(annotations: ToolAnnotations | None = None):
    """Register a tool; expected failures become ToolErrors so the model sees the message
    (the SDK hides the message of any other exception)."""

    def decorate(fn):
        @functools.wraps(fn)
        async def wrapper(*args, **kwargs):
            try:
                return await fn(*args, **kwargs)
            except (ValueError, OSError, sqlite3.Error, httpx.HTTPError) as e:
                raise ToolError(f"{type(e).__name__}: {e}") from e

        return mcp.tool(annotations=annotations)(wrapper)

    return decorate


# ---------------------------------------------------------------- helpers

def _text(body: dict | None, limit: int = 1500) -> str:
    content = (body or {}).get("content") or ""
    if (body or {}).get("contentType") == "html":
        content = re.sub(r"<br\s*/?>|</p>|</div>", "\n", content, flags=re.I)
        content = html.unescape(re.sub(r"<[^>]+>", "", content))
    content = re.sub(r"\n{3,}", "\n\n", content).strip()
    return content if len(content) <= limit else content[:limit] + "…"


def _user_brief(u: dict) -> dict:
    return {
        "id": u.get("id"),
        "displayName": u.get("displayName"),
        "mail": u.get("mail"),
        "userPrincipalName": u.get("userPrincipalName"),
        "jobTitle": u.get("jobTitle"),
        "department": u.get("department"),
    }


def _message_brief(m: dict) -> dict:
    sender = (m.get("from") or {}).get("user") or (m.get("from") or {}).get("application") or {}
    out = {
        "id": m.get("id"),
        "created": m.get("createdDateTime"),
        "from": sender.get("displayName"),
        "type": m.get("messageType"),
        "subject": m.get("subject"),
        "text": _text(m.get("body")),
    }
    if m.get("replies"):
        out["replies"] = [_message_brief(r) for r in m["replies"]]
    return out


def _iso(value: str | None, default: datetime) -> str:
    if not value:
        return default.strftime("%Y-%m-%dT%H:%M:%SZ")
    if len(value) == 10:  # YYYY-MM-DD
        return value + "T00:00:00Z"
    return value


async def _resolve_meeting(meeting: str) -> dict:
    """meeting: a Teams join URL or an onlineMeeting id."""
    if meeting.lower().startswith("http"):
        url = meeting.replace("'", "''")
        data = await graph.get("/me/onlineMeetings", {"$filter": f"JoinWebUrl eq '{url}'"})
        items = data.get("value") or []
        if not items:
            raise ValueError(
                "No online meeting found for that join URL. Attendance is only available "
                "for meetings you organized (or that an admin app-access policy covers)."
            )
        return items[0]
    return await graph.get(f"/me/onlineMeetings/{meeting}")


async def _attendance(meeting_id: str, occurrence_date: str | None) -> list[tuple[dict, list[dict]]]:
    """Attendance reports (one per meeting occurrence) with their records.
    occurrence_date=YYYY-MM-DD picks that day's reports; otherwise the latest one."""
    reports = await graph.get_all(f"/me/onlineMeetings/{meeting_id}/attendanceReports")
    reports.sort(key=lambda r: r.get("meetingStartDateTime") or "")
    if occurrence_date:
        reports = [r for r in reports if (r.get("meetingStartDateTime") or "").startswith(occurrence_date)]
    else:
        reports = reports[-1:]
    out = []
    for rep in reports:
        records = await graph.get_all(
            f"/me/onlineMeetings/{meeting_id}/attendanceReports/{rep['id']}/attendanceRecords"
        )
        out.append((rep, records))
    return out


def _records_to_rows(reports: list[tuple[dict, list[dict]]], room: str) -> list[dict]:
    """Merge records (possibly across several reports) into one row per person."""
    merged: dict[str, dict] = {}
    for _, records in reports:
        for rec in records:
            ident = rec.get("identity") or {}
            email = rec.get("emailAddress")
            key = (email or ident.get("id") or ident.get("displayName") or "").lower()
            if not key:
                continue
            intervals = rec.get("attendanceIntervals") or []
            joins = [i["joinDateTime"] for i in intervals if i.get("joinDateTime")]
            leaves = [i["leaveDateTime"] for i in intervals if i.get("leaveDateTime")]
            row = merged.setdefault(key, {
                "room": room,
                "email": email,
                "user_id": ident.get("id"),
                "display_name": ident.get("displayName"),
                "first_join": None,
                "last_leave": None,
                "duration_seconds": 0,
                "role": rec.get("role"),
            })
            if joins:
                row["first_join"] = min(filter(None, [row["first_join"], *joins]))
            if leaves:
                row["last_leave"] = max(filter(None, [row["last_leave"], *leaves]))
            row["duration_seconds"] += rec.get("totalAttendanceInSeconds") or 0
    return list(merged.values())


# ---------------------------------------------------------------- auth

@tool(READ_ONLY)
async def auth_status() -> dict:
    """Check whether the server is signed in to Microsoft 365 (and finish a pending device-code sign-in)."""
    return await asyncio.to_thread(auth.status)


@tool()
async def auth_sign_in() -> dict:
    """Start Microsoft 365 sign-in. In device mode this returns a URL and a code: show both to
    the user, wait for them to sign in, then call auth_status. Tokens are cached locally."""
    status = await asyncio.to_thread(auth.status)
    if status["status"] == "signed_in":
        return status
    if settings.auth_mode == "interactive":
        return await asyncio.to_thread(auth.interactive_sign_in)
    return await asyncio.to_thread(auth.start_device_flow)


@tool(DESTRUCTIVE)
async def auth_sign_out() -> dict:
    """Sign out and delete the locally cached tokens."""
    await asyncio.to_thread(auth.sign_out)
    return {"status": "signed_out"}


# ---------------------------------------------------------------- directory (AD)

@tool(READ_ONLY)
async def whoami() -> dict:
    """Profile of the signed-in user."""
    return _user_brief(await graph.get("/me"))


@tool(READ_ONLY)
async def search_users(query: str, limit: int = 25) -> list[dict]:
    """Search the directory for users by name, email or UPN (prefix match)."""
    q = query.replace('"', "")
    params = {
        "$search": f'"displayName:{q}" OR "mail:{q}" OR "userPrincipalName:{q}"',
        "$select": "id,displayName,mail,userPrincipalName,jobTitle,department",
        "$count": "true",
        "$top": str(min(limit, 100)),
    }
    users = await graph.get_all("/users", params, headers=EVENTUAL, limit=limit)
    return [_user_brief(u) for u in users]


@tool(READ_ONLY)
async def get_user(user: str) -> dict:
    """Directory details for one user (object id or UPN/email), including manager."""
    select = (
        "id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,"
        "officeLocation,companyName,employeeId,accountEnabled,createdDateTime"
    )
    u = await graph.get(f"/users/{user}", {"$select": select})
    try:
        mgr = await graph.get(f"/users/{user}/manager", {"$select": "id,displayName,mail"})
        u["manager"] = {k: mgr.get(k) for k in ("id", "displayName", "mail")}
    except Exception:  # noqa: BLE001 - many students have no manager
        u["manager"] = None
    u.pop("@odata.context", None)
    return u


@tool(READ_ONLY)
async def get_user_groups(user: str) -> list[dict]:
    """Groups/teams a user (object id or UPN/email) is a direct member of."""
    items = await graph.get_all(
        f"/users/{user}/memberOf", {"$select": "id,displayName,mail,groupTypes,resourceProvisioningOptions"}
    )
    return [
        {
            "id": g["id"],
            "displayName": g.get("displayName"),
            "mail": g.get("mail"),
            "isTeam": "Team" in (g.get("resourceProvisioningOptions") or []),
        }
        for g in items
        if g.get("@odata.type") == "#microsoft.graph.group"
    ]


# ---------------------------------------------------------------- groups & teams

@tool(READ_ONLY)
async def list_groups(search: str | None = None, only_teams: bool = False, limit: int = 100) -> list[dict]:
    """List Microsoft 365 / security groups in the tenant, optionally filtered by display name
    (word-prefix search) and/or to Teams-enabled groups only."""
    params = {
        "$select": "id,displayName,mail,description,groupTypes,resourceProvisioningOptions,createdDateTime",
        "$count": "true",
        "$top": str(min(limit, 999)),
    }
    if search:
        params["$search"] = f'"displayName:{search.replace(chr(34), "")}"'
    if only_teams:
        params["$filter"] = "resourceProvisioningOptions/Any(x:x eq 'Team')"
    groups = await graph.get_all("/groups", params, headers=EVENTUAL, limit=limit)
    return [
        {
            "id": g["id"],
            "displayName": g.get("displayName"),
            "mail": g.get("mail"),
            "description": g.get("description"),
            "isTeam": "Team" in (g.get("resourceProvisioningOptions") or []),
            "created": g.get("createdDateTime"),
        }
        for g in groups
    ]


@tool(READ_ONLY)
async def get_group_members(group_id: str, limit: int = 2000) -> dict:
    """Users in a group/team (id, name, email/UPN, job title, department)."""
    members = await graph.get_all(
        f"/groups/{group_id}/members",
        {"$select": "id,displayName,mail,userPrincipalName,jobTitle,department", "$top": "999"},
        limit=limit,
    )
    users = [_user_brief(m) for m in members if m.get("@odata.type") == "#microsoft.graph.user"]
    return {"count": len(users), "members": users}


@tool(READ_ONLY)
async def list_my_teams() -> list[dict]:
    """Teams the signed-in user is a member of."""
    teams = await graph.get_all("/me/joinedTeams", {"$select": "id,displayName,description"})
    return [{"id": t["id"], "displayName": t.get("displayName"), "description": t.get("description")} for t in teams]


@tool(READ_ONLY)
async def list_team_channels(team_id: str) -> list[dict]:
    """Channels in a team."""
    chans = await graph.get_all(f"/teams/{team_id}/channels", {"$select": "id,displayName,description,membershipType"})
    return [{k: c.get(k) for k in ("id", "displayName", "description", "membershipType")} for c in chans]


@tool(READ_ONLY)
async def list_channel_messages(team_id: str, channel_id: str, limit: int = 30, include_replies: bool = False) -> list[dict]:
    """Recent messages in a team channel (newest first), optionally with replies."""
    params = {"$top": str(min(limit, 50))}
    if include_replies:
        params["$expand"] = "replies"
    msgs = await graph.get_all(f"/teams/{team_id}/channels/{channel_id}/messages", params, limit=limit)
    return [_message_brief(m) for m in msgs]


# ---------------------------------------------------------------- chats & meetings

@tool(READ_ONLY)
async def list_chats(
    chat_type: str | None = "meeting",
    topic_contains: str | None = None,
    created_after: str | None = None,
    scan_limit: int = 300,
) -> list[dict]:
    """List the user's chats. Use chat_type='meeting' to get meeting chats - this includes
    breakout-room chats (topics like 'Room 1'), whose joinWebUrl can be passed to
    sync_meeting_attendance. created_after: ISO date/time. chat_type: meeting|group|oneOnOne|None."""
    chats = await graph.get_all("/me/chats", {"$top": "50"}, limit=scan_limit)
    out = []
    for c in chats:
        if chat_type and c.get("chatType") != chat_type:
            continue
        if topic_contains and topic_contains.lower() not in (c.get("topic") or "").lower():
            continue
        if created_after and (c.get("createdDateTime") or "") < created_after:
            continue
        info = c.get("onlineMeetingInfo") or {}
        out.append({
            "id": c["id"],
            "topic": c.get("topic"),
            "chatType": c.get("chatType"),
            "created": c.get("createdDateTime"),
            "lastUpdated": c.get("lastUpdatedDateTime"),
            "joinWebUrl": info.get("joinWebUrl"),
            "calendarEventId": info.get("calendarEventId"),
        })
    if chat_type == "meeting":  # fill join URLs the list call didn't return
        for item in out[:40]:
            if not item["joinWebUrl"]:
                full = await graph.get(f"/chats/{item['id']}")
                item["joinWebUrl"] = (full.get("onlineMeetingInfo") or {}).get("joinWebUrl")
    out.sort(key=lambda x: x["created"] or "", reverse=True)
    return out


@tool(READ_ONLY)
async def list_chat_messages(chat_id: str, limit: int = 50) -> list[dict]:
    """Messages in a chat (e.g. a meeting or breakout-room chat), newest first."""
    msgs = await graph.get_all(f"/chats/{chat_id}/messages", {"$top": str(min(limit, 50))}, limit=limit)
    return [_message_brief(m) for m in msgs if m.get("messageType") == "message"]


@tool(READ_ONLY)
async def list_calendar_events(
    start: str | None = None,
    end: str | None = None,
    group_id: str | None = None,
    subject_contains: str | None = None,
    online_only: bool = True,
    limit: int = 200,
) -> list[dict]:
    """Calendar events (recurring meetings expanded) between start and end (ISO date or
    date-time, UTC; default: last 7 days to next 7 days). group_id reads a group calendar
    instead of yours. Returns the Teams joinUrl needed for attendance tools."""
    now = datetime.now(timezone.utc)
    params = {
        "startDateTime": _iso(start, now - timedelta(days=7)),
        "endDateTime": _iso(end, now + timedelta(days=7)),
        "$select": "id,subject,start,end,organizer,isOnlineMeeting,onlineMeeting,attendees,seriesMasterId,webLink",
        "$orderby": "start/dateTime",
        "$top": "100",
    }
    path = f"/groups/{group_id}/calendar/calendarView" if group_id else "/me/calendarView"
    events = await graph.get_all(path, params, headers={"Prefer": 'outlook.timezone="UTC"'}, limit=limit)
    out = []
    for e in events:
        if online_only and not e.get("isOnlineMeeting"):
            continue
        if subject_contains and subject_contains.lower() not in (e.get("subject") or "").lower():
            continue
        out.append({
            "id": e["id"],
            "subject": e.get("subject"),
            "start": (e.get("start") or {}).get("dateTime"),
            "end": (e.get("end") or {}).get("dateTime"),
            "organizer": ((e.get("organizer") or {}).get("emailAddress") or {}).get("address"),
            "attendeeCount": len(e.get("attendees") or []),
            "joinUrl": (e.get("onlineMeeting") or {}).get("joinUrl"),
            "seriesMasterId": e.get("seriesMasterId"),
        })
    return out


@tool(READ_ONLY)
async def find_online_meeting(meeting: str) -> dict:
    """Resolve a Teams join URL (or onlineMeeting id) to the online meeting: id, subject,
    times, chat thread id. Only works for meetings you organized."""
    m = await _resolve_meeting(meeting)
    return {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "start": m.get("startDateTime"),
        "end": m.get("endDateTime"),
        "joinWebUrl": m.get("joinWebUrl"),
        "chatThreadId": (m.get("chatInfo") or {}).get("threadId"),
        "organizer": (((m.get("participants") or {}).get("organizer") or {}).get("upn")),
    }


@tool(READ_ONLY)
async def get_meeting_attendance(meeting: str, occurrence_date: str | None = None, max_records: int = 500) -> dict:
    """Read the attendance report of a meeting you organized, without storing it.
    meeting: join URL or onlineMeeting id. occurrence_date: YYYY-MM-DD for a specific
    occurrence of a recurring meeting (default: latest report)."""
    m = await _resolve_meeting(meeting)
    reports = await _attendance(m["id"], occurrence_date)
    return {
        "meeting": {"id": m["id"], "subject": m.get("subject")},
        "reports": [
            {
                "id": rep["id"],
                "start": rep.get("meetingStartDateTime"),
                "end": rep.get("meetingEndDateTime"),
                "totalParticipants": rep.get("totalParticipantCount"),
                "records": [
                    {
                        "name": (r.get("identity") or {}).get("displayName"),
                        "email": r.get("emailAddress"),
                        "role": r.get("role"),
                        "seconds": r.get("totalAttendanceInSeconds"),
                        "intervals": [
                            {"join": i.get("joinDateTime"), "leave": i.get("leaveDateTime")}
                            for i in r.get("attendanceIntervals") or []
                        ],
                    }
                    for r in records[:max_records]
                ],
            }
            for rep, records in reports
        ],
    }


# ---------------------------------------------------------------- tracking (local DB)

@tool()
async def sync_group_roster(group_id: str, label: str | None = None) -> dict:
    """Store a group's members locally so attendance can be matched to year/subject groups.
    label: short name like 'Y2 Math' (defaults to the group's display name). Re-run to refresh."""
    group = await graph.get(f"/groups/{group_id}", {"$select": "id,displayName"})
    members = await graph.get_all(
        f"/groups/{group_id}/members", {"$select": "id,displayName,mail,userPrincipalName", "$top": "999"}
    )
    rows = [
        {"email": m.get("mail") or m.get("userPrincipalName"), "display_name": m.get("displayName"), "user_id": m["id"]}
        for m in members
        if m.get("@odata.type") == "#microsoft.graph.user"
    ]
    n = store.save_group(group_id, label or group.get("displayName") or group_id, group.get("displayName"), rows)
    return {"group_id": group_id, "label": label or group.get("displayName"), "members_saved": n}


@tool()
async def sync_meeting_attendance(
    meeting: str,
    meeting_label: str,
    room: str = "Main",
    occurrence_date: str | None = None,
    group_id: str | None = None,
    include_organizers: bool = False,
) -> dict:
    """Pull a meeting's attendance report from Graph and store it as presence in `room`.
    For breakout rooms call this once per room meeting (join URL from list_chats) using the
    same meeting_label as the main meeting. group_id links the meeting to its class group.
    Re-running overwrites the same meeting/room/person rows."""
    m = await _resolve_meeting(meeting)
    reports = await _attendance(m["id"], occurrence_date)
    if not reports:
        return {"saved": 0, "note": "No attendance report found (meeting not started yet, or wrong occurrence_date)."}
    rows = _records_to_rows(reports, room)
    if not include_organizers:
        # students are often presenters in breakout rooms, so only drop organizers
        rows = [r for r in rows if (r.get("role") or "").lower() not in ("organizer", "coorganizer")]
    first = reports[0][0]
    if room == "Main":
        store.save_meeting(meeting_label, m.get("subject"), m.get("joinWebUrl"),
                           first.get("meetingStartDateTime"), reports[-1][0].get("meetingEndDateTime"), group_id)
    else:
        store.save_meeting(meeting_label, group_id=group_id)
    n = store.save_presence(meeting_label, rows, "graph")
    return {"meeting_label": meeting_label, "room": room, "reports_used": len(reports), "saved": n}


@tool()
async def import_attendance_file(
    path: str, meeting_label: str, room: str = "Main", group_id: str | None = None
) -> dict:
    """Import an attendance file downloaded from Teams (or any CSV/TSV with Name + Email
    columns). If the file has a 'Room'/'Breakout room' column it is used per row, otherwise
    every row goes into `room`. Use the same meeting_label as the rest of the meeting."""
    rows = await asyncio.to_thread(parse_attendance_file, path, room)
    store.save_meeting(meeting_label, group_id=group_id)
    n = store.save_presence(meeting_label, rows, "file")
    rooms = sorted({r["room"] for r in rows})
    return {"meeting_label": meeting_label, "saved": n, "rooms": rooms}


@tool()
async def record_room_assignment(meeting_label: str, room: str, people: list[str], group_id: str | None = None) -> dict:
    """Manually record that people (emails/UPNs) were in a room of a meeting, e.g. from the
    breakout-room assignment screen when no attendance report is available."""
    rows = [{"room": room, "email": p.strip()} for p in people if p.strip()]
    store.save_meeting(meeting_label, group_id=group_id)
    return {"meeting_label": meeting_label, "room": room, "saved": store.save_presence(meeting_label, rows, "manual")}


@tool(READ_ONLY)
async def tracking_overview() -> dict:
    """Groups (with member counts) and meetings (with room/person counts) stored locally."""
    return store.overview()


@tool(READ_ONLY)
async def meeting_room_report(meeting_label: str) -> dict:
    """Who was in which room of a meeting, with each person's groups (year/subject)."""
    return store.meeting_report(meeting_label)


@tool(READ_ONLY)
async def group_attendance(meeting_label: str, group_id: str | None = None) -> dict:
    """For a meeting: which members of a group attended (and in which rooms), who was
    absent, and who attended without being in the group. group_id defaults to the meeting's linked group."""
    return store.group_coverage(meeting_label, group_id)


@tool(READ_ONLY)
async def student_history(query: str) -> list[dict]:
    """A student's groups and every meeting/room they were recorded in. query: part of email or name."""
    return store.person_history(query)


@tool(READ_ONLY)
async def query_tracking_db(sql: str) -> list[dict]:
    """Run a read-only SQL SELECT on the local tracking DB (SQLite). Tables:
    groups(group_id,label,display_name,synced_at), people(person_key,email,display_name,user_id),
    group_members(group_id,person_key), meetings(meeting_label,subject,join_url,start_time,end_time,group_id),
    room_presence(meeting_label,room,person_key,display_name,first_join,last_leave,duration_seconds,source,recorded_at)."""
    return store.readonly_query(sql)


@tool(DESTRUCTIVE)
async def delete_tracked_meeting(meeting_label: str, room: str | None = None) -> dict:
    """Delete stored presence for a meeting (or only one room of it)."""
    return {"deleted_rows": store.delete_meeting(meeting_label, room)}


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
