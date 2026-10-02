"""Local SQLite store for room tracking: who was in which room of which meeting,
and which group (year/subject) each person belongs to."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

SCHEMA = """
CREATE TABLE IF NOT EXISTS groups (
    group_id     TEXT PRIMARY KEY,
    label        TEXT NOT NULL,
    display_name TEXT,
    synced_at    TEXT
);
CREATE TABLE IF NOT EXISTS people (
    person_key   TEXT PRIMARY KEY,   -- lower-case email/UPN, or 'aad:<object id>'
    email        TEXT,
    display_name TEXT,
    user_id      TEXT
);
CREATE INDEX IF NOT EXISTS ix_people_user ON people(user_id);
CREATE TABLE IF NOT EXISTS group_members (
    group_id   TEXT NOT NULL REFERENCES groups(group_id) ON DELETE CASCADE,
    person_key TEXT NOT NULL,
    PRIMARY KEY (group_id, person_key)
);
CREATE TABLE IF NOT EXISTS meetings (
    meeting_label TEXT PRIMARY KEY,  -- user-chosen, e.g. 'Math Y2 2026-10-01'
    subject       TEXT,
    join_url      TEXT,
    start_time    TEXT,
    end_time      TEXT,
    group_id      TEXT               -- the class/group this meeting is for (optional)
);
CREATE TABLE IF NOT EXISTS room_presence (
    meeting_label    TEXT NOT NULL REFERENCES meetings(meeting_label) ON DELETE CASCADE,
    room             TEXT NOT NULL,
    person_key       TEXT NOT NULL,
    display_name     TEXT,
    first_join       TEXT,
    last_leave       TEXT,
    duration_seconds INTEGER,
    source           TEXT NOT NULL,  -- 'graph' | 'file' | 'manual'
    recorded_at      TEXT NOT NULL,
    PRIMARY KEY (meeting_label, room, person_key)
);
"""


def make_key(email: str | None = None, user_id: str | None = None, name: str | None = None) -> str:
    if email:
        return email.strip().lower()
    if user_id:
        return f"aad:{user_id}"
    if name:
        return f"name:{name.strip().lower()}"
    raise ValueError("person needs an email, user id or name")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Store:
    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        with self._conn() as c:
            c.executescript(SCHEMA)

    @contextmanager
    def _conn(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    # ---- writes -----------------------------------------------------------

    def resolve_key(self, conn: sqlite3.Connection, email: str | None, user_id: str | None, name: str | None) -> str:
        """Prefer an existing person matched by AAD object id, so attendance (which may report
        mail) and rosters (which may report UPN) line up on the same person."""
        if user_id:
            row = conn.execute("SELECT person_key FROM people WHERE user_id = ?", (user_id,)).fetchone()
            if row:
                return row["person_key"]
        return make_key(email, user_id, name)

    def _upsert_person(self, conn, key: str, email: str | None, name: str | None, user_id: str | None) -> None:
        conn.execute(
            """INSERT INTO people (person_key, email, display_name, user_id) VALUES (?, ?, ?, ?)
               ON CONFLICT(person_key) DO UPDATE SET
                 email = COALESCE(excluded.email, people.email),
                 display_name = COALESCE(excluded.display_name, people.display_name),
                 user_id = COALESCE(excluded.user_id, people.user_id)""",
            (key, email.lower() if email else None, name, user_id),
        )

    def save_group(self, group_id: str, label: str, display_name: str | None, members: list[dict]) -> int:
        """Replace the stored roster of a group. members: dicts with email, display_name, user_id."""
        with self._conn() as c:
            c.execute(
                """INSERT INTO groups (group_id, label, display_name, synced_at) VALUES (?, ?, ?, ?)
                   ON CONFLICT(group_id) DO UPDATE SET label = excluded.label,
                     display_name = excluded.display_name, synced_at = excluded.synced_at""",
                (group_id, label, display_name, _now()),
            )
            c.execute("DELETE FROM group_members WHERE group_id = ?", (group_id,))
            for m in members:
                key = self.resolve_key(c, m.get("email"), m.get("user_id"), m.get("display_name"))
                self._upsert_person(c, key, m.get("email"), m.get("display_name"), m.get("user_id"))
                c.execute("INSERT OR IGNORE INTO group_members VALUES (?, ?)", (group_id, key))
        return len(members)

    def save_meeting(self, label: str, subject=None, join_url=None, start=None, end=None, group_id=None) -> None:
        with self._conn() as c:
            c.execute(
                """INSERT INTO meetings (meeting_label, subject, join_url, start_time, end_time, group_id)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT(meeting_label) DO UPDATE SET
                     subject = COALESCE(excluded.subject, meetings.subject),
                     join_url = COALESCE(excluded.join_url, meetings.join_url),
                     start_time = COALESCE(excluded.start_time, meetings.start_time),
                     end_time = COALESCE(excluded.end_time, meetings.end_time),
                     group_id = COALESCE(excluded.group_id, meetings.group_id)""",
                (label, subject, join_url, start, end, group_id),
            )

    def save_presence(self, meeting_label: str, rows: list[dict], source: str) -> int:
        """rows: dicts with room, email, user_id, display_name, first_join, last_leave, duration_seconds.
        Re-saving the same (meeting, room, person) overwrites, so syncs are idempotent."""
        with self._conn() as c:
            for r in rows:
                key = self.resolve_key(c, r.get("email"), r.get("user_id"), r.get("display_name"))
                self._upsert_person(c, key, r.get("email"), r.get("display_name"), r.get("user_id"))
                c.execute(
                    """INSERT INTO room_presence (meeting_label, room, person_key, display_name, first_join,
                         last_leave, duration_seconds, source, recorded_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                       ON CONFLICT(meeting_label, room, person_key) DO UPDATE SET
                         display_name = excluded.display_name, first_join = excluded.first_join,
                         last_leave = excluded.last_leave, duration_seconds = excluded.duration_seconds,
                         source = excluded.source, recorded_at = excluded.recorded_at""",
                    (meeting_label, r["room"], key, r.get("display_name"), r.get("first_join"),
                     r.get("last_leave"), r.get("duration_seconds"), source, _now()),
                )
        return len(rows)

    def delete_meeting(self, meeting_label: str, room: str | None = None) -> int:
        with self._conn() as c:
            if room:
                cur = c.execute("DELETE FROM room_presence WHERE meeting_label = ? AND room = ?", (meeting_label, room))
                return cur.rowcount
            cur = c.execute("DELETE FROM room_presence WHERE meeting_label = ?", (meeting_label,))
            c.execute("DELETE FROM meetings WHERE meeting_label = ?", (meeting_label,))
            return cur.rowcount

    # ---- reads ------------------------------------------------------------

    def overview(self) -> dict:
        with self._conn() as c:
            groups = c.execute(
                """SELECT g.group_id, g.label, g.display_name, g.synced_at, COUNT(m.person_key) AS members
                   FROM groups g LEFT JOIN group_members m USING (group_id) GROUP BY g.group_id ORDER BY g.label"""
            ).fetchall()
            meetings = c.execute(
                """SELECT mt.meeting_label, mt.subject, mt.start_time, mt.group_id,
                          COUNT(DISTINCT p.room) AS rooms, COUNT(DISTINCT p.person_key) AS people
                   FROM meetings mt LEFT JOIN room_presence p USING (meeting_label)
                   GROUP BY mt.meeting_label ORDER BY mt.start_time DESC, mt.meeting_label"""
            ).fetchall()
        return {"groups": [dict(r) for r in groups], "meetings": [dict(r) for r in meetings]}

    def meeting_report(self, meeting_label: str) -> dict:
        with self._conn() as c:
            meeting = c.execute("SELECT * FROM meetings WHERE meeting_label = ?", (meeting_label,)).fetchone()
            if not meeting:
                raise ValueError(f"Unknown meeting_label '{meeting_label}'")
            rows = c.execute(
                """SELECT p.room, p.person_key, COALESCE(pe.display_name, p.display_name) AS name,
                          p.first_join, p.last_leave, p.duration_seconds, p.source,
                          (SELECT GROUP_CONCAT(g.label, ', ') FROM group_members gm
                             JOIN groups g USING (group_id) WHERE gm.person_key = p.person_key) AS groups
                   FROM room_presence p LEFT JOIN people pe USING (person_key)
                   WHERE p.meeting_label = ? ORDER BY p.room, name""",
                (meeting_label,),
            ).fetchall()
        rooms: dict[str, list[dict]] = {}
        for r in rows:
            d = dict(r)
            rooms.setdefault(d.pop("room"), []).append(d)
        return {
            "meeting": dict(meeting),
            "rooms": [{"room": k, "count": len(v), "people": v} for k, v in rooms.items()],
        }

    def person_history(self, query: str) -> list[dict]:
        like = f"%{query.lower()}%"
        with self._conn() as c:
            people = c.execute(
                """SELECT person_key, email, display_name, user_id FROM people
                   WHERE person_key LIKE ? OR LOWER(COALESCE(display_name, '')) LIKE ? LIMIT 20""",
                (like, like),
            ).fetchall()
            out = []
            for pe in people:
                groups = c.execute(
                    "SELECT g.label FROM group_members gm JOIN groups g USING (group_id) WHERE gm.person_key = ?",
                    (pe["person_key"],),
                ).fetchall()
                presence = c.execute(
                    """SELECT p.meeting_label, mt.start_time, p.room, p.first_join, p.last_leave,
                              p.duration_seconds, p.source
                       FROM room_presence p JOIN meetings mt USING (meeting_label)
                       WHERE p.person_key = ? ORDER BY mt.start_time, p.meeting_label, p.first_join""",
                    (pe["person_key"],),
                ).fetchall()
                out.append({**dict(pe), "groups": [g["label"] for g in groups], "presence": [dict(p) for p in presence]})
        return out

    def group_coverage(self, meeting_label: str, group_id: str | None) -> dict:
        """For one meeting: each member of the group -> rooms they were in (empty = absent)."""
        with self._conn() as c:
            if not group_id:
                row = c.execute("SELECT group_id FROM meetings WHERE meeting_label = ?", (meeting_label,)).fetchone()
                group_id = row["group_id"] if row else None
            if not group_id:
                raise ValueError("No group_id given and the meeting has no group linked.")
            rows = c.execute(
                """SELECT gm.person_key, pe.display_name,
                          GROUP_CONCAT(p.room, ', ') AS rooms, SUM(p.duration_seconds) AS total_seconds
                   FROM group_members gm
                   LEFT JOIN people pe USING (person_key)
                   LEFT JOIN room_presence p ON p.person_key = gm.person_key AND p.meeting_label = ?
                   WHERE gm.group_id = ?
                   GROUP BY gm.person_key ORDER BY pe.display_name""",
                (meeting_label, group_id),
            ).fetchall()
            outsiders = c.execute(
                """SELECT DISTINCT p.person_key, p.display_name, p.room FROM room_presence p
                   WHERE p.meeting_label = ? AND p.person_key NOT IN
                     (SELECT person_key FROM group_members WHERE group_id = ?)""",
                (meeting_label, group_id),
            ).fetchall()
        present = [dict(r) for r in rows if r["rooms"]]
        absent = [{"person_key": r["person_key"], "display_name": r["display_name"]} for r in rows if not r["rooms"]]
        return {
            "group_id": group_id,
            "present_count": len(present),
            "absent_count": len(absent),
            "present": present,
            "absent": absent,
            "attendees_not_in_group": [dict(r) for r in outsiders],
        }

    def readonly_query(self, sql: str, max_rows: int = 500) -> list[dict]:
        if not sql.lstrip().lower().startswith(("select", "with")):
            raise ValueError("Only SELECT / WITH queries are allowed.")
        conn = sqlite3.connect(f"file:{self.path.as_posix()}?mode=ro", uri=True)
        conn.row_factory = sqlite3.Row
        try:
            return [dict(r) for r in conn.execute(sql).fetchmany(max_rows)]
        finally:
            conn.close()
