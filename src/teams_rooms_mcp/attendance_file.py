"""Parse attendance files downloaded from Teams (Meeting > Attendance > Download).

Teams exports are usually UTF-16 tab-separated with sections like
"1. Summary", "2. Participants", "3. In-Meeting Activities". Plain CSV/TSV
files with similar headers (optionally with a Room column) also work.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path

NAME_COLS = {"name", "full name", "participant name", "display name"}
EMAIL_COLS = {"email", "email address", "e-mail"}
UPN_COLS = {"participant id (upn)", "upn", "user principal name"}
JOIN_COLS = {"first join", "first join time", "join time"}
LEAVE_COLS = {"last leave", "last leave time", "leave time"}
DURATION_COLS = {"in-meeting duration", "duration", "attendance duration"}
ROOM_COLS = {"room", "breakout room", "room name"}

SECTION_RE = re.compile(r"^\s*\d+\.\s+\S")


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    if raw[:400].count(b"\x00") > 50:
        return raw.decode("utf-16-le")
    return raw.decode("utf-8-sig", errors="replace")


def parse_duration(value: str | None) -> int | None:
    if not value or not value.strip():
        return None
    v = value.strip().lower()
    if re.fullmatch(r"\d+:\d{1,2}:\d{1,2}", v):
        h, m, s = (int(x) for x in v.split(":"))
        return h * 3600 + m * 60 + s
    total, found = 0, False
    for amount, unit in re.findall(r"(\d+)\s*(h|hr|hour|m|min|minute|s|sec|second)", v):
        found = True
        total += int(amount) * {"h": 3600, "m": 60, "s": 1}[unit[0]]
    return total if found else None


def _pick(row: dict[str, str], names: set[str]) -> str | None:
    for k, v in row.items():
        if k in names and v and v.strip():
            return v.strip()
    return None


def parse_attendance_file(path: str | Path, default_room: str) -> list[dict]:
    path = Path(path).expanduser()
    lines = _read_text(path).splitlines()
    sample = "\n".join(lines[:50])
    delim = "\t" if sample.count("\t") >= sample.count(",") else ","

    header: list[str] | None = None
    start = 0
    for i, line in enumerate(lines):
        cells = [c.strip().lower() for c in next(csv.reader([line], delimiter=delim), [])]
        if NAME_COLS & set(cells) and (EMAIL_COLS | UPN_COLS) & set(cells):
            header, start = cells, i + 1
            break
    if header is None:
        raise ValueError(
            "Couldn't find a participants header (needs a Name column and an Email or "
            "'Participant ID (UPN)' column)."
        )

    people: list[dict] = []
    for cells in csv.reader(lines[start:], delimiter=delim):
        if not cells or all(not c.strip() for c in cells) or SECTION_RE.match(cells[0]):
            break  # end of the participants section
        row = dict(zip(header, cells))
        email = _pick(row, EMAIL_COLS) or _pick(row, UPN_COLS)
        name = _pick(row, NAME_COLS)
        if not (email or name):
            continue
        people.append({
            "room": _pick(row, ROOM_COLS) or default_room,
            "email": email,
            "display_name": name,
            "first_join": _pick(row, JOIN_COLS),
            "last_leave": _pick(row, LEAVE_COLS),
            "duration_seconds": parse_duration(_pick(row, DURATION_COLS)),
        })
    return people
