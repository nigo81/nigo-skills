#!/usr/bin/env python3
"""Export a complete OpenCode session from the local SQLite database.

Usage:
    python3 export_session.py <session_id> [output.json]
    python3 export_session.py --list              # List recent sessions
    python3 export_session.py --list -n 20        # List N recent sessions

The script reads directly from opencode.db, bypassing compaction checkpoints,
so it returns the FULL conversation history including pre-compaction messages.
"""

import argparse
import json
import os
import sqlite3
import sys
import tempfile
from datetime import datetime


def find_db_path():
    """Locate opencode.db across platforms (Linux / macOS / Windows)."""
    home = os.path.expanduser("~")
    candidates = [
        # Linux/macOS XDG
        os.path.join(home, ".local", "share", "opencode", "opencode.db"),
        os.path.join(home, ".config", "opencode", "opencode.db"),
        # macOS legacy
        os.path.join(home, "Library", "Application Support", "opencode", "opencode.db"),
    ]
    # Windows: %LOCALAPPDATA% / %APPDATA%
    localappdata = os.environ.get("LOCALAPPDATA")
    appdata = os.environ.get("APPDATA")
    if localappdata:
        candidates.append(os.path.join(localappdata, "opencode", "opencode.db"))
    if appdata:
        candidates.append(os.path.join(appdata, "opencode", "opencode.db"))

    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def get_db():
    db_path = find_db_path()
    if not db_path:
        print("Error: opencode.db not found. Searched:\n  "
              + "\n  ".join([
                  "~/.local/share/opencode/opencode.db",
                  "~/.config/opencode/opencode.db",
                  "~/Library/Application Support/opencode/opencode.db",
                  "%LOCALAPPDATA%\\opencode\\opencode.db (Windows)",
                  "%APPDATA%\\opencode\\opencode.db (Windows)",
              ]), file=sys.stderr)
        sys.exit(1)
    return sqlite3.connect(db_path)


def list_sessions(db, limit=10):
    """List recent sessions."""
    cursor = db.execute("""
        SELECT id, title, 
               datetime(time_created/1000, 'unixepoch', 'localtime') as created,
               datetime(COALESCE(time_updated, time_created)/1000, 'unixepoch', 'localtime') as updated
        FROM session 
        ORDER BY time_created DESC 
        LIMIT ?
    """, (limit,))
    sessions = []
    for row in cursor:
        sessions.append({
            "id": row[0],
            "title": row[1] or "(untitled)",
            "created": row[2],
            "updated": row[3],
        })
    return sessions


def export_session(db, session_id):
    """Export full session data including pre-compaction messages."""
    # 1. Get session metadata
    cursor = db.execute("""
        SELECT id, title, time_created, time_updated, time_compacting, share_url
        FROM session WHERE id = ?
    """, (session_id,))
    row = cursor.fetchone()
    if not row:
        # Try partial match
        cursor = db.execute("""
            SELECT id, title, time_created, time_updated, time_compacting, share_url
            FROM session WHERE id LIKE ?
        """, (f"%{session_id}%",))
        row = cursor.fetchone()
        if not row:
            print(f"Error: Session '{session_id}' not found", file=sys.stderr)
            sys.exit(1)
        session_id = row[0]
        print(f"Found session: {session_id}", file=sys.stderr)

    session_data = {
        "id": row[0],
        "title": row[1] or "(untitled)",
        "time_created": row[2],
        "time_updated": row[3],
        "time_compacting": row[4],
        "share_url": row[5],
    }

    # 2. Get all messages
    cursor = db.execute("""
        SELECT id, data, time_created
        FROM message 
        WHERE session_id = ?
        ORDER BY time_created, id
    """, (session_id,))
    message_rows = cursor.fetchall()

    messages = []
    for msg_row in message_rows:
        msg_id, msg_data_raw, msg_time = msg_row
        msg_data = json.loads(msg_data_raw) if msg_data_raw else {}
        
        role = msg_data.get("role", "unknown")
        agent = msg_data.get("agent", "")
        model = msg_data.get("model", "")
        summary = msg_data.get("summary", "")

        # 3. Get parts for this message
        part_cursor = db.execute("""
            SELECT data
            FROM part 
            WHERE message_id = ?
            ORDER BY id
        """, (msg_id,))
        parts = []
        for (part_data_raw,) in part_cursor.fetchall():
            part_data = json.loads(part_data_raw) if part_data_raw else {}
            part_type = part_data.get("type", "text")
            parts.append({
                "type": part_type,
                "data": part_data,
            })

        # Convert timestamp
        ts = datetime.fromtimestamp(msg_time / 1000).strftime("%Y-%m-%d %H:%M:%S") if msg_time else ""

        messages.append({
            "id": msg_id,
            "role": role,
            "timestamp": ts,
            "agent": agent,
            "model": model,
            "summary": summary,
            "parts": parts,
        })

    # 4. Compute stats
    user_msgs = [m for m in messages if m["role"] == "user"]
    asst_msgs = [m for m in messages if m["role"] == "assistant"]
    
    session_data["stats"] = {
        "total_messages": len(messages),
        "user_messages": len(user_msgs),
        "assistant_messages": len(asst_msgs),
        "start_time": messages[0]["timestamp"] if messages else "",
        "end_time": messages[-1]["timestamp"] if messages else "",
    }

    return {
        "session": session_data,
        "messages": messages,
    }


def main():
    parser = argparse.ArgumentParser(description="Export OpenCode session from local database")
    parser.add_argument("session_id", nargs="?", help="Session ID to export")
    parser.add_argument("output", nargs="?", help="Output JSON file path")
    parser.add_argument("--list", "-l", action="store_true", help="List recent sessions")
    parser.add_argument("-n", "--limit", type=int, default=10, help="Number of sessions to list")
    args = parser.parse_args()

    db = get_db()

    if args.list:
        sessions = list_sessions(db, args.limit)
        print(f"{'ID':<30} {'Title':<40} {'Created':<20}")
        print("-" * 90)
        for s in sessions:
            print(f"{s['id']:<30} {s['title'][:40]:<40} {s['created']:<20}")
        return

    if not args.session_id:
        parser.print_help()
        return

    data = export_session(db, args.session_id)

    output = args.output or os.path.join(tempfile.gettempdir(), f"opencode_session_{data['session']['id'][:20]}.json")
    with open(output, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Session: {data['session']['title']}", file=sys.stderr)
    print(f"Messages: {data['session']['stats']['total_messages']} "
          f"({data['session']['stats']['user_messages']} user, "
          f"{data['session']['stats']['assistant_messages']} assistant)", file=sys.stderr)
    print(f"Output: {output}", file=sys.stderr)
    db.close()


if __name__ == "__main__":
    main()
