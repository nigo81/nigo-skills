#!/usr/bin/env python3
"""Generate a standalone HTML file from session JSON data and a template.

Usage:
    python3 generate_html.py <session.json> [output.html]
    python3 generate_html.py --session <session_id> [output.html]

Reads session JSON (from export_session.py) and injects it into template.html,
producing a standalone HTML file viewable in any browser.
"""

import json
import os
import re
import subprocess
import sys
import tempfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "..", "assets", "template.html")
EXPORT_SCRIPT = os.path.join(SCRIPT_DIR, "export_session.py")
PLACEHOLDER = "__SESSION_DATA__"


def sanitize_filename(title):
    """清理标题为合法文件名。"""
    safe = re.sub(r'[/\\:*?"<>|]', '_', title).strip('. ')
    if len(safe) > 80:
        safe = safe[:80]
    return safe or 'session'


def generate(json_path, output_path=None):
    """Inject session JSON into template, producing standalone HTML."""
    # Read JSON data
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Read template
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        template = f.read()

    if PLACEHOLDER not in template:
        print(f"Error: Placeholder '{PLACEHOLDER}' not found in template", file=sys.stderr)
        sys.exit(1)

    # Inject JSON data into template
    # Escape </ to prevent </script> in JSON data from breaking the HTML script tag
    json_str = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
    html = template.replace(PLACEHOLDER, json_str)

    # Replace page title
    session_title = data.get("session", {}).get("title", "session")
    html = html.replace("__PAGE_TITLE__", session_title + " · OpenCode")

    # Determine output path: default to {title}.html in current directory
    if not output_path:
        title = data.get("session", {}).get("title", "session")
        output_path = sanitize_filename(title) + ".html"

    # Write output
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    title = data.get("session", {}).get("title", "(untitled)")
    msg_count = data.get("session", {}).get("stats", {}).get("total_messages", "?")

    print(f"Generated: {output_path}", file=sys.stderr)
    print(f"Session: {title}", file=sys.stderr)
    print(f"Messages: {msg_count}", file=sys.stderr)
    print(f"File size: {os.path.getsize(output_path):,} bytes", file=sys.stderr)

    return output_path


def export_and_generate(session_id, output_path=None):
    """Export session from DB then generate HTML in one step."""
    temp_json = os.path.join(tempfile.gettempdir(), f"opencode_gen_{session_id[:20]}.json")

    result = subprocess.run(
        [sys.executable, EXPORT_SCRIPT, session_id, temp_json],
        capture_output=True, text=True
    )

    if result.returncode != 0 or not os.path.exists(temp_json):
        print(f"Error exporting session: {result.stderr}", file=sys.stderr)
        sys.exit(1)

    generate(temp_json, output_path)
    os.unlink(temp_json)


def main():
    if len(sys.argv) < 2:
        print("Usage:", file=sys.stderr)
        print("  python3 generate_html.py <session.json> [output.html]", file=sys.stderr)
        print("  python3 generate_html.py --session <session_id> [output.html]", file=sys.stderr)
        sys.exit(1)

    args = sys.argv[1:]

    if args[0] == "--session":
        if len(args) < 2:
            print("Error: --session requires a session ID", file=sys.stderr)
            sys.exit(1)
        session_id = args[1]
        output_path = args[2] if len(args) > 2 else None
        export_and_generate(session_id, output_path)
    else:
        json_path = args[0]
        output_path = args[1] if len(args) > 1 else None
        generate(json_path, output_path)


if __name__ == "__main__":
    main()
