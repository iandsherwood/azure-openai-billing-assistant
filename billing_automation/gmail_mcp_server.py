#!/usr/bin/env python3
"""
Custom Gmail MCP Server (configurable account)
==================================================
Exposes Gmail tools to Cowork/Claude Code via the MCP protocol.
Uses the same credentials.json / gmail_token.pickle as run_billing.py.

To add to Claude: add the following to ~/.claude/settings.json
(see add_personal_gmail_mcp.json in this folder for the exact config)

Dependencies:
    pip3 install mcp google-api-python-client google-auth-oauthlib
"""

import base64
import json
import pickle
import re
import sys
from pathlib import Path
from datetime import datetime

# ── MCP framework ─────────────────────────────────────────────────────────────
try:
    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    from mcp.types import Tool, TextContent
    import mcp.types as types
except ImportError:
    sys.exit("Run: pip3 install mcp")

# ── Google API ────────────────────────────────────────────────────────────────
try:
    from googleapiclient.discovery import build
    from google.auth.transport.requests import Request
except ImportError:
    sys.exit("Run: pip3 install google-api-python-client")

# ── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent.resolve()
TOKEN_FILE  = SCRIPT_DIR / "gmail_token.pickle"
ACCOUNT     = "your-email@gmail.com"

# ── Gmail client ──────────────────────────────────────────────────────────────
def get_gmail_service():
    if not TOKEN_FILE.exists():
        raise RuntimeError(
            f"Gmail not authenticated. Run: python3 run_billing.py --auth"
        )
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return build("gmail", "v1", credentials=creds)


def decode_body(payload: dict) -> str:
    mime  = payload.get("mimeType", "")
    data  = payload.get("body", {}).get("data", "")
    parts = payload.get("parts", [])

    if data and "text" in mime:
        return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
    for part in parts:
        result = decode_body(part)
        if result:
            return result
    return ""


# ── MCP server ────────────────────────────────────────────────────────────────
server = Server("gmail-personal")


@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="gmail_personal_search",
            description=(
                f"Search emails in {ACCOUNT} using Gmail search syntax. "
                "Returns a list of matching messages with subject, sender, date, and snippet."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Gmail search query, e.g. 'from:wm.com newer_than:30d'"
                    },
                    "max_results": {
                        "type": "integer",
                        "description": "Max messages to return (default 10)",
                        "default": 10
                    }
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="gmail_personal_read",
            description=(
                f"Read the full body of a specific email in {ACCOUNT} by message ID."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {
                        "type": "string",
                        "description": "Gmail message ID (from gmail_personal_search results)"
                    }
                },
                "required": ["message_id"]
            }
        ),
        Tool(
            name="gmail_personal_profile",
            description=f"Get the profile info for the connected account ({ACCOUNT}).",
            inputSchema={"type": "object", "properties": {}}
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    try:
        service = get_gmail_service()
    except RuntimeError as e:
        return [TextContent(type="text", text=f"Error: {e}")]

    # ── gmail_personal_search ─────────────────────────────────────────────────
    if name == "gmail_personal_search":
        query       = arguments.get("query", "")
        max_results = arguments.get("max_results", 10)

        result = service.users().messages().list(
            userId="me", q=query, maxResults=max_results
        ).execute()
        messages = result.get("messages", [])

        if not messages:
            return [TextContent(type="text", text="No messages found.")]

        output = []
        for msg_stub in messages:
            msg = service.users().messages().get(
                userId="me", id=msg_stub["id"], format="metadata",
                metadataHeaders=["Subject", "From", "Date"]
            ).execute()
            headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
            snippet = msg.get("snippet", "")
            output.append({
                "id":      msg_stub["id"],
                "subject": headers.get("Subject", "(no subject)"),
                "from":    headers.get("From", ""),
                "date":    headers.get("Date", ""),
                "snippet": snippet[:200]
            })

        return [TextContent(type="text", text=json.dumps(output, indent=2))]

    # ── gmail_personal_read ───────────────────────────────────────────────────
    elif name == "gmail_personal_read":
        msg_id = arguments.get("message_id", "")
        msg = service.users().messages().get(
            userId="me", id=msg_id, format="full"
        ).execute()
        headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
        body    = decode_body(msg.get("payload", {}))

        output = {
            "id":      msg_id,
            "subject": headers.get("Subject", "(no subject)"),
            "from":    headers.get("From", ""),
            "date":    headers.get("Date", ""),
            "body":    body[:4000]  # cap at 4k chars
        }
        return [TextContent(type="text", text=json.dumps(output, indent=2))]

    # ── gmail_personal_profile ────────────────────────────────────────────────
    elif name == "gmail_personal_profile":
        profile = service.users().getProfile(userId="me").execute()
        return [TextContent(type="text", text=json.dumps(profile, indent=2))]

    return [TextContent(type="text", text=f"Unknown tool: {name}")]


# ── Entry point ───────────────────────────────────────────────────────────────
async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
