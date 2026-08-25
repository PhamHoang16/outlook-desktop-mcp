"""
Outlook Desktop MCP - Phase 3 MCP Server Test
==============================================
Uses the MCP SDK client to connect to the server over stdio
and exercise all tools.
"""
import json
import asyncio
import logging
import os
import sys

_REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(_REPO_ROOT, "src"))

logging.basicConfig(level=logging.WARNING, stream=sys.stderr)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


async def run_tests():
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    server_params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "outlook_desktop_mcp.server"],
        cwd=_REPO_ROOT,
        env={**os.environ, "PYTHONPATH": os.path.join(_REPO_ROOT, "src")},
    )

    log("=" * 60)
    log("Outlook Desktop MCP - Phase 3 Server Test")
    log("=" * 60)
    log("\nConnecting to server...")

    passed = 0
    total = 0

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            log("Server initialized.\n")

            # ----- Test 1: Tool Discovery -----
            total += 1
            log("--- Test 1: Tool Discovery ---")
            try:
                tools_result = await session.list_tools()
                tool_names = [t.name for t in tools_result.tools]
                log(f"  Found {len(tool_names)} tools:")
                for t in tools_result.tools:
                    desc = t.description.split("\n")[0][:75] if t.description else ""
                    log(f"    - {t.name}: {desc}")

                expected = [
                    "send_email", "list_emails", "read_email", "mark_as_read",
                    "mark_as_unread", "move_email", "reply_email",
                    "list_folders", "search_emails",
                ]
                missing = [n for n in expected if n not in tool_names]
                assert not missing, f"Missing tools: {missing}"
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 2: list_folders -----
            total += 1
            log("\n--- Test 2: list_folders ---")
            try:
                result = await session.call_tool("list_folders", {"max_depth": 1})
                content = result.content[0].text
                folders = json.loads(content)
                folder_names = [f["name"] for f in folders]
                log(f"  Top-level folders: {folder_names}")
                assert len(folders) > 0
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 3: list_emails -----
            total += 1
            log("\n--- Test 3: list_emails (inbox, 3) ---")
            try:
                result = await session.call_tool("list_emails", {"folder": "inbox", "count": 3})
                content = result.content[0].text
                emails = json.loads(content)
                log(f"  Got {len(emails)} emails:")
                for e in emails:
                    log(f"    - {e['subject'][:60]}")
                assert len(emails) > 0
                first_entry_id = emails[0]["entry_id"]
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")
                first_entry_id = None

            # ----- Test 4: read_email -----
            total += 1
            log("\n--- Test 4: read_email (by entry_id) ---")
            try:
                assert first_entry_id, "No entry_id from previous test"
                result = await session.call_tool("read_email", {"entry_id": first_entry_id})
                content = result.content[0].text
                email_data = json.loads(content)
                log(f"  Subject: {email_data['subject']}")
                log(f"  From: {email_data['sender_name']}")
                log(f"  Body: {email_data['body'][:100]}...")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 5: search_emails -----
            total += 1
            log("\n--- Test 5: search_emails ('Feedback') ---")
            try:
                result = await session.call_tool("search_emails", {
                    "query": "Feedback", "folder": "inbox", "count": 5
                })
                content = result.content[0].text
                results = json.loads(content)
                log(f"  Found {len(results)} results")
                for r in results:
                    log(f"    - {r['subject'][:60]}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 6: send_email -----
            total += 1
            log("\n--- Test 6: send_email ---")
            try:
                result = await session.call_tool("send_email", {
                    "to": "user@example.com",
                    "subject": "Outlook Desktop MCP - Phase 3 MCP Test",
                    "body": "Sent through the MCP server via stdio. If you see this, the MCP layer works!",
                })
                content = result.content[0].text
                log(f"  Result: {content}")
                assert "sent" in content.lower()
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

            # ----- Test 7: read_email by subject_search -----
            total += 1
            log("\n--- Test 7: read_email (by subject_search) ---")
            try:
                result = await session.call_tool("read_email", {
                    "subject_search": "Feedback", "folder": "inbox"
                })
                content = result.content[0].text
                email_data = json.loads(content)
                if "error" in email_data:
                    log(f"  No match (expected if no 'Feedback' emails): {email_data['error']}")
                else:
                    log(f"  Found: {email_data['subject']}")
                passed += 1
                log("  PASS")
            except Exception as e:
                log(f"  FAIL: {e}")

    log("")
    log("=" * 60)
    log(f"Results: {passed}/{total} passed")
    log("=" * 60)
    return passed == total


if __name__ == "__main__":
    success = asyncio.run(run_tests())
    sys.exit(0 if success else 1)
