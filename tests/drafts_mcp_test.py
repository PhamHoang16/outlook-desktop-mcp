"""
Outlook Desktop MCP - Drafts MCP Integration Test
=====================================================
Requires Classic Outlook on Windows.

  set RUN_OUTLOOK_INTEGRATION=1
  pytest tests/drafts_mcp_test.py -v

Creates a draft (with HTML + a tiny inline PNG), lists drafts, reads it back,
updates it, then deletes it. Does NOT send any email.
"""
import base64
import json
import os
import sys
import tempfile

import pytest

pytestmark = [
    pytest.mark.skipif(sys.platform != "win32", reason="MCP COM server requires Windows"),
    pytest.mark.asyncio,
]


@pytest.fixture(autouse=True)
def _require_integration(skip_without_outlook_integration):
    pass


# 1x1 transparent PNG
_PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
)


@pytest.fixture
def inline_image_path():
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
        f.write(_PNG_BYTES)
        path = f.name
    yield path
    try:
        os.unlink(path)
    except OSError:
        pass


async def test_draft_lifecycle(stdio_server_params, inline_image_path):
    from mcp.client.session import ClientSession
    from mcp.client.stdio import stdio_client

    params = stdio_server_params("outlook_desktop_mcp.server")

    async with stdio_client(params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            tools = {t.name for t in (await session.list_tools()).tools}
            for expected in (
                "list_drafts", "get_draft", "create_draft",
                "update_draft", "send_draft", "delete_draft",
            ):
                assert expected in tools

            created = await session.call_tool("create_draft", {
                "to": "draft-test@example.invalid",
                "subject": "MCP draft test",
                "body": "fallback plain text",
                "html_body": "<p>Hello <strong>world</strong> {{LOGO}}</p>",
                "inline_images": [{
                    "path": inline_image_path,
                    "cid": "logo1",
                    "placeholder": "{{LOGO}}",
                }],
            })
            created_text = created.content[0].text
            assert not created_text.startswith("Error"), created_text
            data = json.loads(created_text)
            entry_id = data["entry_id"]
            assert data["status"] == "created"
            assert data["attachment_count"] >= 1

            listed = json.loads(
                (await session.call_tool("list_drafts", {"count": 50})).content[0].text
            )
            assert any(d["entry_id"] == entry_id for d in listed)

            fetched = json.loads(
                (await session.call_tool("get_draft", {"entry_id": entry_id})).content[0].text
            )
            assert "cid:logo1" in fetched["html_body"]
            inline = [a for a in fetched["attachments"] if a["inline"]]
            assert inline, "expected at least one inline attachment"
            assert any(a["content_id"] == "logo1" for a in fetched["attachments"])

            updated = await session.call_tool("update_draft", {
                "entry_id": entry_id,
                "subject": "MCP draft test (edited)",
            })
            updated_text = updated.content[0].text
            assert not updated_text.startswith("Error"), updated_text

            deleted_text = (await session.call_tool(
                "delete_draft", {"entry_id": entry_id}
            )).content[0].text
            assert deleted_text.startswith("Draft deleted")
