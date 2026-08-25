"""
Outlook Desktop MCP - Contacts MCP Integration Test
=====================================================
Requires Classic Outlook on Windows. Run with:
  set RUN_OUTLOOK_INTEGRATION=1
  pytest tests/contacts_mcp_test.py -v
"""
import json
import sys

import pytest

pytestmark = [
    pytest.mark.skipif(sys.platform != "win32", reason="MCP COM server requires Windows"),
    pytest.mark.asyncio,
]


@pytest.fixture(autouse=True)
def _require_integration(skip_without_outlook_integration):
    pass


async def test_contacts_tools_discovery_and_cache(stdio_server_params):
    from mcp.client.session import ClientSession
    from mcp.client.stdio import stdio_client

    params = stdio_server_params("outlook_desktop_mcp.server")

    async with stdio_client(params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            tools_result = await session.list_tools()
            tool_names = {t.name for t in tools_result.tools}
            for expected in ("list_contacts", "search_contacts", "resolve_recipient"):
                assert expected in tool_names

            result1 = await session.call_tool(
                "list_contacts", {"count": 5}
            )
            content1 = result1.content[0].text
            assert not content1.startswith("Error")
            contacts = json.loads(content1)
            assert isinstance(contacts, list)

            result2 = await session.call_tool(
                "list_contacts", {"count": 5}
            )
            content2 = result2.content[0].text
            assert content1 == content2

            empty_search = await session.call_tool(
                "search_contacts", {"query": "   ", "count": 5}
            )
            assert empty_search.content[0].text.startswith("Error:")

            if contacts:
                query = contacts[0].get("full_name", "")[:3]
                if query:
                    search1 = await session.call_tool(
                        "search_contacts", {"query": query, "count": 5}
                    )
                    search2 = await session.call_tool(
                        "search_contacts", {"query": query, "count": 5}
                    )
                    assert search1.content[0].text == search2.content[0].text

            user_name = contacts[0].get("full_name") if contacts else None
            if user_name:
                r1 = await session.call_tool(
                    "resolve_recipient", {"name": user_name}
                )
                r2 = await session.call_tool(
                    "resolve_recipient", {"name": user_name}
                )
                text = r1.content[0].text
                assert r1.content[0].text == r2.content[0].text
                data = json.loads(text)
                if data.get("resolved"):
                    assert data.get("email")
