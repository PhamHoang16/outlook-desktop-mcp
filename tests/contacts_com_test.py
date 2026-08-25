"""
Outlook Desktop MCP - Contacts COM Integration Test
======================================================
Requires Classic Outlook on Windows. Run with:
  set RUN_OUTLOOK_INTEGRATION=1
  pytest tests/contacts_com_test.py -v
"""
import json
import sys

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32",
    reason="COM tests require Windows",
)


@pytest.fixture(autouse=True)
def _require_integration(skip_without_outlook_integration):
    pass


def test_list_contacts_com():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        folder = namespace.GetDefaultFolder(10)  # olFolderContacts
        items = folder.Items
        items.Sort("[FullName]", True)
        assert items.Count >= 0
        if items.Count > 0:
            item = items.Item(1)
            assert item.FullName is not None or item.Email1Address is not None
    finally:
        pythoncom.CoUninitialize()


def test_search_and_resolve_com():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        user = namespace.CurrentUser
        name = user.Name
        recipient = outlook.CreateRecipient(name)
        assert recipient.Resolve() is True
        assert recipient.AddressEntry.Address
        data = {
            "resolved": True,
            "email": recipient.AddressEntry.Address,
        }
        assert data["resolved"]
        json.dumps(data)
    finally:
        pythoncom.CoUninitialize()
