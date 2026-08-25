"""Unit tests for contact formatting helpers."""
from types import SimpleNamespace

from outlook_desktop_mcp.utils.formatting import format_contact_summary


def test_format_contact_summary_extracts_fields():
    item = SimpleNamespace(
        EntryID="abc123",
        FullName="Jane Doe",
        Email1Address="jane@example.com",
        CompanyName="Contoso",
        MobileTelephoneNumber="555-0100",
        JobTitle="Engineer",
    )
    result = format_contact_summary(item)
    assert result["full_name"] == "Jane Doe"
    assert result["email"] == "jane@example.com"
    assert result["company"] == "Contoso"
    assert result["phone"] == "555-0100"
    assert result["job_title"] == "Engineer"
    assert result["entry_id"] == "abc123"
