"""
Outlook Desktop MCP Server
===========================
Exposes Microsoft Outlook Desktop (Classic) as an MCP server over stdio.
Uses COM automation — no Microsoft Graph, no Entra app registration.
Just run this on Windows with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp.server
"""
import sys
import json
import logging

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from datetime import datetime, timedelta

import os

from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_CONTACTS,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
    format_task_summary,
    format_task_full,
    format_contact_summary,
)
from outlook_desktop_mcp.utils.attachment_safety import (
    UnsafeAttachmentPath,
    ensure_save_directory,
    resolve_attachment_path,
    validate_readable_attachment,
)
from outlook_desktop_mcp.utils.draft_html import (
    InvalidInlineImage,
    prepare_inline_html,
    suggest_attachment_basename,
)
from outlook_desktop_mcp.utils.errors import (
    format_bridge_exception,
    format_com_error,
    tool_error_json,
)
from outlook_desktop_mcp.utils.dasl import dasl_date_literal
from outlook_desktop_mcp.utils.com_compat import (
    coerce_outlook_datetime,
    compute_task_reminder_time,
    save_then_move,
)
from outlook_desktop_mcp.utils.contact_cache import ContactCache
from outlook_desktop_mcp.utils.contact_helpers import (
    clamp_contact_count,
    contact_matches_query,
    normalize_account_key,
    normalize_search_query,
    pick_best_contact_match,
    should_cache_contact_result,
)

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


# --- Security helpers ---

def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value.

    Escapes SQL wildcards (``%``, ``_``) and the ``[`` LIKE character-class
    opener so user input is matched as a literal substring, then escapes
    quote characters required by DASL syntax.
    """
    query = query.replace("[", "[[]")
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


# Outlook item Class constants (olObjectClass — distinct from olItemType used in CreateItem)
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48
_OL_CLASS_CONTACT = 40
_OL_CLASS_MEETING_REQUEST = 53  # olMeetingRequest

_OL_MEETING_RECEIVED = 3  # OlMeetingStatus
_OL_MEETING_RECEIVED_AND_CANCELED = 7

# olSendUsingAccount DispID (PR_OL_ACCOUNT_KEY hook)
_PR_SEND_USING_ACCOUNT = 64209


def _check_item_class(item, expected_class: int, label: str) -> str | None:
    """Return an error string if item is the wrong type, else None."""
    if item.Class != expected_class:
        return f"Error: Entry ID does not refer to a {label}."
    return None


# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook Desktop on "
        "Windows via COM automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders (including archive), reply to emails, and list the "
        "complete folder hierarchy.\n\n"
        "All operations use the locally authenticated Outlook profile — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
        "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move, attachments\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "respond to invites, search events\n"
        "- Tasks: create, list, complete, update, delete to-do items\n"
        "- Categories: list and set color categories on any item\n"
        "- Rules: list and manage mail rules\n"
        "- Out of Office: check auto-reply status\n"
        "- Folders: list folder hierarchy with item counts\n"
        "- Contacts: list/search Contacts folder, resolve names (GAL on Windows); "
        "7-day in-memory cache (256 entries max, failures not cached)\n"
        "- Drafts: list/get/create/update/send/delete drafts with optional HTML "
        "body and inline images (embedded via cid: references)"
    ),
)

bridge = OutlookBridge()
contact_cache = ContactCache()


# --- Helper: resolve store by account name ---


class AmbiguousAccountError(ValueError):
    """Raised when an account name matches multiple stores."""


def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore.
    Otherwise prefers exact (case-insensitive) match on Store.DisplayName.
    Falls back to a substring match only if no exact match exists. If multiple
    stores match the substring, raises :class:`AmbiguousAccountError` to
    prevent accidental cross-account writes.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    exact: list = []
    partial: list = []
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        name_lower = store.DisplayName.lower()
        if name_lower == account_lower:
            exact.append(store)
        elif account_lower in name_lower:
            partial.append(store)
    if exact:
        return exact[0]
    if len(partial) == 1:
        return partial[0]
    if len(partial) > 1:
        names = ", ".join(p.DisplayName for p in partial)
        raise AmbiguousAccountError(
            f"Account '{account}' is ambiguous; matched: {names}. "
            "Use the full display name to disambiguate."
        )
    return None


def _require_store(namespace, account: str = ""):
    """Resolve store, raising ValueError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise ValueError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


def _resolve_account_object(outlook, store):
    """Return the Outlook ``Account`` object whose DeliveryStore matches ``store``.

    Per-account ``DeliveryStore`` access can raise (e.g. POP3 accounts without
    a delivery store); skip those individually instead of aborting the whole
    scan, which would mask a later matching account.
    """
    try:
        accounts = outlook.Session.Accounts
    except Exception:
        return None
    for acc in accounts:
        try:
            if acc.DeliveryStore.StoreID == store.StoreID:
                return acc
        except Exception:
            continue
    return None


def _store_from_store_id(namespace, store_id: str):
    """Return the Store object matching ``store_id``, else DefaultStore."""
    for i in range(1, namespace.Stores.Count + 1):
        store = namespace.Stores.Item(i)
        if store.StoreID == store_id:
            return store
    return namespace.DefaultStore


def _get_item_for_account(namespace, entry_id: str, account: str = ""):
    """Resolve a COM item by entry_id, always scoping to a StoreID.

    EntryIDs are only unique within a store. When multiple mailboxes are
    loaded, ``GetItemFromID`` without a store can resolve the wrong item.
    """
    if account:
        store = _require_store(namespace, account)
    else:
        if namespace.Stores.Count > 1:
            names = [
                namespace.Stores.Item(i + 1).DisplayName
                for i in range(namespace.Stores.Count)
            ]
            raise ValueError(
                "account is required when multiple stores are loaded "
                f"({', '.join(names)}). Use list_accounts."
            )
        store = namespace.DefaultStore
    return namespace.GetItemFromID(entry_id, store.StoreID)


# --- Helper: resolve folder by name ---

def _walk_folders(parent, name_lower: str):
    """Recursively search subfolders of parent for a folder matching name_lower."""
    for i in range(parent.Folders.Count):
        try:
            f = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:
            continue
    return None


def _resolve_folder(namespace, folder_name: str, store=None):
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    # Slash-delimited path: traverse segment by segment
    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found = None
            for i in range(current.Folders.Count):
                try:
                    f = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    # Built-in Outlook folders
    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Root-level search (fast path)
    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:
            continue

    # Recursive fallback: search entire folder tree
    return _walk_folders(root, folder_lower)


# =====================================================================
# TOOL: list_accounts
# =====================================================================

@mcp.tool()
async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace):
        default_id = namespace.DefaultStore.StoreID
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            results.append({
                "display_name": store.DisplayName,
                "store_id": store.StoreID,
                "is_default": store.StoreID == default_id,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list)
    except Exception as e:
        return format_bridge_exception(e, action="listing accounts")


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _send(outlook, namespace, to, subject, body, cc, bcc, html_body, account):
        store = _require_store(namespace, account)
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        if account:
            sender_account = _resolve_account_object(outlook, store)
            if sender_account is None:
                raise ValueError(
                    f"Account '{account}' is a store but has no matching mail account "
                    f"(no DeliveryStore match). Refusing to send from the default identity."
                )
            mail._oleobj_.Invoke(*(_PR_SEND_USING_ACCOUNT, 0, 8, 0, sender_account))
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        mail.Send()
        return f"Email sent: '{subject}' to {to}"

    try:
        return await bridge.call(_send, to, subject, body, cc, bcc, html_body, account)
    except Exception as e:
        return format_bridge_exception(e, action="sending email")


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", or any custom folder name visible in
            list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        items = target.Items
        restrictions = []
        if unread_only:
            restrictions.append("[UnRead] = True")
        if start_date:
            start = _parse_date(start_date)
            restrictions.append(f"[ReceivedTime] >= {dasl_date_literal(start)}")
        if end_date:
            end = _parse_date(end_date)
            restrictions.append(f"[ReceivedTime] <= {dasl_date_literal(end)}")
        elif start_date:
            restrictions.append(f"[ReceivedTime] <= {dasl_date_literal(datetime.now())}")

        if restrictions:
            items = items.Restrict(" AND ".join(restrictions))
        items.Sort("[ReceivedTime]", True)

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, count, unread_only, start_date, end_date, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing emails")


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
    account: str = "",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The unique Outlook EntryID of the email. Most reliable way
            to identify a specific email. Get this from list_emails or
            search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    def _read(outlook, namespace, entry_id, subject_search, folder, account):
        if entry_id:
            item = _get_item_for_account(namespace, entry_id, account)
            return json.dumps(format_email_full(item), indent=2, default=str)

        if not subject_search:
            return json.dumps({"error": "Provide either entry_id or subject_search"})

        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(subject_search)
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        return json.dumps(format_email_full(items.Item(1)), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder, account)
    except Exception as e:
        return format_bridge_exception(e, action="reading email")


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str, account: str = "") -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return f"Marked as read: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="marking email as read")


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str, account: str = "") -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return f"Marked as unread: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="marking email as unread")


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
    account: str = "",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: The unique Outlook EntryID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.
        account: Optional. Account display name (or substring) to resolve
            the target folder in. Default: primary account.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    def _move(outlook, namespace, entry_id, target_folder, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item_store = _store_from_store_id(namespace, item.Parent.StoreID)
        if account:
            requested = _require_store(namespace, account)
            if requested.StoreID != item_store.StoreID:
                raise ValueError(
                    f"Account '{account}' does not match the mailbox containing "
                    f"this message ({item_store.DisplayName})."
                )
        dest = _resolve_folder(namespace, target_folder, item_store)
        if not dest:
            return f"Error: Target folder '{target_folder}' not found. Use list_folders to see available folders."

        moved = item.Move(dest)
        new_id = getattr(moved, "EntryID", None) if moved is not None else None
        return json.dumps({
            "status": "moved",
            "subject": subject,
            "target_folder": target_folder,
            "entry_id": new_id,
        }, indent=2, default=str)

    try:
        return await bridge.call(_move, entry_id, target_folder, account)
    except Exception as e:
        return format_bridge_exception(e, action="moving email")


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    account: str = "",
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        # 2 = olFormatHTML, 3 = olFormatRichText — preserve HTML when source is HTML
        body_format = getattr(item, "BodyFormat", 1)
        if body_format in (2, 3):
            safe_body = (
                body.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\n", "<br>")
            )
            reply_item.HTMLBody = f"<p>{safe_body}</p>" + (reply_item.HTMLBody or "")
        else:
            reply_item.Body = body + "\n\n" + reply_item.Body
        if account:
            # Also bind the reply to the requested sending account if available.
            store = _require_store(namespace, account)
            sender_account = _resolve_account_object(outlook, store)
            if sender_account is not None:
                reply_item._oleobj_.Invoke(
                    *(_PR_SEND_USING_ACCOUNT, 0, 8, 0, sender_account)
                )
        reply_item.Send()
        return f"Reply sent to '{subject}' (reply_all={reply_all})"

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, account)
    except Exception as e:
        return format_bridge_exception(e, action="replying to email")


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, and subfolders (if any).
    """
    def _list(outlook, namespace, folder, max_depth, account):
        max_depth = min(max(1, max_depth), 10)
        store = _require_store(namespace, account)

        if folder:
            start = _resolve_folder(namespace, folder, store)
            if not start:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            base_path = folder
        else:
            start = store.GetRootFolder()
            base_path = ""

        def walk(f, depth, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            result = {
                "name": f.Name,
                "full_path": current_path,
                "item_count": f.Items.Count,
                "unread_count": f.UnReadItemCount,
            }
            if depth < max_depth:
                children = []
                for i in range(f.Folders.Count):
                    try:
                        child = f.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1, current_path))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(start.Folders.Count):
            try:
                child = start.Folders.Item(i + 1)
                folders.append(walk(child, 1, base_path))
            except Exception:
                continue
        return json.dumps(folders, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, max_depth, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing folders")


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
    start_date: str = "",
    end_date: str = "",
    account: str = "",
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects and bodies using Outlook's DASL filter.
    Results are sorted by received time (newest first). Each result
    includes entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date, account):
        if not query.strip():
            return json.dumps({"error": "query must not be empty"})
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(query)
        dasl_parts = [
            f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
        ]
        if start_date:
            start = _parse_date(start_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" >= {dasl_date_literal(start)}"
            )
        if end_date:
            end = _parse_date(end_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= {dasl_date_literal(end)}"
            )
        elif start_date:
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= {dasl_date_literal(datetime.now())}"
            )

        filter_str = "@SQL=" + " AND ".join(dasl_parts)
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_email_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, folder, count, start_date, end_date, account)
    except Exception as e:
        return format_bridge_exception(e, action="searching emails")


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# --- Helper: parse ISO date string ---

def _parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'."""
    return datetime.fromisoformat(date_str)


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    account: str = "",
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Includes recurring event occurrences. Each summary has
    entry_id, subject, start, end, duration, location, organizer, attendees,
    and status info.

    Use entry_id from results with get_event, update_event, delete_event,
    or respond_to_meeting.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = calendar.Items

        # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)

        restrict = (
            f"[Start] >= {dasl_date_literal(start)} "
            f"AND [Start] <= {dasl_date_literal(end)}"
        )
        filtered = items.Restrict(restrict)

        results = []
        n = 0
        for item in filtered:
            n += 1
            try:
                results.append(format_event_summary(item))
            except Exception:
                continue
            if n >= count:
                break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, start_date, end_date, count, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing events")


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="reading event")


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    account: str = "",
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        account: Optional. Account display name (or substring) to create
            the event in. Default: primary account.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes, account):
        start_dt = coerce_outlook_datetime(start)
        end_dt = coerce_outlook_datetime(end)
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        if start_dt is not None:
            appt.Start = start_dt
        if end_dt is not None:
            appt.End = end_dt
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.AllDayEvent = all_day
        if reminder_minutes > 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
        else:
            appt.ReminderSet = False
        target_folder = None
        if account:
            store = _require_store(namespace, account)
            target_folder = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        appt = save_then_move(appt, target_folder)
        appt.Save()
        return json.dumps({
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "entry_id": appt.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes, account,
        )
    except Exception as e:
        return format_bridge_exception(e, action="creating event")


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    account: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and immediately sends meeting requests to
    all specified attendees. The meeting will appear on the organizer's
    calendar and attendees will receive an invitation they can accept,
    decline, or tentatively accept.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation that the meeting was created and invitations sent.
    """
    def _create(outlook, namespace, subject, start, end, required_attendees,
                location, body, optional_attendees, account):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        if account:
            store = _require_store(namespace, account)
            sender_account = _resolve_account_object(outlook, store)
            if sender_account is None:
                raise ValueError(
                    f"Account '{account}' has no matching mail account "
                    "(no DeliveryStore match). Refusing to send from default identity."
                )
            appt._oleobj_.Invoke(*(_PR_SEND_USING_ACCOUNT, 0, 8, 0, sender_account))
        start_dt = coerce_outlook_datetime(start)
        end_dt = coerce_outlook_datetime(end)
        appt.Subject = subject
        if start_dt is not None:
            appt.Start = start_dt
        if end_dt is not None:
            appt.End = end_dt
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        added_required: list[str] = []
        added_optional: list[str] = []
        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED
                added_required.append(addr)

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL
                    added_optional.append(addr)

        all_resolved = appt.Recipients.ResolveAll()
        unresolved: list[str] = []
        if not all_resolved:
            for i in range(appt.Recipients.Count):
                r = appt.Recipients.Item(i + 1)
                if not bool(r.Resolved):
                    unresolved.append(r.Name)
            return json.dumps({
                "status": "rejected",
                "reason": "unresolved_recipients",
                "unresolved": unresolved,
                "required": added_required,
                "optional": added_optional,
                "message": (
                    "Refusing to send meeting invitation: some recipients could "
                    "not be resolved by Outlook/GAL. Fix or remove them and retry."
                ),
            }, indent=2, default=str)

        appt.Send()
        return json.dumps({
            "status": "sent",
            "subject": subject,
            "required_attendees": added_required,
            "optional_attendees": added_optional,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, required_attendees, location, body,
            optional_attendees, account,
        )
    except Exception as e:
        return format_bridge_exception(e, action="creating meeting")


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. For meetings
    you organize, attendees will receive an update notification.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        start_dt = coerce_outlook_datetime(start) if start else None
        end_dt = coerce_outlook_datetime(end) if end else None
        if subject:
            item.Subject = subject
        if start_dt is not None:
            item.Start = start_dt
        if end_dt is not None:
            item.End = end_dt
        if location:
            item.Location = location
        if body:
            item.Body = body
        item.Save()
        return json.dumps({
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "entry_id": item.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body, account,
        )
    except Exception as e:
        return format_bridge_exception(e, action="updating event")


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str, account: str = "") -> str:
    """Delete a calendar event or cancel/decline a meeting.

    For personal appointments, the event is simply deleted. For meetings you
    organized, this cancels the meeting and sends cancellation notices to all
    attendees. For meetings you received, this sends a decline response to the
    organizer and removes the event from your calendar (falls back to a plain
    delete if the decline cannot be sent).

    Args:
        entry_id: The unique Outlook EntryID of the event to delete/cancel.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the event subject, or an error.
    """
    def _delete(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        subject = item.Subject
        meeting_status = item.MeetingStatus

        if meeting_status == OL_MEETING:
            item.MeetingStatus = OL_MEETING_CANCELED
            item.Send()
            return f"Meeting canceled: '{subject}' (cancellation sent to attendees)"

        if meeting_status == _OL_MEETING_RECEIVED:
            try:
                response_item = item.Respond(OL_RESPONSE_DECLINED, True, True)
                response_item.Send()
                item.Delete()
                return (
                    f"Meeting declined and removed: '{subject}' "
                    "(decline notice sent to organizer)"
                )
            except Exception as e:
                logger.warning(
                    "Could not auto-decline received meeting %r: %s. Falling back to plain delete.",
                    subject, e,
                )

        item.Delete()
        return f"Event deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="deleting event")


# =====================================================================
# TOOL 16: respond_to_meeting
# =====================================================================

@mcp.tool()
async def respond_to_meeting(
    entry_id: str,
    response: str,
    account: str = "",
) -> str:
    """Respond to a meeting invitation (accept, decline, or tentative).

    Accepts either a meeting request item (entry_id from list_emails on the
    Inbox) or an existing calendar appointment. Sends your response to the
    organizer; the meeting will be added to (or updated on) your calendar.

    Args:
        entry_id: The unique Outlook EntryID of the meeting to respond to.
            Get this from list_events or search_events.
        response: Your response. Must be one of: "accept", "decline",
            or "tentative".
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation of your response, or an error.
    """
    def _respond(outlook, namespace, entry_id, response, account):
        response_map = {
            "accept": OL_RESPONSE_ACCEPTED,
            "decline": OL_RESPONSE_DECLINED,
            "tentative": OL_RESPONSE_TENTATIVE,
        }
        response_lower = response.lower().strip()
        if response_lower not in response_map:
            return f"Error: response must be 'accept', 'decline', or 'tentative'. Got: '{response}'"

        item = _get_item_for_account(namespace, entry_id, account)
        item_class = getattr(item, "Class", None)

        if item_class == _OL_CLASS_MEETING_REQUEST:
            try:
                appointment = item.GetAssociatedAppointment(True)
            except Exception as e:
                return f"Error: meeting request has no associated appointment: {e}"
            subject = appointment.Subject
            # fNoUI=True, fAdditionalTextDialog=True: no dialog, return response
            # item that we then Send() ourselves (AppointmentItem.Respond docs).
            response_item = appointment.Respond(response_map[response_lower], True, True)
            response_item.Send()
            return f"Responded '{response_lower}' to meeting request: '{subject}'"

        if item_class == _OL_CLASS_APPOINTMENT:
            subject = item.Subject
            response_item = item.Respond(response_map[response_lower], True, True)
            response_item.Send()
            return f"Responded '{response_lower}' to meeting: '{subject}'"

        return (
            "Error: Entry ID does not refer to a meeting request or appointment "
            f"(class={item_class})."
        )

    try:
        return await bridge.call(_respond, entry_id, response, account)
    except Exception as e:
        return format_bridge_exception(e, action="responding to meeting")


# =====================================================================
# TOOL 17: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    account: str = "",
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)

        restrict = (
            f"[Start] >= {dasl_date_literal(start)} "
            f"AND [Start] <= {dasl_date_literal(end)}"
        )
        filtered = items.Restrict(restrict)

        query_lower = query.lower()
        results = []
        for item in filtered:
            if query_lower in (item.Subject or "").lower():
                try:
                    results.append(format_event_summary(item))
                except Exception:
                    continue
                if len(results) >= count:
                    break

        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, start_date, end_date, count, account)
    except Exception as e:
        return format_bridge_exception(e, action="searching events")


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    account: str = "",
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of task summary objects.
    """
    def _list(outlook, namespace, include_completed, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        if not include_completed:
            items = items.Restrict("[Complete] = False")
        items.Sort("[DueDate]")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_task_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_completed, count, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing tasks")


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    def _get(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        return json.dumps(format_task_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="reading task")


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
    reminder_minutes: int = 0,
    account: str = "",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".
        reminder_minutes: Optional. Minutes before due date to remind.
            Default 0 (no reminder).
        account: Optional. Account display name (or substring) to create
            the task in. Default: primary account.

    Returns:
        Confirmation with task subject and entry_id.
    """
    def _create(outlook, namespace, subject, body, due_date, importance,
                reminder_minutes, account):
        due_dt = coerce_outlook_datetime(due_date) if due_date else None
        task = outlook.CreateItem(OL_TASK_ITEM)
        task.Subject = subject
        if body:
            task.Body = body
        if due_dt is not None:
            task.DueDate = due_dt
        imp_map = {"low": 0, "normal": 1, "high": 2}
        task.Importance = imp_map.get(importance.lower(), 1)
        # TaskItem has ReminderTime (absolute datetime), NOT
        # ReminderMinutesBeforeStart (that's on AppointmentItem only).
        reminder_dt = compute_task_reminder_time(due_dt, reminder_minutes)
        if reminder_dt is not None:
            task.ReminderSet = True
            task.ReminderTime = reminder_dt
        else:
            task.ReminderSet = False
        target_folder = None
        if account:
            store = _require_store(namespace, account)
            target_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        task = save_then_move(task, target_folder)
        task.Save()
        return json.dumps({
            "status": "created",
            "subject": task.Subject,
            "entry_id": task.EntryID,
            "due_date": str(task.DueDate) if due_date else None,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, body, due_date, importance, reminder_minutes,
            account,
        )
    except Exception as e:
        return format_bridge_exception(e, action="creating task")


@mcp.tool()
async def complete_task(entry_id: str, account: str = "") -> str:
    """Mark a task as complete.

    Sets the task status to complete and percent to 100%.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _complete(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        item.Status = OL_TASK_COMPLETE
        item.PercentComplete = 100
        item.Save()
        return f"Task completed: '{item.Subject}'"

    try:
        return await bridge.call(_complete, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="completing task")


@mcp.tool()
async def delete_task(entry_id: str, account: str = "") -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The unique Outlook EntryID of the task to delete.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _delete(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        subject = item.Subject
        item.Delete()
        return f"Task deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="deleting task")


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str, account: str = "") -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of the email or event to check for attachments.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    def _list(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        results = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            results.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": att.Size,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing attachments")


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of the email or event containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    try:
        save_dir = ensure_save_directory(save_directory)
    except UnsafeAttachmentPath as e:
        return f"Error: {e}"

    def _save(outlook, namespace, entry_id, attachment_index, save_dir, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            return f"Error: Only {item.Attachments.Count} attachment(s), requested index {attachment_index}"

        att = item.Attachments.Item(attachment_index)
        original_name = att.FileName or "attachment"
        try:
            save_path = resolve_attachment_path(save_dir, original_name)
        except UnsafeAttachmentPath as e:
            return f"Error: {e}"

        att.SaveAsFile(save_path)
        return json.dumps({
            "status": "saved",
            "filename": os.path.basename(save_path),
            "original_filename": original_name,
            "path": save_path,
            "size": att.Size,
        }, indent=2, default=str)

    try:
        return await bridge.call(_save, entry_id, attachment_index, save_dir, account)
    except Exception as e:
        return format_bridge_exception(e, action="saving attachment")


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of category objects with name and color index.
    """
    def _list(outlook, namespace, account):
        # Categories are profile-wide, not per-store, but we accept the param for consistency
        results = []
        for i in range(namespace.Categories.Count):
            cat = namespace.Categories.Item(i + 1)
            results.append({"name": cat.Name, "color": cat.Color})
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing categories")


@mcp.tool()
async def set_category(
    entry_id: str,
    categories: str,
    account: str = "",
) -> str:
    """Set categories on an email, event, or task.

    Replaces any existing categories on the item. Use comma-separated
    values for multiple categories.

    Args:
        entry_id: The EntryID of the item to categorize.
        categories: Category name(s), comma-separated. Example:
            "Important" or "Work, Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the item subject and applied categories.
    """
    def _set(outlook, namespace, entry_id, categories, account):
        item = _get_item_for_account(namespace, entry_id, account)
        item.Categories = categories
        item.Save()
        return (
            f"Categories set on '{item.Subject}': "
            f"'{item.Categories or '(none)'}'"
        )

    try:
        return await bridge.call(_set, entry_id, categories, account)
    except Exception as e:
        return format_bridge_exception(e, action="setting categories")


# =====================================================================
# RULES TOOLS
# =====================================================================

@mcp.tool()
async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            rules = store.GetRules()
        except Exception as exc:
            return tool_error_json(
                f"This store does not expose Rules ({format_com_error(exc)}). "
                "Inbox rules require an Exchange / Microsoft 365 mailbox; "
                "local PST stores have no Rules collection.",
                code="rules_unavailable",
            )
        if rules is None:
            return tool_error_json(
                "This store has no Rules collection (typically a non-Exchange "
                "store such as a local PST).",
                code="rules_unavailable",
            )
        results = []
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            results.append({
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing rules")


@mcp.tool()
async def toggle_rule(
    rule_name: str,
    enabled: bool,
    account: str = "",
) -> str:
    """Enable or disable a mail rule by name.

    CAUTION: This modifies live mail rules immediately. Confirm the rule name
    with list_rules before calling.

    Args:
        rule_name: The exact name of the rule to toggle. Use list_rules
            to see available rule names.
        enabled: True to enable the rule, False to disable it.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation with the rule name and new status.
    """
    def _toggle(outlook, namespace, rule_name, enabled, account):
        store = _require_store(namespace, account)
        try:
            rules = store.GetRules()
        except Exception as exc:
            return tool_error_json(
                f"This store does not expose Rules ({format_com_error(exc)}).",
                code="rules_unavailable",
            )
        if rules is None:
            return tool_error_json(
                "This store has no Rules collection (typically a non-Exchange store).",
                code="rules_unavailable",
            )
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            if rule.Name == rule_name:
                logger.warning(
                    "toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled
                )
                rule.Enabled = enabled
                rules.Save()
                status = "enabled" if enabled else "disabled"
                return f"Rule '{rule_name}' {status}"
        return f"Error: Rule '{rule_name}' not found. Use list_rules to see available rules."

    try:
        return await bridge.call(_toggle, rule_name, enabled, account)
    except Exception as e:
        return format_bridge_exception(e, action="toggling rule")


# =====================================================================
# CONTACT / ADDRESS BOOK TOOLS
# =====================================================================


def _get_contacts_folder(store):
    """Return the default Contacts folder for a store."""
    return store.GetDefaultFolder(OL_FOLDER_CONTACTS)


def _iter_contact_items(items, limit: int):
    """Yield up to limit ContactItem objects from an Items collection."""
    count = 0
    total = items.Count
    for i in range(1, total + 1):
        if count >= limit:
            break
        try:
            item = items.Item(i)
            if item.Class == _OL_CLASS_CONTACT:
                yield item
                count += 1
        except Exception:
            continue


def _contact_dasl_filter(query: str) -> str:
    safe = _safe_dasl(query.replace("\xa0", " "))
    parts = [
        f'("urn:schemas:contacts:cn" LIKE \'%{safe}%\')',
        f'("urn:schemas:contacts:email1" LIKE \'%{safe}%\')',
        f'("urn:schemas:contacts:email2" LIKE \'%{safe}%\')',
        f'("urn:schemas:contacts:email3" LIKE \'%{safe}%\')',
    ]
    return "@SQL=" + " OR ".join(parts)


def _search_contacts_in_python(folder, query: str, count: int) -> list[dict]:
    scan_limit = min(max(count * 50, count), 2000)
    items = folder.Items
    try:
        items.Sort("[FullName]", False)
    except Exception:
        pass
    results: list[dict] = []
    for item in _iter_contact_items(items, scan_limit):
        summary = format_contact_summary(item)
        if contact_matches_query(summary, query):
            results.append(summary)
            if len(results) >= count:
                break
    return results


def _resolve_via_contacts_folder(folder, name: str) -> dict | None:
    matches = _search_contacts_in_python(folder, name, count=20)
    contact = pick_best_contact_match(matches, name)
    if not contact or not contact.get("email"):
        return None
    return {
        "resolved": True,
        "name": contact.get("full_name") or name,
        "email": contact["email"],
        "address_type": "SMTP",
        "source": "contacts_folder",
    }


@mcp.tool()
async def list_contacts(count: int = 50, account: str = "") -> str:
    """List contacts from the Outlook Contacts folder (not the full GAL).

    Returns contacts sorted by full name (A–Z). Results are cached in memory
    for 7 days (up to 256 distinct queries); restart the MCP server to refresh.

    Args:
        count: Maximum contacts to return (1-200). Default 50.
        account: Optional. Account display name (or substring). Use list_accounts
            to see available accounts.

    Returns:
        JSON array of contact summaries.
    """
    count = clamp_contact_count(count)

    def _list(outlook, namespace, count, account):
        store = _require_store(namespace, account)
        cache_key = contact_cache.make_key(
            "list_contacts",
            store_id=store.StoreID,
            count=count,
            account=normalize_account_key(account),
        )
        cached = contact_cache.get(cache_key)
        if cached is not None:
            return cached
        folder = _get_contacts_folder(store)
        items = folder.Items
        items.Sort("[FullName]", False)
        results = []
        for item in _iter_contact_items(items, count):
            results.append(format_contact_summary(item))
        result = json.dumps(results, indent=2, default=str)
        if should_cache_contact_result(result, tool="list_contacts"):
            contact_cache.set(cache_key, result)
        return result

    try:
        return await bridge.call(_list, count, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing contacts")


@mcp.tool()
async def search_contacts(
    query: str,
    count: int = 20,
    account: str = "",
) -> str:
    """Search the Outlook Contacts folder by name or email (not full GAL).

    Matches full name and email fields in the Contacts folder only. Use
    resolve_recipient for Global Address List name resolution on Windows.
    Results are cached for 7 days per unique query.

    Args:
        query: Search term (case-insensitive substring). Must not be empty.
        count: Maximum results (1-200). Default 20.
        account: Optional. Account display name (or substring).

    Returns:
        JSON array of matching contact summaries.
    """
    normalized = normalize_search_query(query)
    if normalized is None:
        return "Error: query must not be empty"
    count = clamp_contact_count(count)

    def _search(outlook, namespace, query, count, account):
        store = _require_store(namespace, account)
        cache_key = contact_cache.make_key(
            "search_contacts",
            store_id=store.StoreID,
            query=query,
            count=count,
            account=normalize_account_key(account),
        )
        cached = contact_cache.get(cache_key)
        if cached is not None:
            return cached
        folder = _get_contacts_folder(store)
        results: list[dict] = []
        try:
            items = folder.Items.Restrict(_contact_dasl_filter(query))
            try:
                items.Sort("[FullName]", False)
            except Exception:
                pass
            for item in _iter_contact_items(items, count):
                results.append(format_contact_summary(item))
        except Exception as exc:
            logger.warning(
                "Contacts Restrict failed (%s); falling back to Python scan",
                exc,
            )
            results = _search_contacts_in_python(folder, query, count)
        result = json.dumps(results, indent=2, default=str)
        if should_cache_contact_result(result, tool="search_contacts"):
            contact_cache.set(cache_key, result)
        return result

    try:
        return await bridge.call(_search, normalized, count, account)
    except Exception as e:
        return format_bridge_exception(e, action="searching contacts")


@mcp.tool()
async def resolve_recipient(name: str, account: str = "") -> str:
    """Resolve a display name or alias to an email address.

    Uses Outlook CreateRecipient/Resolve (includes Global Address List when
    available). Only successful resolutions are cached (7 days).

    Args:
        name: Display name, alias, or partial email to resolve.
            Examples: "Jane Doe", "jdoe", "jane@contoso.com".
        account: Optional. Validates account exists; resolution uses default
            Outlook session (not store-specific).

    Returns:
        JSON object with resolved, name, email, and address_type fields.
    """
    normalized_name = name.strip()
    if not normalized_name:
        return "Error: name must not be empty"
    def _resolve(outlook, namespace, name, account):
        store = _require_store(namespace, account)
        cache_key = contact_cache.make_key(
            "resolve_recipient",
            store_id=store.StoreID,
            name=name,
            account=normalize_account_key(account),
        )
        cached = contact_cache.get(cache_key)
        if cached is not None:
            return cached
        folder = _get_contacts_folder(store)
        gal_error: str | None = None
        try:
            recipient = outlook.CreateRecipient(name)
            if recipient.Resolve():
                entry = recipient.AddressEntry
                result = json.dumps({
                    "resolved": True,
                    "name": entry.Name or name,
                    "email": entry.Address or "",
                    "address_type": getattr(entry, "Type", "") or "",
                    "source": "gal",
                }, indent=2, default=str)
                if should_cache_contact_result(result, tool="resolve_recipient"):
                    contact_cache.set(cache_key, result)
                return result
        except Exception as exc:
            gal_error = format_com_error(exc)
            logger.warning("GAL resolve failed for %r: %s", name, gal_error)

        local = _resolve_via_contacts_folder(folder, name)
        if local is not None:
            result = json.dumps(local, indent=2, default=str)
            if should_cache_contact_result(result, tool="resolve_recipient"):
                contact_cache.set(cache_key, result)
            return result

        payload = {
            "resolved": False,
            "name": name,
            "email": "",
            "address_type": "",
            "message": (
                "Could not resolve via Global Address List or local Contacts."
            ),
        }
        if gal_error:
            payload["gal_error"] = gal_error
        return json.dumps(payload, indent=2)

    try:
        return await bridge.call(_resolve, normalized_name, account)
    except Exception as e:
        return format_bridge_exception(e, action="resolving recipient")


# =====================================================================
# DRAFT EMAIL TOOLS
# =====================================================================

# Outlook attachment Position=0 hides the attachment from the body footer.
_OL_ATTACH_BY_VALUE = 1
_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_ATTACH_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


def _attach_inline_image(mail, abs_path: str, cid: str) -> bool:
    """Attach a file as an inline image with the given Content-ID.

    Returns True if both PR_ATTACH_CONTENT_ID and PR_ATTACH_HIDDEN were set
    (so the recipient should see it as truly inline), False if either MAPI
    SetProperty call failed (in which case the attachment is still present
    but will likely render as a regular attachment).
    """
    display = suggest_attachment_basename(abs_path)
    attachment = mail.Attachments.Add(abs_path, _OL_ATTACH_BY_VALUE, 0, display)
    try:
        attachment.DisplayName = display
    except Exception:
        pass
    try:
        attachment.PropertyAccessor.SetProperty(_PR_ATTACH_CONTENT_ID, cid)
        attachment.PropertyAccessor.SetProperty(_PR_ATTACH_HIDDEN, True)
        return True
    except Exception as e:
        logger.warning("Could not set inline-image properties for %s: %s", abs_path, e)
        return False


def _apply_draft_body(
    mail,
    body,
    html_body,
    inline_images,
) -> list[dict]:
    """Populate a draft mail item with body, HTML body, and inline images.

    ``body`` and ``html_body`` may be ``None`` to leave the existing value
    untouched, or any string (including ``""``) to set/clear that field.
    ``inline_images`` may be ``None``/``[]`` (no change) or a list of entries
    accepted by :func:`utils.draft_html.prepare_inline_html`.

    Returns a list of dicts describing any inline-image attachments that
    could not have their Content-ID metadata set; callers should surface
    these to the user so they know inline rendering will fall back to a
    regular attachment.
    """
    if body is not None:
        mail.Body = body

    failures: list[dict] = []
    if html_body is not None or inline_images:
        effective_body = body if body is not None else (mail.Body or "")
        effective_html = html_body if html_body is not None else (
            getattr(mail, "HTMLBody", "") or ""
        )
        final_html, cid_pairs = prepare_inline_html(
            effective_body, effective_html, inline_images
        )
        mail.HTMLBody = final_html
        for cid, abs_path in cid_pairs:
            ok = _attach_inline_image(mail, abs_path, cid)
            if not ok:
                failures.append({"cid": cid, "path": abs_path})
    return failures


def _add_regular_attachments(mail, attachments) -> list[str]:
    """Add file paths as ordinary attachments.

    Each path is validated by :func:`validate_readable_attachment` to ensure
    the server never reads files outside the user's home or temp directory
    (defends against an LLM being tricked into exfiltrating credentials,
    SSH keys, etc. via ``send_draft``/``send_email``).

    Returns the list of basenames added.
    """
    added: list[str] = []
    if not attachments:
        return added
    for index, entry in enumerate(attachments):
        if not isinstance(entry, str) or not entry:
            raise ValueError("attachments must be a list of file path strings")
        try:
            abs_path = validate_readable_attachment(
                entry, label=f"attachments[{index}]"
            )
        except UnsafeAttachmentPath as e:
            raise ValueError(str(e)) from e
        mail.Attachments.Add(
            abs_path, _OL_ATTACH_BY_VALUE, 1, suggest_attachment_basename(abs_path)
        )
        added.append(os.path.basename(abs_path))
    return added


def _draft_summary(mail) -> dict:
    return {
        "entry_id": mail.EntryID,
        "subject": mail.Subject or "(no subject)",
        "to": getattr(mail, "To", "") or "",
        "cc": getattr(mail, "CC", "") or "",
        "bcc": getattr(mail, "BCC", "") or "",
        "has_html": bool((getattr(mail, "HTMLBody", "") or "").strip()),
        "attachment_count": int(mail.Attachments.Count),
    }


@mcp.tool()
async def list_drafts(count: int = 20, account: str = "") -> str:
    """List unsent draft emails from the Drafts folder.

    Returns a JSON array sorted by last-modified time (newest first). Each
    entry has entry_id, subject, to, cc, bcc, attachment_count, has_html.

    Args:
        count: Maximum drafts to return (1-200). Default 20.
        account: Optional. Account display name (or substring) to target.

    Returns:
        JSON array of draft summary objects.
    """
    def _list(outlook, namespace, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_DRAFTS)
        items = folder.Items
        try:
            items.Sort("[LastModificationTime]", True)
        except Exception:
            pass
        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                item = items.Item(i + 1)
                if item.Class != _OL_CLASS_MAIL:
                    continue
                results.append({
                    "entry_id": item.EntryID,
                    "subject": item.Subject or "(no subject)",
                    "to": getattr(item, "To", "") or "",
                    "cc": getattr(item, "CC", "") or "",
                    "bcc": getattr(item, "BCC", "") or "",
                    "has_html": bool((getattr(item, "HTMLBody", "") or "").strip()),
                    "attachment_count": int(item.Attachments.Count),
                    "last_modified": str(item.LastModificationTime),
                })
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, count, account)
    except Exception as e:
        return format_bridge_exception(e, action="listing drafts")


@mcp.tool()
async def get_draft(entry_id: str, account: str = "") -> str:
    """Read full details of a single draft email, including HTML body and attachments.

    Args:
        entry_id: The unique Outlook EntryID of the draft.
        account: Optional. Account display name (or substring) when needed.

    Returns:
        JSON object with entry_id, subject, to/cc/bcc, body, html_body, and
        a list of attachments (filename, size, inline flag, content_id).
    """
    def _get(outlook, namespace, entry_id, account):
        item = _get_item_for_account(namespace, entry_id, account)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "draft mail item"):
            return err

        attachments = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            cid = ""
            hidden = False
            try:
                cid = att.PropertyAccessor.GetProperty(_PR_ATTACH_CONTENT_ID) or ""
            except Exception:
                pass
            try:
                hidden = bool(att.PropertyAccessor.GetProperty(_PR_ATTACH_HIDDEN))
            except Exception:
                pass
            attachments.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": int(getattr(att, "Size", 0) or 0),
                "inline": bool(cid) or hidden,
                "content_id": cid,
            })

        return json.dumps({
            "entry_id": item.EntryID,
            "subject": item.Subject or "(no subject)",
            "to": getattr(item, "To", "") or "",
            "cc": getattr(item, "CC", "") or "",
            "bcc": getattr(item, "BCC", "") or "",
            "body": item.Body or "",
            "html_body": getattr(item, "HTMLBody", "") or "",
            "attachments": attachments,
            "last_modified": str(item.LastModificationTime),
        }, indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="reading draft")


@mcp.tool()
async def create_draft(
    to: str,
    subject: str,
    body: str = "",
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    inline_images: list | None = None,
    attachments: list | None = None,
    account: str = "",
) -> str:
    """Create a draft email and save it to the Drafts folder (does NOT send).

    Supports plain text, rich HTML, and inline images (embedded in the body
    via cid: references). Use send_draft to send it later.

    Args:
        to: One or more To recipients, separated by semicolons.
            Example: "alice@example.com; bob@example.com"
        subject: Subject line.
        body: Plain-text body. Always saved as text fallback. If html_body is
            empty and inline_images is empty, this is the rendered body.
        cc: Optional. CC recipients (semicolon-separated).
        bcc: Optional. BCC recipients (semicolon-separated).
        html_body: Optional. HTML body. When set, overrides the rendering of
            body. Use cid: references to embed inline images from
            inline_images (e.g. <img src="cid:logo1">).
        inline_images: Optional list of inline-image entries. Each may be:
            (a) a file path string — auto-assigned CID, appended to the end of
                the HTML inside a <p>;
            (b) a dict like {"path": "/abs/path.png", "cid": "logo1",
                "placeholder": "{{LOGO}}"} — CID is referenced as cid:logo1
                and the placeholder (if present in html_body) is replaced
                with <img src="cid:logo1">.
            Paths must be local files (no UNC, no missing files).
        attachments: Optional list of file paths for regular attachments.
            Same path rules as inline_images.
        account: Optional. Account display name (or substring) for the draft.

    Returns:
        JSON object describing the saved draft (entry_id, subject, etc.).
    """
    try:
        # Pre-validate paths without holding the COM thread.
        prepare_inline_html(body, html_body, inline_images)
    except InvalidInlineImage as e:
        return f"Error: {e}"

    def _create(outlook, namespace, to, subject, body, cc, bcc, html_body,
                inline_images, attachments, account):
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        if account:
            store = _require_store(namespace, account)
            sender_account = _resolve_account_object(outlook, store)
            if sender_account is None:
                raise ValueError(
                    f"Account '{account}' is a store but has no matching mail account "
                    "(no DeliveryStore match). Refusing to create a draft that would "
                    "later send from the default identity."
                )
            drafts = store.GetDefaultFolder(OL_FOLDER_DRAFTS)
            # Save first so Move() has a real item to relocate, then re-pick up
            # the moved item.
            mail.Save()
            moved = mail.Move(drafts)
            if moved is not None:
                mail = moved
            mail._oleobj_.Invoke(*(_PR_SEND_USING_ACCOUNT, 0, 8, 0, sender_account))
        mail.To = to or ""
        mail.Subject = subject
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        try:
            # On create, treat empty html_body as "no HTML supplied" rather
            # than "clear HTML" — there's nothing pre-existing to clear.
            html_arg = html_body if html_body else None
            inline_failures = _apply_draft_body(mail, body, html_arg, inline_images)
            _add_regular_attachments(mail, attachments)
        except (InvalidInlineImage, ValueError) as e:
            return f"Error: {e}"
        mail.Save()
        result = {"status": "created", **_draft_summary(mail)}
        if inline_failures:
            result["inline_image_failures"] = inline_failures
            result["warning"] = (
                "Some inline images could not be tagged with Content-ID; they will "
                "render as regular attachments instead of being embedded inline."
            )
        return json.dumps(result, indent=2, default=str)

    try:
        return await bridge.call(
            _create, to, subject, body, cc, bcc, html_body,
            inline_images, attachments, account,
        )
    except Exception as e:
        return format_bridge_exception(e, action="creating draft")


@mcp.tool()
async def update_draft(
    entry_id: str,
    to: str | None = None,
    subject: str | None = None,
    body: str | None = None,
    cc: str | None = None,
    bcc: str | None = None,
    html_body: str | None = None,
    inline_images: list | None = None,
    attachments: list | None = None,
    replace_attachments: bool = False,
    account: str = "",
) -> str:
    """Update an existing draft. Any field left at its default (None) is unchanged.

    Use replace_attachments=True to remove existing attachments before adding
    the new ones; otherwise the new attachments and inline images are added
    on top of what's already on the draft.

    Args mirror create_draft. inline_images/attachments are added (unless
    replace_attachments is true). To clear the HTML body, pass html_body="".

    Returns:
        JSON object describing the updated draft.
    """
    if inline_images is not None:
        try:
            prepare_inline_html(body or "", html_body or "", inline_images)
        except InvalidInlineImage as e:
            return f"Error: {e}"

    def _update(outlook, namespace, entry_id, to, subject, body, cc, bcc,
                html_body, inline_images, attachments, replace_attachments,
                account):
        mail = _require_draft(namespace, entry_id, account)
        if to is not None:
            mail.To = to
        if subject is not None:
            mail.Subject = subject
        if cc is not None:
            mail.CC = cc
        if bcc is not None:
            mail.BCC = bcc

        if replace_attachments:
            while mail.Attachments.Count > 0:
                try:
                    mail.Attachments.Remove(1)
                except Exception:
                    break

        inline_failures: list = []
        if body is not None or html_body is not None or inline_images:
            try:
                inline_failures = _apply_draft_body(
                    mail, body, html_body, inline_images
                )
            except InvalidInlineImage as e:
                return f"Error: {e}"

        try:
            _add_regular_attachments(mail, attachments)
        except ValueError as e:
            return f"Error: {e}"

        mail.Save()
        result = {"status": "updated", **_draft_summary(mail)}
        if inline_failures:
            result["inline_image_failures"] = inline_failures
        return json.dumps(result, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, to, subject, body, cc, bcc, html_body,
            inline_images, attachments, replace_attachments, account,
        )
    except Exception as e:
        return format_bridge_exception(e, action="updating draft")


def _require_draft(namespace, entry_id: str, account: str):
    """Resolve a draft mail item; raise if it's already sent or not in Drafts."""
    mail = _get_item_for_account(namespace, entry_id, account)
    if mail.Class != _OL_CLASS_MAIL:
        raise ValueError("Entry ID does not refer to a mail item.")
    if bool(getattr(mail, "Sent", False)):
        raise ValueError(
            "This mail item has already been sent — refusing to operate on it via "
            "a draft tool. Use the regular email tools instead."
        )
    try:
        store = _require_store(namespace, account) if account else namespace.DefaultStore
        drafts_id = store.GetDefaultFolder(OL_FOLDER_DRAFTS).EntryID
        parent_id = mail.Parent.EntryID
        if parent_id != drafts_id:
            raise ValueError(
                "This mail item is not in the Drafts folder — refusing to operate "
                "on it via a draft tool."
            )
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(
            "Could not verify item is in the Drafts folder; refusing draft "
            f"operation: {format_com_error(e)}"
        ) from e
    return mail


@mcp.tool()
async def send_draft(entry_id: str, account: str = "") -> str:
    """Send a previously saved draft email.

    Rejects items that have already been sent or that are not in the Drafts
    folder; use the regular reply/send tools for those.

    Args:
        entry_id: The unique Outlook EntryID of the draft to send.
        account: Optional. Account display name (or substring) when needed.

    Returns:
        Confirmation with the sent subject, or an error.
    """
    def _send(outlook, namespace, entry_id, account):
        mail = _require_draft(namespace, entry_id, account)
        subject = mail.Subject or "(no subject)"
        recipients = mail.To or ""
        mail.Send()
        return json.dumps({
            "status": "sent",
            "subject": subject,
            "to": recipients,
        }, indent=2, default=str)

    try:
        return await bridge.call(_send, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="sending draft")


@mcp.tool()
async def delete_draft(entry_id: str, account: str = "") -> str:
    """Delete a draft email (moves it to Deleted Items — not a hard delete).

    Rejects items that have already been sent or that are not in the Drafts
    folder. To recover, look in the Deleted Items folder.

    Args:
        entry_id: The unique Outlook EntryID of the draft.
        account: Optional. Account display name (or substring) when needed.

    Returns:
        Confirmation with the deleted subject, or an error.
    """
    def _delete(outlook, namespace, entry_id, account):
        mail = _require_draft(namespace, entry_id, account)
        subject = mail.Subject or "(no subject)"
        mail.Delete()
        return f"Draft moved to Deleted Items: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return format_bridge_exception(e, action="deleting draft")


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================

@mcp.tool()
async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with the OOF status.
    """
    def _get(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            oof_state = store.PropertyAccessor.GetProperty(prop_tag)
            return json.dumps({
                "out_of_office": bool(oof_state),
                "status": "on" if oof_state else "off",
            }, indent=2)
        except Exception:
            return json.dumps({
                "out_of_office": None,
                "status": "unknown",
                "note": "Could not read OOF property. Check Outlook settings directly.",
            }, indent=2)

    try:
        return await bridge.call(_get, account)
    except Exception as e:
        return format_bridge_exception(e, action="checking OOF status")


# =====================================================================
# Entry point
# =====================================================================

def main():
    logger.info("Starting Outlook Desktop MCP server...")
    bridge.start()
    logger.info("COM bridge ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
