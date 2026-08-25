# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)]()

**Turn your running Outlook Desktop into an MCP server.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

## Quick Start

**1. Install** (requires Python 3.12+):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook and start a Claude Code session.** That's it — tools are available immediately.

## How It Works — Platform Routing

When the server starts, it checks which operating system it is running on and takes one of two paths:

```
                        outlook-desktop-mcp starts
                                  |
                          sys.platform check
                         /                  \
                   "win32"                "darwin"
                      |                      |
              ┌───────┴────────┐    ┌────────┴────────┐
              │  server.py     │    │  server_mac.py   │
              │  COM Bridge    │    │  AppleScript     │
              │  (39 tools)    │    │  Bridge          │
              │                │    │  (32 tools)      │
              └───────┬────────┘    └────────┬─────────┘
                      |                      |
              OUTLOOK.EXE via         Microsoft Outlook
              COM / STA thread        via osascript
                      |                      |
              Exchange / M365         Exchange / M365
```

**Both paths use your locally running Outlook app and its existing authenticated session.** No cloud credentials, no Graph API tokens — the server inherits whatever account Outlook is signed into.

### Why two paths?

Windows Outlook (Classic) exposes a rich COM automation interface — the Outlook Object Model (`MSOUTL.OLB`). This has been the standard way to programmatically control Outlook on Windows for over 20 years. It provides deep access to mail rules, categories, MAPI properties, and the full folder hierarchy.

Mac Outlook does not support COM. Instead, it exposes an AppleScript dictionary that can be driven via the `osascript` command. The AppleScript interface covers the core operations — email, calendar, tasks — but does not expose rules, categories, or certain advanced MAPI features. This is a limitation of what Microsoft chose to include in Outlook for Mac's scripting dictionary, not a limitation of this project.

The server is structured as two parallel implementations with identical tool names and signatures, so MCP clients see the same interface regardless of platform. Tools that are not available on a given platform are simply not registered.

## Requirements

### Windows

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office **or Office 2019**. The new "modern" Outlook (`olk.exe`) does **not** support COM. For **offline / intranet-only Office 2019**, see [Office 2019 — offline and air-gapped deployment](#office-2019--offline-and-air-gapped-deployment).
- **Python 3.12+** (x64 or ARM64)
- **Outlook must be running** when the MCP server starts

Both x64 and ARM64 Windows are supported. On ARM64, all dependencies (`pywin32`, `mcp`, `pydantic-core`, `cryptography`, `cffi`, `rpds-py`) have prebuilt `win_arm64` wheels — see the [ARM64 install notes](#arm64-windows) below for the one extra `pip` flag you need.

#### Outlook "Programmatic Access" security prompts

When the MCP server first touches Outlook's COM API, Outlook may show a dialog: *"A program is trying to access email address information stored in Outlook"* (the Object Model Guard / Programmatic Access prompt). This is Outlook's protection against malicious automation.

You have three options:

1. **Click "Allow access for 10 minutes"** every time you start a session. Fine for casual use.
2. **Get the antivirus status to "Valid"** in *File > Options > Trust Center > Trust Center Settings > Programmatic Access*. When that line reads `Valid`, the "Never warn me about suspicious activity" radio becomes selectable and the prompts go away. On most personal machines with current Defender this works out of the box.
3. **Apply the registry policy** in [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg). From an **elevated** PowerShell or Command Prompt (Win+X → *Terminal (Admin)*), run:

   ```powershell
   reg import "C:\path\to\outlook-desktop-mcp\docs\suppress-outlook-oom-prompts.reg"
   ```

   Then fully quit Outlook (check Task Manager for stray `OUTLOOK.EXE` processes) and reopen it. This writes `AdminSecurityMode=3` and approves all `PromptOOM*` categories under `HKLM\Software\Policies\Microsoft\Office\16.0\Outlook\Security`, which Outlook honors regardless of AV status. On Intune/MDM-managed corporate devices, `HKCU\Software\Policies\...\Outlook` is locked and the HKLM keys may be overwritten on next policy sync — if `reg import` fails or the prompts come back, ask IT to push the equivalent settings via Group Policy.

> **Windows 11 ARM64 note:** Defender does not register with Outlook's `IOfficeAntiVirus` interface on ARM64, so Trust Center shows *"Antivirus status: Invalid"* and the *"Never warn me about suspicious activity"* radio stays greyed out **even when Outlook is launched as Administrator**. Option 2 is unavailable on ARM64; the `reg import` from option 3 is the only durable suppression path.

### macOS

- **Microsoft Outlook for Mac** — version 16.x or later
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

#### Required macOS permissions

The first time a tool runs, macOS will show **two permission prompts** that you must approve:

1. **Privacy & Automation** — a system dialog asks: *"python3.12 wants to control Microsoft Outlook"*. Click **Allow** to let the server send AppleScript commands to Outlook.

2. **Accessibility** — to read your Exchange/M365 inbox, the server uses macOS UI scripting (System Events). This requires Accessibility access for `python3.12`:
   - Open **System Settings > Privacy & Security > Accessibility**
   - Find **python3.12** in the list (it appears after the first prompt)
   - Toggle it **on**

   Without Accessibility enabled, calendar, tasks, and local folder tools will work, but listing Exchange inbox messages will return empty results.

Both permissions are one-time setup — macOS remembers them for future sessions.

## Office 2019 — offline and air-gapped deployment

This section is for **Microsoft Office 2019 (16.0) with Outlook Desktop (Classic)** on Windows in **offline, intranet-only, or air-gapped** environments. The MCP server does not call Microsoft Graph; it drives the **locally running** `OUTLOOK.EXE` via COM and reads whatever Outlook has already cached on disk.

### What works offline

| Capability | Offline behavior |
|------------|------------------|
| Mail, calendar, tasks, drafts | Works against the **local store / OST** for the signed-in mailbox |
| Attachments | Read/save from cached items (paths must be allowed by your policy) |
| `list_contacts` / `search_contacts` | Scans the **Contacts folder** in the selected store (not the full GAL) |
| `resolve_recipient` | Uses Outlook **CreateRecipient/Resolve** → **Global Address List (GAL)** when the **Offline Address Book (OAB)** is present; otherwise falls back to the Contacts folder |

You still need **Outlook running** and **at least one successful sign-in** so profiles, stores, and (for Exchange) the OST exist. Pure “never connected” installs have no mailbox data to automate.

### Use Classic Outlook, not “New Outlook”

Office 2019 ships **Outlook Desktop (Classic)** — `OUTLOOK.EXE` with COM automation. The separate **New Outlook** (`olk.exe`) does **not** expose the Object Model this project uses. In *File → Office Account*, confirm you are not switched to New Outlook.

Registry and Group Policy paths for security settings use the **16.0** hive (Office 2016 and later share this major version), for example:

`HKLM\Software\Policies\Microsoft\Office\16.0\Outlook\Security`

The bundled policy file [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg) targets that path and applies to **Office 2019**.

### Step 1 — Allow programmatic access (required)

Before any tool touches mail or the address book, Outlook may block COM with the **Programmatic Access** / Object Model Guard dialog. On locked-down or offline PCs this often surfaces as COM **`0x80020009`** (*DISP_E_EXCEPTION*) on contact search or GAL resolve.

Configure one of these (details in [Outlook "Programmatic Access" security prompts](#outlook-programmatic-access-security-prompts) above):

1. Approve the prompt when it appears (short-lived).
2. Set Trust Center **Antivirus status** to **Valid** and enable *Never warn me about suspicious activity* (when your environment allows it).
3. Import [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg) from an **elevated** shell, then restart Outlook completely.

On **MDM/Intune-managed** machines, user registry may be read-only; ask IT to deploy the same `16.0\Outlook\Security` policy keys via Group Policy.

### Step 2 — Sync the Offline Address Book (for `resolve_recipient`)

Corporate name resolution (`resolve_recipient`, meeting attendees) depends on the **GAL** copy Outlook keeps locally:

1. Connect to the corporate network (or VPN) at least once, or use your org’s approved OAB distribution point.
2. In Outlook: **Send/Receive → Download Address Book** (wording may vary by language pack).
3. Confirm **File → Account Settings → Address Books** lists an offline address book for your Exchange account.

Without a current OAB, `resolve_recipient` may return `"resolved": false` with a `gal_error` mentioning address-book access — even when **Contacts** search works.

### Step 3 — Contacts folder vs GAL (tool expectations)

| Tool | Data source | Offline tip |
|------|-------------|-------------|
| `list_contacts` | **Contacts** folder only | Populate contacts in Outlook or sync from your org’s contact source |
| `search_contacts` | **Contacts** folder only (DASL `Restrict`) | If search fails with `0x80020009`, fix programmatic access first; ensure the Contacts folder is not empty |
| `resolve_recipient` | **GAL** via `CreateRecipient`, then **Contacts** fallback | Requires OAB for directory names; use full SMTP (`user@domain`) when you already know the address |

**Caching:** Successful contact results are cached in the MCP process for **7 days** (256 entries, LRU). After bulk contact or OAB updates, **restart the MCP server** (and Claude Code session) so tools see fresh data.

### Step 4 — Multi-mailbox / multi-store (optional)

If you have several accounts in one Outlook profile, pass the `account` argument (substring of the display name from `list_accounts`) on contact tools so the server targets the correct **store**. Mail tools use folder names; drafts and some item lookups may need consistent `entry_id` + store context on multi-store profiles.

### Common COM errors (Office 2019 offline)

| HRESULT | Typical cause | What to do |
|---------|---------------|------------|
| `0x80020009` | DASL `Restrict` or address-book COM call rejected | Trust Center programmatic access; sync OAB; retry after Outlook restart |
| `0x80070005` / `0x80004005` | Access denied on address information | Approve the programmatic access prompt for address data |
| `0x8004010F` | Stale `entry_id` after move/delete | Re-list the folder and use a fresh ID |
| `0x80010108` | Outlook RPC disconnected | Quit `OUTLOOK.EXE` (Task Manager), reopen, restart MCP |

Error text from tools includes short hints when the server recognizes these codes.

### Verification checklist (before using MCP in production)

Run on the **same Windows user session** that will host Claude Code / the MCP server:

1. **Outlook Classic** is open and the correct mailbox is selected.
2. **Programmatic access** is approved (no repeated OOM guard dialogs on send/read).
3. **OAB** downloaded if you rely on `resolve_recipient` for directory names.
4. **Contacts** folder contains expected people (for `search_contacts` / local fallback).
5. Optional integration smoke test (machine with Outlook):

   ```powershell
   set RUN_OUTLOOK_INTEGRATION=1
   pytest tests/contacts_mcp_test.py tests/drafts_mcp_test.py -q
   ```

   **GitHub Actions:** Hosted runners cannot run these tests (no Office). Use a
   [self-hosted Windows runner with Office 2019](docs/office2019-self-hosted-runner.md)
   and workflow **Integration (Office 2019 / self-hosted)**. Monitor runs with
   `scripts/watch-integration-office2019.sh` or `gh run watch`.

### What this project does *not* solve offline

- **No Graph / OAuth** — cannot reach cloud APIs when the network is down; only local Outlook state.
- **No replacement for Exchange connectivity** — sending meeting invites to unresolved GAL names still requires Outlook to resolve attendees.
- **No bypass of IT security policy** — if Group Policy blocks programmatic access, registry merges may be reverted; work with administrators.

For development setup and CI, see [CONTRIBUTING.md](CONTRIBUTING.md). For suppressing security prompts via policy, see [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg).

## Available Tools by Platform

### Email

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `send_email` | yes | yes | Send an email with To/CC/BCC, plain text or HTML body |
| `list_emails` | yes | yes | List recent emails from any folder, with optional unread filter |
| `read_email` | yes | yes | Read full email content by entry ID or subject search |
| `search_emails` | yes | yes* | Search subjects + bodies (Windows). macOS scans subjects only. |
| `reply_email` | yes | yes | Reply or reply-all, preserving the conversation thread |
| `mark_as_read` | yes | yes | Mark a specific email as read |
| `mark_as_unread` | yes | yes | Mark a specific email as unread |
| `move_email` | yes | yes | Move an email to Archive, Trash, or any folder |
| `list_folders` | yes | yes | Browse the folder hierarchy with item counts |

### Calendar

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_events` | yes | yes | List upcoming events within a date range |
| `get_event` | yes | yes | Read full event details by entry ID |
| `create_event` | yes | yes | Create a personal calendar appointment |
| `create_meeting` | yes | yes | Create a meeting and send invitations to attendees |
| `update_event` | yes | yes | Modify an existing event's subject, time, location, etc. |
| `delete_event` | yes | yes | Delete an appointment or cancel a meeting |
| `respond_to_meeting` | yes | — | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | yes | yes | Search calendar events by keyword within a date range |

### Tasks

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_tasks` | yes | yes | List pending or completed tasks, sorted by due date |
| `get_task` | yes | yes | Read full task details including body and completion status |
| `create_task` | yes | yes | Create a new task with subject, due date, importance |
| `complete_task` | yes | yes | Mark a task as complete |
| `delete_task` | yes | yes | Remove a task |

### Attachments

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_attachments` | yes | yes | List all attachments on an email or calendar event |
| `save_attachment` | yes | yes | Download an attachment to a local directory |

### Drafts (rich text + inline images)

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_drafts` | yes | yes | List unsent drafts, sorted by last modified |
| `get_draft` | yes | yes | Read a draft including HTML body and attachment metadata |
| `create_draft` | yes | yes* | Save a new draft with optional HTML body and inline images |
| `update_draft` | yes | yes* | Edit fields, body, or attachments on an existing draft |
| `send_draft` | yes | yes | Send a previously saved draft |
| `delete_draft` | yes | yes | Permanently delete a draft |

`create_draft` / `update_draft` accept:

- `body` (plain text, always saved as fallback)
- `html_body` (rich HTML — supports formatting, tables, links, etc.)
- `inline_images`: list of file paths or `{"path": "/abs/x.png", "cid": "logo1", "placeholder": "{{LOGO}}"}` dicts. The HTML may reference them via `<img src="cid:logo1">` or `{{LOGO}}` placeholders. Unreferenced images are appended at the end of the body.
- `attachments`: list of file paths for ordinary (non-inline) attachments.

\*On Outlook for Mac, inline images are added as ordinary attachments — AppleScript cannot reliably set the per-attachment Content-ID required for true inline rendering. Recipients will see them as separate attachments.

### Contacts (address book)

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_contacts` | yes | yes* | List items from the **Contacts folder** (sorted A–Z) |
| `search_contacts` | yes | yes* | Search **Contacts folder** by name/email (not full GAL) |
| `resolve_recipient` | yes | yes* | Resolve a name to email; **Windows uses GAL** via COM |

**Caching:** Successful results are cached in memory for **7 days** (max **256** entries, LRU eviction). Failed resolutions and empty macOS results are not cached. Restart the MCP server to refresh after editing contacts.

\*macOS: AppleScript `contacts` only; scans up to **500** (list) or **1000** (search/resolve) local contacts. No GAL. If AppleScript contacts are unavailable, tools return a clear error.

### Categories, Rules, Out of Office (Windows only)

These tools rely on COM-specific APIs (MAPI property accessors, the Rules object model, and the Categories collection) that Outlook for Mac does not expose through AppleScript.

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_categories` | yes | — | List all available color categories in Outlook |
| `set_category` | yes | — | Set or clear categories on any email, event, or task |
| `list_rules` | yes | — | List all mail rules with enabled/disabled status |
| `toggle_rule` | yes | — | Enable or disable a mail rule by name |
| `get_out_of_office` | yes | — | Check whether Out of Office auto-reply is on or off |

**Total: 39 tools on Windows, 31 tools on macOS** (including 6 new draft tools on each platform).

## Architecture Details

### Windows: COM Bridge (`com_bridge.py`)

All Outlook COM operations run on a dedicated thread using the Single-Threaded Apartment (STA) model, as required by COM. The async MCP event loop dispatches tool calls to this thread via a queue and awaits results, keeping COM threading rules respected and the MCP protocol non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and works directly with the Outlook Object Model — `GetItemFromID`, `CreateItem`, `Items.Restrict` with DASL filters, and so on.

### macOS: AppleScript Bridge (`applescript_bridge.py`)

Each tool call builds an AppleScript string and executes it as a subprocess via `osascript`. There is no persistent connection — every call is stateless.

```
MCP tool call (async)
  → build AppleScript string
    → asyncio.create_subprocess_exec("osascript", "-e", script)
    → parse stdout text into structured data
  → JSON response back to MCP client
```

Each tool constructs a single AppleScript that fetches all needed data in one `osascript` call (no per-message subprocess loops). Results come back as delimited text, which the server parses into the same JSON structure the Windows server produces.

**Key differences from Windows:**

- Entry IDs on macOS are **numeric** (e.g. `42`), not hex strings. They identify items within their folder context.
- Folder references use AppleScript's **locale-independent keywords** (`inbox`, `sent items`, `drafts`, `deleted items`) rather than localized folder names.
- Search uses AppleScript's `whose` clause (e.g. `messages whose subject contains "query"`) instead of DASL filters.
- User input is escaped for safe embedding in AppleScript strings to prevent script injection.

## Install from Source

### Windows (x64)

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### Windows (ARM64)

The `[cli]` extra of `mcp` transitively pulls in `cryptography`, and pip's default resolver may pick a version that lacks a `win_arm64` wheel — which then fails to build because it requires a Rust toolchain plus OpenSSL. Install without the `cli` extra and force wheels-only resolution:

```powershell
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
& "C:\Program Files\Python312-arm64\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: pywin32 mcp
python -m pip install --no-deps -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

The base `mcp` package is sufficient for running the stdio server — the `[cli]` extra is only needed for the `mcp` developer CLI tools (`mcp dev`, `mcp inspector`), which aren't used at runtime.

Register from source the same way as x64:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### macOS

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" -e .
```

Register from source:

```bash
claude mcp add outlook-desktop -- /path/to/outlook-desktop-mcp/.venv/bin/python -m outlook_desktop_mcp
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*

Windows-only examples:

- *"What categories do I have? Set this email to 'Follow-up'"*
- *"List my mail rules"*
- *"Am I set as Out of Office?"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Platform detection → routes to correct server
    server.py                # Windows MCP server (39 tools, COM automation)
    server_mac.py            # macOS MCP server (32 tools, AppleScript)
    com_bridge.py            # Async-to-COM threading bridge (Windows)
    applescript_bridge.py    # Async osascript execution (macOS)
    tools/
      _folder_constants.py   # Outlook enums and constants (Windows)
    utils/
      formatting.py          # Email/event/task data extraction (Windows)
      errors.py              # COM error formatting (Windows)
      applescript_helpers.py # AppleScript escaping, date formatting (macOS)
  tests/
    phase1_com_test.py       # Email COM validation
    phase3_mcp_test.py       # Email MCP test
    calendar_com_test.py     # Calendar COM validation
    calendar_mcp_test.py     # Calendar MCP test
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## License

See [LICENSE](LICENSE) file.
