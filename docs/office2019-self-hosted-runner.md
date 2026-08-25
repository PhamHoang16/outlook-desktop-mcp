# Office 2019 integration tests on GitHub Actions (self-hosted runner)

GitHub-hosted `windows-latest` runners **do not include** Microsoft Office. You cannot install a licensed Office 2019 there in a supported way for CI.

To run **real Outlook COM / MCP integration tests** in Actions, use a **self-hosted Windows runner** with **Classic Outlook (Office 16.0 / 2019)** already installed and signed in.

Workflow file: [`.github/workflows/integration-office2019.yml`](../.github/workflows/integration-office2019.yml)

## 1. Prepare the Windows machine

| Requirement | Notes |
|-------------|--------|
| Windows 10/11 x64 | Same arch as Python 3.12 |
| **Office 2019** — **Classic Outlook** (`OUTLOOK.EXE`) | Not New Outlook (`olk.exe`) |
| Mailbox configured | At least one profile with OST; OAB synced if you test `resolve_recipient` |
| Trust Center | Programmatic access allowed — see [README Office 2019 offline](../README.md#office-2019--offline-and-air-gapped-deployment) |
| Network | Runner needs HTTPS to `github.com` (can be intranet proxy; document proxy env if needed) |

Recommended: dedicated VM or physical PC used only for CI, same user session that runs the runner service.

## 2. Install Actions Runner

1. Repo → **Settings** → **Actions** → **Runners** → **New self-hosted runner** → **Windows**.
2. Follow GitHub’s download/configure commands on the machine.
3. When prompted for labels, include:

   ```
   self-hosted
   Windows
   outlook-office2019
   ```

   The workflow uses: `runs-on: [self-hosted, Windows, outlook-office2019]`.

4. Install the runner as a **service** (recommended) so Outlook can stay logged in under the service account, **or** run interactively under the user that owns the Outlook profile.

> **Important:** COM automation uses the **interactive Outlook profile** of the user running the runner. If the runner service runs as `NT AUTHORITY\NETWORK SERVICE`, Outlook may not see your mailbox. Prefer running the runner under your domain/user account with Outlook already configured.

## 3. One-time Outlook / Python setup on the runner

```powershell
# Python 3.12 x64
py -3.12 -m pip install --upgrade pip

# In your clone of the repo (Actions checkout does this each run; useful for manual debug):
cd C:\actions-runner\_work\outlook-desktop-mcp\outlook-desktop-mcp
py -3.12 -m pip install -e ".[dev]"
py -3.12 -m pywin32_postinstall -install
```

Open **Classic Outlook** once under the runner user and dismiss any first-run prompts.

## 4. Run the workflow

- **Actions** → **Integration (Office 2019 / self-hosted)** → **Run workflow**
- Optional inputs:
  - `ref` — branch/SHA to test
  - `pytest_args` — e.g. `-v --tb=short`

Pushes to `main` that touch integration tests or server code also trigger this workflow (integration job still requires the self-hosted runner).

## 5. Monitor results (CLI)

From any machine with `gh` authenticated to the repo:

```bash
# List recent integration runs
gh run list --workflow=integration-office2019.yml --limit 10

# Watch the latest run until it finishes
gh run watch $(gh run list --workflow=integration-office2019.yml --limit 1 --json databaseId -q '.[0].databaseId')

# View failed job logs
gh run view <run-id> --log-failed

# Download JUnit artifact after a run
gh run download <run-id> -n integration-office2019-junit
```

If the integration job stays **Queued**, no online runner has the `outlook-office2019` label — check **Settings → Actions → Runners**.

## 6. Security notes

- Self-hosted runners execute **arbitrary code from the repo** with access to Outlook and local mail data. Restrict who can push/workflow_dispatch on protected branches.
- Prefer a **private** runner on an isolated machine, not a shared admin workstation.
- Do not store mailbox passwords in repo secrets; rely on the existing Outlook profile.

## 7. Troubleshooting

| Symptom | Likely fix |
|---------|------------|
| Job queued forever | Runner offline or missing `outlook-office2019` label |
| `OUTLOOK.EXE not found` | Install Office 2019 Classic, not web-only |
| COM Error `0x80020009` | Trust Center programmatic access + OAB |
| `0x80010108` RPC | Restart Outlook; restart runner service |
| Tests skip | `RUN_OUTLOOK_INTEGRATION` not set — workflow sets it; check step env |
| Empty GAL resolve | Sync Offline Address Book on the runner machine |
