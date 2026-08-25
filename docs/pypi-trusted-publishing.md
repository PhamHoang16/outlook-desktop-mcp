# PyPI publish (GitHub Actions)

The **Publish to PyPI** workflow (`.github/workflows/publish.yml`) runs on every push to `main` after unit tests and security checks pass.

## Failure: `invalid-publisher`

If the publish job fails with:

```text
Trusted publishing exchange failure: invalid-publisher
sub: repo:<owner>/outlook-desktop-mcp:environment:main
```

(or without `environment` in `sub` after workflow update), PyPI received a valid GitHub OIDC token but **has no Trusted Publisher** row that matches this repository and workflow.

This is **not** a Python/test failure — configure PyPI (and optionally a repo secret) as below.

## Option A — Trusted publishing (recommended)

You must be a **maintainer** of the [outlook-desktop-mcp](https://pypi.org/project/outlook-desktop-mcp/) project on PyPI.

1. Log in at https://pypi.org → **Publishing** → project **outlook-desktop-mcp** → **Manage publishers**.
2. **Add a new pending publisher**:
   - **PyPI Project Name:** `outlook-desktop-mcp`
   - **Owner:** `BobbyNie` (your GitHub user or org that owns the repo)
   - **Repository name:** `outlook-desktop-mcp`
   - **Workflow name:** `publish.yml` (filename only, under `.github/workflows/`)
   - **Environment name:** leave **empty** (current workflow does not use a GitHub Environment)

3. Save, then re-run the failed workflow: **Actions** → **Publish to PyPI** → **Re-run failed jobs**.

### If you previously used `environment: main`

Older workflow revisions set `environment: main` on the publish job. In that case PyPI **Environment name** must be exactly `main`, and GitHub must define an environment named `main` under **Settings → Environments**. The current workflow omits this to reduce fork setup friction.

## Option B — API token (fallback)

If you cannot use trusted publishing yet:

1. Create a PyPI **API token** (scope: entire project or `outlook-desktop-mcp`).
2. In GitHub: **Settings → Secrets and variables → Actions** → **New repository secret**
   - Name: `PYPI_API_TOKEN`
   - Value: `pypi-...` token (including `pypi-` prefix)

The publish step uses the secret when present; otherwise it uses OIDC trusted publishing.

## Version gate

Before upload, CI checks that `pyproject.toml` `version` is **greater than** the version on PyPI. Bump version on `main` before expecting a new release (e.g. `0.5.0` → `0.5.1`).

## Claims reference (debugging)

When troubleshooting, compare PyPI publisher fields to the claims in the Actions log:

| Claim | Example (this repo) |
|-------|---------------------|
| `repository` | `BobbyNie/outlook-desktop-mcp` |
| `repository_owner` | `BobbyNie` |
| `workflow_ref` | `.../publish.yml@refs/heads/main` |
| `ref` | `refs/heads/main` |
| `environment` | empty (no GitHub Environment) |

See https://docs.pypi.org/trusted-publishers/troubleshooting/
