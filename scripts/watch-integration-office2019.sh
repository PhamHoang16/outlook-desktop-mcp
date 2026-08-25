#!/usr/bin/env bash
# Watch the latest Office 2019 integration workflow run until it completes.
set -euo pipefail
WORKFLOW="integration-office2019.yml"
RUN_ID="${1:-$(gh run list --workflow="$WORKFLOW" --limit 1 --json databaseId -q '.[0].databaseId')}"
if [[ -z "$RUN_ID" || "$RUN_ID" == "null" ]]; then
  echo "No runs found for workflow $WORKFLOW" >&2
  exit 1
fi
echo "Watching run $RUN_ID ($WORKFLOW) ..."
gh run watch "$RUN_ID" --exit-status
echo "--- Summary ---"
gh run view "$RUN_ID" --json conclusion,status,url,displayTitle -q '"\(.displayTitle): \(.status) / \(.conclusion)\n\(.url)"'
