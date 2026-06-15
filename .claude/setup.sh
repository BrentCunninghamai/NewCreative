#!/usr/bin/env bash
# Prepare the dev environment for a (possibly fresh) Claude Code session.
#
# Runs from the SessionStart hook in .claude/settings.json. Idempotent: safe to
# run repeatedly. Uses a virtualenv because this project's deps (azure-identity
# -> cryptography) don't work against the OS-packaged site-packages in some
# container images.
set -euo pipefail

# Resolve repo root relative to this script so it works from any cwd.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

if [ ! -d .venv ]; then
  python3 -m venv .venv
fi

.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet -e ".[dev]"

# Best-effort: install the GitHub Copilot CLI if npm is available. Non-fatal so
# a missing/offline npm never blocks session startup.
if command -v npm >/dev/null 2>&1; then
  npm install -g @github/copilot >/dev/null 2>&1 || true
fi

echo "dev environment ready (.venv)"
