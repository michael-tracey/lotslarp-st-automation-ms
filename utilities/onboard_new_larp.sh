#!/usr/bin/env bash
#
# onboard_new_larp.sh — Spin up a fresh copy of this Apps Script project for a
# new LARP.
#
# It uses clasp to (1) create a brand-new Google Sheet with its own bound
# Apps Script container and (2) push this repo's code into it. After it runs,
# the new LARP's storytellers just open the sheet and run
# "Storyteller Menu > Maintenance > Initialise Project".
#
# This does NOT use the Apps Script API directly (which would need interactive
# user OAuth and can't use the service account in utilities/) — it leans on
# clasp, the same toolchain this repo already uses.
#
# Usage:
#   ./utilities/onboard_new_larp.sh "<Project Title>"
#
# Run it from the root of a FRESH clone of this repo. Because .clasp.json is
# gitignored, a fresh clone has no script binding — this script mints a new one.
# It refuses to run if a .clasp.json already exists, so it can never overwrite
# or push into an existing (e.g. production) project by mistake.

set -euo pipefail

# --- Locate the repo root (the dir holding the .js/.html that clasp pushes) ---
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# --- Args ---
TITLE="${1:-}"
if [[ -z "$TITLE" ]]; then
    echo "Usage: $0 \"<Project Title>\"" >&2
    echo "  e.g. $0 \"Nightfall LARP — ST Automation\"" >&2
    exit 1
fi

# --- Preflight: clasp present ---
if ! command -v clasp >/dev/null 2>&1; then
    echo "ERROR: clasp is not installed." >&2
    echo "  Install it with:  npm install -g @google/clasp" >&2
    exit 1
fi

# --- Safety: never clobber an existing binding ---
if [[ -f "$REPO_ROOT/.clasp.json" ]]; then
    echo "ERROR: $REPO_ROOT/.clasp.json already exists." >&2
    echo "  This checkout is already linked to an Apps Script project, and this" >&2
    echo "  script refuses to overwrite it." >&2
    echo "" >&2
    echo "  To onboard a new LARP, clone a fresh copy and run this there:" >&2
    echo "    git clone git@github.com:michael-tracey/lotslarp-st-automation-ms.git new-larp" >&2
    echo "    cd new-larp" >&2
    echo "    ./utilities/onboard_new_larp.sh \"$TITLE\"" >&2
    exit 1
fi

cd "$REPO_ROOT"

# --- Preflight: clasp logged in ---
if ! clasp show-authorized-user >/dev/null 2>&1; then
    echo "You are not logged in to clasp. Launching 'clasp login'..."
    clasp login
fi

echo ""
echo "==> Creating a new Google Sheet + bound Apps Script project:"
echo "    \"$TITLE\""
echo ""

# clasp create writes .clasp.json here and prints the new Sheet + script URLs.
clasp create --type sheets --title "$TITLE"

echo ""
echo "==> Pushing project code to the new script..."
clasp push -f

# --- Resolve the new script URL from .clasp.json for convenience ---
SCRIPT_ID=""
if command -v python3 >/dev/null 2>&1; then
    SCRIPT_ID="$(python3 -c 'import json,sys; print(json.load(open(".clasp.json")).get("scriptId",""))' 2>/dev/null || true)"
fi

echo ""
echo "============================================================"
echo " ✅  New project created and code pushed."
echo "============================================================"
if [[ -n "$SCRIPT_ID" ]]; then
    echo " Script editor: https://script.google.com/d/$SCRIPT_ID/edit"
fi
echo " (The new Google Sheet URL is printed in the 'clasp create' output above.)"
echo ""
echo " Next steps for the new LARP:"
echo "   1. Open the new Google Sheet (link above)."
echo "   2. Reload it, then run:"
echo "        Storyteller Menu > Maintenance > Initialise Project"
echo "   3. Set the Script Properties (Extensions > Apps Script > Project"
echo "      Settings > Script Properties). See README.md section 2.2 for the"
echo "      full list (webhooks, form ID, downtime month/year, etc.)."
echo "   4. Create the downtime Google Form and reinstall triggers:"
echo "        Storyteller Menu > Maintenance > Reinstall Form Trigger"
echo "        Storyteller Menu > Maintenance > Reinstall Edit Trigger"
echo "============================================================"
