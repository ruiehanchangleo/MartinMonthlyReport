#!/bin/bash
# Script to set up the daily XTM per-user statistics snapshot
# This caches each active project's per-user stats so that when a project is
# later archived, reports can restore real user names instead of "Archived User".
# Configures a LaunchAgent to run every day at 6:00 PM.

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PLIST_NAME="com.xtm.snapshot.plist"
PLIST_SOURCE="${SCRIPT_DIR}/${PLIST_NAME}"
PLIST_DEST="${HOME}/Library/LaunchAgents/${PLIST_NAME}"

echo "Setting up XTM Daily Snapshot Automation"
echo "========================================="

# Check if Python script exists
if [ ! -f "${SCRIPT_DIR}/generate_report.py" ]; then
    echo "Error: generate_report.py not found in ${SCRIPT_DIR}"
    exit 1
fi

# Check if config exists
if [ ! -f "${SCRIPT_DIR}/xtm_config.json" ]; then
    echo "Error: xtm_config.json not found in ${SCRIPT_DIR}"
    exit 1
fi

# Create LaunchAgents directory if it doesn't exist
mkdir -p "${HOME}/Library/LaunchAgents"

# Unload existing agent if present
if launchctl list | grep -q "com.xtm.snapshot"; then
    echo "Unloading existing snapshot agent..."
    launchctl unload "${PLIST_DEST}" 2>/dev/null || true
fi

# Copy plist to LaunchAgents directory
echo "Installing snapshot LaunchAgent..."
cp "${PLIST_SOURCE}" "${PLIST_DEST}"

# Load the agent
echo "Loading snapshot agent..."
launchctl load "${PLIST_DEST}"

# Verify it's loaded
if launchctl list | grep -q "com.xtm.snapshot"; then
    echo ""
    echo "✓ Daily snapshot automation successfully configured!"
    echo ""
    echo "Schedule: Every day at 6:00 PM"
    echo "Action: Cache per-user statistics for all active projects"
    echo ""
    echo "Why: When a project is archived, the XTM API no longer returns the"
    echo "     per-user breakdown. The daily snapshot captures it while the"
    echo "     project is active, so reports keep real user names afterward."
    echo ""
    echo "Logs will be written to:"
    echo "  - ${SCRIPT_DIR}/xtm_snapshot.log"
    echo "  - ${SCRIPT_DIR}/xtm_snapshot_error.log"
    echo ""
    echo "To manually run a snapshot now:"
    echo "  python3 ${SCRIPT_DIR}/generate_report.py --snapshot"
    echo ""
    echo "To check automation status:"
    echo "  launchctl list | grep xtm"
    echo ""
    echo "To view recent logs:"
    echo "  tail -50 ${SCRIPT_DIR}/xtm_snapshot.log"
    echo ""
    echo "To disable snapshot automation:"
    echo "  launchctl unload ${PLIST_DEST}"
else
    echo ""
    echo "✗ Failed to load snapshot agent"
    echo "Check the plist file for errors:"
    echo "  plutil -lint ${PLIST_DEST}"
    exit 1
fi
