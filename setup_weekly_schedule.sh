#!/bin/bash
# Script to set up weekly automated XTM report generation
# This script configures a LaunchAgent to run every Monday at 9:00 AM

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PLIST_NAME="com.xtm.weeklyreport.plist"
PLIST_SOURCE="${SCRIPT_DIR}/${PLIST_NAME}"
PLIST_DEST="${HOME}/Library/LaunchAgents/${PLIST_NAME}"

echo "Setting up XTM Weekly Report Automation"
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
if launchctl list | grep -q "com.xtm.weeklyreport"; then
    echo "Unloading existing weekly report agent..."
    launchctl unload "${PLIST_DEST}" 2>/dev/null || true
fi

# Copy plist to LaunchAgents directory
echo "Installing weekly report LaunchAgent..."
cp "${PLIST_SOURCE}" "${PLIST_DEST}"

# Load the agent
echo "Loading weekly report agent..."
launchctl load "${PLIST_DEST}"

# Verify it's loaded
if launchctl list | grep -q "com.xtm.weeklyreport"; then
    echo ""
    echo "✓ Weekly report automation successfully configured!"
    echo ""
    echo "Schedule: Every Monday at 9:00 AM"
    echo "Action: Generate report for previous 7 days and auto-send via Outlook"
    echo ""
    echo "Logs will be written to:"
    echo "  - ${SCRIPT_DIR}/xtm_weekly_report.log"
    echo "  - ${SCRIPT_DIR}/xtm_weekly_report_error.log"
    echo ""
    echo "To manually test the weekly report:"
    echo "  python3 ${SCRIPT_DIR}/generate_report.py --weekly --auto-send"
    echo ""
    echo "To check automation status:"
    echo "  launchctl list | grep xtm"
    echo ""
    echo "To view recent logs:"
    echo "  tail -50 ${SCRIPT_DIR}/xtm_weekly_report.log"
    echo ""
    echo "To disable weekly automation:"
    echo "  launchctl unload ${PLIST_DEST}"
else
    echo ""
    echo "✗ Failed to load weekly report agent"
    echo "Check the plist file for errors:"
    echo "  plutil -lint ${PLIST_DEST}"
    exit 1
fi
