#!/bin/bash

# Setup script for XTM Monthly Report automation

echo "Setting up XTM Monthly Report automation..."

# Copy the plist file to LaunchAgents
PLIST_FILE="com.xtm.monthlyreport.plist"
LAUNCH_AGENTS_DIR="$HOME/Library/LaunchAgents"

# Create LaunchAgents directory if it doesn't exist
mkdir -p "$LAUNCH_AGENTS_DIR"

# Copy the plist file
cp "$PLIST_FILE" "$LAUNCH_AGENTS_DIR/"

# Load the LaunchAgent
launchctl unload "$LAUNCH_AGENTS_DIR/$PLIST_FILE" 2>/dev/null
launchctl load "$LAUNCH_AGENTS_DIR/$PLIST_FILE"

echo ""
echo "✓ Automation setup complete!"
echo ""
echo "The report will be automatically generated and sent on the 1st of each month at 9:00 AM."
echo ""
echo "Useful commands:"
echo "  - Check status:    launchctl list | grep xtm"
echo "  - View logs:       tail -f xtm_report.log"
echo "  - Test manually:   python3 generate_report.py --auto-send"
echo "  - Disable:         launchctl unload ~/Library/LaunchAgents/$PLIST_FILE"
echo "  - Re-enable:       launchctl load ~/Library/LaunchAgents/$PLIST_FILE"
echo ""
