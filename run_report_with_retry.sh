#!/bin/bash
#
# Wrapper script to run XTM report generation with retry logic
# This ensures the automation doesn't fail on transient issues
#

set -o pipefail

# Configuration
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PYTHON_SCRIPT="$SCRIPT_DIR/generate_report.py"
LOG_FILE="$SCRIPT_DIR/xtm_report.log"
MAX_ATTEMPTS=3
RETRY_DELAY=300  # 5 minutes between retries

# Change to script directory
cd "$SCRIPT_DIR" || exit 1

# Log start
echo "======================================" >> "$LOG_FILE"
echo "Starting XTM report generation at $(date)" >> "$LOG_FILE"
echo "======================================" >> "$LOG_FILE"

# Self-healing: Clear LaunchAgent error state if present
clear_error_state() {
    local plist_path="$HOME/Library/LaunchAgents/com.xtm.monthlyreport.plist"

    # Check if the LaunchAgent has a failed exit code
    if launchctl list com.xtm.monthlyreport &>/dev/null; then
        local exit_status=$(launchctl list com.xtm.monthlyreport 2>/dev/null | grep "LastExitStatus" | awk '{print $3}' | tr -d ';')

        # If exit status is non-zero and non-empty, clear the error state
        if [ -n "$exit_status" ] && [ "$exit_status" != "0" ]; then
            echo "⚠ Detected LaunchAgent error state (exit code: $exit_status). Clearing..." >> "$LOG_FILE"

            # Unload and reload to clear the error state
            launchctl unload "$plist_path" 2>/dev/null || true
            sleep 1
            launchctl load "$plist_path" 2>/dev/null || true

            echo "✓ LaunchAgent error state cleared" >> "$LOG_FILE"
        fi
    fi
}

# Clear any error state before starting
clear_error_state

# Function to run the report
run_report() {
    python3 "$PYTHON_SCRIPT" --auto-send
    return $?
}

# Retry logic
attempt=1
success=false

while [ $attempt -le $MAX_ATTEMPTS ]; do
    echo "Attempt $attempt of $MAX_ATTEMPTS..." >> "$LOG_FILE"

    if run_report; then
        echo "✓ Report generation succeeded on attempt $attempt" >> "$LOG_FILE"
        success=true
        break
    else
        exit_code=$?
        echo "✗ Report generation failed on attempt $attempt (exit code: $exit_code)" >> "$LOG_FILE"

        if [ $attempt -lt $MAX_ATTEMPTS ]; then
            echo "Waiting $RETRY_DELAY seconds before retry..." >> "$LOG_FILE"
            sleep $RETRY_DELAY
        fi
    fi

    attempt=$((attempt + 1))
done

# Final status
if [ "$success" = true ]; then
    echo "======================================" >> "$LOG_FILE"
    echo "Report generation completed successfully at $(date)" >> "$LOG_FILE"
    echo "======================================" >> "$LOG_FILE"
    exit 0
else
    echo "======================================" >> "$LOG_FILE"
    echo "Report generation FAILED after $MAX_ATTEMPTS attempts at $(date)" >> "$LOG_FILE"
    echo "======================================" >> "$LOG_FILE"

    # Send notification about failure
    osascript -e 'display notification "XTM report generation failed after 3 attempts. Check logs." with title "XTM Report Error" sound name "Basso"' 2>/dev/null || true

    exit 1
fi
