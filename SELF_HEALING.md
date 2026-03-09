# Self-Healing Automation

## Overview

The XTM report automation now includes a **self-healing mechanism** that automatically detects and clears LaunchAgent error states before each run.

## How It Works

### Problem It Solves

macOS LaunchAgents can enter a failed state (exit code 78: EX_CONFIG) due to:
- Configuration file modifications while the agent is loaded
- System state inconsistencies
- Unexpected crashes or interruptions

Once in this state, launchd refuses to run the service again until the state is manually cleared.

### Automatic Recovery

Before each scheduled run, the wrapper script (`run_report_with_retry.sh`) now:

1. **Checks LaunchAgent Status** - Queries launchctl for the current exit status
2. **Detects Error State** - If exit status is non-zero, an error state is detected
3. **Clears Error State** - Automatically unloads and reloads the LaunchAgent
4. **Logs Action** - Records the recovery in `xtm_report.log`

### Implementation

```bash
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
```

## Benefits

- **Zero Downtime** - Automation recovers automatically without manual intervention
- **Silent Recovery** - Only logs when recovery is needed
- **Fail-Safe** - Uses `|| true` to prevent recovery errors from breaking the script
- **Audit Trail** - All recovery actions are logged

## Monitoring

Check for self-healing activity in the logs:

```bash
grep "LaunchAgent error state" xtm_report.log
```

Example log output:
```
⚠ Detected LaunchAgent error state (exit code: 78). Clearing...
✓ LaunchAgent error state cleared
```

## Manual Override

If needed, you can still manually clear the error state:

```bash
./setup_schedule.sh  # Unloads, reloads, and reconfigures
```

## Updated: March 9, 2026
