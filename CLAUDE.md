# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Automated monthly reporting system for XTM Cloud translation management. Generates Excel reports with translation metrics, excludes specific users, converts locale codes to readable language names, and emails reports via Microsoft Outlook.

## Core Commands

### Generate Reports

```bash
# Generate report with draft email (for review)
python3 generate_report.py

# Generate and automatically send email via Outlook
python3 generate_report.py --auto-send

# Debug: View user statistics (shows all users including excluded ones)
python3 debug_user_stats.py

# Export detailed user statistics to Excel
python3 export_user_report.py

# Test API connectivity
python3 debug_api.py

# Test single project data retrieval
python3 test_single_project.py
```

### Automation Setup

```bash
# Set up monthly automated sending (1st of month at 9:00 AM)
./setup_schedule.sh

# Test the automation configuration
./test_automation.sh

# Check if automation is running
launchctl list | grep xtm

# View logs
tail -f xtm_report.log
```

## Architecture

### Key Design Decisions

**Per-User Statistics API**: Uses `/projects/{id}/statistics` endpoint (not `/metrics`) to retrieve word counts broken down by user and workflow step. This enables filtering out specific users' work.

**Excluded Users**: By default, work from `leo.chang@familysearch.org` and `LeoAdmin` is excluded from all reports. This is handled in `get_project_statistics()` by filtering the `usersStatistics` array before aggregation.

**Locale Translation**: All locale codes (e.g., `es_ES`, `zh_TW`) are converted to readable language names (e.g., "Spanish", "Chinese (Traditional)") using the `LOCALE_TO_LANGUAGE` dictionary (66 mappings). This happens during data aggregation via `_locale_to_language_name()`.

**Custom Column Ordering**: Workflow steps appear in a specific order: Language, translate, correct, final review, Total. This is enforced in `create_workflow_sheet()` using a predefined `workflow_order` list.

### Main Class: XTMReportGenerator

**Core Methods:**

- `get_project_statistics(project_id, excluded_users)`: Fetches per-user statistics and filters out excluded users
- `aggregate_monthly_data(start_month, end_month)`: Aggregates all project data for the reporting period, summing words from filtered users
- `create_workflow_sheet(wb, sheet_name, data, title)`: Creates Excel sheet with AutoFilter enabled, custom column ordering, and bar charts
- `send_email_via_outlook(report_path, monthly_data, ytd_data)`: Handles email via Outlook with automatic launching and fallback to Apple Mail
- `_ensure_outlook_running()`: Launches Outlook if not running (for --auto-send mode)

**Data Flow:**

1. Load config from `xtm_config.json`
2. For each project, call `/projects/{id}/statistics`
3. Filter out excluded users from `usersStatistics` array
4. Aggregate word counts by language and workflow step
5. Convert locale codes to language names
6. Generate Excel with two sheets: Monthly and Year-to-Date
7. Enable AutoFilter on all data sheets
8. Add bar charts showing total words per language
9. Launch Outlook (if needed) and create/send email

### Excel Report Structure

**Two Sheets Generated:**

1. **Monthly**: Current month data (e.g., "2026-01")
2. **Year-to-Date**: Cumulative data from January to current month

**Each Sheet Contains:**

- Title rows with period information
- Column headers: Language, translate, correct, final review, Total
- Data rows sorted by total words (descending)
- AutoFilter enabled on all columns (for sorting/filtering)
- Bar chart showing total words per language
- Summary row at bottom with totals

### Email Automation (macOS)

**Outlook Priority**: Always tries Microsoft Outlook first, falls back to Apple Mail if Outlook unavailable.

**Auto-Launch**: In `--auto-send` mode, script detects if Outlook is running and launches it automatically if needed (waits up to 30 seconds for startup).

**AppleScript**: Uses AppleScript to control Outlook/Mail, creating messages with recipients and attachments.

**Draft vs Send**: Without `--auto-send`, creates draft for review. With `--auto-send`, sends immediately.

## Configuration Files

**xtm_config.json**: Contains API credentials, OneDrive path, and email recipients. The `auth_token` is sensitive.

**com.xtm.monthlyreport.plist**: LaunchAgent configuration for monthly scheduling (1st of month at 9:00 AM).

**xtm-docs.json**: Complete XTM REST API OpenAPI 3.0 specification (688KB).

## Utility Scripts

**debug_user_stats.py**: Shows all users (including excluded ones) with their projects, languages, and word counts. Useful for verifying exclusion logic.

**export_user_report.py**: Generates detailed Excel report with three sheets: User Summary, Languages by User, and Project Details. Highlights excluded users in red.

**test_single_project.py**: Tests API connection and data retrieval for a single project.

**debug_api.py**: Basic API connectivity test.

## Important Implementation Notes

### Excluded Users

To modify excluded users, update the default parameter in `get_project_statistics()`:
```python
def get_project_statistics(self, project_id: int, excluded_users: List[str] = None):
    if excluded_users is None:
        excluded_users = ["leo.chang@familysearch.org", "LeoAdmin"]
```

The exclusion is case-insensitive.

### Adding Languages

Add new locale codes to the `LOCALE_TO_LANGUAGE` dictionary at the top of the `XTMReportGenerator` class. If a locale is not in the dictionary, it will be displayed as-is.

### Workflow Step Ordering

To change column order, modify the `workflow_order` list in `create_workflow_sheet()`:
```python
workflow_order = ['translate', 'correct', 'final review']
```

Steps not in this list will appear after these (sorted alphabetically).

### Date Range

The script reports on the **previous complete month**. When run on February 1st, it generates January's complete data. When run on March 1st, it generates February's complete data. Month is calculated in `__init__()`:
```python
first_day_current_month = self.report_date.replace(day=1)
last_day_previous_month = first_day_current_month - timedelta(days=1)
self.report_month = last_day_previous_month.strftime('%Y-%m')
```

The Year-to-Date report covers January through the end of the previous month.

### Email Body

Email content is defined in `send_email_via_outlook()`. Includes monthly summary, YTD summary, and top 3 languages for each period.

## Logs

- **xtm_report.log**: All operations (INFO level and above)
- **xtm_report_error.log**: Errors only (from LaunchAgent stderr)

Both files are in the project root directory.

## Security

- `xtm_config.json` contains the XTM API authentication token
- Uses XTM-Basic authentication: `Authorization: XTM-Basic <token>`
- Token should be rotated regularly and never committed to public repositories
- OneDrive path is user-specific and hardcoded in the config

## macOS-Specific

This tool is designed for macOS and uses:
- AppleScript for Outlook/Mail control
- LaunchAgent for scheduling (not cron)
- POSIX file paths in AppleScript
- System Events for app detection

Windows or Linux would require significant changes to email automation.
