# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Automated reporting system for XTM Cloud translation management. Generates both HTML and Excel reports with translation metrics, excludes specific users, converts locale codes to readable language names, and emails reports via Microsoft Outlook. Supports both monthly reports (with YTD data) and weekly reports (previous 7 days).

## Core Commands

### Generate Reports

```bash
# Monthly report with draft email (for review)
python3 generate_report.py

# Monthly report - automatically send email via Outlook
python3 generate_report.py --auto-send

# Weekly report (previous 7 days) with draft email
python3 generate_report.py --weekly

# Weekly report - automatically send email via Outlook
python3 generate_report.py --weekly --auto-send

# Snapshot per-user stats for all active projects (preserves names after archival)
python3 generate_report.py --snapshot

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

# Set up weekly automated sending (every Monday at 9:00 AM)
./setup_weekly_schedule.sh

# Set up daily snapshot (every day at 6:00 PM) so archived projects keep real user names
./setup_snapshot_schedule.sh

# Test the monthly automation configuration
./test_automation.sh

# Test the weekly report generation
./test_weekly_report.sh

# Test resilience features (retry, health checks, etc.)
python3 test_resilience.py

# Check if automation is running
launchctl list | grep xtm

# View monthly logs
tail -f xtm_report.log

# View weekly logs
tail -f xtm_weekly_report.log
```

### Resilience Features

The automation includes multiple layers of error handling to ensure it never fails:

1. **API Retry Logic**: All API requests automatically retry up to 5 times with exponential backoff (2s, 4s, 8s, 16s, 32s)
2. **Shell-Level Retry**: The wrapper script retries the entire process 3 times with 5-minute delays between attempts
3. **Graceful Degradation**: If individual projects fail, the script continues processing remaining projects
4. **Multiple Fallback Locations**: If the primary save location fails, tries Desktop → Current Directory → Temp Directory
5. **Health Checks**: Validates API connectivity, disk space, permissions, and configuration before starting
6. **Failure Notifications**: Sends macOS system notifications and email alerts when critical failures occur
7. **Partial Data Handling**: Generates reports even when only partial data is available
8. **Email Fallback**: Tries Outlook → Apple Mail → Just saves file if email fails

## Architecture

### Key Design Decisions

**Volunteer Hours (login/logout time)**: The reports include a "Volunteer Hours in XTM" section (HTML) and sheet (Excel) showing per-volunteer **active hours** = time from LOGIN to the last recorded action in each session. This data is NOT in the XTM REST API — the only REST time field (`manualTime` on `/projects/{id}/workflows/time-trackings/jobs`) is empty because volunteers never log manual time, and there is no session/activity endpoint. Real login/logout history lives only in the PM-GUI backend at `POST /project-manager-gui/getUserLoginHistory.serv`, which requires a browser session cookie plus a per-session `uust` header (the REST API token does not work there). `fetch_login_history.js` (Playwright/Node) drives a headed Chromium with a persistent profile (`.cache/pw-profile`), auto-submits the pre-filled login form (credentials saved in the profile), captures `uust` from the app's own requests, and pages through all records. `volunteer_hours.py` runs that fetcher via subprocess, pairs LOGIN/LOGOUT events per user, applies `EXCLUDED_USERS`, drops sessions > 16h (stale/left-open tabs), and returns per-volunteer active hours. `generate_report._compute_volunteer_hours()` calls it for the report period (whole month for monthly, the 7-day window for weekly) and stores the result on `self._volunteer_hours`; it is best-effort and never blocks report generation. Request dates are **DD-MM-YYYY** but the response `DATE` field is **MM-DD-YYYY** (an XTM quirk). NOTE: the fetch needs an interactive login and cannot run fully unattended — for scheduled LaunchAgent runs it only works if a GUI session is available to pop the browser; otherwise the report generates without the hours section. `.cache/pw-profile` holds the saved credentials/session and is gitignored — never commit it. `xtm_login_history_capture.js` is the DevTools-style capture tool used to reverse-engineer the endpoint.

**Per-User Statistics API**: Uses `/projects/{id}/statistics` endpoint (not `/metrics`) to retrieve word counts broken down by user and workflow step. This enables filtering out specific users' work.

**Archived-Project Snapshots**: When a project is archived, `/projects/{id}/statistics` stops returning the per-user breakdown, so reports can only fall back to `/metrics` (no user attribution) and bucket that work under a generic "Archived User". To prevent this, `snapshot_active_projects()` (run via `--snapshot`, scheduled daily by `com.xtm.snapshot.plist`) caches each active project's raw `/statistics` + `/status` to `.cache/snapshots/project_{id}.json`. During aggregation, if a project's live statistics are empty, `_restore_stats_from_snapshot()` restores the per-user data from the snapshot (re-applying the current excluded-user filter) before the metrics fallback runs — so real names are preserved. Snapshots are only written when statistics are non-empty, so an already-archived project never overwrites a good snapshot. This only helps projects snapshotted while still active; projects archived before any snapshot existed still use the "Archived User" fallback.

**Excluded Users**: By default, work from `leo.chang@familysearch.org`, `LeoAdmin`, `Robert.Sena@churchofjesuschrist.org`, `MartinADMIN`, and `Tester BSP BSP` is excluded from all reports. This is handled in `get_project_statistics()` by filtering the `usersStatistics` array before aggregation. (The live default list lives in the `EXCLUDED_USERS` class constant.)

**Full Volunteer Roster**: The user report lists *every* volunteer in XTM — defined as everyone returned by `GET /users` minus `EXCLUDED_USERS` — not just whoever logged work in the period. `get_volunteers()` fetches the roster plus each user's assigned target languages from `/users/{id}/language-combinations` (parallel; memoized in-process and cached to `.cache/volunteers.json` for 24h). After aggregation, `_inject_zero_volunteers()` adds a zero row per assigned language for any volunteer with no work this period (matched by `username`; volunteers who did work are untouched). Injection runs *after* the month cache is saved and YTD is built, so caches keep only real work. Charts stay limited to volunteers/languages with actual data — zero rows appear in tables and filters only.

**Locale Translation**: All locale codes (e.g., `es_ES`, `zh_TW`) are converted to readable language names (e.g., "Spanish", "Chinese (Traditional)") using the `LOCALE_TO_LANGUAGE` dictionary (66 mappings). This happens during data aggregation via `_locale_to_language_name()`.

**Custom Column Ordering**: Workflow steps appear in a specific order: Language, translate, correct, final review, Total. This is enforced in both HTML and Excel reports.

**Dual Report Format**: The system generates two complementary report formats:
- **HTML Report**: Interactive browser-based report with Chart.js charts, sortable tables, and dynamic filtering
- **Excel Report**: Traditional spreadsheet with built-in charts, AutoFilter, and easy data manipulation

### Main Class: XTMReportGenerator

**Core Methods:**

- `get_project_statistics(project_id, excluded_users)`: Fetches per-user statistics and filters out excluded users
- `aggregate_monthly_data(start_month, end_month)`: Aggregates all project data for the reporting period, summing words from filtered users
- `aggregate_ytd_data(start_month, end_month, current_month_data)`: Aggregates YTD data, uses JSON cache for past months
- `create_combined_html_report(monthly_data, ytd_monthly_breakdown, output_path)`: Creates interactive HTML report with embedded and dynamic charts
- `create_excel_report(monthly_data, ytd_monthly_breakdown, output_path)`: Creates Excel workbook with monthly and YTD sheets
- `_create_monthly_sheet(wb, monthly_data)`: Creates Excel monthly sheet with bar charts
- `_create_ytd_sheet(wb, ytd_monthly_breakdown)`: Creates Excel YTD sheet with line charts
- `send_email_via_outlook(html_path, excel_path, monthly_data, ytd_data)`: Handles email via Outlook with both attachments and fallback to Apple Mail
- `_ensure_outlook_running()`: Launches Outlook if not running (for --auto-send mode)

**Data Flow:**

1. Load config from `xtm_config.json`
2. For each project, call `/projects/{id}/statistics` (with parallel API calls for speed)
3. Filter out excluded users from `usersStatistics` array
4. Aggregate word counts by language and workflow step
5. Convert locale codes to language names
6. Generate HTML report:
   - Interactive charts using Chart.js (with matplotlib static fallbacks)
   - Sortable and filterable tables
   - Responsive design for mobile and desktop
7. Generate Excel report:
   - Monthly sheet: Current month summary with bar charts
   - YTD sheet: Monthly breakdown with line charts showing trends
   - AutoFilter enabled on all sheets
8. Launch Outlook (if needed) and create/send email with both attachments
9. Save JSON cache for future YTD queries

### Report Structure

**HTML Report Features:**
- Interactive Chart.js visualizations (with matplotlib PNG fallbacks)
- Sortable tables (click column headers)
- Dynamic filtering by language and user
- Search boxes for quick filtering
- Responsive design for mobile and desktop
- Two main sections: Monthly and Year-to-Date
- Each section includes language and user breakdowns

**Excel Workbook Sheets:**

1. **Monthly - YYYY-MM**: Current month data
   - Column headers: Language, translate, correct, final review, Total
   - Data rows sorted by total words (descending)
   - AutoFilter enabled on all columns
   - Bar chart showing total words per language

2. **YTD - YYYY-MM to YYYY-MM**: Year-to-Date monthly breakdown
   - Column headers: Language, Month1, Month2, ..., Total
   - Shows words processed per language per month
   - Data rows sorted by total words (descending)
   - AutoFilter enabled on all columns
   - Line chart showing monthly trends for top 10 languages

3. **User Stats - YYYY-MM**: Per-user breakdown for current month
   - Column headers: User, Language, workflow steps, Total
   - Data rows sorted by total words (descending)
   - AutoFilter enabled on all columns
   - Bar chart showing top 20 users

4. **User Stats - YTD**: Per-user cumulative totals for year-to-date
   - Column headers: User, Language, Month1, Month2, ..., Total
   - Shows words processed per user per month
   - Data rows sorted by total words (descending)
   - AutoFilter enabled on all columns
   - Line chart showing trends for top 10 users

### Email Automation (macOS)

**Dual Attachments**: Both HTML and Excel reports are attached to the email.

**Outlook Priority**: Always tries Microsoft Outlook first, falls back to Apple Mail if Outlook unavailable.

**Auto-Launch**: In `--auto-send` mode, script detects if Outlook is running and launches it automatically if needed (waits up to 30 seconds for startup).

**AppleScript**: Uses AppleScript to control Outlook/Mail, creating messages with recipients and both attachments.

**Draft vs Send**: Without `--auto-send`, creates draft for review. With `--auto-send`, sends immediately.

## Configuration Files

**xtm_config.json**: Contains API credentials, OneDrive path, and email recipients. The `auth_token` is sensitive.

**com.xtm.monthlyreport.plist**: LaunchAgent configuration for monthly scheduling (1st of month at 9:00 AM).

**com.xtm.weeklyreport.plist**: LaunchAgent configuration for weekly scheduling (every Monday at 9:00 AM).

**com.xtm.snapshot.plist**: LaunchAgent configuration for the daily per-user statistics snapshot (every day at 6:00 PM). Logs to `xtm_snapshot.log` / `xtm_snapshot_error.log`.

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
        excluded_users = ["leo.chang@familysearch.org", "LeoAdmin", "Robert.Sena@churchofjesuschrist.org", "MartinADMIN", "Tester BSP BSP"]
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

**Monthly Reports**: The script reports on the **previous complete month**. When run on February 1st, it generates January's complete data. When run on March 1st, it generates February's complete data. Month is calculated in `__init__()`:
```python
first_day_current_month = self.report_date.replace(day=1)
last_day_previous_month = first_day_current_month - timedelta(days=1)
self.report_month = last_day_previous_month.strftime('%Y-%m')
```

The Year-to-Date report covers January through the end of the previous month.

**Weekly Reports**: When using `--weekly` flag, the script reports on the **previous 7 days**. When run on Monday, it covers the previous 7 days (ending yesterday). Calculated in `__init__()`:
```python
end_date = self.report_date - timedelta(days=1)  # Yesterday
start_date = end_date - timedelta(days=6)  # 7 days total including end_date
```

Weekly reports include:
- HTML report with charts and filterable tables
- Excel workbook with weekly data sheet and user statistics
- Email subject: "XTM Weekly Report - Week of YYYY-MM-DD"
- File naming: `XTM_Weekly_Report_YYYY-MM-DD_<generation_date>.html/xlsx`

Weekly reports do NOT include YTD data (only the weekly period).

### Email Body

Email content is defined in `send_email_via_outlook()`. 
- **Monthly**: Includes monthly summary, YTD summary, and top 3 languages for each period
- **Weekly**: Includes only weekly summary and top 3 languages for the week

## Logs

**Monthly Reports:**
- **xtm_report.log**: All operations (INFO level and above)
- **xtm_report_error.log**: Errors only (from LaunchAgent stderr)

**Weekly Reports:**
- **xtm_weekly_report.log**: All operations (INFO level and above)
- **xtm_weekly_report_error.log**: Errors only (from LaunchAgent stderr)

All log files are in the project root directory.

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

## Troubleshooting

### Automation Not Running

```bash
# Check if LaunchAgent is loaded
launchctl list | grep xtm

# If not loaded, reload it
./setup_schedule.sh

# Check logs for errors
tail -50 xtm_report.log
tail -50 xtm_report_error.log
```

### API Failures

The system automatically retries API failures, but if they persist:
- Check `xtm_config.json` auth_token is valid
- Verify XTM Cloud API is accessible: `curl -H "Authorization: XTM-Basic <token>" https://your-instance.xtm-intl.com/rest-api/projects`
- Check `xtm_report.log` for `xtm-trace-id` values to share with XTM support

### Email Not Sending

The system tries multiple email methods automatically:
1. Microsoft Outlook (preferred)
2. Apple Mail (fallback)
3. Saves report locally (last resort)

If email consistently fails, check:
- Outlook/Mail is properly configured with your account
- Recipients in `xtm_config.json` are valid email addresses
- macOS permissions allow the script to control Mail/Outlook

### Reports Not Saving

The system tries multiple save locations:
1. OneDrive path from config (preferred)
2. Desktop (fallback)
3. Current working directory (fallback)
4. Temp directory (last resort)

If all fail, check disk space and permissions.

### No Data in Reports

Check:
- Projects exist in XTM Cloud for the reporting period
- Projects have `lastCompletionDate` set (work was completed)
- Excluded users list in `generate_report.py` isn't filtering out all users
- Date range is correct (reports on previous complete month)
