# Weekly Reports Feature

This document describes the weekly reporting functionality added to the XTM Report Generator.

## Overview

In addition to monthly reports, the system now supports weekly reports that cover the previous 7 days. Weekly reports are automatically sent every Monday at 9:00 AM.

## Key Features

### Date Range
- **Coverage**: Previous 7 days ending yesterday
- **Example**: When run on Monday, May 12, 2026, covers May 5-11, 2026
- **Week Label**: "Week of 2026-05-05" (using the start date)

### Report Contents

**HTML Report**:
- Bar charts showing translation volume by language and workflow step
- User productivity bar charts
- Sortable and filterable tables
- No YTD data (weekly only)

**Excel Report**:
- Weekly data sheet with workflow breakdown by language
- User statistics sheet with per-user, per-language details
- Built-in bar charts
- AutoFilter enabled on all sheets
- No YTD sheets (weekly only)

### File Naming

- HTML: `XTM_Weekly_Report_2026-05-05_20260512.html`
- Excel: `XTM_Weekly_Report_2026-05-05_20260512.xlsx`

### Email

- **Subject**: "XTM Weekly Report - Week of 2026-05-05"
- **Content**: Weekly summary with total words and top languages
- **Attachments**: Both HTML and Excel reports
- **Recipients**: Same as configured in `xtm_config.json`

## Usage

### Manual Generation

```bash
# Generate weekly report with draft email (for review)
python3 generate_report.py --weekly

# Generate and automatically send via Outlook
python3 generate_report.py --weekly --auto-send
```

### Automated Weekly Reports

```bash
# Set up weekly automation (runs every Monday at 9:00 AM)
./setup_weekly_schedule.sh

# Test the setup
./test_weekly_report.sh

# Check if running
launchctl list | grep xtm.weeklyreport

# View logs
tail -f xtm_weekly_report.log
```

### Disable Weekly Automation

```bash
launchctl unload ~/Library/LaunchAgents/com.xtm.weeklyreport.plist
```

## Implementation Details

### Code Changes

The weekly functionality was added to the existing `generate_report.py` by:

1. **New `--weekly` flag**: Switches between monthly and weekly mode
2. **Date calculation**: Computes previous 7 days instead of previous month
3. **New method**: `aggregate_weekly_data()` aggregates data by exact date range
4. **Conditional logic**: HTML/Excel generation adapts based on mode
5. **Email templates**: Different subject and body for weekly vs monthly

### Configuration Files

- **com.xtm.weeklyreport.plist**: LaunchAgent with Weekday=1 (Monday)
- **xtm_weekly_report.log**: Separate log file for weekly reports
- **xtm_weekly_report_error.log**: Error log for weekly automation

### Resilience

Weekly reports inherit all the resilience features from monthly reports:
- API retry logic with exponential backoff
- Health checks before generation
- Graceful degradation on partial failures
- Multiple fallback save locations
- Email fallback (Outlook → Apple Mail → Save locally)

## Logs

Weekly reports write to separate log files:
- `xtm_weekly_report.log` - All operations
- `xtm_weekly_report_error.log` - Errors only

View recent logs:
```bash
tail -50 xtm_weekly_report.log
```

## Troubleshooting

### Weekly automation not running

```bash
# Check if loaded
launchctl list | grep weeklyreport

# If not loaded, reload
./setup_weekly_schedule.sh

# Check for errors
tail -50 xtm_weekly_report_error.log
```

### Testing before Monday

Weekly reports can be run any day - they always cover the previous 7 days:

```bash
python3 generate_report.py --weekly
```

### No data in weekly report

- Check that projects had work completed in the previous 7 days
- Verify excluded users list isn't filtering out all work
- Check API connectivity: `python3 debug_api.py`

## Comparison: Weekly vs Monthly

| Feature | Monthly | Weekly |
|---------|---------|--------|
| Date Range | Previous complete month | Previous 7 days |
| YTD Data | Yes | No |
| Schedule | 1st of month, 9:00 AM | Every Monday, 9:00 AM |
| File Naming | `XTM_Report_2026-04_...` | `XTM_Weekly_Report_2026-05-05_...` |
| Email Subject | "XTM Monthly Report - 2026-04" | "XTM Weekly Report - Week of 2026-05-05" |
| Excel Sheets | 4 (Monthly, YTD, User Monthly, User YTD) | 2 (Weekly, User Stats) |
| HTML Sections | 2 (Monthly + YTD) | 1 (Weekly only) |
| Log Files | `xtm_report.log` | `xtm_weekly_report.log` |
| LaunchAgent | `com.xtm.monthlyreport.plist` | `com.xtm.weeklyreport.plist` |

## Future Enhancements

Potential improvements:
- Custom date ranges via command-line arguments
- Week-over-week comparison trends
- Configurable week start day (currently ends on yesterday)
- Separate email recipient lists for weekly vs monthly
