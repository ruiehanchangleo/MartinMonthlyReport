# Weekly Report Setup - Quick Start

## What Was Added

Weekly reporting capability has been added to your XTM Report Generator. This allows you to generate reports for the previous 7 days automatically every Monday.

## Quick Setup (3 Steps)

### 1. Test the Weekly Report Manually

```bash
cd /Users/jayjay5032/MartinMonthlyReport
python3 generate_report.py --weekly
```

This will generate a draft email with reports for the previous 7 days. Review the output to make sure it looks correct.

### 2. Set Up Automatic Weekly Reports

```bash
./setup_weekly_schedule.sh
```

This installs a LaunchAgent that runs every Monday at 9:00 AM and automatically sends the weekly report via Outlook.

### 3. Verify It's Running

```bash
launchctl list | grep xtm
```

You should see both:
- `com.xtm.monthlyreport` (1st of month at 9:00 AM)
- `com.xtm.weeklyreport` (every Monday at 9:00 AM)

## What You Get

### Weekly Report (Every Monday at 9:00 AM)

- **Date Range**: Previous 7 days (ending yesterday)
- **Example**: Monday, May 12, 2026 → covers May 5-11, 2026
- **Reports Generated**:
  - HTML report with interactive charts and tables
  - Excel workbook with weekly data and user statistics
- **Email**: Automatically sent via Outlook to configured recipients
- **File Naming**: `XTM_Weekly_Report_2026-05-05_<date>.html/xlsx`

### Monthly Report (Unchanged)

- **Date Range**: Previous complete month + YTD
- **Schedule**: 1st of month at 9:00 AM
- **Everything else works exactly as before**

## Files Added

- `com.xtm.weeklyreport.plist` - LaunchAgent configuration
- `setup_weekly_schedule.sh` - Setup script
- `test_weekly_report.sh` - Testing script
- `WEEKLY_REPORTS.md` - Detailed documentation
- `WEEKLY_SETUP_SUMMARY.md` - This file

## Files Modified

- `generate_report.py` - Added `--weekly` flag and weekly logic
- `CLAUDE.md` - Updated documentation

## Logs

Weekly reports write to separate log files (so they don't mix with monthly logs):
- `xtm_weekly_report.log` - All operations
- `xtm_weekly_report_error.log` - Errors only

View logs:
```bash
tail -f xtm_weekly_report.log
```

## Common Commands

```bash
# Manual generation (draft email)
python3 generate_report.py --weekly

# Manual generation (auto-send)
python3 generate_report.py --weekly --auto-send

# Check automation status
launchctl list | grep xtm

# View recent weekly logs
tail -50 xtm_weekly_report.log

# Disable weekly automation
launchctl unload ~/Library/LaunchAgents/com.xtm.weeklyreport.plist

# Re-enable weekly automation
launchctl load ~/Library/LaunchAgents/com.xtm.weeklyreport.plist
```

## Testing

To test without waiting for Monday:

```bash
# Test weekly report generation
./test_weekly_report.sh

# Or manually
python3 generate_report.py --weekly
```

The weekly report can be run on any day - it always covers the previous 7 days.

## Troubleshooting

### No email sent?
- Check Outlook is running: `ps aux | grep Outlook`
- Check logs: `tail -50 xtm_weekly_report_error.log`
- Recipients configured: check `email_recipients` in `xtm_config.json`

### No data in report?
- Check projects have completed work in the last 7 days
- Verify API connectivity: `python3 debug_api.py`
- Check excluded users aren't filtering everything

### Automation not running?
- Verify LaunchAgent is loaded: `launchctl list | grep weeklyreport`
- Reload if needed: `./setup_weekly_schedule.sh`
- Check system logs: `tail -50 xtm_weekly_report_error.log`

## Support

For detailed information, see:
- `WEEKLY_REPORTS.md` - Complete feature documentation
- `CLAUDE.md` - Full system documentation
- Logs: `xtm_weekly_report.log`
