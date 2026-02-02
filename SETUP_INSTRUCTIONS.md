# Automatic Monthly Report Setup Instructions

## Overview
This guide will help you set up the XTM monthly report to automatically generate and send on the 1st of each month at 9:00 AM.

## Prerequisites

1. **Microsoft Outlook** must be installed and configured on your Mac
2. **Python 3** and required packages must be installed
3. **Outlook must be running** when the scheduled task executes

## Setup Steps

### Option 1: Automatic Setup (Recommended)

Simply run the setup script:

```bash
cd /Users/jayjay5032/Desktop/MartinMonthlyReport
./setup_schedule.sh
```

This will:
- Install the LaunchAgent configuration
- Schedule the report to run on the 1st of each month at 9:00 AM
- Enable automatic email sending

### Option 2: Manual Setup

If you prefer manual setup:

1. Copy the plist file to LaunchAgents:
```bash
cp com.xtm.monthlyreport.plist ~/Library/LaunchAgents/
```

2. Load the LaunchAgent:
```bash
launchctl load ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

## Important Notes

### Outlook Auto-Launch
- **The script will automatically launch Outlook** if it's not running
- For best results, you can still configure Outlook to start at login:
  - System Preferences > Users & Groups > Login Items
  - Add Microsoft Outlook to the list (optional)
- The script will wait up to 30 seconds for Outlook to start before proceeding

### Security Permissions
- First run will prompt for automation permissions
- Allow Terminal/Python to control Microsoft Outlook:
  - System Preferences > Security & Privacy > Privacy > Automation
  - Enable access for Terminal or Python

### Testing

Test the automatic sending manually:
```bash
cd /Users/jayjay5032/Desktop/MartinMonthlyReport
python3 generate_report.py --auto-send
```

This will generate and send the report immediately.

To test without sending (draft only):
```bash
python3 generate_report.py
```

## Schedule Details

- **Runs:** 1st day of every month
- **Time:** 9:00 AM
- **Action:** Generates report and sends automatically via Outlook
- **Logs:** Written to `xtm_report.log` and `xtm_report_error.log`

## Management Commands

### Check if the task is running:
```bash
launchctl list | grep xtm
```

### View logs:
```bash
tail -f /Users/jayjay5032/Desktop/MartinMonthlyReport/xtm_report.log
```

### Disable automatic sending:
```bash
launchctl unload ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

### Re-enable automatic sending:
```bash
launchctl load ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

### Change schedule time:
Edit the plist file and reload:
```bash
nano ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
# Change Hour and Minute values
launchctl unload ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
launchctl load ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

## Troubleshooting

### Email not sending
1. Check Outlook is running
2. Check automation permissions (System Preferences > Security & Privacy)
3. Review error log: `cat xtm_report_error.log`
4. Test manually with: `python3 generate_report.py --auto-send`

### Task not running on schedule
1. Check if loaded: `launchctl list | grep xtm`
2. Check system logs: `log show --predicate 'process == "launchd"' --last 1h | grep xtm`
3. Ensure computer is awake at 9:00 AM on the 1st

### Computer is asleep at 9:00 AM
- The task will run when the computer wakes up
- Alternatively, adjust the time in the plist file to when you're typically working

## Changing Report Month

By default, the script reports on the **current month**. To change it to report on the **previous month** (more typical for monthly reports):

Edit `generate_report.py` lines 62-66:

```python
# Change from (current month):
self.report_month = self.report_date.strftime('%Y-%m')

# To (previous month):
first_day_current_month = self.report_date.replace(day=1)
last_day_previous_month = first_day_current_month - timedelta(days=1)
self.report_month = last_day_previous_month.strftime('%Y-%m')
self.report_month_name = last_day_previous_month.strftime('%B %Y')
```

## Support

If you encounter issues:
1. Check the log files in the project directory
2. Test manual execution first
3. Verify all prerequisites are met
4. Ensure Outlook is running and configured
