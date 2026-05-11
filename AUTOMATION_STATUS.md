# XTM Report Automation Status

## ✅ Active Automations

### 1. Monthly Report
- **Schedule**: 1st of every month at 9:00 AM
- **LaunchAgent**: `com.xtm.monthlyreport`
- **Command**: `python3 generate_report.py --auto-send`
- **Coverage**: Previous complete month + Year-to-Date
- **Logs**: 
  - `xtm_report.log`
  - `xtm_report_error.log`
- **Status**: ✓ Loaded and Active

### 2. Weekly Report (NEW!)
- **Schedule**: Every Monday at 9:00 AM
- **LaunchAgent**: `com.xtm.weeklyreport`
- **Command**: `python3 generate_report.py --weekly --auto-send`
- **Coverage**: Previous 7 days
- **Logs**:
  - `xtm_weekly_report.log`
  - `xtm_weekly_report_error.log`
- **Status**: ✓ Loaded and Active

## Check Automation Status

```bash
# View all XTM automations
launchctl list | grep xtm

# Should show:
# -	0	com.xtm.weeklyreport
# -	0	com.xtm.monthlyreport
```

## View Logs

```bash
# Weekly report logs
tail -f xtm_weekly_report.log

# Monthly report logs
tail -f xtm_report.log

# Check for errors
tail -50 xtm_weekly_report_error.log
tail -50 xtm_report_error.log
```

## Manual Testing

```bash
# Test weekly report (draft email)
python3 generate_report.py --weekly

# Test weekly report (auto-send)
python3 generate_report.py --weekly --auto-send

# Test monthly report (draft email)
python3 generate_report.py

# Test monthly report (auto-send)
python3 generate_report.py --auto-send
```

## Disable/Enable Automation

```bash
# Disable weekly automation
launchctl unload ~/Library/LaunchAgents/com.xtm.weeklyreport.plist

# Re-enable weekly automation
launchctl load ~/Library/LaunchAgents/com.xtm.weeklyreport.plist

# Disable monthly automation
launchctl unload ~/Library/LaunchAgents/com.xtm.monthlyreport.plist

# Re-enable monthly automation
launchctl load ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

## Report Output

All reports are saved to:
`/Users/jayjay5032/Library/CloudStorage/OneDrive-ChurchofJesusChrist/XTM_Reports/`

### Monthly Reports
- File pattern: `XTM_Report_2026-04_20260501.html/xlsx`
- Includes: Monthly data + YTD data + User statistics

### Weekly Reports
- File pattern: `XTM_Weekly_Report_2026-05-05_20260512.html/xlsx`
- Includes: Weekly data + User statistics

## Email Recipients

Both reports go to the same recipients configured in `xtm_config.json`:
```json
"email_recipients": [
    "recipient1@example.com",
    "recipient2@example.com"
]
```

## Next Monday (First Weekly Report)

Your first automated weekly report will be sent on:
**Monday, May 12, 2026 at 9:00 AM**

It will cover: **May 5-11, 2026** (previous 7 days)

## Troubleshooting

If reports don't send automatically:

1. **Check LaunchAgent is loaded**:
   ```bash
   launchctl list | grep weeklyreport
   ```

2. **Check logs for errors**:
   ```bash
   tail -50 xtm_weekly_report_error.log
   ```

3. **Test manually**:
   ```bash
   python3 generate_report.py --weekly --auto-send
   ```

4. **Reload LaunchAgent**:
   ```bash
   launchctl unload ~/Library/LaunchAgents/com.xtm.weeklyreport.plist
   launchctl load ~/Library/LaunchAgents/com.xtm.weeklyreport.plist
   ```

## Summary

✅ Weekly reports are now configured and will run automatically every Monday
✅ Monthly reports continue to run as before on the 1st of each month
✅ Both automations are independent and won't interfere with each other
✅ All resilience features (retry, fallback, health checks) are active for both
