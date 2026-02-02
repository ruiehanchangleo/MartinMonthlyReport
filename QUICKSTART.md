# Quick Start - Automatic Monthly Report

## 🚀 Setup in 3 Steps

### 1. Test the Report (Manual)
```bash
cd /Users/jayjay5032/Desktop/MartinMonthlyReport
python3 generate_report.py
```
This creates a draft email for you to review.

### 2. Test Auto-Send
```bash
python3 generate_report.py --auto-send
```
This will:
- ✅ Automatically launch Outlook if not running
- ✅ Generate the report
- ✅ Send the email to all recipients

### 3. Enable Monthly Automation
```bash
./setup_schedule.sh
```
This schedules the report to automatically generate and send on the **1st of every month at 9:00 AM**.

## ✅ What's Included

Your automated report includes:
- **Monthly Sheet**: Current month workflow data by language
- **Year-to-Date Sheet**: Cumulative data from January to current month
- **Bar Charts**: Visual representation of total words per language
- **Auto-send**: Emails sent automatically via Outlook

## 🔧 How It Works

1. **Outlook Auto-Launch**: Script detects if Outlook is running and launches it automatically if needed
2. **Smart Fallback**: If Outlook fails, it falls back to Apple Mail
3. **Reliable Scheduling**: Uses macOS LaunchAgent for scheduling (more reliable than cron)
4. **Logging**: All operations logged to `xtm_report.log` for troubleshooting

## 📅 Schedule Details

- **When**: 1st day of every month
- **Time**: 9:00 AM
- **Action**: Generate report + send email
- **Recipients**: Defined in `xtm_config.json`

## 🛠️ Management Commands

```bash
# Check if automation is running
launchctl list | grep xtm

# View recent logs
tail -f xtm_report.log

# Test manually with auto-send
python3 generate_report.py --auto-send

# Disable automation
launchctl unload ~/Library/LaunchAgents/com.xtm.monthlyreport.plist

# Re-enable automation
launchctl load ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

## ⚠️ Important Notes

1. **Your Mac must be awake** at 9:00 AM on the 1st for the task to run
   - If asleep, it will run when the computer wakes up
   - Adjust the time in `com.xtm.monthlyreport.plist` if needed

2. **First run permissions**
   - You'll be prompted to allow Terminal/Python to control Outlook
   - Click "Allow" or "OK" when prompted

3. **Recipients**
   - Email recipients are defined in `xtm_config.json`
   - Edit that file to add/remove recipients

## 📖 More Information

- **Detailed setup guide**: See `SETUP_INSTRUCTIONS.md`
- **Troubleshooting**: See `README.md`
- **Code documentation**: See `CLAUDE.md`

## 🧪 Testing

Run the automated test script:
```bash
./test_automation.sh
```

This checks:
- ✅ Outlook is installed
- ✅ Python and packages are installed
- ✅ Config file exists
- ✅ OneDrive path is valid
- ✅ Report generates successfully

## 🎯 Next Steps

1. Run `./test_automation.sh` to verify everything works
2. Review the test report
3. Run `./setup_schedule.sh` to enable automation
4. Done! Your reports will be sent automatically every month.

---

**Need help?** Check the log files in this directory or review the detailed documentation in `SETUP_INSTRUCTIONS.md`.
