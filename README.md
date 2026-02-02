# XTM Monthly Report Generator

Automated monthly reporting tool for XTM Cloud translation management system.

## Prerequisites

- Python 3.7 or higher
- macOS with either Microsoft Outlook or Apple Mail installed (for email functionality)
- OneDrive configured and synced to the path specified in `xtm_config.json`

## Installation

1. Install required Python packages:

```bash
pip install -r requirements.txt
```

## Configuration

The `xtm_config.json` file contains all necessary configuration:

- **base_url**: XTM API endpoint
- **auth_token**: API authentication token
- **auth_type**: Authentication method
- **onedrive_path**: Local path to OneDrive folder for report storage
- **email_recipients**: List of email addresses to receive reports

## Usage

### Generate Monthly Report

Run the script to generate the current month's report:

```bash
python generate_report.py
```

The script will:
1. Connect to XTM Cloud API
2. Retrieve project data, workflow metrics, and translation volumes
3. Generate an Excel report with multiple sheets
4. Save the report to your OneDrive folder
5. Open a draft email in Microsoft Outlook (or Apple Mail as fallback) with the report attached

### Review Before Sending

The script creates a draft email without sending it automatically. This allows you to:
- Review the report content
- Verify recipients
- Add additional context or notes
- Send manually when ready

**Note**: The script will try Microsoft Outlook first, then fall back to Apple Mail if Outlook is not available.

## Report Contents

The generated Excel report includes four sheets:

1. **Summary**: Overall project statistics and key metrics
2. **Translation Volume**: Translation volume broken down by language pair
3. **Workflow Metrics**: Performance metrics for each workflow step, including total words processed
4. **Project Details**: Complete listing of all projects with status and details

## Output

Reports are saved with the naming convention:
```
XTM_Monthly_Report_YYYY-MM_YYYYMMDD.xlsx
```

Example: `XTM_Monthly_Report_2024-01_20240205.xlsx`

## Logging

The script logs all activities to `xtm_report.log` for troubleshooting and audit purposes.

## Troubleshooting

### Common Issues

**API Authentication Errors**:
- Verify the auth_token in `xtm_config.json` is valid
- Check that the XTM API base_url is correct

**Email Client Not Opening**:
- Ensure Microsoft Outlook or Apple Mail is installed on macOS
- Grant necessary permissions when prompted (System Preferences > Security & Privacy > Automation)
- If AppleScript is blocked, manually enable it in System Preferences
- The script will automatically try both Outlook and Apple Mail

**OneDrive Path Issues**:
- Verify the onedrive_path in configuration exists
- Ensure you have write permissions to the folder
- OneDrive must be synced and the path must exist locally

**No Data Retrieved**:
- Check XTM API connectivity
- Verify your API token has appropriate permissions
- Review `xtm_report.log` for detailed error messages

**macOS Security Prompts**:
- First run may prompt for permission to control Outlook/Mail
- Click "OK" or "Allow" when prompted
- If denied, go to System Preferences > Security & Privacy > Privacy > Automation and enable access

## Scheduling (Optional)

To run this report automatically on a monthly schedule on macOS:

### macOS Cron
Add to crontab with `crontab -e`:
```bash
# Run on the 1st of every month at 9 AM
0 9 1 * * cd /Users/jayjay5032/Desktop/MartinMonthlyReport && /usr/bin/python3 generate_report.py
```

### macOS LaunchAgent (Recommended)
Create a LaunchAgent plist file at `~/Library/LaunchAgents/com.xtm.monthlyreport.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.xtm.monthlyreport</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>/Users/jayjay5032/Desktop/MartinMonthlyReport/generate_report.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>/Users/jayjay5032/Desktop/MartinMonthlyReport</string>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Day</key>
        <integer>1</integer>
        <key>Hour</key>
        <integer>9</integer>
        <key>Minute</key>
        <integer>0</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>/Users/jayjay5032/Desktop/MartinMonthlyReport/xtm_report.log</string>
    <key>StandardErrorPath</key>
    <string>/Users/jayjay5032/Desktop/MartinMonthlyReport/xtm_report_error.log</string>
</dict>
</plist>
```

Then load it:
```bash
launchctl load ~/Library/LaunchAgents/com.xtm.monthlyreport.plist
```

## Security Notes

- The `xtm_config.json` file contains sensitive authentication credentials
- Never commit this file to public repositories
- Restrict file permissions appropriately
- Rotate API tokens regularly according to security policies
