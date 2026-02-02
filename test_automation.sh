#!/bin/bash

# Test script for XTM Monthly Report automation

echo "=== XTM Monthly Report - Automation Test ==="
echo ""

# Check if Outlook is installed
if [ ! -d "/Applications/Microsoft Outlook.app" ]; then
    echo "❌ Microsoft Outlook is not installed"
    exit 1
fi
echo "✓ Microsoft Outlook is installed"

# Check if Python 3 is available
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed"
    exit 1
fi
echo "✓ Python 3 is available"

# Check if required Python packages are installed
echo ""
echo "Checking Python packages..."
python3 -c "import requests, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "❌ Required Python packages are missing"
    echo "Run: pip install -r requirements.txt"
    exit 1
fi
echo "✓ Required Python packages are installed"

# Check if config file exists
if [ ! -f "xtm_config.json" ]; then
    echo "❌ xtm_config.json not found"
    exit 1
fi
echo "✓ Configuration file exists"

# Check if OneDrive folder exists
ONEDRIVE_PATH=$(python3 -c "import json; print(json.load(open('xtm_config.json'))['onedrive_path'])")
if [ ! -d "$ONEDRIVE_PATH" ]; then
    echo "❌ OneDrive path does not exist: $ONEDRIVE_PATH"
    exit 1
fi
echo "✓ OneDrive path exists"

echo ""
echo "=== Running Test (Draft Mode) ==="
echo "This will create a draft email without sending..."
echo ""

python3 generate_report.py

if [ $? -eq 0 ]; then
    echo ""
    echo "✓ Test completed successfully!"
    echo ""
    echo "Next steps:"
    echo "1. Review the draft email in Outlook/Mail"
    echo "2. If it looks good, run: ./setup_schedule.sh"
    echo "3. Or test auto-send with: python3 generate_report.py --auto-send"
else
    echo ""
    echo "❌ Test failed. Check the log files for details:"
    echo "   - xtm_report.log"
    echo "   - xtm_report_error.log"
    exit 1
fi
