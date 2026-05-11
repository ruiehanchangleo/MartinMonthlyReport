#!/bin/bash
# Test script for weekly report generation

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

echo "Testing XTM Weekly Report Generation"
echo "====================================="
echo ""
echo "This will generate a weekly report for the previous 7 days"
echo "and create a draft email (not auto-send)."
echo ""

cd "${SCRIPT_DIR}"

# Run the report generator in weekly mode without auto-send
python3 generate_report.py --weekly

echo ""
echo "Test complete. Check the output above for any errors."
echo ""
echo "To test with auto-send:"
echo "  python3 generate_report.py --weekly --auto-send"
