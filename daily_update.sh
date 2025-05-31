#!/bin/bash
# Daily script to update Razorpay payment links in Google Sheets

# Change to the directory where the script is located
cd "$(dirname "$0")"

# Load environment variables from .env file if it exists
if [ -f .env ]; then
    export $(grep -v '^#' .env | xargs)
fi

# Run the Python script
python3 razorpay_to_sheets.py

# Log the execution
echo "$(date): Razorpay to Google Sheets update completed with exit code $?" >> update_log.txt 