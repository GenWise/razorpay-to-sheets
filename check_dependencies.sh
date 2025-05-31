#!/bin/bash
# Check if all dependencies are properly installed

echo "Checking Python dependencies..."
python3 -c "import requests, gspread, google.oauth2.service_account, dotenv" 2>/dev/null

if [ $? -ne 0 ]; then
    echo "Error: Some dependencies are missing. Installing required packages..."
    pip install requests gspread google-auth google-auth-oauthlib python-dotenv
else
    echo "All Python dependencies are installed correctly!"
fi

echo "Checking for .env file..."
if [ -f .env ]; then
    echo "Found .env file."
else
    echo "Warning: .env file not found. Creating from template..."
    cp env.example .env
    echo "Please edit .env file with your Razorpay API keys and Google Sheet ID."
fi

echo "Checking for service account JSON file..."
SERVICE_ACCOUNT_FILE=$(grep GOOGLE_SERVICE_ACCOUNT_FILE .env | cut -d= -f2)
if [ -f "$SERVICE_ACCOUNT_FILE" ]; then
    echo "Found service account file: $SERVICE_ACCOUNT_FILE"
else
    echo "Warning: Service account file not found."
    echo "Please download your Google Service Account JSON file and save it as $SERVICE_ACCOUNT_FILE"
    echo "Visit: https://console.cloud.google.com/iam-admin/serviceaccounts"
fi

echo ""
echo "Dependency check complete. Fix any warnings above before running the main script." 