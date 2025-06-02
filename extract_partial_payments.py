#!/usr/bin/env python3
"""
Extract Partial Payments from Google Sheet
This script extracts all payment links where the amount paid is less than the total amount
"""

import os
import sys
import logging
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("partial_payments.log"),
        logging.StreamHandler()
    ]
)

# Try to import dotenv_loader
try:
    from dotenv_loader import load_env_vars
    # Load environment variables from .env file
    load_env_vars()
except ImportError:
    logging.warning("dotenv_loader.py not found. Make sure environment variables are set manually.")

# Google Sheets Configuration
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
OUTPUT_FILE = "partial_payments.csv"

def connect_to_sheet():
    """Connect to Google Sheet and return the worksheet"""
    if not GOOGLE_SHEET_ID:
        raise ValueError("Google Sheet ID not found in environment variables")
    
    # Check if service account file exists
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account file '{SERVICE_ACCOUNT_FILE}' not found. Please download it from Google Cloud Console.")
    
    # Set up Google Sheets authentication
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes
    )
    
    gc = gspread.authorize(credentials)
    
    # Open the spreadsheet and select the first worksheet
    try:
        spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
        worksheet = spreadsheet.sheet1
        logging.info(f"Connected to spreadsheet: {spreadsheet.title}")
        logging.info(f"Using worksheet: {worksheet.title}")
        return worksheet
    except Exception as e:
        raise Exception(f"Failed to connect to Google Sheet: {str(e)}")

def extract_partial_payments(worksheet):
    """Extract payment links where amount paid is less than total amount"""
    try:
        # Get all data from the worksheet
        data = worksheet.get_all_records()
        logging.info(f"Retrieved {len(data)} records from Google Sheet")
        
        # Convert to pandas DataFrame for easier filtering
        df = pd.DataFrame(data)
        
        # Check if the required columns exist
        amount_col = "Amount (₹)"
        paid_col = "Amount Paid (₹)"
        
        # Find the actual column names if they don't match exactly
        if amount_col not in df.columns:
            # Try to find similar column names
            for col in df.columns:
                if "amount" in col.lower() and "paid" not in col.lower():
                    amount_col = col
                    break
        
        if paid_col not in df.columns:
            # Try to find similar column names
            for col in df.columns:
                if "amount" in col.lower() and "paid" in col.lower():
                    paid_col = col
                    break
        
        # Check if we found the columns
        if amount_col not in df.columns or paid_col not in df.columns:
            raise ValueError(f"Could not find amount columns in the sheet. Available columns: {', '.join(df.columns)}")
        
        # Convert amount columns to float
        df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
        df[paid_col] = pd.to_numeric(df[paid_col], errors='coerce')
        
        # Filter for records where amount paid is less than total amount
        partial_payments = df[df[paid_col] < df[amount_col]].copy()
        
        # Add a Due Amount column
        due_col = "Due Amount (₹)"
        partial_payments[due_col] = partial_payments[amount_col] - partial_payments[paid_col]
        
        # Sort by due amount (highest to lowest)
        partial_payments = partial_payments.sort_values(by=due_col, ascending=False)
        
        logging.info(f"Found {len(partial_payments)} payment links with partial payments")
        
        # Select relevant columns
        columns_to_export = [
            "ID", amount_col, paid_col, due_col, 
            "Status", "Short URL", "Reference ID", "Customer Email", "Customer Contact"
        ]
        
        # Filter columns that exist in the DataFrame
        available_columns = [col for col in columns_to_export if col in partial_payments.columns]
        result = partial_payments[available_columns]
        
        return result
    
    except Exception as e:
        logging.error(f"Error extracting partial payments: {str(e)}")
        raise

def main():
    try:
        # Connect to Google Sheet
        worksheet = connect_to_sheet()
        
        # Extract partial payments
        partial_payments = extract_partial_payments(worksheet)
        
        if len(partial_payments) > 0:
            # Export to CSV
            partial_payments.to_csv(OUTPUT_FILE, index=False)
            logging.info(f"Exported {len(partial_payments)} partial payments to {OUTPUT_FILE}")
            
            # Display summary
            print("\nPartial Payments Summary:")
            print(f"Total partial payments: {len(partial_payments)}")
            
            # Get the due amount column name
            due_col = "Due Amount (₹)"
            if due_col in partial_payments.columns:
                total_due = partial_payments[due_col].sum()
                print(f"Total due amount: ₹{total_due:.2f}")
            
            # Display the first 5 records
            print("\nTop 5 partial payments (by due amount):")
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 1000)
            print(partial_payments.head(5).to_string())
            
            print(f"\nFull details exported to {OUTPUT_FILE}")
        else:
            print("No partial payments found.")
        
        return 0
    
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        print(f"Error: {str(e)}")
        return 1

if __name__ == "__main__":
    exit(main()) 