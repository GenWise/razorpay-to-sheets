#!/usr/bin/env python3
"""
Extract Partial Payments from Google Sheet
This script extracts payment links where:
1. Amount paid is less than total amount
2. Status is "created"
Then stores results in a new tab and emails a summary
"""

import os
import sys
import logging
import pandas as pd
import gspread
import smtplib
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
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

# Email Configuration
EMAIL_SENDER = os.environ.get("EMAIL_SENDER")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_RECIPIENT = "rajesh@genwise.in"

def connect_to_sheet():
    """Connect to Google Sheet and return the spreadsheet"""
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
    
    # Open the spreadsheet
    try:
        spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
        logging.info(f"Connected to spreadsheet: {spreadsheet.title}")
        return spreadsheet
    except Exception as e:
        raise Exception(f"Failed to connect to Google Sheet: {str(e)}")

def extract_partial_payments(worksheet):
    """Extract payment links where amount paid is less than total amount and status is 'created'"""
    try:
        # Get all data from the worksheet
        data = worksheet.get_all_records()
        logging.info(f"Retrieved {len(data)} records from Google Sheet")
        
        # Convert to pandas DataFrame for easier filtering
        df = pd.DataFrame(data)
        
        # Check if the required columns exist
        amount_col = "Amount (₹)"
        paid_col = "Amount Paid (₹)"
        status_col = "Status"
        
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
        
        if status_col not in df.columns:
            # Try to find similar column names
            for col in df.columns:
                if "status" in col.lower():
                    status_col = col
                    break
        
        # Check if we found the columns
        required_cols = [amount_col, paid_col, status_col]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Could not find required columns: {', '.join(missing_cols)}. Available columns: {', '.join(df.columns)}")
        
        # Convert amount columns to float
        df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
        df[paid_col] = pd.to_numeric(df[paid_col], errors='coerce')
        
        # Filter for records where:
        # 1. Amount paid is less than total amount
        # 2. Status is "created"
        partial_payments = df[(df[paid_col] < df[amount_col]) & (df[status_col] == "created")].copy()
        
        # Add a Due Amount column
        due_col = "Due Amount (₹)"
        partial_payments[due_col] = partial_payments[amount_col] - partial_payments[paid_col]
        
        # Sort by due amount (highest to lowest)
        partial_payments = partial_payments.sort_values(by=due_col, ascending=False)
        
        logging.info(f"Found {len(partial_payments)} payment links with status 'created' and partial payments")
        
        # Select relevant columns
        columns_to_export = [
            "ID", amount_col, paid_col, due_col, 
            status_col, "Short URL", "Reference ID", "Customer Email", "Customer Contact"
        ]
        
        # Filter columns that exist in the DataFrame
        available_columns = [col for col in columns_to_export if col in partial_payments.columns]
        result = partial_payments[available_columns]
        
        return result
    
    except Exception as e:
        logging.error(f"Error extracting partial payments: {str(e)}")
        raise

def create_or_update_sheet_tab(spreadsheet, data, tab_name="Partial Payments"):
    """Create or update a tab in the Google Sheet with the partial payments data"""
    try:
        # Check if the worksheet already exists
        try:
            worksheet = spreadsheet.worksheet(tab_name)
            # Clear existing content if it exists
            worksheet.clear()
            logging.info(f"Cleared existing '{tab_name}' tab")
        except gspread.exceptions.WorksheetNotFound:
            # Create a new worksheet if it doesn't exist
            worksheet = spreadsheet.add_worksheet(title=tab_name, rows=len(data)+1, cols=len(data.columns))
            logging.info(f"Created new '{tab_name}' tab")
        
        # Update the worksheet with the data
        # First, add the headers
        worksheet.update('A1', [data.columns.tolist()])
        
        # Then add the data
        if not data.empty:
            worksheet.update('A2', data.values.tolist())
        
        logging.info(f"Updated '{tab_name}' tab with {len(data)} rows of data")
        return worksheet
    except Exception as e:
        logging.error(f"Error updating sheet tab: {str(e)}")
        raise

def generate_summary(data):
    """Generate summary of total amount due, split by Reference ID's starting with 'July' and the rest"""
    # Check if Reference ID column exists
    ref_id_col = "Reference ID"
    if ref_id_col not in data.columns:
        # Try to find a similar column
        for col in data.columns:
            if "reference" in col.lower():
                ref_id_col = col
                break
    
    # If we still don't have a Reference ID column, return a simple summary
    if ref_id_col not in data.columns:
        total_due = data["Due Amount (₹)"].sum()
        return {
            "total_due": total_due,
            "july_due": 0,
            "other_due": total_due,
            "july_count": 0,
            "other_count": len(data)
        }
    
    # Split data into July references and others
    july_data = data[data[ref_id_col].astype(str).str.startswith("July")]
    other_data = data[~data[ref_id_col].astype(str).str.startswith("July")]
    
    # Calculate totals
    july_due = july_data["Due Amount (₹)"].sum() if not july_data.empty else 0
    other_due = other_data["Due Amount (₹)"].sum() if not other_data.empty else 0
    total_due = july_due + other_due
    
    return {
        "total_due": total_due,
        "july_due": july_due,
        "other_due": other_due,
        "july_count": len(july_data),
        "other_count": len(other_data)
    }

def send_email_summary(summary, sheet_url):
    """Send email with summary of partial payments"""
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        logging.warning("Email sender or password not found in environment variables. Skipping email.")
        return
    
    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECIPIENT
        msg['Subject'] = f"Partial Payments Summary - {datetime.datetime.now().strftime('%Y-%m-%d')}"
        
        # Email body
        body = f"""
        <html>
        <body>
            <h2>Partial Payments Summary</h2>
            <p>Here's a summary of payment links with status "created" and partial payments:</p>
            
            <table border="1" cellpadding="5" cellspacing="0">
                <tr>
                    <td><b>Category</b></td>
                    <td><b>Count</b></td>
                    <td><b>Amount Due (₹)</b></td>
                </tr>
                <tr>
                    <td>July Reference IDs</td>
                    <td>{summary['july_count']}</td>
                    <td>₹{summary['july_due']:,.2f}</td>
                </tr>
                <tr>
                    <td>Other Reference IDs</td>
                    <td>{summary['other_count']}</td>
                    <td>₹{summary['other_due']:,.2f}</td>
                </tr>
                <tr>
                    <td><b>Total</b></td>
                    <td><b>{summary['july_count'] + summary['other_count']}</b></td>
                    <td><b>₹{summary['total_due']:,.2f}</b></td>
                </tr>
            </table>
            
            <p>For complete details, view the <a href="{sheet_url}">Google Sheet</a>.</p>
            
            <p>This is an automated message from the Razorpay Payment Links Tracker.</p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'html'))
        
        # Connect to Gmail SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        
        # Send the email
        server.send_message(msg)
        server.quit()
        
        logging.info(f"Email summary sent to {EMAIL_RECIPIENT}")
    
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")
        print(f"Error sending email: {str(e)}")

def main():
    try:
        # Connect to Google Sheet
        spreadsheet = connect_to_sheet()
        worksheet = spreadsheet.sheet1
        
        # Extract partial payments with status "created"
        partial_payments = extract_partial_payments(worksheet)
        
        if len(partial_payments) > 0:
            # Create or update the "Partial Payments" tab
            create_or_update_sheet_tab(spreadsheet, partial_payments)
            
            # Generate summary
            summary = generate_summary(partial_payments)
            
            # Get the Google Sheet URL
            sheet_url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}"
            
            # Send email summary
            send_email_summary(summary, sheet_url)
            
            # Export to CSV (optional)
            partial_payments.to_csv(OUTPUT_FILE, index=False)
            logging.info(f"Exported {len(partial_payments)} partial payments to {OUTPUT_FILE}")
            
            # Display summary
            print("\nPartial Payments Summary:")
            print(f"Total partial payments: {len(partial_payments)}")
            print(f"Total due amount: ₹{summary['total_due']:,.2f}")
            print(f"July references: {summary['july_count']} items, ₹{summary['july_due']:,.2f}")
            print(f"Other references: {summary['other_count']} items, ₹{summary['other_due']:,.2f}")
            
            # Display the first 5 records
            print("\nTop 5 partial payments (by due amount):")
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 1000)
            print(partial_payments.head(5).to_string())
            
            print(f"\nFull details exported to Google Sheet tab 'Partial Payments'")
            print(f"Email summary sent to {EMAIL_RECIPIENT}")
        else:
            print("No partial payments with status 'created' found.")
        
        return 0
    
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        print(f"Error: {str(e)}")
        return 1

if __name__ == "__main__":
    exit(main()) 