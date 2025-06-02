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
import argparse
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
        currency_col = "Currency"
        
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
        
        if currency_col not in df.columns:
            # Try to find similar column names
            for col in df.columns:
                if "currency" in col.lower():
                    currency_col = col
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
        
        # If currency column doesn't exist, add a default one
        if currency_col not in partial_payments.columns:
            partial_payments[currency_col] = "INR"
            currency_col = "Currency"  # Ensure we use the standard name
        
        # Sort by due amount (highest to lowest)
        partial_payments = partial_payments.sort_values(by=due_col, ascending=False)
        
        logging.info(f"Found {len(partial_payments)} payment links with status 'created' and partial payments")
        
        # Select relevant columns
        columns_to_export = [
            "ID", amount_col, paid_col, due_col, currency_col,
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
    """Generate summary of total amount due, split by Reference ID's starting with 'July' and the rest, and by currency"""
    # Check if Reference ID column exists
    ref_id_col = "Reference ID"
    if ref_id_col not in data.columns:
        # Try to find a similar column
        for col in data.columns:
            if "reference" in col.lower():
                ref_id_col = col
                break
    
    # Check if Currency column exists
    currency_col = "Currency"
    if currency_col not in data.columns:
        # Try to find a similar column
        for col in data.columns:
            if "currency" in col.lower():
                currency_col = col
                break
    
    # If currency column doesn't exist, assume all are INR
    has_currency = currency_col in data.columns
    
    # Create a summary dictionary
    summary = {
        "total": {"count": 0, "amount": 0},
        "july": {"count": 0, "amount": 0},
        "other": {"count": 0, "amount": 0},
        "by_currency": {}
    }
    
    # If we don't have a Reference ID column, return a simple summary
    if ref_id_col not in data.columns:
        total_due = data["Due Amount (₹)"].sum()
        summary["total"]["count"] = len(data)
        summary["total"]["amount"] = total_due
        summary["other"]["count"] = len(data)
        summary["other"]["amount"] = total_due
        
        # Add currency breakdown if available
        if has_currency:
            for currency, group in data.groupby(currency_col):
                if currency not in summary["by_currency"]:
                    summary["by_currency"][currency] = {
                        "total": {"count": 0, "amount": 0},
                        "july": {"count": 0, "amount": 0},
                        "other": {"count": 0, "amount": 0}
                    }
                summary["by_currency"][currency]["total"]["count"] = len(group)
                summary["by_currency"][currency]["total"]["amount"] = group["Due Amount (₹)"].sum()
                summary["by_currency"][currency]["other"]["count"] = len(group)
                summary["by_currency"][currency]["other"]["amount"] = group["Due Amount (₹)"].sum()
        else:
            # Default to INR if no currency column
            summary["by_currency"]["INR"] = {
                "total": {"count": len(data), "amount": total_due},
                "july": {"count": 0, "amount": 0},
                "other": {"count": len(data), "amount": total_due}
            }
        
        return summary
    
    # Split data into July references and others
    july_data = data[data[ref_id_col].astype(str).str.startswith("July")]
    other_data = data[~data[ref_id_col].astype(str).str.startswith("July")]
    
    # Calculate totals
    july_due = july_data["Due Amount (₹)"].sum() if not july_data.empty else 0
    other_due = other_data["Due Amount (₹)"].sum() if not other_data.empty else 0
    total_due = july_due + other_due
    
    # Update summary
    summary["total"]["count"] = len(data)
    summary["total"]["amount"] = total_due
    summary["july"]["count"] = len(july_data)
    summary["july"]["amount"] = july_due
    summary["other"]["count"] = len(other_data)
    summary["other"]["amount"] = other_due
    
    # Add currency breakdown
    if has_currency:
        # Group by currency
        for currency, group in data.groupby(currency_col):
            if currency not in summary["by_currency"]:
                summary["by_currency"][currency] = {
                    "total": {"count": 0, "amount": 0},
                    "july": {"count": 0, "amount": 0},
                    "other": {"count": 0, "amount": 0}
                }
            
            # Total for this currency
            summary["by_currency"][currency]["total"]["count"] = len(group)
            summary["by_currency"][currency]["total"]["amount"] = group["Due Amount (₹)"].sum()
            
            # July data for this currency
            july_group = group[group[ref_id_col].astype(str).str.startswith("July")]
            summary["by_currency"][currency]["july"]["count"] = len(july_group)
            summary["by_currency"][currency]["july"]["amount"] = july_group["Due Amount (₹)"].sum() if not july_group.empty else 0
            
            # Other data for this currency
            other_group = group[~group[ref_id_col].astype(str).str.startswith("July")]
            summary["by_currency"][currency]["other"]["count"] = len(other_group)
            summary["by_currency"][currency]["other"]["amount"] = other_group["Due Amount (₹)"].sum() if not other_group.empty else 0
    else:
        # Default to INR if no currency column
        summary["by_currency"]["INR"] = {
            "total": {"count": len(data), "amount": total_due},
            "july": {"count": len(july_data), "amount": july_due},
            "other": {"count": len(other_data), "amount": other_due}
        }
    
    return summary

def send_email_summary(summary, sheet_url):
    """Send email with summary of partial payments"""
    logging.info("Attempting to send email summary...")
    
    # Check if .env file exists
    if not os.path.exists('.env'):
        logging.error("No .env file found. Please create one based on env.example.")
        print("ERROR: No .env file found. Please create one based on env.example.")
        return False
    
    # Check for email credentials
    if not EMAIL_SENDER or EMAIL_SENDER == '':
        logging.error("EMAIL_SENDER not found or empty in environment variables. Email cannot be sent.")
        print("ERROR: EMAIL_SENDER not configured in .env file")
        print("Please add the following to your .env file:")
        print("EMAIL_SENDER=your_email@gmail.com")
        print("EMAIL_PASSWORD=your_app_password")
        return False
        
    if not EMAIL_PASSWORD or EMAIL_PASSWORD == '':
        logging.error("EMAIL_PASSWORD not found or empty in environment variables. Email cannot be sent.")
        print("ERROR: EMAIL_PASSWORD not configured in .env file")
        print("Please add the following to your .env file:")
        print("EMAIL_SENDER=your_email@gmail.com")
        print("EMAIL_PASSWORD=your_app_password")
        return False
    
    # Check for non-ASCII characters in email credentials
    if any(ord(c) > 127 for c in EMAIL_SENDER):
        logging.error("EMAIL_SENDER contains non-ASCII characters. This can cause encoding issues.")
        print("ERROR: EMAIL_SENDER contains non-ASCII characters. Please remove any special characters.")
        return False
    
    if any(ord(c) > 127 for c in EMAIL_PASSWORD):
        logging.error("EMAIL_PASSWORD contains non-ASCII characters. This can cause encoding issues.")
        print("ERROR: EMAIL_PASSWORD contains non-ASCII characters. Please use only ASCII characters.")
        return False
    
    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECIPIENT
        msg['Subject'] = f"Partial Payments Summary - {datetime.datetime.now().strftime('%Y-%m-%d')}"
        
        logging.info(f"Preparing email from {EMAIL_SENDER} to {EMAIL_RECIPIENT}")
        
        # Email body
        body = f"""
        <html>
        <body>
            <h2>Partial Payments Summary</h2>
            <p>Here's a summary of payment links with status "created" and partial payments:</p>
            
            <h3>Overall Summary</h3>
            <table border="1" cellpadding="5" cellspacing="0">
                <tr>
                    <td><b>Category</b></td>
                    <td><b>Count</b></td>
                    <td><b>Amount Due</b></td>
                </tr>
                <tr>
                    <td>July Reference IDs</td>
                    <td>{summary['july']['count']}</td>
                    <td>Rs. {summary['july']['amount']:,.2f}</td>
                </tr>
                <tr>
                    <td>Other Reference IDs</td>
                    <td>{summary['other']['count']}</td>
                    <td>Rs. {summary['other']['amount']:,.2f}</td>
                </tr>
                <tr>
                    <td><b>Total</b></td>
                    <td><b>{summary['total']['count']}</b></td>
                    <td><b>Rs. {summary['total']['amount']:,.2f}</b></td>
                </tr>
            </table>
            
            <h3>Breakdown by Currency</h3>
        """
        
        # Add currency breakdown tables
        for currency, data in summary["by_currency"].items():
            body += f"""
            <h4>{currency}</h4>
            <table border="1" cellpadding="5" cellspacing="0">
                <tr>
                    <td><b>Category</b></td>
                    <td><b>Count</b></td>
                    <td><b>Amount Due</b></td>
                </tr>
                <tr>
                    <td>July Reference IDs</td>
                    <td>{data['july']['count']}</td>
                    <td>{currency} {data['july']['amount']:,.2f}</td>
                </tr>
                <tr>
                    <td>Other Reference IDs</td>
                    <td>{data['other']['count']}</td>
                    <td>{currency} {data['other']['amount']:,.2f}</td>
                </tr>
                <tr>
                    <td><b>Total</b></td>
                    <td><b>{data['total']['count']}</b></td>
                    <td><b>{currency} {data['total']['amount']:,.2f}</b></td>
                </tr>
            </table>
            """
        
        body += f"""
            <p>For complete details, view the <a href="{sheet_url}">Google Sheet</a>.</p>
            
            <p>This is an automated message from the Razorpay Payment Links Tracker.</p>
        </body>
        </html>
        """
        
        # Use UTF-8 encoding for the email body
        msg.attach(MIMEText(body, 'html', 'utf-8'))
        logging.info("Email content prepared successfully")
        
        # Connect to Gmail SMTP server
        logging.info("Connecting to Gmail SMTP server...")
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.set_debuglevel(1)  # Add debug level for more detailed SMTP logs
            logging.info("SMTP connection established")
            
            server.ehlo()  # SMTP protocol start
            logging.info("EHLO command sent")
            
            server.starttls()
            logging.info("STARTTLS command sent")
            
            server.ehlo()  # SMTP protocol restart after TLS
            logging.info("Second EHLO command sent")
            
            # Log authentication attempt (without showing the password)
            logging.info(f"Attempting to login with account: {EMAIL_SENDER}")
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            logging.info("SMTP authentication successful")
            
            # Send the email
            logging.info("Sending email message...")
            server.send_message(msg)
            logging.info("Email sent successfully")
            
            server.quit()
            logging.info("SMTP connection closed")
            print(f"Email summary successfully sent to {EMAIL_RECIPIENT}")
            return True
        
        except smtplib.SMTPAuthenticationError as auth_error:
            logging.error(f"SMTP Authentication Error: {auth_error}")
            print(f"ERROR: Email authentication failed. If using Gmail, make sure you're using an App Password, not your regular password.")
            return False
        
        except smtplib.SMTPException as smtp_error:
            logging.error(f"SMTP Error: {smtp_error}")
            print(f"ERROR: SMTP error occurred: {smtp_error}")
            return False
    
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")
        print(f"ERROR: Failed to send email: {str(e)}")
        return False

def test_email_connection():
    """Test email connection and authentication"""
    logging.info("Testing email connection...")
    
    # Check if .env file exists
    if not os.path.exists('.env'):
        logging.error("No .env file found. Please create one based on env.example.")
        print("ERROR: No .env file found. Please create one based on env.example.")
        return False
    
    # Check for email credentials
    if not EMAIL_SENDER or EMAIL_SENDER == '':
        logging.error("EMAIL_SENDER not found or empty in environment variables. Email cannot be sent.")
        print("ERROR: EMAIL_SENDER not configured in .env file")
        print("Please add the following to your .env file:")
        print("EMAIL_SENDER=your_email@gmail.com")
        print("EMAIL_PASSWORD=your_app_password")
        return False
        
    if not EMAIL_PASSWORD or EMAIL_PASSWORD == '':
        logging.error("EMAIL_PASSWORD not found or empty in environment variables. Email cannot be sent.")
        print("ERROR: EMAIL_PASSWORD not configured in .env file")
        print("Please add the following to your .env file:")
        print("EMAIL_SENDER=your_email@gmail.com")
        print("EMAIL_PASSWORD=your_app_password")
        return False
    
    # Check for non-ASCII characters in email credentials
    if any(ord(c) > 127 for c in EMAIL_SENDER):
        logging.error("EMAIL_SENDER contains non-ASCII characters. This can cause encoding issues.")
        print("ERROR: EMAIL_SENDER contains non-ASCII characters. Please remove any special characters.")
        return False
    
    if any(ord(c) > 127 for c in EMAIL_PASSWORD):
        logging.error("EMAIL_PASSWORD contains non-ASCII characters. This can cause encoding issues.")
        print("ERROR: EMAIL_PASSWORD contains non-ASCII characters. Please use only ASCII characters.")
        return False
    
    try:
        # Connect to Gmail SMTP server
        print("Connecting to Gmail SMTP server...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.set_debuglevel(1)  # Add debug level for more detailed SMTP logs
        print("SMTP connection established")
        
        server.ehlo()  # SMTP protocol start
        print("EHLO command sent")
        
        server.starttls()
        print("STARTTLS command sent")
        
        server.ehlo()  # SMTP protocol restart after TLS
        print("Second EHLO command sent")
        
        # Log authentication attempt (without showing the password)
        print(f"Attempting to login with account: {EMAIL_SENDER}")
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        print("SMTP authentication successful")
        
        # Send a test email
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = EMAIL_RECIPIENT
        msg['Subject'] = "Test Email - Razorpay Payment Links Tracker"
        
        body = """
        <html>
        <body>
            <h2>Email Test Successful</h2>
            <p>This is a test email from the Razorpay Payment Links Tracker.</p>
            <p>If you're receiving this, email functionality is working correctly.</p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'html', 'utf-8'))
        
        print("Sending test email...")
        server.send_message(msg)
        print("Test email sent successfully")
        
        server.quit()
        print("SMTP connection closed")
        
        print(f"\nSUCCESS: Test email sent to {EMAIL_RECIPIENT}")
        return True
        
    except smtplib.SMTPAuthenticationError as auth_error:
        logging.error(f"SMTP Authentication Error: {auth_error}")
        print(f"ERROR: Email authentication failed. If using Gmail, make sure you're using an App Password, not your regular password.")
        return False
    
    except smtplib.SMTPException as smtp_error:
        logging.error(f"SMTP Error: {smtp_error}")
        print(f"ERROR: SMTP error occurred: {smtp_error}")
        return False
    
    except Exception as e:
        logging.error(f"Error testing email: {str(e)}")
        print(f"ERROR: Failed to test email: {str(e)}")
        return False

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Extract partial payments from Google Sheet')
    parser.add_argument('--test-email', action='store_true', help='Test email functionality only')
    args = parser.parse_args()
    
    # If test-email flag is provided, only test email functionality
    if args.test_email:
        return 0 if test_email_connection() else 1
    
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
            email_sent = send_email_summary(summary, sheet_url)
            
            # Export to CSV (optional)
            partial_payments.to_csv(OUTPUT_FILE, index=False)
            logging.info(f"Exported {len(partial_payments)} partial payments to {OUTPUT_FILE}")
            
            # Display summary
            print("\nPartial Payments Summary:")
            print(f"Total partial payments: {len(partial_payments)}")
            print(f"Total due amount: Rs. {summary['total']['amount']:,.2f}")
            print(f"July references: {summary['july']['count']} items, Rs. {summary['july']['amount']:,.2f}")
            print(f"Other references: {summary['other']['count']} items, Rs. {summary['other']['amount']:,.2f}")
            
            # Display currency breakdown
            print("\nBreakdown by Currency:")
            for currency, data in summary["by_currency"].items():
                print(f"\n{currency}:")
                print(f"  Total: {data['total']['count']} items, {currency} {data['total']['amount']:,.2f}")
                print(f"  July: {data['july']['count']} items, {currency} {data['july']['amount']:,.2f}")
                print(f"  Other: {data['other']['count']} items, {currency} {data['other']['amount']:,.2f}")
            
            # Display the first 5 records
            print("\nTop 5 partial payments (by due amount):")
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 1000)
            print(partial_payments.head(5).to_string())
            
            print(f"\nFull details exported to Google Sheet tab 'Partial Payments'")
            if email_sent:
                print(f"Email summary sent to {EMAIL_RECIPIENT}")
            else:
                print(f"WARNING: Email summary could not be sent to {EMAIL_RECIPIENT}. Check logs for details.")
        else:
            print("No partial payments with status 'created' found.")
        
        return 0
    
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        print(f"Error: {str(e)}")
        return 1

if __name__ == "__main__":
    exit(main()) 