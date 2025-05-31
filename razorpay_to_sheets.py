#!/usr/bin/env python3
"""
Razorpay Payment Links to Google Sheets
Fetches all Razorpay payment links and updates a Google Sheet
"""

import os
import sys
import time
import json
import logging
from datetime import datetime, timezone, UTC
import argparse

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("razorpay_sync.log"),
        logging.StreamHandler()
    ]
)

# Check for required packages and provide helpful error messages
required_packages = {
    "requests": "HTTP requests to Razorpay API",
    "gspread": "Google Sheets API access",
    "google.oauth2.service_account": "Google authentication"
}

missing_packages = []
for package, purpose in required_packages.items():
    try:
        if "." in package:
            # Handle module.submodule format
            main_package, submodule = package.split(".", 1)
            __import__(main_package)
            # Try to import the submodule from the main package
            module = sys.modules[main_package]
            try:
                for part in submodule.split("."):
                    module = getattr(module, part)
            except AttributeError:
                missing_packages.append(f"{package} (needed for {purpose})")
        else:
            __import__(package)
    except ImportError:
        missing_packages.append(f"{package} (needed for {purpose})")

if missing_packages:
    logging.error("Missing required Python packages. Please install them using pip:")
    logging.error(f"pip install {' '.join(p.split()[0] for p in missing_packages)}")
    logging.error("\nMissing packages and their purposes:")
    for package in missing_packages:
        logging.error(f"- {package}")
    sys.exit(1)

# Now that we've checked for packages, import them
import requests
from google.oauth2.service_account import Credentials
import gspread

# Try to import dotenv_loader
try:
    from dotenv_loader import load_env_vars
    # Load environment variables from .env file
    load_env_vars()
except ImportError:
    logging.warning("dotenv_loader.py not found. Make sure environment variables are set manually.")

# Razorpay API Configuration
RAZORPAY_BASE_URL = "https://api.razorpay.com/v1/payment_links"
RAZORPAY_KEY_ID = os.environ.get("RAZORPAY_KEY_ID")
RAZORPAY_KEY_SECRET = os.environ.get("RAZORPAY_KEY_SECRET")

# Google Sheets Configuration
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")

# Check if we're in debug mode
DEBUG = os.environ.get("DEBUG", "").lower() in ["true", "1", "yes"]

def debug_dump(data, filename="debug_output.json"):
    """Dump data to a file for debugging"""
    if DEBUG:
        with open(filename, "w") as f:
            json.dump(data, f, indent=2)
        logging.debug(f"Debug data dumped to {filename}")

def validate_razorpay_credentials():
    """Validate Razorpay API credentials and check connection"""
    logging.info("Validating Razorpay API credentials...")
    
    if not RAZORPAY_KEY_ID or not RAZORPAY_KEY_SECRET:
        logging.error("Razorpay API keys not found in environment variables")
        logging.error("Please check your .env file or environment variables")
        logging.error(f"RAZORPAY_KEY_ID: {'✓ Set' if RAZORPAY_KEY_ID else '✗ Missing'}")
        logging.error(f"RAZORPAY_KEY_SECRET: {'✓ Set' if RAZORPAY_KEY_SECRET else '✗ Missing'}")
        return False
    
    # Mask the key for logging (show only first 4 and last 4 characters)
    key_id_masked = f"{RAZORPAY_KEY_ID[:4]}...{RAZORPAY_KEY_ID[-4:]}" if len(RAZORPAY_KEY_ID) > 8 else "****"
    logging.info(f"Using Razorpay Key ID: {key_id_masked}")
    
    # Test connection with a minimal request (count=1)
    try:
        logging.info("Testing connection to Razorpay API...")
        response = requests.get(
            RAZORPAY_BASE_URL,
            auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET),
            params={"count": 1}
        )
        
        if response.status_code == 200:
            data = response.json()
            logging.info("✓ Successfully connected to Razorpay API")
            logging.info(f"API Response: Status 200 OK, Content Type: {response.headers.get('Content-Type', 'unknown')}")
            
            # Check for payment_links key (Razorpay API structure)
            if 'payment_links' in data:
                logging.info(f"Found {len(data['payment_links'])} payment links in sample response")
                if DEBUG:
                    debug_dump(data, "razorpay_sample_response.json")
                return True
            else:
                logging.warning("Connected to API but 'payment_links' key not found in response")
                logging.debug(f"Response keys: {list(data.keys())}")
                debug_dump(data, "razorpay_unexpected_response.json")
                return False
        else:
            logging.error(f"Failed to connect to Razorpay API: Status {response.status_code}")
            logging.error(f"Response: {response.text}")
            return False
            
    except Exception as e:
        logging.error(f"Error connecting to Razorpay API: {str(e)}")
        return False

def fetch_all_payment_links(from_ts=None, to_ts=None):
    """Fetch all payment links from Razorpay with pagination"""
    if not validate_razorpay_credentials():
        raise ValueError("Failed to validate Razorpay API credentials")
    
    all_links = []
    skip = 0
    count = 100  # Maximum allowed by Razorpay API
    request_count = 0
    
    logging.info("Starting to fetch all payment links...")
    logging.info(f"Date range: {from_ts or 'all time'} to {to_ts or 'present'}")
    
    start_time = time.time()
    
    while True:
        request_count += 1
        params = {"count": count, "skip": skip}
        
        # Add optional date range filters if provided
        if from_ts:
            params["from"] = from_ts
        if to_ts:
            params["to"] = to_ts
        
        logging.info(f"Request #{request_count}: Fetching payment links (skip={skip}, count={count})")
        
        try:
            response = requests.get(
                RAZORPAY_BASE_URL,
                auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET),
                params=params
            )
            
            if response.status_code != 200:
                logging.error(f"API request #{request_count} failed with status {response.status_code}")
                logging.error(f"Response: {response.text}")
                raise Exception(f"API request failed with status {response.status_code}: {response.text}")
            
            data = response.json()
            
            # Debug dump the first and last response if in debug mode
            if DEBUG and (request_count == 1 or len(data.get('payment_links', [])) < count):
                debug_dump(data, f"razorpay_response_{request_count}.json")
            
            # Get payment links from the 'payment_links' field (Razorpay API structure)
            links = data.get("payment_links", [])
            item_count = len(links)
            
            logging.info(f"Retrieved {item_count} payment links in request #{request_count}")
            
            all_links.extend(links)
            
            # If we got fewer items than requested, we've reached the end
            if item_count < count:
                logging.info(f"Received fewer items ({item_count}) than requested ({count}), ending pagination")
                break
            
            skip += count
            logging.info(f"Fetched {len(all_links)} payment links so far...")
            
            # Small delay to avoid rate limiting
            time.sleep(0.5)
            
        except Exception as e:
            logging.error(f"Error in request #{request_count}: {str(e)}")
            raise
    
    end_time = time.time()
    duration = end_time - start_time
    
    logging.info(f"Completed fetching payment links in {duration:.2f} seconds")
    logging.info(f"Total API requests: {request_count}")
    logging.info(f"Total payment links fetched: {len(all_links)}")
    
    # Log a sample payment link if available
    if all_links and DEBUG:
        sample = all_links[0]
        logging.debug("Sample payment link:")
        for key in ['id', 'created_at', 'amount', 'status']:
            if key in sample:
                logging.debug(f"  {key}: {sample[key]}")
    
    return all_links

def format_timestamp(timestamp, default=""):
    """Convert a Unix timestamp to ISO format datetime string with proper UTC handling"""
    if not timestamp or timestamp == 0:
        return default
    try:
        # Use the recommended approach (Python 3.11+)
        if hasattr(datetime, 'UTC'):  # Python 3.11+
            dt = datetime.fromtimestamp(timestamp, UTC)
        else:  # Fallback for older Python versions
            dt = datetime.fromtimestamp(timestamp, timezone.utc)
        return dt.isoformat()
    except (ValueError, TypeError, OverflowError):
        return default

def process_payment_links(links):
    """Process and transform payment link data for Google Sheets"""
    logging.info(f"Processing {len(links)} payment links for Google Sheets...")
    
    processed_data = []
    
    # Expanded header row with all fields
    header = [
        "ID", "Created At (UTC)", "Updated At (UTC)", "Amount (₹)", "Amount Paid (₹)", 
        "Status", "Currency", "Description", "Reference ID", "Short URL", 
        "UPI Link", "WhatsApp Link", "Accept Partial", "First Min Partial Amount (₹)",
        "Customer Email", "Customer Contact", "Order ID", "User ID",
        "Cancelled At (UTC)", "Expire By (UTC)", "Expired At (UTC)",
        "Reminder Enable", "Reminder Status", "Payments Count",
        "Payments Details", "Notes"
    ]
    processed_data.append(header)
    
    statuses = {}
    
    for link in links:
        try:
            # Convert Unix timestamps to UTC string (handle 0 values)
            created_at = format_timestamp(link.get("created_at"))
            updated_at = format_timestamp(link.get("updated_at"))
            cancelled_at = format_timestamp(link.get("cancelled_at"))
            expire_by = format_timestamp(link.get("expire_by"))
            expired_at = format_timestamp(link.get("expired_at"))
            
            # Convert amounts from paise to rupees
            amount = float(link.get("amount", 0)) / 100
            amount_paid = float(link.get("amount_paid", 0)) / 100
            first_min_partial_amount = float(link.get("first_min_partial_amount", 0)) / 100
            
            # Extract customer information (with fallbacks to empty string)
            customer = link.get("customer", {})
            customer_email = customer.get("email", "") if customer else ""
            customer_contact = customer.get("contact", "") if customer else ""
            
            # Handle booleans
            upi_link = "Yes" if link.get("upi_link", False) else "No"
            whatsapp_link = "Yes" if link.get("whatsapp_link", False) else "No"
            accept_partial = "Yes" if link.get("accept_partial", False) else "No"
            reminder_enable = "Yes" if link.get("reminder_enable", False) else "No"
            
            # Handle reminders - could be object with status, empty, or None
            reminders = link.get("reminders", {})
            reminder_status = ""
            if reminders:
                if isinstance(reminders, dict):
                    reminder_status = reminders.get("status", "")
                elif isinstance(reminders, list) and len(reminders) > 0:
                    reminder_status = str(reminders)
                else:
                    reminder_status = str(reminders)
            
            # Handle payments array
            payments = link.get("payments", [])
            payments = [] if payments is None else payments
            payments_count = len(payments)
            
            # Create a summary of payment details
            payment_details = []
            for payment in payments:
                payment_amount = float(payment.get("amount", 0)) / 100
                payment_method = payment.get("method", "")
                payment_status = payment.get("status", "")
                payment_id = payment.get("payment_id", "")
                payment_created = format_timestamp(payment.get("created_at"))
                
                payment_details.append(f"{payment_id}: {payment_amount}₹ via {payment_method} ({payment_status}) on {payment_created}")
            
            payment_details_str = " | ".join(payment_details)
            
            # Handle notes (could be array, object, or None)
            notes = link.get("notes", [])
            notes_str = ""
            if notes:
                if isinstance(notes, list):
                    notes_str = ", ".join(str(note) for note in notes) if notes else ""
                else:
                    notes_str = str(notes)
            
            # Track statuses for summary
            status = link.get("status", "unknown")
            statuses[status] = statuses.get(status, 0) + 1
            
            row = [
                link.get("id", ""),
                created_at,
                updated_at,
                amount,
                amount_paid,
                status,
                link.get("currency", ""),
                link.get("description", ""),
                link.get("reference_id", ""),
                link.get("short_url", ""),
                upi_link,
                whatsapp_link,
                accept_partial,
                first_min_partial_amount,
                customer_email,
                customer_contact,
                link.get("order_id", ""),
                link.get("user_id", ""),
                cancelled_at,
                expire_by,
                expired_at,
                reminder_enable,
                reminder_status,
                payments_count,
                payment_details_str,
                notes_str
            ]
            
            processed_data.append(row)
        except Exception as e:
            logging.error(f"Error processing payment link {link.get('id', 'unknown')}: {str(e)}")
            if DEBUG:
                logging.debug(f"Problematic payment link data: {json.dumps(link, indent=2)}")
    
    # Log status summary
    logging.info("Payment link status summary:")
    for status, count in statuses.items():
        logging.info(f"  {status}: {count}")
    
    logging.info(f"Processed {len(processed_data)-1} payment links with {len(header)} columns")
    return processed_data

def update_google_sheet(data):
    """Update Google Sheet with payment links data"""
    logging.info("Updating Google Sheet...")
    
    if not GOOGLE_SHEET_ID:
        logging.error("Google Sheet ID not found in environment variables")
        raise ValueError("Google Sheet ID not found in environment variables")
    
    # Mask the sheet ID for logging
    sheet_id_masked = f"{GOOGLE_SHEET_ID[:4]}...{GOOGLE_SHEET_ID[-4:]}" if len(GOOGLE_SHEET_ID) > 8 else "****"
    logging.info(f"Using Google Sheet ID: {sheet_id_masked}")
    
    # Check if service account file exists
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        logging.error(f"Service account file '{SERVICE_ACCOUNT_FILE}' not found")
        raise FileNotFoundError(f"Service account file '{SERVICE_ACCOUNT_FILE}' not found. Please download it from Google Cloud Console.")
    
    logging.info(f"Using service account file: {SERVICE_ACCOUNT_FILE}")
    
    # Set up Google Sheets authentication
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=scopes
        )
        
        gc = gspread.authorize(credentials)
        logging.info("Google Sheets authentication successful")
        
        # Open the spreadsheet and select the first worksheet
        spreadsheet = gc.open_by_key(GOOGLE_SHEET_ID)
        worksheet = spreadsheet.sheet1
        
        logging.info(f"Connected to spreadsheet: {spreadsheet.title}")
        logging.info(f"Using worksheet: {worksheet.title}")
        
        # Clear existing data
        logging.info("Clearing existing data from worksheet")
        worksheet.clear()
        
        # Get the range for our data (e.g., A1:K100)
        # We need the column count from the first row and the total row count
        if data:
            num_rows = len(data)
            num_cols = len(data[0])
            
            # Convert column number to letter (1=A, 2=B, etc.)
            def col_num_to_letter(n):
                result = ""
                while n > 0:
                    n, remainder = divmod(n - 1, 26)
                    result = chr(65 + remainder) + result
                return result
            
            last_col_letter = col_num_to_letter(num_cols)
            range_name = f"A1:{last_col_letter}{num_rows}"
            
            logging.info(f"Updating worksheet with {num_rows} rows and {num_cols} columns (range: {range_name})")
            
            # Update with new data using the correct range
            worksheet.update(range_name=range_name, values=data)
            
            logging.info(f"Google Sheet updated successfully with {len(data)-1} payment links")
            logging.info(f"Sheet URL: https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}")
        else:
            logging.warning("No data to update in Google Sheet")
        
    except Exception as e:
        logging.error(f"Failed to update Google Sheet: {str(e)}")
        raise Exception(f"Failed to update Google Sheet: {str(e)}")

def main():
    parser = argparse.ArgumentParser(description="Fetch Razorpay payment links and update Google Sheet")
    parser.add_argument("--from_date", help="Start date in YYYY-MM-DD format")
    parser.add_argument("--to_date", help="End date in YYYY-MM-DD format")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode")
    args = parser.parse_args()
    
    # Set debug mode if requested
    global DEBUG
    if args.debug:
        DEBUG = True
        logging.getLogger().setLevel(logging.DEBUG)
        logging.debug("Debug mode enabled")
    
    logging.info("Starting Razorpay to Google Sheets sync")
    logging.info(f"Python version: {sys.version}")
    
    # Convert date strings to Unix timestamps if provided
    from_ts = None
    to_ts = None
    
    if args.from_date:
        logging.info(f"Using from_date: {args.from_date}")
        from_dt = datetime.strptime(args.from_date, "%Y-%m-%d")
        from_ts = int(from_dt.timestamp())
    
    if args.to_date:
        logging.info(f"Using to_date: {args.to_date}")
        to_dt = datetime.strptime(args.to_date, "%Y-%m-%d")
        # Set to end of day
        to_dt = to_dt.replace(hour=23, minute=59, second=59)
        to_ts = int(to_dt.timestamp())
    
    try:
        # Fetch all payment links
        payment_links = fetch_all_payment_links(from_ts, to_ts)
        
        if not payment_links:
            logging.warning("No payment links were found. Check your Razorpay account and API keys.")
            logging.info("Continuing to update Google Sheet with headers only.")
        
        # Process data for Google Sheets
        processed_data = process_payment_links(payment_links)
        
        # Update Google Sheet
        update_google_sheet(processed_data)
        
        logging.info("Razorpay to Google Sheets sync completed successfully")
        
    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 