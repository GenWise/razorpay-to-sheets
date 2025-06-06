# Razorpay to Google Sheets

A Python application that fetches all Razorpay payment links from the live environment and dumps them into a Google Sheet. The application supports both daily automated updates and on-demand manual updates.

## Features

- Fetches all Razorpay payment links from the live environment
- Handles pagination to retrieve all records
- Converts Unix timestamps to readable UTC datetime
- Converts amounts from paise to rupees
- Updates a Google Sheet with the fetched data
- Supports optional date range filtering
- Can be run manually or scheduled as a cron job
- Provides detailed logging and debugging
- Exports all available payment link fields, including nested data

## Payment Link Fields Exported

The application extracts and exports the following fields from each payment link:

1. **ID** - Unique identifier for the payment link
2. **Created At (UTC)** - When the payment link was created
3. **Updated At (UTC)** - When the payment link was last updated
4. **Amount (₹)** - Total amount in rupees
5. **Amount Paid (₹)** - Amount paid so far in rupees
6. **Status** - Current status (created, paid, partially_paid, cancelled, expired)
7. **Currency** - Currency code (e.g., INR)
8. **Description** - Payment link description
9. **Reference ID** - Reference ID for the payment
10. **Short URL** - The shortened URL for the payment link
11. **UPI Link** - Whether UPI link is enabled (Yes/No)
12. **WhatsApp Link** - Whether WhatsApp link is enabled (Yes/No)
13. **Accept Partial** - Whether partial payments are accepted (Yes/No)
14. **First Min Partial Amount (₹)** - Minimum partial payment amount
15. **Customer Email** - Customer's email address
16. **Customer Contact** - Customer's contact number
17. **Order ID** - Associated order ID
18. **User ID** - User ID that created the payment link
19. **Cancelled At (UTC)** - When the payment link was cancelled (if applicable)
20. **Expire By (UTC)** - When the payment link is set to expire
21. **Expired At (UTC)** - When the payment link expired (if applicable)
22. **Reminder Enable** - Whether reminders are enabled (Yes/No)
23. **Reminder Status** - Status of reminders
24. **Payments Count** - Number of payments made against this link
25. **Payments Details** - Summary of all payments (IDs, amounts, methods, statuses)
26. **Notes** - Any notes attached to the payment link

## Extract Partial Payments

The `extract_partial_payments.py` script allows you to identify payment links where:
1. The amount paid is less than the total amount
2. The status is "created" (active links awaiting payment)

### Features

- Reads data directly from the Google Sheet
- Identifies payment links with status "created" where amount paid < total amount
- Calculates the due amount for each payment link
- Sorts results by due amount (highest to lowest)
- Creates/updates a "Partial Payments" tab in the same Google Sheet
- Provides a breakdown of payments by currency and reference ID type
- Sends an email summary to the specified recipient with:
  - Total due amount
  - Due amount split by Reference IDs starting with "July" vs others
  - Breakdown by currency (INR, AED, etc.)
  - Link to the Google Sheet
- Exports results to a CSV file (optional)

### Usage

```bash
python extract_partial_payments.py
```

The script will:
1. Connect to the Google Sheet using the service account credentials
2. Extract all payment links with status "created" and partial payments
3. Create or update a "Partial Payments" tab in the same Google Sheet
4. Send an email summary to the configured recipient
5. Display a summary of the results
6. Export the full list to `partial_payments.csv`

### Email Configuration

To enable email functionality, add the following to your `.env` file:

```
EMAIL_SENDER=your_email@gmail.com
EMAIL_PASSWORD=your_app_password
```

For Gmail, you need to use an App Password, not your regular password:

1. Go to https://myaccount.google.com/apppasswords
2. Sign in with your Google account
3. Select "Mail" as the app and "Other" as the device
4. Enter a name like "Razorpay Script"
5. Click "Generate" and copy the 16-character password
6. Use this as your EMAIL_PASSWORD

**Important**: Make sure both your email address and password contain only ASCII characters (no special characters or Unicode). Non-ASCII characters can cause encoding issues when sending emails.

### Troubleshooting Email Issues

If you're having trouble with email sending, you can test the email functionality separately:

```bash
python extract_partial_payments.py --test-email
```

This will:
1. Test the connection to the Gmail SMTP server
2. Verify your email credentials
3. Send a test email to the configured recipient
4. Show detailed logs of the email sending process

Common issues:

1. **Non-ASCII characters**: If you see an error like `'ascii' codec can't encode character`, check your email and password for special characters and replace them.

2. **Authentication errors**: Make sure you're using an App Password for Gmail, not your regular password. Regular passwords won't work if you have 2-factor authentication enabled.

3. **Security settings**: If you're not using an App Password, you may need to enable "Less secure app access" in your Google account settings.

4. **Environment variables**: Verify that EMAIL_SENDER and EMAIL_PASSWORD are correctly set in your .env file and don't contain any whitespace or special characters.

5. **Gmail limits**: Gmail has sending limits. If you're sending too many emails, you might hit these limits.

### Output Example

```
Partial Payments Summary:
Total partial payments: 7
Total due amount: Rs. 253,900.00
July references: 5 items, Rs. 24,900.00
Other references: 2 items, Rs. 229,000.00

Breakdown by Currency:

AED:
  Total: 4 items, AED 4,200.00
  July: 4 items, AED 4,200.00
  Other: 0 items, AED 0.00

INR:
  Total: 3 items, INR 249,700.00
  July: 1 items, INR 20,700.00
  Other: 2 items, INR 229,000.00

Top 5 partial payments (by due amount):
                       ID  Amount (₹)  Amount Paid (₹)  Due Amount (₹) Currency   Status                    Short URL Reference ID             Customer Email Customer Contact
36   plink_Q3neODhnEhHdss      159000                0          159000      INR  created   https://rzp.io/rzp/ZTYh45M           F1   jayanthanmohan@gmail.com       9500848488
123  plink_PYxkvmzfPM8b6W       70000                0           70000      INR  created  https://rzp.io/rzp/Ilniud0k        T60S3  monajaisinghani@gmail.com       9891183507
9    plink_QbmfxDOcH1Jerr       20700                0           20700      INR  created  https://rzp.io/rzp/jCrJPLnK     July-003    veeravictoria@gmail.com     971501603843
8    plink_QbmsLkeHHw1hQQ        1500                0            1500      AED  created  https://rzp.io/rzp/9Pwidule     July-004        rucha.oka@gmail.com
0    plink_QcF6LHq43BkgIG         900                0             900      AED  created   https://rzp.io/rzp/3jVUYxU     July-008   rudolph.serrao@gmail.com     971526453705

Full details exported to Google Sheet tab 'Partial Payments'
Email summary sent to rajesh@genwise.in
```

## Setup

### Prerequisites

- Python 3.6+
- Razorpay live API keys
- Google Service Account with access to Google Sheets
- Google Sheet ID

### Installation

1. Clone this repository:
   ```
   git clone <repository-url>
   cd rzrpy
   ```

2. Run the dependency check script to ensure all requirements are met:
   ```
   chmod +x check_dependencies.sh
   ./check_dependencies.sh
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

4. Create a `.env` file based on the provided `env.example`:
   ```
   cp env.example .env
   ```

5. Edit the `.env` file and add your Razorpay API keys and Google Sheet ID:
   ```
   # Razorpay API credentials (from Razorpay Dashboard)
   RAZORPAY_KEY_ID=your_razorpay_key_id
   RAZORPAY_KEY_SECRET=your_razorpay_key_secret
   
   # Google Sheets configuration
   GOOGLE_SHEET_ID=your_google_sheet_id
   GOOGLE_SERVICE_ACCOUNT_FILE=service_account.json
   ```

6. Place your Google Service Account JSON file in the project directory with the name specified in your `.env` file (default is `service_account.json`)

### Google Sheets Setup

1. Create a new Google Sheet or use an existing one
2. Share the sheet with the email address of your Google Service Account (with Editor permissions)
3. Copy the Sheet ID from the URL (it's the long string between `/d/` and `/edit` in the URL)

## Usage

### Running Manually (On-Demand)

To fetch all payment links and update the Google Sheet:

```bash
python razorpay_to_sheets.py
```

To fetch payment links within a specific date range:

```bash
python razorpay_to_sheets.py --from_date 2023-01-01 --to_date 2023-01-31
```

To run with detailed debugging enabled:

```bash
python razorpay_to_sheets.py --debug
```

### Scheduling Daily Updates

1. Make sure the daily update script is executable:
   ```bash
   chmod +x daily_update.sh
   ```

2. Add a cron job to run the script daily:
   ```bash
   crontab -e
   ```

3. Add the following line to run the job daily at 1:00 AM:
   ```
   0 1 * * * /path/to/rzrpy/daily_update.sh
   ```

## Utility Scripts

The project includes several utility scripts to help with setup and maintenance:

1. **check_dependencies.sh** - Checks if all required Python packages are installed and validates environment setup
   ```
   ./check_dependencies.sh
   ```

2. **clean.sh** - Removes Python cache files, debug logs, and output files
   ```
   ./clean.sh
   ```

3. **daily_update.sh** - Script for scheduled automatic updates via cron
   ```
   ./daily_update.sh
   ```

## Troubleshooting

### Common Issues

1. **Missing API Keys**: Ensure that your Razorpay API keys are correctly set in the `.env` file.

2. **Google Sheets Access**: If you get permission errors with Google Sheets, make sure:
   - The service account email has been granted Editor access to the Google Sheet
   - The service account JSON file is correctly placed and referenced

3. **API Rate Limits**: If you hit Razorpay API rate limits, try using the date range filters to make smaller, more focused requests.

4. **Import Errors**: If you see import errors:
   - Run `./check_dependencies.sh` to verify all required packages are installed
   - Run `./clean.sh` to remove Python cache files that might be causing issues

5. **Deprecation Warnings**: The script may display some deprecation warnings from the gspread library. These are just warnings and don't affect functionality.

6. **No Data in Google Sheet**: If headers appear but no data is shown:
   - Use the `--debug` flag to get detailed logging: `python razorpay_to_sheets.py --debug`
   - Check the generated log file `razorpay_sync.log` for error messages
   - Verify your Razorpay API keys are correct
   - Check that you're using the LIVE environment keys, not TEST

### Debugging

For advanced troubleshooting, you can use the debug mode:

```bash
python razorpay_to_sheets.py --debug
```

This will:
1. Enable detailed logging to both console and `razorpay_sync.log` file
2. Dump API responses to JSON files for inspection
3. Show more information about each step of the process

After debugging, clean up the generated files with:

```bash
./clean.sh
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.