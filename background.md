# Claude 3.7 Sonnet + Cursor Prompt

## 1. Goal
Fetch **all** Razorpay payment links (any status) in the **live** environment, dump them into a Google Sheet, and ensure that sheet can be updated:
- **Daily** (automated)
- **On-demand** (manual trigger)

## 2. Context & Setup

1. **Razorpay Live**  
   - Use **live** API keys from Razorpay Dashboard.  
   - Base URL: `https://api.razorpay.com/v1/payment_links`  
   - Pagination limit is 100 items per call (must loop/skip until all are fetched).

2. **Google Sheets Integration**  
   - A service‐account Key ID (and Email) is available.  
   - Sheet ID (the string in the Google Sheet URL) is known.  

## 3. Data Fields to Fetch
For each payment link record, collect all fields (including sub-fields), including:
- `id`
- `created_at` (Unix epoch seconds → convert to UTC string)
- `amount` (total, in paise → divide by 100 for ₹)
- `amount_paid` (in paise → divide by 100 for ₹)
- `amount_due` (computed: `amount − amount_paid`, ₹)
- `status` (enum: `created`, `issued`, `pending`, `paid`, `cancelled`, `expired`, `partially_paid`)
- `customer.email` (empty string if none)
- `order_id`
- `reference_id`
- `short_url`
- `upi_link` (boolean)

## 4. Filtering & Pagination
- **No status filter** (fetch all statuses).  
- **Pagination:**  
  1. Call  
     ```
     GET https://api.razorpay.com/v1/payment_links?count=100&skip=SKIP_VALUE
     ```  
  2. Extract `items` from JSON.  
  3. If `items.length < 100`, stop. Else do `skip += 100` and repeat.

- _Optional:_ To restrict by date range, pass `from=START_TS` and `to=END_TS` (Unix seconds).

## 5. Google Sheets Workflow
1. **Auth**: Use service‐account Key ID   
2. **Select Worksheet**: Open by Sheet ID and target the first worksheet (`.sheet1`).  
3. **Clear Existing Data**: `worksheet.clear()` (ensures fresh data).  
4. **Write Rows**:  (this might change based on additional fields you may extract)
   - First row = header:  
     ```csv
     ID | Created At (UTC) | Total (₹) | Paid (₹) | Due (₹) | Status | Customer Email | Order ID | Reference ID | Short URL | UPI Link
     ```  
   - For each link: convert epoch → ISO UTC, amounts → ₹, UPI boolean → string.

5. **Trigger**  
   - For “daily” updates, schedule this same prompt via cron or Cloud Function.  
   - For “on-demand,” paste this prompt into Claude whenever you want real-time data.

---

