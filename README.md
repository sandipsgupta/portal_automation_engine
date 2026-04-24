# Portal Automation Engine

Automated data entry into the Tata Consumer portal (Salesforce Experience Cloud) for payment/collection settlement. Processes invoice payments from Excel sheets, searches the portal, and enters data via Playwright automation.

**Portal URL**: `https://mavic-tataconsumer.my.site.com/dms/s/`

---

## Quick Start

### Prerequisites
- Python 3.8+
- Playwright browsers installed: `python -m playwright install chromium`
- Excel file with invoice data (see Data Format below)
- Portal credentials in `.env` file
- For `.xlsx` support: `pip install openpyxl` (optional, CSV supported by default)

### Installation
```bash
pip install -r requirements.txt
python -m playwright install chromium
```

### Setup
Create `.env` file:
```
PORTAL_USERNAME=your_username
PORTAL_PASSWORD=your_password
```

Update `config.json`:
```json
{
  "portal_url": "https://mavic-tataconsumer.my.site.com/dms/s/login/",
  "settlement_url": "https://mavic-tataconsumer.my.site.com/dms/s/manual-invoice-payment-settlement",
  "sheet_path": "D:/TATA_SHIPMENTS_April - Test.xlsx",
  "headless": false,
  "debug": false,
  "selected_portal": "tata_consumer"
}
```

### Execution Modes
- **CLI Mode**: `python main.py` (uses `config.json`; includes an interactive pause before the final page Save)
- **GUI Mode**: `python app.py` (Tkinter interface for configuration and monitoring + CSV run log)

### Run Output
- **CSV Log (GUI mode)**: `logs/run_tata_consumer_{timestamp}.csv` — Entry-by-entry status
- **Screenshots**: `screenshots/` directory
  - `01_after_search.png` — Table after search
  - `02_modal_*.png` — Modal state per payment
  - `03_done_*.png` — Invoice completion
  - `04_final_save.png` — Portal after save
- **Console**: Real-time progress

---

## For AI Assistants / Contributors

If you're using this repo with ChatGPT/Copilot/Claude/Codex/Cursor, this section is the "mental model" to keep changes safe.

- **Entry points**: `main.py` (CLI + engine) and `app.py` (GUI wrapper + logger).
- **Core flow**: `load_rows_from_csv()` builds a flat list of payment entries + a Bill Date search range; `RetailerPortalEngine` logs in, searches, detects duplicates, and posts payments.
- **Key invariant**: Portal search uses **Bill Date only** (invoice date range), while the payment date written into the modal comes from `DELIVERY DATE`/`CHEQUE DATE`.
- **Duplicate safety**: The engine checks the portal table cells and skips payment modes that already have a non-zero value (safe to re-run).
- **Be careful changing selectors/timing**: This is a Salesforce Lightning UI; small selector/timing tweaks can regress automation. Prefer small, isolated changes and keep screenshots/debug output useful.
- **Secrets**: Keep real credentials in `.env` only; do not commit or share `.env` contents.

## Project Structure

### Files
- **`main.py`**: Core Playwright automation engine + CLI entrypoint
- **`app.py`**: Tkinter GUI frontend for configuration and monitoring
- **`config.json`**: Runtime settings (persisted by GUI)
- **`.env`**: Portal credentials (username/password)
- **`requirements.txt`**: Python dependencies
- **`logs/`**: Execution logs (CSV format)
- **`screenshots/`**: Portal screenshots during automation
- **`build/`**: PyInstaller build artifacts
- **`PortalAutomationEngine.spec`**: PyInstaller spec for building the Windows GUI executable
- **`build.bat`**: Convenience build script (Windows)

### Entry Points
- **`main.py`**: CLI automation runner (headless optional via `config.json`)
- **`app.py`**: GUI launcher with configuration management

### Core Components
- **`RetailerPortalEngine`** class: Browser automation, login, search, data entry
- **`load_rows_from_csv()`**: Parse Excel/CSV, extract invoice and payment data
- **Date parsers**: Support DD-Mon-YYYY, DD/MM/YYYY, YYYY-MM-DD formats

---

## Data Processing Pipeline

### Input Format (Excel)
Required columns:
- `Bill Number` — Invoice ID
- `Bill Date` — Invoice date (used for portal search range)
- `DELIVERY DATE` — Payment date for Cash/UPI (if blank, those payment entries are skipped)
- `CHEQUE DATE` — Payment date for Cheque (optional, falls back to `DELIVERY DATE`)
- `CHEQUE NUMBER` — For Cheque payments (optional)
- `Cash` — Amount (skipped if 0)
- `Cheque Amount` — Amount (skipped if 0)
- `UPI Amount` — Amount (skipped if 0)
- `Credit Amount` — Credit-only rows are skipped

### Processing Steps

1. **Read Sheet**: Parse Excel/CSV, extract all invoice rows
2. **Skip Rows**:
   - Credit-only entries (no payable amount)
   - Rows without a usable payment date for that payment mode (e.g., blank `DELIVERY DATE`)
3. **Normalize Data**:
   - Float amounts: `263.0` → `263`
   - Dates: Parse any format, normalize to DD-MM-YYYY
4. **Calculate Date Range**:
   - Uses **BILL DATE ONLY** for portal search range (to find invoices by invoice date)
   - `DELIVERY DATE` used for Cash/UPI payment dates
   - `CHEQUE DATE` used for Cheque payment dates (falls back to Delivery Date)
5. **Group by Invoice**: All payments for one invoice processed in one pass

### Output Format
```python
rows = [
  {
    "invoice_no": "IN17781827000153",
    "payment_mode": "Cheque",
    "cheque_no": "7602",
    "amount": "3754",
    "date": "13-04-2026"
  },
  # ... more payments
]
from_date = "04-04-2026"  # Bill date range
to_date = "06-04-2026"
```

---

## Portal Workflow

### Settlement Page Search

1. **Navigate** to settlement URL
2. **Fill Date Filters**:
   - Invoice From Date: DD-Mon-YYYY format
   - Invoice To Date: DD-Mon-YYYY format
3. **Click Top Search** (date validation button)
4. **Click Bottom Search** (loads results table)
5. **Wait for Table** (50 rows per page, pagination supported)

### Invoice Processing (Per Invoice)

1. **Find Row**: Search table for invoice number, paginate if needed
2. **Enter Cash** (if present):
   - Check if already entered (duplicate detection)
   - Fill directly in table cell (no modal needed)
   - Stabilize: 1.2s + wait for masks
3. **Enter Modal Payments** (UPI, Cheque, NEFT):
   - Check if already entered (skip if found)
   - Click row (+) button → Modal opens
   - Fill: Payment Mode, Cheque No, Amount, Date, Transaction ID
   - Click "Add Payment" → "Save"
   - Wait for modal to close, re-locate row
4. **Screenshot** after all payments done
5. **Move to Next Invoice**

### Final Save
- Click "Save" button on main page (commits all entries)
- Handles empty page gracefully (all entries skipped/duplicates)
- If all payments were duplicates, final save may be skipped (normal behavior)

---

## Duplicate Detection

**Objective**: Safe to run multiple times without re-entering data

**How It Works**:
- Before entering any payment, checks if that payment mode already has a value in the table
- Reads column header to find correct column index
- Checks cell value in target row
- **Non-zero value** = Already entered → **Skip**
- **Zero/blank** = New entry → **Process**

**Logging**:
```
⏭  Cash already entered for IN17781827000153 — skipping.
⏭  Cheque already entered for IN17781827000152 — skipping.
```

---

## Key Features

✅ **Robust Error Handling**
- Automatic retry with force-click fallback
- Graceful handling of missing elements
- Clear error messages with diagnostics

✅ **Duplicate Detection**
- Safe repeat runs
- No accidental double-entry

✅ **Sequential Processing**
- One payment at a time (Cash, then modals)
- Portal stabilization between operations
- Isolated error handling per payment

✅ **Portal Adaptation**
- Handles Shadow DOM (uses bounding-box clicks)
- LWC-aware (keyboard.type() for date fields)
- Pagination support (forward-only)
- Spinner/mask wait detection

✅ **Detailed Logging**
- CSV log with timestamp, invoice, amount, status
- Per-invoice screenshots
- Sample table rows on mismatch
- Error traces for troubleshooting

---

## Current Scope & Limitations

- **Supported Portals**: Only Tata Consumer (non-OTP) portal currently active
- **File Formats**: `.xlsx` (requires `openpyxl`), `.csv` (default). Note: legacy `.xls` is not supported by `openpyxl`; convert to `.xlsx` or `.csv`.
- **Payment Types**: Cash (direct entry), UPI/Cheque/NEFT (modal entry)
- **Credit Notes**: Skipped (separate Add Credit Note flow)
- **OTP Portals**: Planned but not yet implemented
- **Parallel Processing**: Sequential only (one invoice at a time)

---

## Configuration

### `config.json`
```json
{
  "portal_url": "https://mavic-tataconsumer.my.site.com/dms/s/login/",
  "settlement_url": "https://mavic-tataconsumer.my.site.com/dms/s/manual-invoice-payment-settlement",
  "sheet_path": "D:/TATA_SHIPMENTS_April - Test.xlsx",
  "headless": false,
  "debug": false,
  "selected_portal": "tata_consumer"
}
```

### `.env`
```
PORTAL_USERNAME=your_username
PORTAL_PASSWORD=your_password
```

### Environment Variables (Optional)
- `PLAYWRIGHT_BROWSERS_PATH`: Custom Playwright browser location (used for PyInstaller exe)

---

## Build (Windows EXE)

This project includes a PyInstaller spec for producing a Windows GUI executable.

```bash
python -m PyInstaller PortalAutomationEngine.spec --noconfirm
```

Or run the convenience script on Windows:

```bat
build.bat
```

Note: When packaged as an `.exe`, the engine sets `PLAYWRIGHT_BROWSERS_PATH` automatically to use the user's installed Playwright browsers (typically under `%LOCALAPPDATA%\\ms-playwright`).

---

## Changelog & Recent Fixes

### Latest Fixes (2026-04-23)
- ✅ Fixed search date range to use **Bill Date only** (not Cheque Date)
  - Was causing incorrect invoice matches (e.g., 4/6 → 4/13 instead of 4/6 → 4/6)
- ✅ Improved spinner blocking detection with 1s pre-wait + retry logic
  - Added force-click fallback for blocked button clicks
- ✅ Added fallback for table visibility checks
  - Falls back to row count if CSS visibility check fails
- ✅ Graceful handling of empty save button
  - Final save wrapped in try/catch, succeeds even with 0 new entries (all duplicates)

---

## Known Issues & Fixes (2026-04-23)

### Issue 1: Incorrect Date Range for Search
**Problem**: Search range included cheque dates, returning wrong invoices (e.g., 4/6 → 4/13 instead of 4/6 → 4/6)
**Root Cause**: Date range calculation included both delivery and cheque dates
**Fix**: Changed to use **BILL DATE ONLY** for portal search range
**Result**: Now searches for invoices by invoice date, not payment dates

### Issue 2: Spinner Blocking Clicks
**Problem**: `<lightning-spinner>` element intercepted pointer events on search buttons
**Fix**: Added 1s pre-wait before click, increased timeout to 15s, added force-click retry
**Result**: Reliable button clicks even with loading spinners

### Issue 3: Table Visibility Check Failure
**Problem**: `.wait_for(state="visible")` failed even though rows existed (CSS display issue)
**Fix**: Fallback to `.count()` check; only error if truly 0 rows
**Result**: Works with both visible and hidden (CSS) rows

### Issue 4: Missing Save Button on Empty Page
**Problem**: Final Save button timeout when all entries skipped (nothing to save)
**Fix**: Wrapped save_page() in try/catch, graceful failure with informative message
**Result**: Completes successfully even with 0 new entries

---

## Performance

### Wait Times (Optimized)
- Pre-click waits: 500ms - 1s
- Modal operations: 2000ms
- Search API response: 8000ms
- Cash stabilization: 1200ms
- Modal save stabilization: 1500ms
- Page transitions: 1000ms - 2000ms

### Scalability
- Handles **100-150 invoices** in sequence
- No batch limits (sequential, not parallel)
- ~30-45 seconds per invoice (1-5 payments each)
- **Total time estimate**: 50-75 minutes for 100 invoices

### Batch Characteristics
- Sequential payment processing (one at a time)
- Isolated error handling (failures don't cascade)
- Safe re-runs with duplicate detection
- Pagination-aware (forward-only search)

---

## Testing

### Before Batch Run
- [ ] Excel has valid invoice numbers (exist in portal)
- [ ] Bill Date column populated
- [ ] Portal login works (credentials in .env)
- [ ] Settlement page loads, manual search works
- [ ] First run on 5 invoices completes successfully

### Verify Duplicate Detection
1. Run Excel once
2. Run same Excel again
3. All invoices should show: `⏭  {payment_mode} already entered — skipping.`
4. Final Save completes with 0 new entries

---

## CSV Log Format

**File**: `logs/run_tata_consumer_{timestamp}.csv`

**Columns**: `timestamp`, `invoice_no`, `payment_mode`, `amount`, `status`, `notes`

**Status Values**:
- `RUN STARTED` — Initial state
- `✅ success` — Payment entered successfully
- `⏭  duplicate` — Payment already entered, skipped
- `❌ error` — Failure with error message

**Example**:
```csv
2026-04-23 19:22:29,IN17781827000153,Cheque,3754,✅ success,"Modal Save clicked"
2026-04-23 19:22:45,IN17781827000152,UPI,1502,⏭  duplicate,"already entered"
```

---

## Troubleshooting

### No invoices found in search results
- Verify date range is correct (check portal manually)
- Confirm invoices exist in portal for those dates
- Check settlement page loads without errors

### Invoices found, but not matching
- Review first few rows in console output: `📋 First row cells: ...`
- Verify invoice numbers in Excel match portal
- Check if invoices are archived/hidden in portal

### Modal fields not found
- Enable debug mode: `"debug": true` in config.json
- Modal diagnostics printed to console
- Review screenshot `02_modal_*.png` for actual modal structure

### Click timeout on buttons
- Check for spinner/loading indicators blocking clicks
- Increase timeout values in config
- Try non-headless mode to visualize

### Save button not found (empty page)
- Normal if all entries are duplicates (already entered)
- Check logs to confirm entries were skipped
- Review CSV log for status

### Excel parsing fails
- For `.xlsx` files, ensure `openpyxl` is installed: `pip install openpyxl`
- Check date formats (supported: DD-Mon-YYYY, DD/MM/YYYY, YYYY-MM-DD)
- Unsupported dates throw: `Unrecognised date format`

---

## Architecture & Design

### Why Sequential Processing?
Each payment is processed one-at-a-time (not batched) to:
- Ensure portal stabilizes between operations
- Isolate errors to single payments
- Enable safe re-runs with duplicate detection
- Provide clear, individual logging per entry

### Why Bill Date Only?
Portal search uses invoice date (Bill Date), not payment dates:
- Cheque dates can be weeks/months later than invoice
- Including cheque dates expands search range incorrectly
- Bill Date matches portal's "Invoice From/To Date" filters
- Ensures correct set of invoices loaded

### Why Duplicate Detection?
Enables safe repeat runs and manual corrections:
- Same Excel can be processed multiple times
- Partial failures can be re-run without duplication
- Manual portal corrections won't be overwritten
- Scaling to 100-150 invoices safely

---

## Related Documentation
This repo is intentionally self-contained (see `main.py` and `app.py`).

---

## Future Enhancements
- Parallel invoice processing (currently sequential)
- Resumable runs (save/restore state)
- Configurable payment type priority
- Advanced retry logic with exponential backoff
- Real-time progress webhook/API
- Support for additional portals (OTP-enabled)

---

**Last Updated**: 2026-04-23  
**Status**: Production-ready for 100-150 invoice batch runs
