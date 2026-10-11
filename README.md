# StatementHub

StatementHub is a local Windows application for searching bank statements from HDFC Bank, Canara Bank, ICICI Bank, and similar CSV, XLS, or XLSX exports in one place. Statements are processed on your own computer in the browser window opened by the app; they are not sent to an online account or cloud service.

## What you need

- A Windows computer
- Python 3.10 or newer
- Bank statement files in `.csv`, `.xls`, or `.xlsx` format

## First-time setup

1. Download and install Python from [python.org](https://www.python.org/downloads/windows/).
2. During Python installation, tick **Add Python to PATH**.
3. Open this application folder.
4. Double-click `run_app.bat`.
5. On the first run, the required local libraries install automatically. A browser window then opens the app.

Keep the small command window opened by `run_app.bat` running while you use the application. Closing it stops the app.

## Start the app later

Double-click `run_app.bat` again. The browser normally opens automatically. If it does not, open the local address shown in the command window, usually `http://localhost:8501`.

## Load statements & daily downloads

1. Select **Add or replace statements** in the browser app.
2. Select all statements you want to search. CSV, XLS, and XLSX files can be selected together.
3. When connected to Supabase:
   - Statements uploaded daily with overlapping dates are **automatically deduplicated**.
   - Only new transactions are inserted; existing ones are skipped.
   - All historical transactions remain stored and searchable across sessions.
4. When running in local mode (without Supabase):
   - Statements load into your current browser session.

## Cloud database & persistence (Supabase)

StatementHub can connect to **Supabase** (free 500MB cloud PostgreSQL) for persistent storage and multi-device access:

1. Create a free account at [supabase.com](https://supabase.com).
2. Open your project's **SQL Editor** and run the schema script provided in [`db.py`](file:///d:/Codex/StatementHub/db.py) (or in the app's setup guide).
3. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`:
   ```toml
   [supabase]
   url = "https://your-project-id.supabase.co"
   key = "your-supabase-key"
   ```
4. Restart the app. The top header will display `🟢 Supabase Connected`.

## Search transactions

Choose one checkbox above the search field. Selecting a checkbox automatically turns off the other two, so only the chosen type is searched.

### UTR / Reference

Choose **UTR / Reference** to search a UTR, UPI reference, NEFT/IMPS/RTGS reference, cheque number, or transaction ID.

- Start typing part of a UTR, such as `3964`.
- Matching full identifiers appear in the same search dropdown.
- Select a suggestion or enter the full reference.
- Canara references such as `CNRBH00157703028` are recognized as transaction identifiers.

### Date

Choose **Date** and enter a date in `DD-MM-YYYY` format, for example `01-09-2026`. The results show every transaction from that day. Dates in the results table also use `DD-MM-YYYY` format.

### Amount

Choose **Amount** and enter an exact transaction amount, for example `25000`, `25,000`, or `₹25,000`. The results show all debit and credit transactions with that exact amount.

## Understanding the results & Bank color coding

The transaction table color-codes rows by the issuing bank for instant visual differentiation:
- **🟢 Canara Bank**: Subtle emerald green tint
- **🟠 ICICI Bank**: Subtle warm coral orange tint
- **🔵 HDFC Bank**: Subtle modern blue tint

A live **Bank Legend** badge bar above the table displays active bank counts. Use **Download displayed results as Excel** to save the current search results.

## Monthly Profit & Cash Flow Dashboard

Switch to the **💰 Monthly Profit & Cash Flow** tab to track your net financial health over time:

- **Top KPI Cards**: Real-time totals for **Total Received (Credits)**, **Total Sent (Debits)**, **Net Profit / Surplus**, and **Profit Margin %**.
- **Interactive Inflows vs. Outflows Chart**: Month-by-month comparative bar chart (Green for Money In, Red for Money Out).
- **Monthly Profit Breakdown Table**: Tabular statement for each month showing Received, Sent, Net Profit/Loss, Margin %, and status (🟢 Profit / 🔴 Deficit).
- **Bank-Wise Monthly Breakdown**: Inspect how each bank (Canara, ICICI, HDFC) contributed to incoming and outgoing cash each month.
- **Excel Report Download**: One-click download of the complete monthly P&L and bank breakdown workbook.
- **Search Screen Quick Ribbon**: Real-time summary cards right under the bank & period filters on the search screen.

## Assign business or party names

You can add or update business and party names directly in the table row:

- Double-click the **Party Name** cell for any transaction.
- **Instant Auto-Completion**: As you start typing (e.g. `PAV`), matching saved party names (e.g. `PAVANI M/S VENKATAPUR`) appear right below the cell. Select with mouse or press **Enter** / **Tab**.
- **Add New Parties on the Fly**: If a party name is not yet registered, simply type it into the cell and press **Enter**. It will save to that transaction and automatically be added to your saved party names list for future auto-completion.
- **Stored Registry**: Open **Stored Party Names** below the table to see all registered party names and transaction counts.
- Clearing the text in the cell removes the assigned party name from that transaction.

Assigned party names are persisted in Supabase and locally in `~/.statementhub/stored_parties.json`. Excel downloads include the assigned party names.

## Duplicate identifiers

When the same nonblank UTR/reference appears more than once in the loaded statements, every matching row is marked `Duplicate UTR = true`. The app keeps all transactions; it never silently removes duplicates.

## Supported statement layouts

The importer finds the transaction-table header automatically and ignores account cover pages, totals, blank lines, and summaries. It normalizes different layouts into one shared format.

| Statement layout | Recognized transaction fields |
| --- | --- |
| HDFC XLS | `Date`, `Narration`, `Chq./Ref.No.`, `Withdrawal Amt.`, `Deposit Amt.`, `Closing Balance` |
| Canara CSV | `Txn Date`, `Value Date`, `Cheque No.`, `Description`, `Debit`, `Credit`, `Balance` |
| ICICI XLSX | `Value Date`, `Tran. Id`, `Cheque. No./Ref. No.`, `Transaction Remarks`, `Withdrawal Amt`, `Deposit Amt`, `Balance` |

Other bank statements may work if they contain a transaction date, description/narration, and debit or credit amount columns.

## Troubleshooting

- **`py` is not recognized:** Reinstall Python and select **Add Python to PATH**, then restart the computer.
- **A statement does not load:** Confirm it is a CSV, XLS, or XLSX statement export and that it includes transaction date, description, and debit/credit columns.
- **The app stops responding:** Keep the command window opened by `run_app.bat` running. Close it only when you have finished.
- **A date finds no results:** Use `DD-MM-YYYY`, for example `01-09-2026`.
