import argparse
import pyodbc
import blpapi
import pandas as pd
import re
import textwrap
from datetime import datetime
from pandas import ExcelWriter
import io
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime
import matplotlib.dates as mdates
from pathlib import Path

# ==== Brand Color =====
COLOR_PRIMARY = "#2D3D5D"

# --- Bloomberg Suffix Mapping ---
BBG_SUFFIX_MAP = {
    "US": "US", "CA": "CN", "MX": "MM", "GB": "LN", "IE": "ID", "FR": "FP",
    "DE": "GR", "CH": "SW", "IT": "IM", "ES": "SM", "NL": "NA", "SE": "SS",
    "NO": "NO", "FI": "FH", "DK": "DC", "BE": "BB", "AT": "AV", "AU": "AU",
    "NZ": "NZ", "JP": "JP", "HK": "HK", "KR": "KS", "CN": "CH", "SG": "SP",
    "PL": "PW", "PT": "PL", "LU": "LX", "GR": "GA", "BR": "BZ", "IN": "IN",
    "IL": "IT", "SA": "AB", "ZA": "SJ", "RU": "RM", "ID": "IJ", "MY": "MK",
    "TH": "TB", "PH": "PM", "EG": "EY", "AR": "AR", "TR": "TI", "TW": "TT",
    "CZ": "CP", "HU": "HB", "SK": "SK", "UA": "UX", "VN": "VN", "UG": "UG"
}

# --- Argument Parsing ---
parser = argparse.ArgumentParser(description="Bloomberg Option and Stock Gain Excel Generator")
parser.add_argument("--symbol", required=True, help="Base ticker symbol (e.g. HBM, VIX)")
parser.add_argument("--country", required=True, help="Underlying country (e.g. US, CA)")
parser.add_argument("--instrument", choices=["Stock", "Index"], required=True, help="Instrument type: Stock or Index")
parser.add_argument("--start", required=True, help="Start Date (YYYYMMDD)")
parser.add_argument("--end", required=True, help="End Date (YYYYMMDD)")
parser.add_argument("--stock", choices=["Y", "N"], default="N", help="Include Stock Gain Calculation (Y/N)")
parser.add_argument("--portfolio", required=True, help="Portfolio to search for (e.g. Voyager Fund)")
args = parser.parse_args()

# --- User Inputs ---
option_base = args.symbol.strip().upper()
instrument_type = args.instrument.strip().capitalize()
underlying_country = args.country.strip().upper()
start = args.start.strip()
end = args.end.strip()
do_stock_calc = args.stock.upper()
portfolio = args.portfolio.strip()
portfolio = args.portfolio.strip()
use_grouping = portfolio.upper() == "ALL"
portfolio_filter = "%%" if use_grouping else f"%{portfolio}%"
# --- Get script directory for saving files ---
script_dir = os.path.dirname(os.path.abspath(__file__))



regex_pattern = f"^(?:[1-6])?{option_base}$"

# --- Construct Bloomberg Ticker ---
if instrument_type == "Index":
    bpipe_ticker = f"{option_base} Index"
    bloomberg_ticker = option_base
else:
    bbg_suffix = BBG_SUFFIX_MAP.get(underlying_country)
    if not bbg_suffix:
        raise Exception(f"No Bloomberg suffix found for country: {underlying_country}")
    bpipe_ticker = f"{option_base} {bbg_suffix} Equity"
    bloomberg_ticker = f"{option_base}.{bbg_suffix}"

base_stock = bloomberg_ticker

# --- Connect to Database ---
cnxn = pyodbc.connect("")   # put the password here 
cursor = cnxn.cursor()








# --- Fetch Open Option Positions at Start Date ---
cursor.execute("""
    SELECT SECURITY
    FROM psc_position_history
    WHERE POSN_DATE_INT = ?
      AND PORTFOLIO LIKE ?
      AND COMPANY_SYMBOL REGEXP ?
      AND UNDERLYING_COUNTRY = ?
""", start, portfolio_filter, regex_pattern, underlying_country)

open_options_start = {row.SECURITY.strip().upper() for row in cursor.fetchall()}

# --- Fetch Open Option Positions at End Date ---
cursor.execute("""
    SELECT SECURITY
    FROM psc_position_history
    WHERE POSN_DATE_INT = ?
      AND PORTFOLIO LIKE ?
      AND COMPANY_SYMBOL REGEXP ?
      AND UNDERLYING_COUNTRY = ?
""", end, portfolio_filter, regex_pattern, underlying_country)

open_options_end = {row.SECURITY.strip().upper() for row in cursor.fetchall()}

# --- Combine both start and end open positions ---
open_options = open_options_start.union(open_options_end)
# --- Option Gain Calculation ---
option_gain = 0.0
used_options = []

if portfolio.upper() == "ALL":
    cursor.execute("""
        SELECT atr.SECURITY,
               atr.TRADE_DATE_INT,
               SUM(atr.SETTLE_CCY_AMT) AS TOTAL_AMT
        FROM psc_all_transactions atr
        WHERE atr.SECURITY_TYPE LIKE '%Option%'
          AND COMPANY_SYMBOL REGEXP ?
          AND atr.UNDERLYING_COUNTRY = ?
          AND atr.TRADE_DATE_INT BETWEEN ? AND ?
        GROUP BY atr.SECURITY, atr.TRADE_DATE_INT
        ORDER BY atr.TRADE_DATE_INT
    """, regex_pattern, underlying_country, start, end)

    rows = cursor.fetchall()
    print(f"\nTotal option transactions fetched: {len(rows)}")

    for row in rows:
        sec = row.SECURITY.strip().upper()
        amt = float(row.TOTAL_AMT)

        if sec not in open_options:
            print(f"[INCLUDED] Option: {sec} | Amt: {amt:.2f}")
            option_gain += amt
            used_options.append(sec)
        else:
            print(f"[EXCLUDED - OPEN] Option: {sec}")
else:
    cursor.execute("""
        SELECT SECURITY, TRANSACTION, SETTLE_CCY_AMT
        FROM psc_all_transactions
        WHERE TRADE_DATE_INT BETWEEN ? AND ?
          AND PORTFOLIO LIKE ?
          AND SECURITY_TYPE LIKE '%EquityOption%'
          AND COMPANY_SYMBOL REGEXP ?
          AND UNDERLYING_COUNTRY = ?
    """, start, end, portfolio_filter, regex_pattern, underlying_country)

    rows = cursor.fetchall()
    print(f"\nTotal option transactions fetched: {len(rows)}")

    for row in rows:
        sec = row.SECURITY.strip().upper()
        txn = row.TRANSACTION.upper()
        amt = float(row.SETTLE_CCY_AMT)

        if sec not in open_options:
            print(f"[INCLUDED] Option: {sec} | Txn: {txn} | Amt: {amt:.2f}")
            option_gain += amt
            used_options.append(sec)
        else:
            print(f"[EXCLUDED - OPEN] Option: {sec} | Txn: {txn}")

# --- Find First Option Trade Date ---
cursor.execute("""
    SELECT TRADE_DATE_INT
    FROM psc_all_transactions
    WHERE PORTFOLIO LIKE ?
      AND TRANSACTION_TYPE COLLATE utf8mb4_unicode_ci <> 'CORPORATE_ACTION'
      AND SECURITY_TYPE LIKE '%EquityOption%'
      AND COMPANY_SYMBOL REGEXP ?
      AND UNDERLYING_COUNTRY = ?
      AND TRADE_DATE_INT BETWEEN ? AND ?
    ORDER BY TRADE_DATE_INT ASC
    LIMIT 1
""", portfolio_filter, regex_pattern, underlying_country, start, end)


row = cursor.fetchone()
first_option_date = row.TRADE_DATE_INT if row else None

# --- Initial Investment Calculation (Buy only) ---
option_initial_investment = 0.0

if first_option_date:
    if portfolio.upper() == "ALL":
        cursor.execute("""
            SELECT atr.SECURITY,
                   atr.TRADE_DATE_INT,
                   SUM(atr.SETTLE_CCY_AMT) AS INITIAL_INVEST
            FROM psc_all_transactions atr
            WHERE atr.SECURITY_TYPE LIKE '%Option%'
              AND COMPANY_SYMBOL REGEXP ?
              AND atr.UNDERLYING_COUNTRY = ?
              AND atr.TRADE_DATE_INT = ?
            GROUP BY atr.SECURITY, atr.TRADE_DATE_INT
        """, regex_pattern, underlying_country, first_option_date)

        for row in cursor.fetchall():
            option_initial_investment += float(row.INITIAL_INVEST)
            print(option_initial_investment)
    else:
        cursor.execute("""
            SELECT SETTLE_CCY_AMT
            FROM psc_all_transactions
            WHERE TRADE_DATE_INT = ?
            AND PORTFOLIO LIKE ?
            AND SECURITY_TYPE LIKE '%EquityOption%'
            AND COMPANY_SYMBOL REGEXP ?
            AND UNDERLYING_COUNTRY = ?
        """, first_option_date, portfolio_filter, regex_pattern, underlying_country)

        for row in cursor.fetchall():
            option_initial_investment += float(row.SETTLE_CCY_AMT)
            print(option_initial_investment)

# --- Final IRR ---
option_irr = (option_gain / option_initial_investment) * 100 if option_initial_investment else 0.0

# Print used options
print("\nOptions used in gain calculation:")
for sec in sorted(set(used_options)):
    print(f" - {sec}")

option_result_label = "OptionGain" if option_gain >= 0 else "OptionLoss"
















# --- Bloomberg Session Initialization ---
sessionOptions = blpapi.SessionOptions()
sessionOptions.setServerHost("localhost")
sessionOptions.setServerPort(8194)
session = blpapi.Session(sessionOptions)

if not session.start():
    raise Exception("Failed to start Bloomberg session.")
if not session.openService("//blp/refdata"):
    raise Exception("Failed to open Bloomberg refdata service.")

service = session.getService("//blp/refdata")

# --- Bloomberg Price Fetch ---
request = service.createRequest("HistoricalDataRequest")
request.getElement("securities").appendValue(bpipe_ticker)
request.getElement("fields").appendValue("PX_LAST")
request.set("startDate", start)
request.set("endDate", end)
request.set("periodicitySelection", "DAILY")
session.sendRequest(request)

# --- Collect Bloomberg Prices ---
data = []
while True:
    event = session.nextEvent()
    for msg in event:
        if msg.hasElement("securityData"):
            fieldDataArray = msg.getElement("securityData").getElement("fieldData")
            for i in range(fieldDataArray.numValues()):
                fieldData = fieldDataArray.getValueAsElement(i)
                date = fieldData.getElementAsDatetime("date")
                px_last = fieldData.getElementAsFloat("PX_LAST")
                data.append({"Date": date, "Close": px_last})
    if event.eventType() == blpapi.Event.RESPONSE:
        break

price_df = pd.DataFrame(data)

if price_df.empty:
    print(f" No Bloomberg price data found for {bpipe_ticker}.")

    # Ask for manual ticker input from frontend (GUI)
    root = tk.Tk()
    root.withdraw()  # Hide root window
    fallback_ticker = simpledialog.askstring(
        title="Manual Bloomberg Ticker",
        prompt="Bloomberg could not find the ticker.\nPlease enter the full Bloomberg ticker (e.g. VIX Index or HBM CN Equity):"
    )
    root.destroy()

    if not fallback_ticker:
        print(" No manual ticker entered. Exiting.")
        exit()

    # Try again with manual ticker
    request = service.createRequest("HistoricalDataRequest")
    request.getElement("securities").appendValue(fallback_ticker.strip())
    request.getElement("fields").appendValue("PX_LAST")
    request.set("startDate", start)
    request.set("endDate", end)
    request.set("periodicitySelection", "DAILY")
    session.sendRequest(request)

    data = []
    while True:
        event = session.nextEvent()
        for msg in event:
            if msg.hasElement("securityData"):
                fieldDataArray = msg.getElement("securityData").getElement("fieldData")
                for i in range(fieldDataArray.numValues()):
                    fieldData = fieldDataArray.getValueAsElement(i)
                    date = fieldData.getElementAsDatetime("date")
                    px_last = fieldData.getElementAsFloat("PX_LAST")
                    data.append({"Date": date, "Close": px_last})
        if event.eventType() == blpapi.Event.RESPONSE:
            break

    price_df = pd.DataFrame(data)

    if price_df.empty:
        print(" Still no price data found. Exiting.")
        exit()

price_df['Date'] = pd.to_datetime(price_df['Date'])


# --- Stock Gain (Optional)---
if do_stock_calc == 'Y':
    print("\n" + "="*60)
    print("         STOCK GAIN CALCULATION DEBUG")
    print("="*60)

    # --- Start Position ---
    cursor.execute("""
        SELECT QUANTITY, CLOSE_PRICE
        FROM psc_position_history
        WHERE POSN_DATE_INT = ?
          AND PORTFOLIO LIKE ?
          AND COMPANY_SYMBOL REGEXP ?
          AND UNDERLYING_COUNTRY = ?
          AND SECURITY_TYPE LIKE '%Stock%'
    """, start, portfolio_filter, regex_pattern, underlying_country)
    row_start = cursor.fetchone()
    quantity_start = float(row_start.QUANTITY) if row_start else 0.0
    close_price_start = float(row_start.CLOSE_PRICE) if row_start else 0.0
    value_at_start = quantity_start * close_price_start

    print(f"\n--- START POSITION ({start}) ---")
    print(f"  Quantity Start     : {quantity_start:,.4f}")
    print(f"  Close Price Start  : ${close_price_start:,.4f}")
    print(f"  Value at Start     : ${value_at_start:,.2f}")
    if not row_start:
        print("  *** WARNING: No start position found in DB! Defaulting to 0 ***")

    # --- End Position ---
    cursor.execute("""
        SELECT QUANTITY, CLOSE_PRICE
        FROM psc_position_history
        WHERE POSN_DATE_INT = ?
          AND PORTFOLIO LIKE ?
          AND COMPANY_SYMBOL REGEXP ?
          AND UNDERLYING_COUNTRY = ?
          AND SECURITY_TYPE LIKE '%Stock%'
    """, end, portfolio_filter, regex_pattern, underlying_country)
    row_end = cursor.fetchone()
    quantity_end = float(row_end.QUANTITY) if row_end else 0.0
    close_price_end = float(row_end.CLOSE_PRICE) if row_end else 0.0

    # Fallback to Bloomberg price if no DB record
    if close_price_end == 0.0 and not price_df.empty:
        close_price_end = price_df.iloc[-1]['Close']
        print(f"\n  *** WARNING: No end position price in DB! Using Bloomberg last price: ${close_price_end:,.4f} ***")

    value_at_end = quantity_end * close_price_end

    print(f"\n--- END POSITION ({end}) ---")
    print(f"  Quantity End       : {quantity_end:,.4f}")
    print(f"  Close Price End    : ${close_price_end:,.4f}")
    print(f"  Value at End       : ${value_at_end:,.2f}")
    if not row_end:
        print("  *** WARNING: No end position found in DB! Defaulting to 0 ***")

    # --- Cash Transactions ---
    cursor.execute("""
        SELECT TRADE_DATE_INT, SETTLE_CCY_AMT
        FROM psc_all_transactions
        WHERE TRADE_DATE_INT BETWEEN ? AND ?
        AND PORTFOLIO LIKE ?
        AND COMPANY_SYMBOL REGEXP ?
        AND UNDERLYING_COUNTRY = ?
        AND SECURITY_TYPE LIKE '%Stock%'
    """, start, end, portfolio_filter, regex_pattern, underlying_country)

    cash_rows = cursor.fetchall()
    cash = 0.0

    print(f"\n--- STOCK TRANSACTIONS ({start} to {end}) ---")
    if not cash_rows:
        print("  *** WARNING: No stock transactions found in this period! ***")
    for r in cash_rows:
        amt = float(r.SETTLE_CCY_AMT)
        cash += amt
        print(f"  Date: {r.TRADE_DATE_INT}  |  SETTLE_CCY_AMT: ${amt:,.2f}")

    print(f"\n  Total Net Cash from Trades : ${cash:,.2f}")

    # --- Stock Gain ---
    stock_gain = (value_at_end - value_at_start) + cash

    print(f"\n--- STOCK GAIN SUMMARY ---")
    print(f"  Value at End       : ${value_at_end:,.2f}")
    print(f"  Value at Start     : ${value_at_start:,.2f}")
    print(f"  Net Cash           : ${cash:,.2f}")
    print(f"  Stock Gain Formula : ({value_at_end:,.2f} - {value_at_start:,.2f}) + {cash:,.2f}")
    print(f"  TOTAL STOCK GAIN   : ${stock_gain:,.2f}")

    # --- Initial Investment ---
    cursor.execute("""
        SELECT TRADE_DATE_INT
        FROM psc_all_transactions
        WHERE PORTFOLIO LIKE ?
        AND SECURITY_TYPE LIKE '%Stock%'
        AND COMPANY_SYMBOL REGEXP ?
        AND UNDERLYING_COUNTRY = ?
        ORDER BY TRADE_DATE_INT ASC
        LIMIT 1
    """, portfolio_filter, regex_pattern, underlying_country)

    row = cursor.fetchone()
    first_stock_date = row.TRADE_DATE_INT if row else None

    stock_initial_investment = 0.0
    if first_stock_date:
        print(f"\n--- INITIAL INVESTMENT (First Trade Date: {first_stock_date}) ---")
        cursor.execute("""
            SELECT TRADE_DATE_INT, SETTLE_CCY_AMT
            FROM psc_all_transactions
            WHERE TRADE_DATE_INT = ?
            AND PORTFOLIO LIKE ?
            AND SECURITY_TYPE LIKE '%Stock%'
            AND COMPANY_SYMBOL REGEXP ?
            AND UNDERLYING_COUNTRY = ?
        """, first_stock_date, portfolio_filter, regex_pattern, underlying_country)

        for row in cursor.fetchall():
            amt = float(row.SETTLE_CCY_AMT)
            stock_initial_investment += amt
            print(f"  Date: {row.TRADE_DATE_INT}  |  SETTLE_CCY_AMT: ${amt:,.2f}")

        print(f"  Total Initial Investment   : ${stock_initial_investment:,.2f}")
        stock_irr = (stock_gain / stock_initial_investment) * 100 if stock_initial_investment != 0 else 0.0
        print(f"  Stock IRR                  : {stock_irr:.2f}%")
    else:
        print("  *** WARNING: No first stock trade date found! IRR set to 0 ***")
        stock_irr = 0.0

    print("="*60 + "\n")

else:
    stock_irr = None


# --- Option transactions for plotting (unchanged except filter for open options) ---
cursor.execute("""
    SELECT SETTLE_CCY_AMT, SECURITY, TRADE_DATE_INT, TRANSACTION_TYPE, TRANSACTION
    FROM psc_all_transactions
    WHERE TRADE_DATE_INT BETWEEN ? AND ?
      AND COMPANY_SYMBOL REGEXP ?
      AND UNDERLYING_COUNTRY = ?
      AND PORTFOLIO LIKE ?
      AND SECURITY_TYPE LIKE '%EquityOption%'
""", start, end, regex_pattern, underlying_country, portfolio_filter)

rows = [tuple(row) for row in cursor.fetchall()]
df = pd.DataFrame(rows, columns=["SETTLE_CCY_AMT", "SECURITY", "TRADE_DATE_INT", "TRANSACTION_TYPE", "TRANSACTION"])
df["Date"] = pd.to_datetime(df["TRADE_DATE_INT"].astype(str))
df["SECURITY"] = df["SECURITY"].str.strip().str.upper()
df = df[~df["SECURITY"].isin(open_options)]
df = df.sort_values("Date")





# --- Filter for Options ---
def is_option_like(sec):
    return bool(re.search(r'\d{2}/\d{2}/\d{2}', sec))

option_trades = df[df['SECURITY'].apply(is_option_like)].copy()

def simplify_transaction(txn):
    txn = txn.upper()
    if "SELL" in txn:
        return "SELL"
    if "BUY" in txn:
        return "BUY"
    return txn

def shorten_label(row):
    sec = row['SECURITY']
    txn = simplify_transaction(row['TRANSACTION'])
    match = re.search(r'([PC])(\d+(?:\.\d+)?)', sec)
    if match:
        opt_type = match.group(1)
        strike = match.group(2)
        return f"{txn}/{opt_type}{strike}"
    return f"{txn} (OPT)"

if not option_trades.empty:
    option_trades['Label'] = option_trades.apply(shorten_label, axis=1)
    # group but join with newlines instead of commas
    option_trades_grouped = (
        option_trades
        .groupby('Date')['Label']
        .apply(lambda x: '\n'.join(sorted(set(x))))
        .reset_index()
    )

else:
    option_trades_grouped = pd.DataFrame(columns=['Date', 'Label'])

merged = pd.merge(price_df, option_trades_grouped, on='Date', how='left')


# --- Plotting ---
plt.rcParams['font.family'] = "Franklin Gothic Book"

fig, ax = plt.subplots(figsize=(29.34, 12.87), dpi=600)
fig.subplots_adjust(left=0.04, right=0.96, top=0.88, bottom=0.12)

y_min = merged['Close'].min()
y_max = merged['Close'].max()
y_range = y_max - y_min
ax.set_ylim(y_min - 0.45 * y_range, y_max + 0.45 * y_range)

# Plot price line
ax.plot(merged['Date'], merged['Close'], label=f'{bloomberg_ticker} Closing Price', color=COLOR_PRIMARY, linewidth=4)

# --- Offsets cycle: long up, short down, short up, long down
offset_pattern = [130, -60, 60, -140]

labels_df = merged[merged['Label'].notna()]

for i, (_, row) in enumerate(labels_df.iterrows()):
    date = row['Date']
    label = row['Label']
    y = row['Close']

    y_offset = offset_pattern[i % len(offset_pattern)]  

    ax.annotate(
        label,
        xy=(date, y),
        xytext=(0, y_offset),
        textcoords='offset points',
        ha='center',
        va='bottom' if y_offset > 0 else 'top',
        fontsize=16,
        fontname="Franklin Gothic Book",
        color=COLOR_PRIMARY,
        alpha=0.95,
        clip_on=True,
        arrowprops=dict(
            arrowstyle='-',
            lw=2.0,
            color="#C6A373",
            alpha=0.9,
            shrinkA=0, shrinkB=0,
        ),
    )


# --- Title ---
start_fmt = datetime.strptime(start, "%Y%m%d").strftime("%Y-%m-%d")
end_fmt = datetime.strptime(end, "%Y%m%d").strftime("%Y-%m-%d")
line1 = f"{bloomberg_ticker}   |   {start_fmt} to {end_fmt}"

option_irr_str = f"Total Option IRR: {'' if option_irr == 0 else f'{option_irr:.2f}%'}"


# Format option gain with dollar sign
option_gain_str = f"Option Gain: ${option_gain:,.2f}"
# Format option gain/loss — escape $ to avoid matplotlib math mode
if option_gain < 0:
    option_gain_str = f"Option Loss: -\\${abs(option_gain):,.2f}"
elif option_gain > 0:
    option_gain_str = f"Option Gain: \\${option_gain:,.2f}"
else:
    option_gain_str = "Option Gain: \\$0.00"

# Format stock gain/loss — escape $ and fix spacing
if stock_irr is not None:
    if stock_gain < 0:
        stock_gain_str = f"Stock Loss: -\\${abs(stock_gain):,.2f}"
    elif stock_gain > 0:
        stock_gain_str = f"Stock Gain: \\${stock_gain:,.2f}"
    else:
        stock_gain_str = "Stock Gain: \\$0.00"
    line2 = f"{option_gain_str}    |    {stock_gain_str}"
else:
    line2 = option_gain_str
    
full_title = f"{line1}\n{line2}"

# --- Axes ---
ax.set_title(full_title, fontsize=18, fontweight='normal', fontname="Franklin Gothic Book", color=COLOR_PRIMARY)
ax.set_xlabel('Date', fontsize=16, fontweight='normal', color=COLOR_PRIMARY)
ax.set_ylabel(f'{bloomberg_ticker} Spot Price', fontsize=16, fontweight='normal', color=COLOR_PRIMARY)

for lbl in ax.get_xticklabels() + ax.get_yticklabels():
    lbl.set_color(COLOR_PRIMARY)    
    lbl.set_fontweight('normal')

ax.grid(True, axis='y', linestyle='--', alpha=0.25, color=COLOR_PRIMARY)  # only horizontal lines





# Add timestamp for uniqueness
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# --- Excel Export ---
excel_output_path = os.path.join(script_dir, f"{option_base}_{underlying_country}_{portfolio}_{start}_{end}_graph.xlsx")

with ExcelWriter(excel_output_path, engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet('Stock Price Chart')

    # Save the plot to a buffer
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    buf.seek(0)

    worksheet.insert_image(
        'A1',
        'stock_chart.png',
        {
            'image_data': buf,
            'x_scale': 0.7,
            'y_scale': 0.7
        }
    )

    worksheet.set_column('A:A', 2)
    worksheet.set_row(0, 15)

print(f"Excel file with graph saved as: {excel_output_path}")

# --- Auto Open Excel File (Windows only) ---
try:
    if os.name == 'nt':
        os.startfile(excel_output_path)
    else:
        subprocess.call(['open', excel_output_path]) 
except Exception as e:
    print(f"Could not open the Excel file automatically: {e}")



# Format x-axis labels as Mon-YY (e.g. Sep-23)
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%y'))

# Optional: set major ticks to monthly
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))

# Make x and y tick labels larger and consistent
ax.tick_params(axis='x', labelsize=18)  # increase font size for x-axis
ax.tick_params(axis='y', labelsize=18)  # increase font size for y-axis

# Apply font to tick labels
for lbl in ax.get_xticklabels() + ax.get_yticklabels():
    lbl.set_fontname("Franklin Gothic Book")
    lbl.set_fontweight('normal')
    lbl.set_color(COLOR_PRIMARY)

# --- PDF Export ---
pdf_output_path = os.path.join(script_dir, f"{option_base}_{underlying_country}_{portfolio}_{start}_{end}_{timestamp}_graph.pdf")
fig.savefig(pdf_output_path, format='pdf', dpi=300, bbox_inches='tight')
print(f"PDF file with graph saved as: {pdf_output_path}")

# --- Auto Open PDF File (Windows only) ---
try:
    if os.name == 'nt':
        os.startfile(pdf_output_path)
    else:
        subprocess.call(['open', pdf_output_path]) 
except Exception as e:
    print(f"Could not open the PDF file automatically: {e}")






"""
==================== Program Overview ====================

This program calculates and visualizes the investment performance of options 
and stocks for a given security, portfolio, and date range. It integrates SQL 
transaction/position data with Bloomberg historical prices and produces both 
numerical IRR results and annotated price charts.

--- Option Gain & IRR ---
1. Fetches all open option positions at the start and end dates to exclude 
   them from gain calculations (only closed trades count).
2. Aggregates option transactions between the start and end dates, filtering 
   out open positions. The net sum of SETTLE_CCY_AMT is the total option gain.
3. Determines the first option trade date and sums all transactions on that 
   day to calculate the initial investment.
4. Computes Option IRR as:
       (Option Gain / Initial Investment) * 100

--- Stock Gain & IRR ---
1. Fetches the quantity of shares and closing price at the start and end dates 
   to determine the portfolio’s starting and ending position values.
2. Sums all stock trade cash flows (SETTLE_CCY_AMT) between the start and end 
   dates to account for realized trades.
3. Calculates stock gain as:
       (Value at End – Value at Start) + Net Cash from Trades
4. Identifies the first stock trade date and sums all SETTLE_CCY_AMT on that day 
   to determine initial investment.
5. Computes Stock IRR as:
       (Stock Gain / Initial Investment) * 100

--- Bloomberg Price Data ---
- Connects to Bloomberg via blpapi to fetch daily closing prices for the chosen ticker.
- If the ticker fails, prompts the user to manually input a fallback Bloomberg ticker.

--- Visualization & Export ---
- Merges option trade labels with Bloomberg price data and generates a line chart 
  with annotated trades.
- Includes Option IRR and Stock IRR in the chart title.
- Exports the results into both Excel (with embedded chart) and PDF files.
- Attempts to auto-open the PDF for convenience.



If you just want the option gain and stock gain printed on the screen, you can remove 
the IRR calculation and change the title to ‘Stock Gain’ and ‘Option Gain,’ since we 
already calculate those for the IRR and you can simply reuse them.
"""
