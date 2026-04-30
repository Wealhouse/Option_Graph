# 📊 Program Overview

This program calculates and visualizes the **investment performance of options and stocks** for a given security, portfolio, and date range.

It integrates:
- SQL transaction and position data  
- Bloomberg historical prices  

…and produces both **numerical IRR results** and **annotated price charts**.

---

## 📈 Option Gain & IRR

1. Fetches all **open option positions** at the start and end dates to exclude them *(only closed trades are counted)*.  
2. Aggregates option transactions within the date range.  
   - The net sum of `SETTLE_CCY_AMT` = **total option gain**  
3. Identifies the **first option trade date** to compute **initial investment**  
4. Calculates IRR:  
   **(Option Gain / Initial Investment) × 100**

---

## 📉 Stock Gain & IRR

1. Fetches **share quantity and price** at start and end dates  
2. Sums all stock trade cash flows (`SETTLE_CCY_AMT`)  
3. Calculates stock gain:  
   **(Value at End – Value at Start) + Net Cash from Trades**  
4. Uses first trade date to compute **initial investment**  
5. Calculates IRR:  
   **(Stock Gain / Initial Investment) × 100**

---

## 📡 Bloomberg Price Data

- Connects via `blpapi` to fetch **daily closing prices**  
- Prompts for a fallback ticker if the original fails  

---

## 📊 Visualization & Export

- Merges trade data with price data  
- Generates **annotated price charts**  
- Displays **Option IRR and Stock IRR** in the title  
- Exports results to:
  - 📄 PDF (auto-opens if possible)  
  - 📊 Excel (with embedded chart)  

---

## 💡 Optional

If you only need gains (no IRR):
- Remove IRR calculations  
- Rename outputs to:
  - **Stock Gain**
  - **Option Gain**
