"""
Monthly Trade Volume Summary Generator
========================================
Reads all order-audit Excel files (the "partXofY" files) from the "Excel Files"
folder, aggregates buy/sell trade volume (Qty x Gross Fill Price) by ISIN for
each calendar month, and writes one output workbook to the "Output" folder.

- Grouping key: ISIN (column D). Instrument Description (column C) is carried
  along for readability. Differences in Instrument Code (column B) - e.g.
  buy-in / cash market / main market variants - are ignored, per spec.
- Month is derived from Trade Date (column AB, format yyyy-mm-dd).
- September is always excluded from the summary.
- Output: one sheet ("Monthly Volume Summary") with one stacked table per
  month (ISIN | Instrument Description | Buy Volume (IDR) | Sell Volume (IDR)),
  plus a second sheet ("USD-IDR FX Reference") listing the average USD/IDR
  rate for each month shown, for reference only (not used in the calculation
  since all trade values are already in IDR).

HOW TO RUN
----------
1. Place this script in a folder that also contains an "Excel Files"
   subfolder with the 8 (or however many) part files, and an "Output"
   subfolder (created automatically if missing).
2. Run:  python generate_summary.py
3. The result appears in Output/Monthly_Trade_Volume_Summary_<timestamp>.xlsx

CONFIG
------
Adjust INPUT_FOLDER / OUTPUT_FOLDER below if your folders live elsewhere.
Adjust FX_RATES_BY_MONTH if your data spans months not already listed there.
"""

import glob
import os
import sys
import datetime as dt
from collections import defaultdict

import pandas as pd

def log(msg):
    print(msg, flush=True)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(SCRIPT_DIR, "Excel Files")
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, "Output")

# Fixed column positions (0-indexed) matching the audit file layout.
# Column letters as given in the spec: B, C, D, N, P, R, AB
COL_INSTRUMENT_CODE = 1    # B
COL_INSTRUMENT_DESC = 2    # C
COL_ISIN = 3               # D
COL_SIDE = 13              # N
COL_QTY_FILLED = 15        # P
COL_GROSS_PRICE = 17       # R
COL_TRADE_DATE = 27        # AB

EXCLUDED_MONTH = 9  # September is always dropped from the summary

# Average USD -> IDR rate per month, pulled from x-rates.com monthly averages
# (interbank/market mid rate) on 2026-09-14. Update/add entries here if your
# data covers months not listed - the script will flag any month it can't
# find a rate for instead of guessing.
FX_RATES_BY_MONTH = {
    "2026-01": 16807.31,
    "2026-02": 16828.17,
    "2026-03": 16919.46,
    "2026-04": 17123.89,
    "2026-05": 17563.99,
    "2026-06": 17910.47,
    "2026-07": 18019.54,
    "2026-08": 17831.21,
}
FX_SOURCE_NOTE = "Source: x-rates.com monthly average, USD/IDR (retrieved 2026-09-14)"

# ------------------------------------------------------------------------
# STEP 1: Load and combine all input files
# ------------------------------------------------------------------------

def find_input_files(folder):
    if not os.path.isdir(folder):
        sys.exit(f"ERROR: Input folder not found: {folder}")
    files = [
        f for f in glob.glob(os.path.join(folder, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")
    ]
    if not files:
        sys.exit(f"ERROR: No .xlsx files found in: {folder}")
    return sorted(files)


def load_file(path, file_label):
    """Read one audit file and return a cleaned DataFrame with just the
    columns we need, using fixed column positions (not header names, since
    header text may vary slightly between files). Every row dropped is
    counted and attributed to a specific reason - nothing is dropped quietly."""
    fname = os.path.basename(path)
    try:
        raw = pd.read_excel(path, header=0, engine="calamine")
    except ImportError:
        # python-calamine not installed - fall back to openpyxl in read_only
        # mode, which streams the file instead of loading the full workbook
        # object model (much lower memory on large files). For a large
        # speed-up on big files, install python-calamine:
        #   pip install python-calamine
        try:
            raw = pd.read_excel(path, header=0, engine="openpyxl",
                                 engine_kwargs={"read_only": True})
        except Exception as e:
            log(f"  [{file_label}] FAILED TO READ '{fname}': {e}. Skipping this file entirely.")
            return None, {"read_error": 1}
    except Exception as e:
        log(f"  [{file_label}] FAILED TO READ '{fname}': {e}. Skipping this file entirely.")
        return None, {"read_error": 1}

    needed_max_idx = max(COL_INSTRUMENT_CODE, COL_INSTRUMENT_DESC, COL_ISIN,
                          COL_SIDE, COL_QTY_FILLED, COL_GROSS_PRICE, COL_TRADE_DATE)
    if raw.shape[1] <= needed_max_idx:
        log(f"  [{file_label}] WARNING: '{fname}' has only {raw.shape[1]} "
            f"columns (need at least {needed_max_idx + 1}). Skipping file.")
        return None, {"read_error": 1}

    total_rows = len(raw)
    log(f"  [{file_label}] '{fname}': {total_rows} data row(s) found. Validating...")

    isin = raw.iloc[:, COL_ISIN].astype(str).str.strip()
    desc = raw.iloc[:, COL_INSTRUMENT_DESC].astype(str).str.strip()
    side_raw = raw.iloc[:, COL_SIDE].astype(str).str.strip().str.upper()
    qty = pd.to_numeric(raw.iloc[:, COL_QTY_FILLED], errors="coerce")
    price = pd.to_numeric(raw.iloc[:, COL_GROSS_PRICE], errors="coerce")
    trade_date = pd.to_datetime(raw.iloc[:, COL_TRADE_DATE], errors="coerce")

    # Diagnose *why* each row would be dropped, before actually dropping anything,
    # so every exclusion is visible instead of silently disappearing.
    reasons = {
        "missing_isin": (raw.iloc[:, COL_ISIN].isna() | (isin == "") | (isin.str.lower() == "nan")).sum(),
        "missing_or_invalid_side": (~side_raw.isin(["B", "S"])).sum(),
        "invalid_qty": qty.isna().sum(),
        "invalid_price": price.isna().sum(),
        "invalid_trade_date": trade_date.isna().sum(),
    }

    valid_mask = (
        (~(raw.iloc[:, COL_ISIN].isna() | (isin == "") | (isin.str.lower() == "nan")))
        & side_raw.isin(["B", "S"])
        & qty.notna()
        & price.notna()
        & trade_date.notna()
    )

    df = pd.DataFrame({
        "InstrumentDescription": desc,
        "ISIN": isin.astype("category"),
        "Side": side_raw.astype("category"),
        "Qty": qty,
        "Price": price,
        "TradeDate": trade_date,
    })[valid_mask].copy()

    kept = len(df)
    dropped = total_rows - kept
    if dropped:
        reason_str = ", ".join(f"{k}={v}" for k, v in reasons.items() if v)
        log(f"  [{file_label}] '{fname}': kept {kept}, dropped {dropped} "
            f"row(s) -> {reason_str}")
    else:
        log(f"  [{file_label}] '{fname}': kept all {kept} row(s), no issues found.")

    df["Volume"] = df["Qty"] * df["Price"]
    df["Month"] = df["TradeDate"].dt.strftime("%Y-%m")
    df["SourceFile"] = fname
    return df, reasons


def load_all_files(folder):
    files = find_input_files(folder)
    log(f"Found {len(files)} input file(s) in '{folder}':")
    for f in files:
        log(f"  - {os.path.basename(f)}")
    log("")
    log("Reading and validating files...")

    frames = []
    total_reasons = defaultdict(int)
    n = len(files)
    for i, f in enumerate(files, start=1):
        fname = os.path.basename(f)
        log(f"  [{i}/{n}] Reading '{fname}' ...")
        t_file = dt.datetime.now()
        df, reasons = load_file(f, file_label=f"{i}/{n}")
        elapsed = (dt.datetime.now() - t_file).total_seconds()
        log(f"  [{i}/{n}] '{fname}' processed in {elapsed:.1f}s")
        for k, v in reasons.items():
            total_reasons[k] += v
        if df is not None and len(df):
            frames.append(df)

    log("")
    if not frames:
        sys.exit("ERROR: No usable rows found across all input files.")
    combined = pd.concat(frames, ignore_index=True)

    total_dropped = sum(v for k, v in total_reasons.items() if k != "read_error")
    if total_dropped:
        log(f"TOTAL rows dropped across all files: {total_dropped}")
        for k, v in total_reasons.items():
            if v and k != "read_error":
                log(f"  - {k}: {v}")
    if total_reasons.get("read_error"):
        log(f"TOTAL files that could not be read at all: {total_reasons['read_error']}")
    log(f"TOTAL valid rows loaded: {len(combined)}")
    log("")

    # Check for ISINs mapped to more than one distinct Instrument Description
    # (whitespace already stripped) - this would previously have silently
    # split the same stock's volume into separate summary rows.
    desc_variants = combined.groupby("ISIN")["InstrumentDescription"].nunique()
    inconsistent = desc_variants[desc_variants > 1]
    if len(inconsistent):
        log("WARNING: The following ISIN(s) appear with more than one distinct "
            "Instrument Description across your files (all will still be combined "
            "under a single ISIN row, using the most frequent description):")
        for isin_code in inconsistent.index:
            variants = combined.loc[combined["ISIN"] == isin_code, "InstrumentDescription"].unique()
            log(f"  - {isin_code}: {list(variants)}")
        log("")

    return combined


# ------------------------------------------------------------------------
# STEP 2: Aggregate by month / ISIN / side
# ------------------------------------------------------------------------

def build_monthly_tables(df):
    """Returns dict: month ('YYYY-MM') -> DataFrame[ISIN, Description, Buy, Sell]
    sorted by month, with September excluded.

    Grouping is strictly by ISIN (per spec) - Instrument Description is only
    a display label, chosen as the most frequently occurring description for
    that ISIN. This avoids silently splitting one stock's volume across
    multiple rows if its description text varies slightly between files."""
    df = df[df["TradeDate"].dt.month != EXCLUDED_MONTH].copy()

    # ISIN/Side were stored as category dtype to save memory during loading.
    # Cast back to plain strings before grouping - categorical groupby keys
    # can silently produce every unused category combination (observed=False
    # is the pandas default), which would otherwise inflate row counts here.
    df["ISIN"] = df["ISIN"].astype(str)
    df["Side"] = df["Side"].astype(str)

    # One canonical description per ISIN (most common variant seen).
    isin_to_desc = (
        df.groupby("ISIN")["InstrumentDescription"]
        .agg(lambda s: s.value_counts().idxmax())
    )

    grouped = (
        df.groupby(["Month", "ISIN", "Side"])["Volume"]
        .sum()
        .reset_index()
    )

    tables = {}
    for month, month_df in grouped.groupby("Month"):
        pivot = month_df.pivot_table(
            index=["ISIN"],
            columns="Side",
            values="Volume",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        if "B" not in pivot.columns:
            pivot["B"] = 0
        if "S" not in pivot.columns:
            pivot["S"] = 0

        pivot["Instrument Description"] = pivot["ISIN"].map(isin_to_desc)
        pivot = pivot.rename(columns={
            "B": "Buy Volume (IDR)",
            "S": "Sell Volume (IDR)",
        })
        pivot = pivot[["ISIN", "Instrument Description", "Buy Volume (IDR)", "Sell Volume (IDR)"]]
        pivot = pivot.sort_values("ISIN").reset_index(drop=True)
        tables[month] = pivot

    return dict(sorted(tables.items()))


# ------------------------------------------------------------------------
# STEP 3: Write output workbook
# ------------------------------------------------------------------------

MONTH_NAMES = {
    "01": "January", "02": "February", "03": "March", "04": "April",
    "05": "May", "06": "June", "07": "July", "08": "August",
    "09": "September", "10": "October", "11": "November", "12": "December",
}


def month_label(month_key):
    year, mm = month_key.split("-")
    return f"{MONTH_NAMES.get(mm, mm)} {year}"


def write_output(tables, out_path):
    log("Writing output workbook...")
    wb = Workbook()

    # --- Sheet 1: Monthly Volume Summary ---
    ws = wb.active
    ws.title = "Monthly Volume Summary"

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    month_font = Font(name="Arial", bold=True, size=13)
    normal_font = Font(name="Arial")
    col_widths = [18, 38, 22, 22]

    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1
    for month_key, table in tables.items():
        log(f"  Writing table for {month_label(month_key)} ({len(table)} ISIN row(s))...")
        ws.cell(row=row, column=1, value=month_label(month_key)).font = month_font
        row += 1

        headers = ["ISIN", "Instrument Description", "Buy Volume (IDR)", "Sell Volume (IDR)"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        row += 1

        for _, r in table.iterrows():
            ws.cell(row=row, column=1, value=r["ISIN"]).font = normal_font
            ws.cell(row=row, column=2, value=r["Instrument Description"]).font = normal_font
            buy_cell = ws.cell(row=row, column=3, value=float(r["Buy Volume (IDR)"]))
            sell_cell = ws.cell(row=row, column=4, value=float(r["Sell Volume (IDR)"]))
            buy_cell.font = normal_font
            sell_cell.font = normal_font
            buy_cell.number_format = "#,##0"
            sell_cell.number_format = "#,##0"
            row += 1

        # Totals row for the month
        total_row = row
        ws.cell(row=total_row, column=2, value="Total").font = Font(name="Arial", bold=True)
        first_data_row = total_row - len(table)
        last_data_row = total_row - 1
        if len(table):
            ws.cell(row=total_row, column=3,
                    value=f"=SUM(C{first_data_row}:C{last_data_row})").font = Font(name="Arial", bold=True)
            ws.cell(row=total_row, column=4,
                    value=f"=SUM(D{first_data_row}:D{last_data_row})").font = Font(name="Arial", bold=True)
        ws.cell(row=total_row, column=3).number_format = "#,##0"
        ws.cell(row=total_row, column=4).number_format = "#,##0"
        row += 2  # blank row separator

    # --- Sheet 2: USD-IDR FX Reference ---
    log("  Writing USD-IDR FX Reference sheet...")
    fx_ws = wb.create_sheet("USD-IDR FX Reference")
    fx_ws.column_dimensions["A"].width = 14
    fx_ws.column_dimensions["B"].width = 22
    fx_ws.column_dimensions["C"].width = 60

    fx_ws.cell(row=1, column=1, value="Month").font = header_font
    fx_ws.cell(row=1, column=2, value="Avg 1 USD = IDR").font = header_font
    fx_ws.cell(row=1, column=3, value="Notes").font = header_font
    for c in (1, 2, 3):
        fx_ws.cell(row=1, column=c).fill = header_fill

    r = 2
    for month_key in tables.keys():
        fx_ws.cell(row=r, column=1, value=month_label(month_key)).font = normal_font
        rate = FX_RATES_BY_MONTH.get(month_key)
        if rate is None:
            fx_ws.cell(row=r, column=2, value="N/A - add rate manually").font = normal_font
        else:
            cell = fx_ws.cell(row=r, column=2, value=rate)
            cell.font = normal_font
            cell.number_format = "#,##0.00"
        fx_ws.cell(row=r, column=3, value=FX_SOURCE_NOTE).font = Font(name="Arial", italic=True, size=9)
        r += 1

    log(f"  Saving file to disk...")
    wb.save(out_path)


# ------------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------------

def main():
    t0 = dt.datetime.now()
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    log("=" * 60)
    log("STEP 1/3: Loading and validating input files")
    log("=" * 60)
    log(f"Input folder:  {INPUT_FOLDER}")
    log(f"Output folder: {OUTPUT_FOLDER}")
    log("")
    combined = load_all_files(INPUT_FOLDER)

    log("=" * 60)
    log("STEP 2/3: Aggregating by month and ISIN")
    log("=" * 60)
    tables = build_monthly_tables(combined)
    if not tables:
        sys.exit("ERROR: No data left after excluding September / invalid rows.")

    for month_key, table in tables.items():
        log(f"  {month_label(month_key)}: {len(table)} distinct ISIN(s)")
    log(f"Months included: {', '.join(tables.keys())} (September excluded)")

    missing_fx = [m for m in tables if m not in FX_RATES_BY_MONTH]
    if missing_fx:
        log(f"  NOTE: No FX rate on file for: {', '.join(missing_fx)}. "
            f"These will show 'N/A' on the FX Reference sheet - add manually if needed.")
    log("")

    log("=" * 60)
    log("STEP 3/3: Writing output workbook")
    log("=" * 60)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(OUTPUT_FOLDER, f"Monthly_Trade_Volume_Summary_{timestamp}.xlsx")
    write_output(tables, out_path)

    elapsed = (dt.datetime.now() - t0).total_seconds()
    log("")
    log(f"Done in {elapsed:.1f}s. Output written to:")
    log(f"  {out_path}")


if __name__ == "__main__":
    main()
