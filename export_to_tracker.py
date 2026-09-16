"""
export_to_tracker.py
---------------------
Appends fresh NAV/index-level rows straight from the history/ cache into
Sheet1 of the ORIGINAL Bloomberg-format tracker workbook (the one with
'Sheet1' / 'Indices Return' / 'Tracker' / 'Manual' tabs), instead of
Bloomberg's own live pull. Nothing above the appended rows is touched -
headers, the old BDH formulas in row 15, the 'Indices Return' formulas,
and the 'Tracker' sheet's formatting all stay exactly as they are.

Why this is safe to append as plain values: Sheet1's own historical rows
(16 onward) are already plain pasted values, not formulas - Bloomberg's
BDH array only ever populated row 15. So this script is just doing, in
code, exactly what's already been done by hand for years.

What it guarantees, to keep the 'Indices Return' date-picker working:
  1. Every calendar day is added - weekends and holidays included - with
     non-trading days carrying forward the last available close, exactly
     like the existing rows already do. (Indices Return does an EXACT
     match on date, so a single gap would show "N/A" for any Analysis
     Date or lookback date that lands on it.)
  2. Values land in the exact same column position as the existing
     tickers (position, not header text, is what 'Indices Return'
     locks onto). The 4 columns this pipeline doesn't fetch
     (Nifty 50 TRI, Nifty Next 50 TRI, Nifty 100 TRI, Nifty 500 TRI)
     are simply left blank on the new rows - ask before filling those in
     some other way.

One deliberate correction vs. the original Bloomberg file: ~10 sector
"TRI" columns there (PSU Bank, PSE, Auto, Metal, FMCG, Energy, IT, ESG,
and the two BSE ones) were tracking PRICE return, not total return, per
Bloomberg ticker. This pipeline's indices_config.csv fetches genuine TRI
for the NSE ones (nse_tri source) and the best available BSE data (see
README) - so those columns' numbers will jump upward on the first
appended row. That's an intentional fix, not a bug - flag it once to
whoever reads the numbers next.

Usage:
    python export_to_tracker.py --config indices_config.csv \\
        --tracker "CA_Equity_Markets_Tracker.xlsx" \\
        --output "CA_Equity_Markets_Tracker_updated.xlsx"

    # Stop the append earlier than "latest data available" if needed:
    python export_to_tracker.py --config indices_config.csv \\
        --tracker "CA_Equity_Markets_Tracker.xlsx" \\
        --output out.xlsx --through 2026-09-10
"""
import argparse
from datetime import datetime, timedelta

import openpyxl

import indices_lib as lib

# Sheet1 column (C..BF), in order, mapped to this pipeline's config label.
# None = a ticker Bloomberg had that this pipeline doesn't fetch (left
# blank on new rows, per "we only need the indices we've already done").
SHEET1_COLUMN_ORDER = [
    ("C", None),                                   # NIFTYTR Index (Nifty 50 TRI) - not fetched
    ("D", "Nifty 50"),
    ("E", "Sensex"),
    ("F", None),                                   # NIFTYJRT Index (Nifty Next 50 TRI) - not fetched
    ("G", "Nifty Next 50"),
    ("H", None),                                   # NSE100TR Index (Nifty 100 TRI) - not fetched
    ("I", "Nifty 100"),
    ("J", "Nifty Midcap 150"),
    ("K", "Nifty Large Midcap 250"),
    ("L", "Nifty Smallcap 250"),
    ("M", "Nifty50 Equal Weight"),
    ("N", "Nifty100 Equal Weight"),
    ("O", None),                                   # NSE500TR Index (Nifty 500 TRI) - not fetched
    ("P", "Nifty 200 TRI"),
    ("Q", "NSE 500"),
    ("R", "Nifty200 Value 30 TRI"),
    ("S", "Nifty 200 Momentum 30 Index TRI"),
    ("T", "Nifty 200 Quality 30 Index TRI"),
    ("U", "Nifty 100 Low Volatility 30"),
    ("V", "Nifty 200 Momentum 30 Index"),
    ("W", "Nifty 100 Quality 30"),
    ("X", "Nifty Alpha 50"),
    ("Y", "Nifty Alpha Low Volatility 30"),
    ("Z", "Nifty India Defence TRI"),
    ("AA", "Nifty PSU Bank TRI"),
    ("AB", "Nifty PSE TRI"),
    ("AC", "Nifty Realty TRI"),
    ("AD", "Nifty Auto TRI"),
    ("AE", "Nifty India Manufacturing TRI"),
    ("AF", "Nifty Infrastructure TRI"),
    ("AG", "Nifty Healthcare TRI"),
    ("AH", "Nifty Metal TRI"),
    ("AI", "BSE Consumer Discretionary TRI"),
    ("AJ", "Nifty Financial Services TRI"),
    ("AK", "Nifty Commodities TRI"),
    ("AL", "Nifty Bank TRI"),
    ("AM", "BSE Utilities TRI"),
    ("AN", "Nifty Private Bank TRI"),
    ("AO", "NIFTY100 ESG TRI"),
    ("AP", "Nifty FMCG TRI"),
    ("AQ", "Nifty Oil & Gas TRI"),
    ("AR", "Nifty Energy TRI"),
    ("AS", "Nifty IT TRI"),
    ("AT", "S&P 500(US)"),
    ("AU", "Nasdaq(US)"),
    ("AV", "Nikkei(Japan)"),
    ("AW", "Dax Index(Germany)"),
    ("AX", "CAC 40 Index(France)"),
    ("AY", "FTSE(UK)"),
    ("AZ", "Hang Seng(Hong Kong)"),
    ("BA", "USD/INR"),
    ("BB", "Gold(USD)"),
    ("BC", "Silver(USD)"),
    ("BD", "Shanghai(China)"),
    ("BE", "BSE 500"),
    ("BF", "Nifty MidSmall 400"),
]


def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="indices_config.csv")
    parser.add_argument("--history-dir", default=lib.HISTORY_DIR)
    parser.add_argument("--tracker", required=True,
                         help="Path to the existing Bloomberg-format tracker workbook")
    parser.add_argument("--output", required=True)
    parser.add_argument("--through", help="YYYY-MM-DD - last date to append. "
                         "Omit to use the latest date available across the cache.")
    args = parser.parse_args()

    config = lib.load_config(args.config)
    label_to_config = {e["label"]: e for e in config}

    covered_labels = [label for _, label in SHEET1_COLUMN_ORDER if label is not None]
    missing_from_config = [l for l in covered_labels if l not in label_to_config]
    if missing_from_config:
        raise SystemExit(
            f"These labels are in SHEET1_COLUMN_ORDER but not in {args.config}: "
            f"{missing_from_config}. Fix the mapping at the top of this script."
        )

    print(f"Loading cached history for {len(covered_labels)} indices ...")
    series_by_label = {label: lib.load_history(label, args.history_dir) for label in covered_labels}
    empty = [label for label, s in series_by_label.items() if not s]
    if empty:
        print(f"  [warn] no cached history for: {empty} - run refresh_history.py first. "
              f"These columns will be left blank on every new row.")

    wb = openpyxl.load_workbook(args.tracker, data_only=True)
    ws = wb["Sheet1"]

    last_row = ws.max_row
    while last_row > 14 and ws.cell(row=last_row, column=2).value is None:
        last_row -= 1
    last_date = ws.cell(row=last_row, column=2).value
    if not isinstance(last_date, datetime):
        raise SystemExit(f"Couldn't find a date in Sheet1!B{last_row} - check the file.")
    last_date_str = last_date.date()
    print(f"Sheet1's last existing row is {last_row}, dated {last_date_str}.")

    if args.through:
        through = datetime.strptime(args.through, "%Y-%m-%d")
    else:
        through = max((s[-1][0] for s in series_by_label.values() if s), default=None)
        if through is None:
            raise SystemExit("No cached history found for any covered index - nothing to append.")
    print(f"Appending every calendar day from {last_date_str + timedelta(days=1)} "
          f"through {through.date()} ...")

    start = last_date + timedelta(days=1)
    if start > through:
        print("Nothing to append - tracker is already at or ahead of the requested date.")
        wb.save(args.output)
        return

    date_number_format = ws.cell(row=last_row, column=2).number_format

    n_rows = 0
    r = last_row
    for d in daterange(start, through):
        r += 1
        n_rows += 1
        date_cell = ws.cell(row=r, column=2, value=d)
        date_cell.number_format = date_number_format
        for col_letter, label in SHEET1_COLUMN_ORDER:
            if label is None:
                continue
            series = series_by_label.get(label) or []
            found = lib.value_on_or_before(series, d)
            if found is None:
                continue
            _, value = found
            ws.cell(row=r, column=openpyxl.utils.column_index_from_string(col_letter),
                     value=round(value, 4))

    wb.save(args.output)
    print(f"\nAppended {n_rows} rows (rows {last_row + 1}-{r}) -> {args.output}")
    print("\nReminders:")
    print("  - Update the 'Data updated as of' cell (Tracker!C71) and the Manual sheet's "
          "disclaimer date by hand - those are static text, not formulas.")
    print("  - Set Tracker!C5 (Analysis Date) to the new latest date and confirm no "
          "'N/A'/'Error' shows up anywhere before sharing the file.")
    print("  - The TRI columns noted in this script's docstring will show a level shift "
          "upward on the first appended row (correct total-return data replacing what "
          "was price-return data) - worth a one-line heads-up to whoever reads the sheet.")


if __name__ == "__main__":
    main()