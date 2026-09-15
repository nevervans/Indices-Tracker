"""
update_indices.py
------------------
Replaces the manual Bloomberg pull for the index-level data in the
CA Equity Markets Tracker. No Bloomberg needed.

Sources used (all free, no login/API key):
  - nse_price / nse_tri : niftyindices.com, via the `jugaad-data` library
                          (this is NSE's own public index-history website)
  - yahoo               : Yahoo Finance, via `yfinance`
  - manual              : not available free (see README) - left blank

For every index in indices_config.csv, this script:
  1. Downloads the full available daily history
  2. Computes the same trailing return periods as the original workbook's
     'Indices Return' sheet: 1 Week, MTD, 1 Month, 3 Month, 6 Month, CYTD,
     FYTD, 1 Year, 2/3/5/10 Year (CAGR for the multi-year ones)
  3. Writes everything into a fresh 'Tracker' sheet, grouped exactly like
     the original (Megacap Indices, Large Cap Indices, ...)

Usage:
    pip install jugaad-data yfinance openpyxl
    python update_indices.py --config indices_config.csv --output tracker_auto.xlsx
    python update_indices.py --config indices_config.csv --output tracker_auto.xlsx --asof 2026-09-04
"""
import argparse
import csv
import socket
import sys
import time
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import Font

FETCH_TIMEOUT_SEC = 25  # hard ceiling per request - prevents silent hangs on a blocked network

PERIOD_COLUMNS = [
    # (header, kind)  kind: "simple" -> pct change, "cagr" -> annualised
    ("1 Week", "simple"),
    ("MTD", "simple"),
    ("1 Month", "simple"),
    ("3 Month", "simple"),
    ("6 Month", "simple"),
    ("CYTD", "simple"),
    ("FYTD", "simple"),
    ("1 Year", "simple"),
    ("2 Year", "cagr"),
    ("3 Year", "cagr"),
    ("5 Year", "cagr"),
    ("10 Year", "cagr"),
]


# ---------------------------------------------------------------- fetching

def _first_present(d: dict, candidates):
    for key in candidates:
        if key in d:
            return d[key]
    return None


def fetch_nse(symbol: str, is_tri: bool, years_back: int = 12):
    """Full daily history for an NSE index via niftyindices.com (jugaad-data).

    NSE's API has changed field names across versions before, so this parses
    defensively across the known variants rather than assuming one fixed key.
    """
    from jugaad_data.nse import index_raw, index_tri_raw

    to_date = date.today()
    from_date = to_date - timedelta(days=365 * years_back)

    date_keys = ["Date", "EOD_TIMESTAMP", "HistoricalDate", "TIMESTAMP"]

    if is_tri:
        raw = index_tri_raw(symbol, symbol, from_date, to_date)
        value_keys = ["TotalReturnsIndex", "TRI", "Close"]
    else:
        raw = index_raw(symbol, from_date, to_date)
        value_keys = ["EOD_CLOSE_INDEX_VAL", "Close", "CLOSE", "closingValue"]

    if not raw:
        return []

    series = []
    for r in raw:
        raw_date = _first_present(r, date_keys)
        raw_val = _first_present(r, value_keys)
        if raw_date is None or raw_val is None:
            raise KeyError(
                f"Unrecognised response shape from niftyindices.com for '{symbol}'. "
                f"Got keys: {list(r.keys())}. Update date_keys/value_keys in fetch_nse()."
            )
        parsed_date = None
        for fmt in ("%d %b %Y", "%d-%b-%Y", "%d-%m-%Y"):
            try:
                parsed_date = datetime.strptime(str(raw_date), fmt)
                break
            except ValueError:
                continue
        if parsed_date is None:
            raise ValueError(f"Could not parse date '{raw_date}' for '{symbol}'")
        series.append((parsed_date, float(str(raw_val).replace(",", ""))))

    series.sort(key=lambda x: x[0])
    return series


def fetch_yahoo(symbol: str, years_back: int = 12):
    """Full daily history for a global index / FX / commodity via Yahoo Finance."""
    import yfinance as yf

    period = f"{years_back}y"
    df = yf.Ticker(symbol).history(period=period, interval="1d", auto_adjust=False)
    if df.empty:
        return []
    series = [(idx.to_pydatetime().replace(tzinfo=None), float(row["Close"]))
              for idx, row in df.iterrows()]
    series.sort(key=lambda x: x[0])
    return series


def fetch_series(row: dict):
    source = row["source"]
    symbol = row["symbol"]
    if source == "nse_price":
        return fetch_nse(symbol, is_tri=False)
    if source == "nse_tri":
        return fetch_nse(symbol, is_tri=True)
    if source == "yahoo":
        return fetch_yahoo(symbol)
    if source == "manual":
        return []
    raise ValueError(f"Unknown source '{source}' for {row['label']}")


def fetch_with_timeout(row: dict):
    """fetch_series wrapped so a silently-blocked connection (firewall/VPN
    dropping packets rather than refusing them) raises a clear error instead
    of hanging forever with no output. Relies on the global socket timeout
    set in main()."""
    try:
        return fetch_series(row)
    except (socket.timeout, TimeoutError) as exc:
        raise RuntimeError(
            f"Timed out with no response from {row['source']}. This almost "
            f"always means a firewall/VPN/proxy is silently dropping the "
            f"connection rather than the site being down - try again from a "
            f"different network (e.g. mobile hotspot) to confirm."
        ) from exc


# ---------------------------------------------------------------- returns

def value_on_or_before(series, target: datetime):
    """series is oldest-first; find the last value at or before target."""
    result = None
    for d, v in series:
        if d <= target:
            result = (d, v)
        else:
            break
    return result


def month_end_before(d: datetime) -> datetime:
    first_of_month = d.replace(day=1)
    return first_of_month - timedelta(days=1)


def fyear_start(d: datetime) -> datetime:
    """Indian fiscal year: Apr 1 - Mar 31. Returns 31-Mar just before FY start."""
    if d.month >= 4:
        return datetime(d.year, 3, 31)
    return datetime(d.year - 1, 3, 31)


def compute_returns(series, asof: datetime):
    """series is oldest-first list of (date, value). Returns dict header->value."""
    if not series:
        return {h: None for h, _ in PERIOD_COLUMNS}

    latest = value_on_or_before(series, asof)
    if latest is None:
        return {h: None for h, _ in PERIOD_COLUMNS}
    latest_date, latest_val = latest

    anchors = {
        "1 Week": latest_date - timedelta(days=7),
        "MTD": month_end_before(latest_date),
        "1 Month": latest_date - timedelta(days=30),
        "3 Month": latest_date - timedelta(days=91),
        "6 Month": latest_date - timedelta(days=182),
        "CYTD": datetime(latest_date.year - 1, 12, 31),
        "FYTD": fyear_start(latest_date),
        "1 Year": latest_date - timedelta(days=365),
        "2 Year": latest_date - timedelta(days=365 * 2),
        "3 Year": latest_date - timedelta(days=365 * 3),
        "5 Year": latest_date - timedelta(days=365 * 5),
        "10 Year": latest_date - timedelta(days=365 * 10),
    }

    results = {}
    for header, kind in PERIOD_COLUMNS:
        hist = value_on_or_before(series, anchors[header])
        if hist is None or hist[1] in (None, 0):
            results[header] = None
            continue
        hist_date, hist_val = hist
        if kind == "simple":
            results[header] = (latest_val / hist_val) - 1
        else:  # cagr
            elapsed_years = (latest_date - hist_date).days / 365.25
            results[header] = (latest_val / hist_val) ** (1 / elapsed_years) - 1 if elapsed_years > 0 else None
    return results


# ---------------------------------------------------------------- output

def load_config(path: str):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def build_workbook(config, asof: datetime, delay_sec: float):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tracker"

    ws["C2"] = "CA Equity Markets Tracker"
    ws["C2"].font = Font(bold=True, size=14)
    ws["B5"] = "Analysis Date"
    ws["C5"] = asof

    headers = ["Index Name"] + [h for h, _ in PERIOD_COLUMNS]
    for col, text in enumerate(headers, start=2):
        cell = ws.cell(row=9, column=col, value=text)
        cell.font = Font(bold=True)

    row = 10
    current_group = None
    for entry in config:
        label = entry["label"]
        group = entry["group"]
        if group != current_group:
            ws.cell(row=row, column=2, value=group).font = Font(bold=True, italic=True)
            current_group = group
            row += 1

        if entry["source"] == "manual":
            ws.cell(row=row, column=2, value=label)
            ws.cell(row=row, column=3, value="manual - see README")
            print(f"[manual] {label}: not available from a free source, skipping.")
            row += 1
            continue

        print(f"Fetching {label} ({entry['source']}: {entry['symbol']}) ...")
        try:
            series = fetch_with_timeout(entry)
        except Exception as exc:
            print(f"  [error] {label}: {exc}")
            ws.cell(row=row, column=2, value=label)
            ws.cell(row=row, column=3, value=f"error: {exc}")
            row += 1
            time.sleep(delay_sec)
            continue

        returns = compute_returns(series, asof)
        ws.cell(row=row, column=2, value=label)
        for col, (header, _) in enumerate(PERIOD_COLUMNS, start=3):
            val = returns.get(header)
            cell = ws.cell(row=row, column=col)
            if val is None:
                cell.value = "-"
            else:
                cell.value = round(val, 4)
                cell.number_format = "0.00%"
        row += 1
        time.sleep(delay_sec)

    for col_idx in range(2, 3 + len(PERIOD_COLUMNS)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 14
    ws.column_dimensions["B"].width = 32

    return wb


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--asof", help="YYYY-MM-DD, defaults to today")
    parser.add_argument("--delay", type=float, default=0.3,
                         help="Seconds to wait between requests (be polite to free sources)")
    args = parser.parse_args()

    asof = datetime.strptime(args.asof, "%Y-%m-%d") if args.asof else datetime.today()
    config = load_config(args.config)

    socket.setdefaulttimeout(FETCH_TIMEOUT_SEC)  # belt-and-braces: catches hangs
    print(f"(network timeout set to {FETCH_TIMEOUT_SEC}s per request - "
          f"a hang past that means something's silently blocking the connection)\n")

    wb = build_workbook(config, asof, args.delay)
    wb.save(args.output)
    print(f"\nDone -> {args.output}")


if __name__ == "__main__":
    main()
