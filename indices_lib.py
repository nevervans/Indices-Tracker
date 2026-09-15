"""
indices_lib.py
--------------
Shared code used by refresh_history.py, query_returns.py, and
update_indices.py: fetching from free sources, date/return math, and
reading/writing the per-index history CSVs.

Nothing in here talks to the network on import - it's just functions.
"""
import csv
import os
import re
from datetime import date, datetime, timedelta

HISTORY_DIR = "history"

PERIOD_COLUMNS = [
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


# --------------------------------------------------------------- utilities

def slugify(label: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "_", label.strip().lower())
    return s.strip("_")


def load_config(path: str):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


# --------------------------------------------------------------- fetching

def _first_present(d: dict, candidates):
    for key in candidates:
        if key in d:
            return d[key]
    return None


def fetch_nse(symbol: str, is_tri: bool, from_date, to_date):
    """Full daily history for an NSE index via niftyindices.com (jugaad-data).

    NSE's API has changed field names across versions before, so this parses
    defensively across the known variants rather than assuming one fixed key.
    """
    from jugaad_data.nse import index_raw, index_tri_raw

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


def fetch_yahoo(symbol: str, from_date, to_date):
    """Daily history for a global index / FX / commodity via Yahoo Finance."""
    import yfinance as yf

    df = yf.Ticker(symbol).history(start=from_date, end=to_date + timedelta(days=1),
                                    interval="1d", auto_adjust=False)
    if df.empty:
        return []
    series = [(idx.to_pydatetime().replace(tzinfo=None), float(row["Close"]))
              for idx, row in df.iterrows()]
    series.sort(key=lambda x: x[0])
    return series


def fetch_series(row: dict, from_date, to_date):
    source = row["source"]
    symbol = row["symbol"]
    if source == "nse_price":
        return fetch_nse(symbol, is_tri=False, from_date=from_date, to_date=to_date)
    if source == "nse_tri":
        return fetch_nse(symbol, is_tri=True, from_date=from_date, to_date=to_date)
    if source == "yahoo":
        return fetch_yahoo(symbol, from_date, to_date)
    if source == "manual":
        return []
    raise ValueError(f"Unknown source '{source}' for {row['label']}")


# --------------------------------------------------------------- storage

def history_path(label: str, history_dir: str = HISTORY_DIR) -> str:
    return os.path.join(history_dir, f"{slugify(label)}.csv")


def load_history(label: str, history_dir: str = HISTORY_DIR):
    """Returns oldest-first list of (datetime, float), or [] if no file yet."""
    path = history_path(label, history_dir)
    if not os.path.exists(path):
        return []
    series = []
    with open(path, newline="") as f:
        for row in csv.reader(f):
            if not row or row[0] == "date":
                continue
            series.append((datetime.strptime(row[0], "%Y-%m-%d"), float(row[1])))
    series.sort(key=lambda x: x[0])
    return series


def save_history(label: str, series, history_dir: str = HISTORY_DIR):
    os.makedirs(history_dir, exist_ok=True)
    path = history_path(label, history_dir)
    series = sorted(series, key=lambda x: x[0])
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "value"])
        for d, v in series:
            w.writerow([d.strftime("%Y-%m-%d"), v])


def merge_series(old_series, new_series):
    """Combine two (date, value) lists, new values win on date clashes."""
    by_date = {d: v for d, v in old_series}
    by_date.update({d: v for d, v in new_series})
    return sorted(by_date.items(), key=lambda x: x[0])


# --------------------------------------------------------------- return math

def value_on_or_before(series, target: datetime):
    """series is oldest-first; find the last (date, value) at or before target."""
    result = None
    for d, v in series:
        if d <= target:
            result = (d, v)
        else:
            break
    return result


def value_on_or_after(series, target: datetime):
    """series is oldest-first; find the first (date, value) at or after target."""
    for d, v in series:
        if d >= target:
            return (d, v)
    return None


def month_end_before(d: datetime) -> datetime:
    first_of_month = d.replace(day=1)
    return first_of_month - timedelta(days=1)


def fyear_start(d: datetime) -> datetime:
    """Indian fiscal year: Apr 1 - Mar 31. Returns 31-Mar just before FY start."""
    if d.month >= 4:
        return datetime(d.year, 3, 31)
    return datetime(d.year - 1, 3, 31)


def compute_return_between(series, start_date: datetime, end_date: datetime, cagr: bool = None):
    """Return the change in value between two dates (nearest available data
    on/before each). If cagr is None, auto-picks CAGR for gaps > ~370 days.
    Returns None if there isn't enough history to answer.
    """
    start = value_on_or_before(series, start_date)
    end = value_on_or_before(series, end_date)
    if start is None or end is None or start[1] in (None, 0):
        return None
    start_d, start_v = start
    end_d, end_v = end
    if end_d <= start_d:
        return None
    elapsed_years = (end_d - start_d).days / 365.25
    use_cagr = cagr if cagr is not None else elapsed_years > 1.01
    if use_cagr:
        return (end_v / start_v) ** (1 / elapsed_years) - 1
    return (end_v / start_v) - 1


def compute_standard_returns(series, asof: datetime):
    """Same 12 period returns as the original workbook's 'Indices Return' sheet."""
    if not series:
        return {h: None for h, _ in PERIOD_COLUMNS}

    latest = value_on_or_before(series, asof)
    if latest is None:
        return {h: None for h, _ in PERIOD_COLUMNS}
    latest_date, _ = latest

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
        results[header] = compute_return_between(
            series, anchors[header], latest_date, cagr=(kind == "cagr")
        )
    return results
