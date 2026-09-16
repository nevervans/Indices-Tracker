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


def fetch_bse(index_name: str, from_date, to_date, work_dir: str = ".bse_downloads",
              chunk_days: int = 90, max_retries: int = 3):
    """Full daily history for a BSE index via BSE's own data API (the 'bse'
    package - api.bseindia.com, not the bot-protected main website).

    BSE's archive can be slow to generate large CSVs, which trips the
    library's internal 10s timeout on big date ranges - so this uses
    smaller chunks (90 days by default) and retries each chunk a few
    times before giving up on it.
    """
    import csv as csv_module
    import os as os_module
    import time as time_module
    from bse import BSE

    os_module.makedirs(work_dir, exist_ok=True)
    from_d = from_date if not isinstance(from_date, datetime) else from_date.date()
    to_d = to_date if not isinstance(to_date, datetime) else to_date.date()

    date_keys = ["Date", "Index Date", "date"]
    value_keys = ["Close", "Close Value", "Closing Value", "close"]

    series = []
    with BSE(download_folder=work_dir) as bse:
        for chunk_start, chunk_end in BSE.split_date_range(from_d, to_d, max_chunk_size=chunk_days):
            fpath = None
            last_error = None
            for attempt in range(1, max_retries + 1):
                try:
                    fpath = bse.fetchHistoricalIndexData(index_name, chunk_start, chunk_end)
                    break
                except Exception as exc:
                    last_error = exc
                    if attempt < max_retries:
                        time_module.sleep(2 * attempt)
            if fpath is None:
                if last_error is not None:
                    print(f"    [warn] chunk {chunk_start}->{chunk_end} failed after "
                          f"{max_retries} attempts ({last_error}), skipping this chunk")
                continue
            with open(fpath, newline="", encoding="utf-8-sig") as f:
                for row in csv_module.DictReader(f):
                    raw_date = _first_present(row, date_keys)
                    raw_val = _first_present(row, value_keys)
                    if raw_date is None or raw_val is None:
                        raise KeyError(
                            f"Unrecognised CSV columns from BSE for '{index_name}'. "
                            f"Got columns: {list(row.keys())}. Update date_keys/value_keys in fetch_bse()."
                        )
                    parsed_date = None
                    for fmt in ("%d-%B-%Y", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
                        try:
                            parsed_date = datetime.strptime(str(raw_date).strip(), fmt)
                            break
                        except ValueError:
                            continue
                    if parsed_date is None:
                        raise ValueError(f"Could not parse BSE date '{raw_date}' for '{index_name}'")
                    series.append((parsed_date, float(str(raw_val).replace(",", ""))))
            fpath.unlink(missing_ok=True)

    series.sort(key=lambda x: x[0])
    return series


def fetch_akshare_index(symbol: str, from_date, to_date, max_retries: int = 4,
                         market_prefix: str = "1"):
    """Full daily history for a Chinese A-share index (e.g. CSI 300 = '000300')
    via Eastmoney's public data feed (push2his.eastmoney.com) - hit directly
    with `requests` rather than through akshare's index_zh_a_hist().

    Why direct: akshare's function goes through Python's `requests` library
    with no custom headers, which sends the default 'python-requests/x.x'
    User-Agent - one of the most commonly blocklisted signatures by
    anti-scraping rules. A plain curl to the identical URL (default curl
    User-Agent) succeeded where akshare's request kept failing with
    'RemoteDisconnected', which points at UA-based filtering rather than an
    IP-level block. Setting a normal browser User-Agent here fixes that.

    market_prefix: "1" for Shanghai-listed indices (CSI 300 = 1.000300),
    "0" for Shenzhen-listed. Defaults to Shanghai since that's what CSI 300
    (and most commonly-tracked broad indices) use.
    """
    import time as time_module
    import requests

    url = "https://push2his.eastmoney.com/api/qt/stock/kline/get"
    params = {
        "secid": f"{market_prefix}.{symbol}",
        "fields1": "f1,f2,f3,f4,f5,f6",
        "fields2": "f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61",
        "klt": "101",  # daily
        "fqt": "0",    # no adjustment - correct for an index, not a stock
        "beg": "0",
        "end": "20500101",
        "ut": "fa5fd1943c7b386f172d6893dbfba10b",
    }
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
    }

    payload = None
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, params=params, headers=headers, timeout=15)
            resp.raise_for_status()
            payload = resp.json()
            break
        except Exception as exc:
            last_error = exc
            if attempt < max_retries:
                time_module.sleep(3 * attempt)

    if payload is None:
        raise ConnectionError(
            f"Eastmoney connection failed {max_retries}x for index '{symbol}': {last_error}. "
            f"This is usually transient - try running refresh_history.py again."
        )

    klines = (payload.get("data") or {}).get("klines")
    if not klines:
        return []

    series = []
    for line in klines:
        parts = line.split(",")
        if len(parts) < 3:
            raise KeyError(
                f"Unrecognised kline row from Eastmoney for '{symbol}': {line!r}. "
                f"Update fetch_akshare_index()."
            )
        parsed_date = datetime.strptime(parts[0], "%Y-%m-%d")
        close = float(parts[2])  # order is date,open,close,high,low,...
        series.append((parsed_date, close))

    from_d = from_date if not isinstance(from_date, datetime) else from_date.date()
    to_d = to_date if not isinstance(to_date, datetime) else to_date.date()
    series = [(d, v) for d, v in series if from_d <= d.date() <= to_d]

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
    if source == "bse":
        return fetch_bse(symbol, from_date, to_date)
    if source == "akshare":
        return fetch_akshare_index(symbol, from_date, to_date)
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
            if not row or row[0].startswith("#") or row[0] == "date":
                continue
            series.append((datetime.strptime(row[0], "%Y-%m-%d"), float(row[1])))
    series.sort(key=lambda x: x[0])
    return series


def get_saved_source_key(label: str, history_dir: str = HISTORY_DIR):
    """Returns the 'source:symbol' string the existing file was fetched with,
    or None if the file doesn't exist or predates this tracking (older files).
    Used to detect when indices_config.csv has switched an index to a
    different source/symbol, so a stale partial history isn't silently
    treated as complete.
    """
    path = history_path(label, history_dir)
    if not os.path.exists(path):
        return None
    with open(path, newline="") as f:
        first_line = f.readline().strip()
    if first_line.startswith("#source:"):
        return first_line[len("#source:"):]
    return None


def save_history(label: str, series, history_dir: str = HISTORY_DIR, source_key: str = None):
    os.makedirs(history_dir, exist_ok=True)
    path = history_path(label, history_dir)
    series = sorted(series, key=lambda x: x[0])
    with open(path, "w", newline="") as f:
        if source_key:
            f.write(f"#source:{source_key}\n")
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