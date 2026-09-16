# CA Equity Indices Tracker (GitHub + Codespaces, no Bloomberg)

Pulls every index from free sources (NSE's own public index site + Yahoo
Finance), saves the **full daily history** for each one, and lets you ask
for returns as of any date, or between any two dates you choose - not just
a fixed set of periods.

## The four scripts

| Script | What it does | Needs internet? |
|---|---|---|
| `refresh_history.py` | Fetches full daily history for every index and saves it to `history/<index>.csv`. First run backfills years; every run after only pulls new dates since the last save (fast). | Yes |
| `query_returns.py` | Computes returns from the saved history - either the standard 1W…10Y table as of any date, or a fully custom start/end window. Includes the actual NAV/index value alongside every return. No re-fetching. | No |
| `export_history.py` | Exports the **raw NAV/index levels themselves** (not returns) to Excel - one column per index, one row per date. Use this when you want the actual historical data, e.g. to chart it or run your own numbers. | No |
| `update_indices.py` | One-shot version of the old behaviour: fetches fresh and builds the standard table in one go, without saving history. Kept for convenience. | Yes |

## Setup
```bash
pip install -r requirements.txt
```

## Typical workflow
```bash
# 1. Build up the history cache (do this once, or just let the weekly Action do it)
python refresh_history.py --config indices_config.csv

# 2. Ask it anything, instantly, no network needed:

# Standard table, every index, as of today
python query_returns.py --all

# Standard table, as of a specific past date
python query_returns.py --all --asof 2025-06-30

# One index only
python query_returns.py --index "Nifty 50" --asof 2025-06-30

# A completely custom window
python query_returns.py --index "Nifty Bank TRI" --start 2024-01-01 --end 2026-01-01

# Custom window, every index, saved to a file
python query_returns.py --all --start 2024-01-01 --end 2026-01-01 --output custom.xlsx

# 3. Want the actual NAV/index values themselves, not returns? Use export_history.py:
python export_history.py --config indices_config.csv --output history_all.xlsx
python export_history.py --config indices_config.csv --output nifty50_only.xlsx --index "Nifty 50"
python export_history.py --config indices_config.csv --output windowed.xlsx --start 2024-01-01 --end 2025-12-31
```
`--index` must match the exact `label` column in `indices_config.csv` (e.g.
`"Nifty 50"`, `"Nifty Bank TRI"`).

## Running in Codespaces
Same as before - open the repo, **Code -> Codespaces -> Create codespace on
main**, dependencies install automatically, then just run the commands
above in the terminal.

## The weekly GitHub Action
`.github/workflows/update-indices.yml` now runs `refresh_history.py` then
`query_returns.py --all --output tracker_auto.xlsx`, and commits **both**
the updated `history/` folder and `tracker_auto.xlsx` back to the repo.
Because `history/` is committed, each week's run only has to fetch the
handful of new trading days since last time - not years of data again.

Runs every Friday 17:00 IST automatically, or manually anytime from the
**Actions** tab -> "Update indices tracker" -> **Run workflow**.

## Sources (all free, no login/API key)
| Type | Source |
|---|---|
| NSE/Nifty-family indices (price + TRI) | niftyindices.com, via `jugaad-data` |
| BSE 500, BSE Utilities TRI, BSE Consumer Discretionary TRI | BSE's own data API (`api.bseindia.com`), via the `bse` package - a separate, unprotected endpoint from BSE's bot-blocked main website |
| S&P 500, Nasdaq, Nikkei, DAX, CAC, FTSE, Hang Seng, USD/INR, Gold, Silver, Sensex | Yahoo Finance, via `yfinance` |
| Shanghai/CSI300 | Eastmoney's public data feed (`push2his.eastmoney.com`), via the `akshare` library - the **real CSI 300 index**, not an ETF proxy. Eastmoney serves data back to each index's actual inception (this endpoint is what powers `akshare`, `efinance`, and is even used by Microsoft's own `qlib` project for this exact index). Replaces the earlier `000300.SS` Yahoo ticker, which only went back to 2021 |

## Every index is now covered
The 3 BSE indices' exact names (`BSE 500`, `BSE Utilities`, `BSE Consumer
Discretionary`) were confirmed by actually calling `bse.fetchIndexNames()`
- turns out BSE's real names are simpler than the `S&P BSE X` guesses
first tried (S&P/Asia Index only renamed a handful of legacy indices, not
these). One open question: BSE's name list has no separate "TRI" entry
for Utilities or Consumer Discretionary, so it's not confirmed whether
`fetchHistoricalIndexData` returns price return or total return for
these two - worth spot-checking a live value against bseindia.com if the
distinction matters for your reporting.

CSI 300 (Shanghai) uses akshare's `index_zh_a_hist(symbol="000300")` -
this one is well-established (same endpoint used by akshare, efinance,
and Microsoft's qlib for this exact index) so it's lower-risk than the
BSE names above, but I likewise couldn't test a live call from here.

## Other things worth sanity-checking before trusting it weekly
1. **Exact NSE index names.** A few rows are marked `VERIFY` - strategy
   indices (e.g. Nifty Alpha Low Volatility 30) sometimes need a short
   code rather than the full name. Confirm with:
   ```python
   from jugaad_data.nse import index_name_list
   print(index_name_list('Equity', 'Historical Index Data'))
   ```
2. **NSE's and BSE's JSON/CSV field names.** Both parsers are built
   defensively (they try several known key-name variants), but I
   couldn't test live calls to either niftyindices.com or
   api.bseindia.com from where I built this. If a fetch throws a
   `KeyError`, it prints the actual response keys/columns it got back -
   paste those back to me and I'll fix the parser in one line.

## If you're updating from an earlier version of this repo
Two fixes since the last update:

1. **Stale Shanghai data.** If you already ran `refresh_history.py` before
   this update, `history/shanghai_china.csv` has old Yahoo-sourced data
   (2021 onward) from before the switch to akshare. New source changes
   are now caught automatically going forward, but this specific file
   predates that tracking, so delete it once and let it re-backfill under
   the real source:
   ```bash
   rm history/shanghai_china.csv
   python refresh_history.py --config indices_config.csv
   ```
2. **BSE timeouts.** The BSE fetch now uses smaller 90-day chunks (down
   from 365) with automatic retries, since BSE's server can be slow to
   generate large CSVs and was tripping the library's internal 10s
   timeout. If a chunk still fails after 3 retries, it's skipped with a
   warning rather than failing the whole index - so a BSE index might
   come back with a small gap rather than nothing. Just re-run
   `refresh_history.py` again to fill any gaps; the incremental logic
   picks up where it left off. Each history file now also records which
   source/symbol it was built from, so if you ever change a source again
   in `indices_config.csv`, stale data gets discarded and re-backfilled
   automatically instead of silently mixing old and new sources.