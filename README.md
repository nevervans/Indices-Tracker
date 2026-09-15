# CA Equity Indices Tracker (automated, GitHub + Codespaces)

Pulls all the indices from free sources (NSE's own public index site +
Yahoo Finance - no Bloomberg) and writes a `Tracker` sheet with the same
return periods as the original workbook. Runs either by hand in
Codespaces, or on its own every week via GitHub Actions.

## 1. Get this into a repo
Easiest path - create a new **empty** repo on github.com (no README/gitignore
so there's nothing to conflict with), then from this folder:
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```
(Private repo is fine and probably what you want here.)

## 2. Run it in Codespaces (no local setup at all)
On the repo's GitHub page: **Code -> Codespaces -> Create codespace on main**.
It opens a full VS Code in the browser and installs `jugaad-data`,
`yfinance`, and `openpyxl` automatically (that's what `.devcontainer/`
does). Once it's ready, just run in the terminal:
```bash
python update_indices.py --config indices_config.csv --output tracker_auto.xlsx
```
Download `tracker_auto.xlsx` from the file explorer (right-click -> Download)
or commit/push it back to the repo.

## 3. Let it update itself every week (optional but the whole point)
`.github/workflows/update-indices.yml` is already wired up:
- Runs every **Friday 17:00 IST** automatically
- Can also be triggered manually anytime: repo -> **Actions** tab ->
  "Update indices tracker" -> **Run workflow**
- Commits the refreshed `tracker_auto.xlsx` straight back into the repo,
  and also attaches it as a downloadable artifact on that run

Change the schedule by editing the `cron` line in that file - the current
one is `"30 11 * * 5"` (11:30 UTC = 17:00 IST, Fridays). No code changes
needed elsewhere; Actions and Codespaces both just run `update_indices.py`.

## What's actually inside
| File | What it's for |
|---|---|
| `update_indices.py` | The script - fetches data, computes returns, writes the sheet |
| `indices_config.csv` | Every index + which free source it comes from |
| `.devcontainer/devcontainer.json` | Makes Codespaces auto-install dependencies |
| `.github/workflows/update-indices.yml` | The weekly auto-run |
| `requirements.txt` | `jugaad-data`, `yfinance`, `openpyxl` |

## Sources (all free, no login/API key)
| Type | Source |
|---|---|
| NSE/Nifty-family indices (price + TRI) | niftyindices.com, via `jugaad-data` |
| S&P 500, Nasdaq, Nikkei, DAX, CAC, FTSE, Hang Seng, Shanghai/CSI300, USD/INR, Gold, Silver, Sensex | Yahoo Finance, via `yfinance` |

## The 3 it can't cover for free
`BSE 500`, `BSE Utilities TRI`, `BSE Consumer Discretionary TRI` are
BSE-branded indices - neither NSE's site nor Yahoo carries them. Marked
`manual` in `indices_config.csv`, left blank in the output.

## Two things worth sanity-checking before trusting it weekly
1. **Exact NSE index names** - a few rows in `indices_config.csv` are
   marked `VERIFY`. Confirm the exact name/short-code from inside
   Codespaces with:
   ```python
   from jugaad_data.nse import index_name_list
   print(index_name_list('Equity', 'Historical Index Data'))
   ```
2. **NSE's JSON field names** - the parser tries several known key-name
   variants defensively, but I haven't been able to test a live call
   against niftyindices.com myself. If a fetch throws a `KeyError`, it
   prints the actual response keys - send those back and it's a one-line
   fix.

Once the numbers check out against a couple of weeks of the old
Bloomberg-based tracker, you can point `--output` at the real tracker
file directly instead of a separate `tracker_auto.xlsx`.
