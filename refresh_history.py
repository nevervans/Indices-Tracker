"""
refresh_history.py
-------------------
Fetches full daily history for every index in indices_config.csv and saves
it to history/<index>.csv (one file per index, date+value).

First run: pulls years of history (slow, one-time).
Every run after: only fetches dates newer than what's already saved (fast).

Usage:
    python refresh_history.py --config indices_config.csv
    python refresh_history.py --config indices_config.csv --years 15   # first-time backfill depth
"""
import argparse
import time
from datetime import date, timedelta

import indices_lib as lib


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    parser.add_argument("--history-dir", default=lib.HISTORY_DIR)
    parser.add_argument("--years", type=int, default=12,
                         help="How far back to backfill on the very first run for each index")
    parser.add_argument("--delay", type=float, default=0.3)
    args = parser.parse_args()

    config = lib.load_config(args.config)
    today = date.today()

    for entry in config:
        label = entry["label"]
        if entry["source"] == "manual":
            print(f"[skip] {label}: manual/no free source")
            continue

        existing = lib.load_history(label, args.history_dir)
        if existing:
            from_date = existing[-1][0].date() + timedelta(days=1)
            if from_date > today:
                print(f"[up to date] {label} (last saved: {existing[-1][0].date()})")
                continue
        else:
            from_date = today - timedelta(days=365 * args.years)

        print(f"Fetching {label} ({entry['source']}: {entry['symbol']}) "
              f"from {from_date} to {today} ...")
        try:
            new_series = lib.fetch_series(entry, from_date, today)
        except Exception as exc:
            print(f"  [error] {label}: {exc}")
            time.sleep(args.delay)
            continue

        if not new_series:
            print(f"  [warn] no new data returned for {label}")
            time.sleep(args.delay)
            continue

        merged = lib.merge_series(existing, new_series)
        lib.save_history(label, merged, args.history_dir)
        print(f"  saved {len(merged)} total rows "
              f"({merged[0][0].date()} -> {merged[-1][0].date()})")
        time.sleep(args.delay)

    print("\nDone.")


if __name__ == "__main__":
    main()
