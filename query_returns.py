"""
query_returns.py
-----------------
Answers "what was the return as of <date>" or "what was the return from
<date A> to <date B>" for any index (or all of them), using the history
already saved by refresh_history.py. No network calls - instant.

Includes the underlying NAV/index values alongside the returns, not just
the percentages. For the full raw daily history instead, use
export_history.py.

Examples:
    # Standard 1W/MTD/.../10Y table, as of today, every index
    python query_returns.py --all

    # Same, but as of a specific date
    python query_returns.py --all --asof 2026-06-15

    # One index only
    python query_returns.py --index "Nifty 50" --asof 2026-06-15

    # A completely custom window for one index
    python query_returns.py --index "Nifty Bank TRI" --start 2024-01-01 --end 2026-01-01

    # Custom window for every index, saved to a file
    python query_returns.py --all --start 2024-01-01 --end 2026-01-01 --output custom_window.xlsx
"""
import argparse
from datetime import datetime

import openpyxl
from openpyxl.styles import Font

import indices_lib as lib

# column "kind" tags used to pick the right cell format on export
TEXT, NUMBER, PERCENT = "text", "number", "percent"


def print_table(rows, columns):
    widths = [max(len(str(r[i])) for r in ([columns] + rows)) for i in range(len(columns))]
    def fmt_row(r):
        return "  ".join(str(v).ljust(w) for v, w in zip(r, widths))
    print(fmt_row(columns))
    print(fmt_row(["-" * w for w in widths]))
    for r in rows:
        print(fmt_row(r))


def fmt_pct(v):
    return "-" if v is None else f"{v:.2%}"


def fmt_num(v):
    return "-" if v is None else f"{v:.2f}"


def save_xlsx(rows, columns, kinds, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Returns"
    for c, h in enumerate(columns, start=1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            kind = kinds[c_idx - 1]
            if kind == TEXT:
                cell.value = val
                continue
            try:
                if kind == PERCENT:
                    cell.value = float(str(val).strip("%")) / 100
                    cell.number_format = "0.00%"
                else:  # NUMBER
                    cell.value = float(val)
                    cell.number_format = "#,##0.00"
            except (ValueError, TypeError):
                # covers "-", "no data", "error: ...", or anything else non-numeric
                cell.value = str(val)
    for c in range(1, len(columns) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 16
    wb.save(path)
    print(f"\nSaved -> {path}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="indices_config.csv")
    parser.add_argument("--history-dir", default=lib.HISTORY_DIR)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--index", help="Exact label from indices_config.csv, e.g. 'Nifty 50'")
    group.add_argument("--all", action="store_true", help="Every index in the config")
    parser.add_argument("--asof", help="YYYY-MM-DD - standard period table as of this date (default: today)")
    parser.add_argument("--start", help="YYYY-MM-DD - use with --end for a custom window instead of --asof")
    parser.add_argument("--end", help="YYYY-MM-DD - use with --start")
    parser.add_argument("--output", help="Optional: save results to this .xlsx path")
    args = parser.parse_args()

    if (args.start and not args.end) or (args.end and not args.start):
        parser.error("--start and --end must be used together")

    config = lib.load_config(args.config)
    labels = [args.index] if args.index else [e["label"] for e in config if e["source"] != "manual"]

    if args.index and args.index not in [e["label"] for e in config]:
        parser.error(f"'{args.index}' is not a label in {args.config}. "
                     f"Check the exact spelling in that file's 'label' column.")

    custom_window = bool(args.start)

    if custom_window:
        start_d = datetime.strptime(args.start, "%Y-%m-%d")
        end_d = datetime.strptime(args.end, "%Y-%m-%d")
        columns = ["Index", "Start Date", "Start Value", "End Date", "End Value",
                   f"Return ({args.start} -> {args.end})"]
        kinds = [TEXT, TEXT, NUMBER, TEXT, NUMBER, PERCENT]
        rows = []
        for label in labels:
            series = lib.load_history(label, args.history_dir)
            if not series:
                rows.append([label, "-", "-", "-", "-", "no data"])
                continue
            start_pt = lib.value_on_or_before(series, start_d)
            end_pt = lib.value_on_or_before(series, end_d)
            ret = lib.compute_return_between(series, start_d, end_d)
            rows.append([
                label,
                start_pt[0].strftime("%Y-%m-%d") if start_pt else "-",
                fmt_num(start_pt[1]) if start_pt else "-",
                end_pt[0].strftime("%Y-%m-%d") if end_pt else "-",
                fmt_num(end_pt[1]) if end_pt else "-",
                fmt_pct(ret),
            ])
    else:
        asof_d = datetime.strptime(args.asof, "%Y-%m-%d") if args.asof else datetime.today()
        columns = ["Index", "As Of Date", "Value"] + [h for h, _ in lib.PERIOD_COLUMNS]
        kinds = [TEXT, TEXT, NUMBER] + [PERCENT] * len(lib.PERIOD_COLUMNS)
        rows = []
        for label in labels:
            series = lib.load_history(label, args.history_dir)
            if not series:
                rows.append([label, "-", "-"] + ["no data"] * len(lib.PERIOD_COLUMNS))
                continue
            latest = lib.value_on_or_before(series, asof_d)
            returns = lib.compute_standard_returns(series, asof_d)
            rows.append([
                label,
                latest[0].strftime("%Y-%m-%d") if latest else "-",
                fmt_num(latest[1]) if latest else "-",
            ] + [fmt_pct(returns[h]) for h, _ in lib.PERIOD_COLUMNS])

    print_table(rows, columns)
    if args.output:
        save_xlsx(rows, columns, kinds, args.output)


if __name__ == "__main__":
    main()