#!/usr/bin/env bash
# update_now.sh
# --------------
# Runs the exact same steps as the scheduled GitHub Action, on demand.
# Use this whenever you want fresher NAVs without waiting for Friday.
#
# Usage:
#   bash update_now.sh
set -e

if [ -d .git/rebase-merge ] || [ -d .git/rebase-apply ]; then
  echo "A rebase is already in progress - resolve it first:"
  echo "  git status              # see what's conflicted"
  echo "  git rebase --abort      # or, to bail out entirely and match origin:"
  echo "  git rebase --abort && git fetch origin && git reset --hard origin/main"
  exit 1
fi

echo "== Syncing with the remote first (avoids racing the scheduled Action) =="
git pull --rebase origin main

echo "== Refreshing cached history =="
python refresh_history.py --config indices_config.csv

echo "== Rebuilding tracker_auto.xlsx =="
python query_returns.py --all --output tracker_auto.xlsx

echo "== Rebuilding CA_Equity_Markets_Tracker.xlsx fresh from the pristine template =="
echo "   (formulas always come from tracker_template_PRISTINE_DO_NOT_EDIT.xlsx, never"
echo "    from last week's output - never edit that template file by hand)"
python build_tracker.py \
  --config indices_config.csv \
  --old-tracker tracker_template_PRISTINE_DO_NOT_EDIT.xlsx \
  --output CA_Equity_Markets_Tracker.xlsx

echo "== Committing and pushing =="
git add history/ tracker_auto.xlsx CA_Equity_Markets_Tracker.xlsx
if git diff --cached --quiet; then
  echo "Nothing changed - already up to date."
else
  git commit -m "Manual indices update $(date +'%Y-%m-%d')"
  git pull --rebase origin main   # in case the Action ran while this script was working
  git push
  echo "Pushed. CA_Equity_Markets_Tracker.xlsx is updated in the repo."
fi