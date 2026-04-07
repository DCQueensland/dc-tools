#!/bin/bash
# DC Tools Auto-Commit Watcher
# Watches ~/dc-tools/ and automatically commits + pushes any change within ~5 seconds.
# Prevents the "I edited locally but forgot to push" failure mode.
#
# Start manually:   ~/dc-tools/auto-commit-watcher.sh &
# Start at login:   loaded by ~/Library/LaunchAgents/com.dcqueensland.dctools-autopush.plist
# Stop:             pkill -f auto-commit-watcher.sh
# Log:              /tmp/dc-tools-autocommit.log

cd "$HOME/dc-tools" || exit 1
LOG=/tmp/dc-tools-autocommit.log
echo "[$(date)] watcher started" >> "$LOG"

# Requires fswatch (brew install fswatch). Falls back to a 10s polling loop if missing.
if command -v fswatch >/dev/null 2>&1; then
  fswatch -o -l 3 \
    --exclude '\.git/' \
    --exclude '\.DS_Store' \
    --exclude 'auto-commit-watcher\.sh' \
    "$HOME/dc-tools" | while read -r _; do
      sleep 2  # debounce - let burst edits settle
      if [ -n "$(git status --porcelain)" ]; then
        git add -A
        git commit -m "Auto-commit $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG" 2>&1
        # post-commit hook handles the push
      fi
    done
else
  echo "[$(date)] fswatch not found - using 10s polling fallback. Install with: brew install fswatch" >> "$LOG"
  while true; do
    if [ -n "$(git status --porcelain)" ]; then
      git add -A
      git commit -m "Auto-commit $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG" 2>&1
    fi
    sleep 10
  done
fi
