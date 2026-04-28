#!/bin/bash
# Daily sync: pull latest from GitHub, copy newest xlsx to ~/Documents
# Triggered by ~/Library/LaunchAgents/com.paullo.nhl-sync.plist at 12:15 PM SGT.

set -e

REPO="/Users/paullo/01_PLO/02_CLAUDE CODE/02_NHL 2026 Playoffs"
DOCS="/Users/paullo/Documents"
LOG="$REPO/AGENT/sync.log"

# Ensure git/PATH are available when launchd invokes us
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"

cd "$REPO"

echo "[$(date '+%Y-%m-%d %H:%M:%S')] Starting sync" >> "$LOG"

# Pull latest from GitHub — retry up to 3 times for transient network failures
PULL_OK=0
for attempt in 1 2 3; do
    if git pull --quiet --rebase origin main 2>>"$LOG"; then
        PULL_OK=1
        break
    fi
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] git pull attempt $attempt failed; retrying in 30s" >> "$LOG"
    sleep 30
done

if [ "$PULL_OK" -ne 1 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] git pull failed after 3 attempts; copying current local file as fallback" >> "$LOG"
    # Don't exit — still copy whatever local file is newest, so user has *something*
fi

# Find the newest versioned xlsx (highest v1.X)
LATEST=$(ls -1 *.xlsx 2>/dev/null | sort -V | tail -1)

if [ -z "$LATEST" ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] No xlsx file found in repo" >> "$LOG"
    exit 1
fi

# Copy the versioned file to Documents (preserves history of versions)
cp -p "$LATEST" "$DOCS/$LATEST"

# Also maintain a stable filename pointing at the latest version
cp -p "$LATEST" "$DOCS/2026 NHL Playoffs_LATEST.xlsx"

echo "[$(date '+%Y-%m-%d %H:%M:%S')] Synced: $LATEST -> Documents/" >> "$LOG"
