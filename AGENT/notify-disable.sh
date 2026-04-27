#!/bin/bash
# Fires once on 2026-06-22 to remind Paul to disable the NHL Tracker cloud routine.
# Triggered by ~/Library/LaunchAgents/com.paullo.nhl-disable-reminder.plist.
# Self-aborts on any other date (launchd Month/Day fires every year, so this guards against accidental future fires).

TODAY=$(date +%Y-%m-%d)
TARGET="2026-06-22"

if [ "$TODAY" != "$TARGET" ]; then
    echo "[$(date)] Not target date ($TARGET); skipping. Today=$TODAY" >> "$HOME/01_PLO/02_CLAUDE CODE/02_NHL 2026 Playoffs/AGENT/notify-disable.log"
    exit 0
fi

# Fire macOS notification (banner + sound)
osascript -e 'display notification "NHL 2026 playoffs ended yesterday. Disable the daily routine at claude.ai/code/routines/trig_01ScCxDdyBWosRDCza58txQZ" with title "🏒 NHL Tracker — Action Needed" subtitle "Disable the daily cron job" sound name "Ping"'

echo "[$(date)] Notification fired." >> "$HOME/01_PLO/02_CLAUDE CODE/02_NHL 2026 Playoffs/AGENT/notify-disable.log"
