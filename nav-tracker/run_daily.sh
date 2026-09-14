#!/usr/bin/env bash
#
# The daily routine: fetch the current price and append it to the workbook.
# Safe to run repeatedly — a day already recorded is left untouched.
#
# Every argument is passed through to `nav_tracker.py update`, so the
# scheduler entry can carry options, e.g.:
#
#     run_daily.sh --workbook /data/nav.xlsx --max-move 15
#
# Exit codes: 0 fine, 2 the page could not be fetched or parsed,
# 3 the scraped value failed its plausibility check.

set -uo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON="${NAV_TRACKER_PYTHON:-python3}"
LOG_DIR="${NAV_TRACKER_LOG_DIR:-$HERE/logs}"
LOG="$LOG_DIR/nav_tracker.log"

mkdir -p "$LOG_DIR"

out="$(mktemp)"
trap 'rm -f "$out"' EXIT

"$PYTHON" "$HERE/nav_tracker.py" update "$@" >"$out" 2>&1
status=$?

{
    printf '\n=== %s (local %s) ===\n' \
        "$(date -u '+%Y-%m-%dT%H:%M:%SZ')" "$(date '+%Y-%m-%d %H:%M %Z')"
    cat "$out"
} >>"$LOG"

# Stay quiet when all is well — cron mails whatever reaches stdout.
if [ "$status" -ne 0 ]; then
    echo "nav_tracker: daily update failed (exit $status) — see $LOG" >&2
    cat "$out" >&2
fi

# Keep the log from growing without bound.
if [ "$(wc -l <"$LOG")" -gt 5000 ]; then
    tail -n 2000 "$LOG" >"$LOG.trimmed" && mv "$LOG.trimmed" "$LOG"
fi

exit "$status"
