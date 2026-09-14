#!/usr/bin/env bash
#
# Install (or remove) the daily 10:00 job on this machine.
#
#     ./schedule/install.sh            # install
#     ./schedule/install.sh uninstall  # remove
#
# macOS gets a launchd agent, Linux a crontab entry. Both fire at 10:00 in
# the machine's local timezone — set that to Europe/Zurich for 10:00 Swiss
# time. For a machine that is not always on, prefer the GitHub Actions
# workflow in .github/workflows/nav-daily.yml instead.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(dirname "$HERE")"
RUN="$ROOT/run_daily.sh"
LOG_DIR="${NAV_TRACKER_LOG_DIR:-$ROOT/logs}"
LABEL="com.varia.navtracker"
MARKER="# nav_tracker — daily NAV update at 10:00"

[ -x "$RUN" ] || { echo "not executable: $RUN" >&2; exit 1; }

install_launchd() {
    local dest="$HOME/Library/LaunchAgents/$LABEL.plist"
    mkdir -p "$(dirname "$dest")" "$LOG_DIR"
    sed -e "s|__RUN_DAILY__|$RUN|g" -e "s|__LOG_DIR__|$LOG_DIR|g" \
        "$HERE/$LABEL.plist" >"$dest"
    launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
    launchctl bootstrap "gui/$(id -u)" "$dest"
    echo "Installed launchd agent -> $dest"
    echo "Verify with: launchctl list | grep $LABEL"
}

uninstall_launchd() {
    local dest="$HOME/Library/LaunchAgents/$LABEL.plist"
    launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
    rm -f "$dest"
    echo "Removed launchd agent."
}

current_crontab() { crontab -l 2>/dev/null || true; }

install_cron() {
    command -v crontab >/dev/null || { echo "crontab not available" >&2; exit 1; }
    if current_crontab | grep -Fq "$RUN"; then
        echo "Already scheduled — nothing to do."
        current_crontab | grep -F "$RUN"
        return
    fi
    {
        current_crontab
        printf '%s\n0 10 * * * %s\n' "$MARKER" "$RUN"
    } | crontab -
    echo "Installed crontab entry: 0 10 * * * $RUN"
    echo "Verify with: crontab -l"
}

uninstall_cron() {
    command -v crontab >/dev/null || return 0
    current_crontab | grep -Fv "$RUN" | grep -Fv "$MARKER" | crontab -
    echo "Removed crontab entry."
}

case "${1:-install}" in
    install)
        case "$(uname -s)" in
            Darwin) install_launchd ;;
            *)      install_cron ;;
        esac
        echo
        echo "The job runs: $RUN"
        echo "Logs:         $LOG_DIR/nav_tracker.log"
        echo "Try it now:   $RUN"
        ;;
    uninstall)
        case "$(uname -s)" in
            Darwin) uninstall_launchd ;;
            *)      uninstall_cron ;;
        esac
        ;;
    *)
        echo "usage: $0 [install|uninstall]" >&2
        exit 1
        ;;
esac
