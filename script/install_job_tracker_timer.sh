#!/usr/bin/env bash

set -Eeuo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
UNIT_SOURCE_DIR="$PROJECT_DIR/systemd"
UNIT_TARGET_DIR="$HOME/.config/systemd/user"

SERVICE_SOURCE="$UNIT_SOURCE_DIR/job-tracker.service"
TIMER_SOURCE="$UNIT_SOURCE_DIR/job-tracker.timer"
SERVICE_TARGET="$UNIT_TARGET_DIR/job-tracker.service"
TIMER_TARGET="$UNIT_TARGET_DIR/job-tracker.timer"

if [[ ! -f "$PROJECT_DIR/config/local.sh" ]]; then
    echo "ERROR: Missing $PROJECT_DIR/config/local.sh" >&2
    echo "Create it from config/local.example.sh and configure rclone first." >&2
    exit 1
fi

if [[ ! -x "$PROJECT_DIR/script/actualizar_job_tracker.sh" ]]; then
    echo "ERROR: Launcher is missing or not executable:" >&2
    echo "  $PROJECT_DIR/script/actualizar_job_tracker.sh" >&2
    echo "Run: chmod +x script/actualizar_job_tracker.sh" >&2
    exit 1
fi

if [[ ! -f "$SERVICE_SOURCE" || ! -f "$TIMER_SOURCE" ]]; then
    echo "ERROR: systemd unit templates not found in $UNIT_SOURCE_DIR" >&2
    exit 1
fi

mkdir -p "$UNIT_TARGET_DIR"
install -m 0644 "$SERVICE_SOURCE" "$SERVICE_TARGET"
install -m 0644 "$TIMER_SOURCE" "$TIMER_TARGET"

systemctl --user daemon-reload
systemctl --user enable --now job-tracker.timer

echo
echo "Timer installed and enabled."
echo "Next scheduled execution:"
systemctl --user list-timers job-tracker.timer --no-pager

echo
echo "To inspect the latest run:"
echo "  journalctl --user -u job-tracker.service -n 100 --no-pager"
echo
echo "To run once immediately:"
echo "  systemctl --user start job-tracker.service"
echo
echo "To disable the timer:"
echo "  systemctl --user disable --now job-tracker.timer"
echo
echo "This is a user-level timer. It normally runs while your user session is active."
echo "For runs while logged out, enable lingering if your system policy allows it:"
echo "  sudo loginctl enable-linger \"$USER\""
