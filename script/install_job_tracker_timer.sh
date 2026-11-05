#!/usr/bin/env bash

set -Eeuo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd -- "$SCRIPT_DIR/.." && pwd)"
CONFIG_FILE="$PROJECT_DIR/config/local.sh"
UNIT_TARGET_DIR="$HOME/.config/systemd/user"
SERVICE_TARGET="$UNIT_TARGET_DIR/job-tracker.service"
TIMER_TARGET="$UNIT_TARGET_DIR/job-tracker.timer"

if [[ ! -f "$CONFIG_FILE" ]]; then
    echo "ERROR: Missing $CONFIG_FILE" >&2
    echo "Create it from config/local.example.sh and configure rclone first." >&2
    exit 1
fi

if [[ ! -x "$PROJECT_DIR/script/actualizar_job_tracker.sh" ]]; then
    echo "ERROR: Launcher is missing or not executable:" >&2
    echo "  $PROJECT_DIR/script/actualizar_job_tracker.sh" >&2
    echo "Run: chmod +x script/actualizar_job_tracker.sh" >&2
    exit 1
fi

# Load local settings. The schedule is deliberately configurable locally,
# while the installer resolves the repository path automatically.
# shellcheck source=/dev/null
source "$CONFIG_FILE"
JOB_TRACKER_SCHEDULE="${JOB_TRACKER_SCHEDULE:-*-*-* 09:00:00}"

if ! systemd-analyze calendar "$JOB_TRACKER_SCHEDULE" >/dev/null; then
    echo "ERROR: Invalid JOB_TRACKER_SCHEDULE: $JOB_TRACKER_SCHEDULE" >&2
    echo 'Use a systemd calendar expression, e.g. "*-*-* 09:00:00".' >&2
    exit 1
fi

mkdir -p "$UNIT_TARGET_DIR"

cat > "$SERVICE_TARGET" <<EOF_SERVICE
[Unit]
Description=Local AI Job Application Tracker

[Service]
Type=oneshot
TimeoutStartSec=infinity
TimeoutStopSec=30s
WorkingDirectory=$PROJECT_DIR
Environment=HOME=%h
Environment=PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin
ExecStart=$PROJECT_DIR/script/actualizar_job_tracker.sh --no-open
EOF_SERVICE

cat > "$TIMER_TARGET" <<EOF_TIMER
[Unit]
Description=Run the Local AI Job Application Tracker daily

[Timer]
OnCalendar=$JOB_TRACKER_SCHEDULE
Persistent=false
AccuracySec=1min
Unit=job-tracker.service

[Install]
WantedBy=timers.target
EOF_TIMER

systemctl --user daemon-reload
systemctl --user enable --now job-tracker.timer

printf '\nTimer installed and enabled.\n'
printf 'Project:  %s\n' "$PROJECT_DIR"
printf 'Schedule: %s (machine local timezone)\n' "$JOB_TRACKER_SCHEDULE"
printf '\nNext scheduled execution:\n'
systemctl --user list-timers job-tracker.timer --no-pager

printf '\nTo inspect the latest run:\n'
printf '  journalctl --user -u job-tracker.service -n 100 --no-pager\n'
printf '\nTo run once immediately:\n'
printf '  systemctl --user start job-tracker.service\n'
printf '\nTo disable the timer:\n'
printf '  systemctl --user disable --now job-tracker.timer\n'
printf '\nThe timer runs under your user session. Lingering is not required\n'
printf 'and is intentionally not enabled by this installer.\n'
