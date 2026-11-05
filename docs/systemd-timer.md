# Daily automation with a systemd user timer

The repository includes an installer for a user-level systemd service and timer. It runs the local Job Application Tracker once per day while the Linux user session is active.

## Prerequisites

- Ubuntu/systemd user session.
- `config/local.sh` exists and contains the local `BASE_DIR`, `RCLONE_REMOTE`, and `RCLONE_FILE_PATH` values.
- `rclone` is configured for the same Linux user that owns the systemd user service.
- Docker is running; the user can run `docker` without an interactive `sudo` prompt.
- The `ollama-aislado` container exists and publishes Ollama's API on `localhost:11434`.
- Python dependencies used by the project are installed in the Python environment used by the launcher.

## Configure schedule

Add or update this setting in `config/local.sh`:

```bash
JOB_TRACKER_SCHEDULE="*-*-* 09:00:00"
```

This is the default: every day at 09:00 in the machine's local timezone. For weekdays only, use:

```bash
JOB_TRACKER_SCHEDULE="Mon..Fri 09:00"
```

Any valid systemd calendar expression is accepted. The installer validates it with `systemd-analyze calendar`.

## Install or update

Run from the repository root:

```bash
chmod +x script/actualizar_job_tracker.sh
chmod +x script/install_job_tracker_timer.sh
./script/install_job_tracker_timer.sh
```

The installer detects the repository's actual location from its own path and generates `~/.config/systemd/user/job-tracker.service` with the correct `WorkingDirectory` and `ExecStart`. It also reads `JOB_TRACKER_SCHEDULE` from `config/local.sh`, reloads systemd, and enables the timer. You do not need to edit systemd unit paths manually, even if the repository is cloned into a different directory.

`Persistent=false` means a missed 09:00 execution is not replayed when you log in later. The user-level timer is intended to run while your user session is active. Lingering is not required and should remain disabled for this workflow.

## Verify and operate

```bash
systemctl --user status job-tracker.timer
systemctl --user list-timers job-tracker.timer --no-pager
systemctl --user cat job-tracker.service
systemctl --user cat job-tracker.timer
```

Run a one-off test immediately:

```bash
systemctl --user start job-tracker.service
```

View logs:

```bash
journalctl --user -u job-tracker.service -n 100 --no-pager
journalctl --user -u job-tracker.service -f
```

Stop or disable the schedule:

```bash
systemctl --user disable --now job-tracker.timer
```

Re-enable it after changing `JOB_TRACKER_SCHEDULE`:

```bash
./script/install_job_tracker_timer.sh
```

## Runtime behavior

- The service runs `script/actualizar_job_tracker.sh --no-open`; it generates `reportes/reporte.html` without trying to launch a browser.
- `TimeoutStartSec=infinity` accommodates long Ollama classification runs.
- The launcher starts `ollama-aislado` only if it was stopped. It stops the container on exit only if that execution started it.
- The launcher uses `flock` on `data/.job_tracker.lock` to prevent concurrent runs.
- The service uses the same `config/local.sh`, rclone remote, local data, and Ollama model as a manual run.
- Search/filter/sort in the generated HTML remain client-side. Persistent edits to actions or company/vacancy overrides still require interactive `--serve` mode.

## Troubleshooting

- **`203/EXEC` / executable not found:** run the installer again from the repository clone. It regenerates the service paths from the actual repository location. Confirm `test -x script/actualizar_job_tracker.sh`.
- **Docker permission denied:** confirm the user can run `docker ps` without `sudo`; if recently added to the `docker` group, log out and back in.
- **rclone remote not found:** run `rclone listremotes` as the same user and verify `~/.config/rclone/rclone.conf` is available.
- **Ollama health check times out:** run `docker port ollama-aislado` and verify host port `11434` is published; expected health URL is `http://localhost:11434/api/tags`.
- **Timer not firing:** confirm the user session is active at the scheduled time, check `timedatectl`, then inspect `systemctl --user list-timers` and the journal.
