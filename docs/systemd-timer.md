# Daily automation with a systemd user timer

The repository includes a user-level systemd service and timer for running the local Job Application Tracker once per day.

## Prerequisites

- Ubuntu/systemd user session.
- The project is located at `~/MEGAsync/Documentos/Job_search/Job2026` (adjust `systemd/job-tracker.service` if your clone is elsewhere).
- `config/local.sh` exists and contains the local `BASE_DIR`, `RCLONE_REMOTE`, and `RCLONE_FILE_PATH` values.
- `rclone` is configured for the same Linux user that owns the systemd user service.
- Docker is running; the user can run `docker` without an interactive `sudo` prompt.
- The `ollama-aislado` container exists and publishes Ollama's API on `localhost:11434`.
- Python dependencies used by the project are installed in the Python environment used by `python3`.

## Install

From the repository root:

```bash
chmod +x script/actualizar_job_tracker.sh
chmod +x script/install_job_tracker_timer.sh
./script/install_job_tracker_timer.sh
```

The installer copies `systemd/job-tracker.service` and `systemd/job-tracker.timer` into `~/.config/systemd/user/`, reloads the user manager, and enables the daily timer.

The default schedule is midnight in the machine's local timezone. `Persistent=true` causes a missed run to be triggered after the user systemd manager starts again. It does not make the user manager start while the user is logged out.

To keep the user manager running after logout and start it at boot, enable lingering (administrator permission required):

```bash
sudo loginctl enable-linger "$USER"
```

For a desktop workflow, it is also acceptable to leave lingering disabled; the timer will run while the user session is active.

## Verify and operate

```bash
systemctl --user status job-tracker.timer
systemctl --user list-timers job-tracker.timer --no-pager
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

Re-enable it:

```bash
systemctl --user enable --now job-tracker.timer
```

## Runtime behavior

- The service runs `script/actualizar_job_tracker.sh --no-open`; it generates `reportes/reporte.html` without trying to launch a browser. `TimeoutStartSec=infinity` prevents systemd from killing a long Ollama classification run at the default startup timeout.
- The launcher starts `ollama-aislado` only if it was stopped. It stops the container on exit only if that execution started it.
- The launcher uses `flock` on `data/.job_tracker.lock` to prevent concurrent runs. A second run exits successfully without starting another processing job.
- The service uses the same `config/local.sh`, rclone remote, local data, and Ollama model as a manual run.
- Search/filter/sort in the generated HTML remain client-side. Persistent edits to actions or company/vacancy overrides still require the interactive `--serve` mode.

## Troubleshooting

- **Docker permission denied:** confirm the user can run `docker ps` without `sudo`; if recently added to the `docker` group, log out and back in.
- **rclone remote not found:** run `rclone listremotes` as the same user and verify `~/.config/rclone/rclone.conf` is available.
- **Ollama health check times out:** run `docker port ollama-aislado` and verify host port `11434` is published; the expected health URL is `http://localhost:11434/api/tags`.
- **Timer is not running while logged out:** enable lingering as described above, and ensure Docker is enabled at boot.
- **Project moved:** update `WorkingDirectory` and `ExecStart` in `systemd/job-tracker.service`, then rerun the installer.
