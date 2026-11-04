# ============================================================
# LOCAL CONFIGURATION TEMPLATE
# Copy this file to config/local.sh and customize it.
# ============================================================

# Project/data root used by the Bash launcher and Python pipeline.
BASE_DIR="$HOME/path/to/job-search-ai-tracker"

# Name of the rclone remote configured on your machine.
RCLONE_REMOTE="your_rclone_remote"

# Path to the Google spreadsheet relative to the rclone remote.
RCLONE_FILE_PATH="Job2026/Job_Emails_2026.xlsx"

# Optional: explicit Python interpreter. If omitted, the launcher uses
# BASE_DIR/.venv/bin/python when available, otherwise python3 from PATH.
# PYTHON_BIN="$HOME/path/to/job-search-ai-tracker/.venv/bin/python"
