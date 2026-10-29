
#!/usr/bin/env bash

set -e

# Resolve the repository location from this script.
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

CONFIG_FILE="$PROJECT_DIR/config/local.sh"

if [[ ! -f "$CONFIG_FILE" ]]; then
    echo "ERROR: Local configuration not found."
    echo "Create it with:"
    echo "  cp \"$PROJECT_DIR/config/local.example.sh\" \"$CONFIG_FILE\""
    exit 1
fi

# Load private local configuration.
source "$CONFIG_FILE"

# Use the configured project path, or the repository path by default.
BASE_DIR="${BASE_DIR:-$PROJECT_DIR}"

DATA_DIR="$BASE_DIR/data"
REPORT_DIR="$BASE_DIR/reportes"
SCRIPT_DIR="$BASE_DIR/script"

# Validate rclone settings.
if [[ -z "${RCLONE_REMOTE:-}" || -z "${RCLONE_FILE_PATH:-}" ]]; then
    echo "ERROR: RCLONE_REMOTE and RCLONE_FILE_PATH must be configured."
    exit 1
fi

REMOTE="${RCLONE_REMOTE}:${RCLONE_FILE_PATH}"
SHEET_FILE="$DATA_DIR/job_emails.xlsx"

# Pass the same project paths to Python.
export BASE_DIR
export DATA_DIR
export REPORT_DIR

echo "=========================================="
echo " JOB APPLICATION TRACKER"
echo "=========================================="

mkdir -p "$DATA_DIR" "$REPORT_DIR"

echo ""
echo "📥 Descargando Google Sheet..."

rclone copyto "$REMOTE" "$SHEET_FILE" --drive-export-formats xlsx

echo ""
echo "🧠 Procesando correos con Granite4.2..."

python3 "$SCRIPT_DIR/procesar_job_emails.py" "$@"

echo ""
echo "=========================================="
echo " Reporte generado"
echo "=========================================="

LATEST_REPORT="$(
    ls -t "$REPORT_DIR"/reporte_*.html 2>/dev/null |
    head -n 1
)"

if [[ -n "$LATEST_REPORT" ]]; then
    echo ""
    echo "📊 Abriendo:"
    echo "$LATEST_REPORT"

    xdg-open "$LATEST_REPORT" >/dev/null 2>&1 &
else
    echo ""
    echo "⚠️ No se encontró el reporte."
fi
