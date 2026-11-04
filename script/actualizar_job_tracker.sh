#!/usr/bin/env bash

set -Eeuo pipefail

# Resolve repository location from this script.
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
# shellcheck disable=SC1090
source "$CONFIG_FILE"

# BASE_DIR contains data/, reportes/, and script/ in the current setup.
BASE_DIR="${BASE_DIR:-$PROJECT_DIR}"
DATA_DIR="$BASE_DIR/data"
REPORT_DIR="$BASE_DIR/reportes"
PYTHON_SCRIPT="$BASE_DIR/script/procesar_job_emails.py"

RCLONE_REMOTE="${RCLONE_REMOTE:-}"
RCLONE_FILE_PATH="${RCLONE_FILE_PATH:-}"
OLLAMA_CONTAINER="${OLLAMA_CONTAINER:-ollama-aislado}"
OLLAMA_HEALTH_URL="${OLLAMA_HEALTH_URL:-http://localhost:11434/api/tags}"
OLLAMA_START_TIMEOUT="${OLLAMA_START_TIMEOUT:-120}"
PYTHON_BIN="${PYTHON_BIN:-}"

# Prefer the project's virtual environment when present; otherwise use python3.
if [[ -z "$PYTHON_BIN" ]]; then
    if [[ -x "$BASE_DIR/.venv/bin/python" ]]; then
        PYTHON_BIN="$BASE_DIR/.venv/bin/python"
    else
        PYTHON_BIN="$(command -v python3 || true)"
    fi
fi
if [[ -z "$PYTHON_BIN" || ! -x "$PYTHON_BIN" ]]; then
    echo "ERROR: Python interpreter not found. Set PYTHON_BIN in config/local.sh." >&2
    exit 1
fi

if [[ -z "$RCLONE_REMOTE" || -z "$RCLONE_FILE_PATH" ]]; then
    echo "ERROR: RCLONE_REMOTE and RCLONE_FILE_PATH must be configured in config/local.sh."
    exit 1
fi

if [[ ! -f "$PYTHON_SCRIPT" ]]; then
    echo "ERROR: Python script not found: $PYTHON_SCRIPT"
    exit 1
fi

REMOTE="${RCLONE_REMOTE}:${RCLONE_FILE_PATH}"
SHEET_FILE="$DATA_DIR/job_emails.xlsx"
export BASE_DIR DATA_DIR REPORT_DIR

# Ensure the lock directory exists and prevent overlapping manual/timer runs.
mkdir -p "$DATA_DIR" "$REPORT_DIR"
if ! command -v flock >/dev/null 2>&1; then
    echo "ERROR: flock is required to prevent overlapping tracker executions." >&2
    exit 1
fi
exec 9>"$DATA_DIR/.job_tracker.lock"
if ! flock -n 9; then
    echo "ℹ️ Otra ejecución del Job Tracker ya está activa. Se omite esta ejecución."
    exit 0
fi

CONTAINER_STARTED_BY_SCRIPT=0

cleanup() {
    local exit_code=$?
    if [[ "$CONTAINER_STARTED_BY_SCRIPT" -eq 1 ]]; then
        echo ""
        echo "🐳 Deteniendo contenedor Docker: $OLLAMA_CONTAINER"
        docker stop "$OLLAMA_CONTAINER" >/dev/null || {
            echo "WARNING: Could not stop $OLLAMA_CONTAINER automatically." >&2
        }
    fi
    exit "$exit_code"
}
trap cleanup EXIT

# Validate Docker and container before starting the workflow.
if ! command -v docker >/dev/null 2>&1; then
    echo "ERROR: Docker CLI was not found."
    exit 1
fi
if ! docker inspect "$OLLAMA_CONTAINER" >/dev/null 2>&1; then
    echo "ERROR: Docker container '$OLLAMA_CONTAINER' does not exist."
    exit 1
fi

CONTAINER_RUNNING="$(docker inspect -f '{{.State.Running}}' "$OLLAMA_CONTAINER")"
if [[ "$CONTAINER_RUNNING" != "true" ]]; then
    echo "🐳 Iniciando contenedor Ollama: $OLLAMA_CONTAINER"
    docker start "$OLLAMA_CONTAINER" >/dev/null
    CONTAINER_STARTED_BY_SCRIPT=1
else
    echo "🐳 El contenedor $OLLAMA_CONTAINER ya estaba ejecutándose; no se detendrá al finalizar."
fi

# Wait for Ollama's HTTP API to become ready.
echo "⏳ Esperando a que Ollama esté disponible..."
ready=0
for ((attempt=1; attempt<=OLLAMA_START_TIMEOUT; attempt++)); do
    if curl -fsS --max-time 2 "$OLLAMA_HEALTH_URL" >/dev/null 2>&1; then
        ready=1
        break
    fi
    sleep 1
done
if [[ "$ready" -ne 1 ]]; then
    echo "ERROR: Ollama did not become available at $OLLAMA_HEALTH_URL within ${OLLAMA_START_TIMEOUT}s."
    exit 1
fi

echo "=========================================="
echo " JOB APPLICATION TRACKER"
echo "=========================================="
echo ""
echo "📥 Descargando Google Sheet..."
rclone copyto "$REMOTE" "$SHEET_FILE" --drive-export-formats xlsx

echo ""
echo "🧠 Procesando correos con Ollama..."
"$PYTHON_BIN" "$PYTHON_SCRIPT" "$@"

echo ""
echo "=========================================="
echo " Reporte generado"
echo "=========================================="
LATEST_REPORT="$REPORT_DIR/reporte.html"
if [[ -f "$LATEST_REPORT" ]]; then
    echo "📊 Reporte: $LATEST_REPORT"
    if [[ " ${*:-} " == *" --no-open "* ]]; then
        echo "Ejecución sin interfaz gráfica: el reporte no se abrió automáticamente."
    elif [[ " ${*:-} " != *" --serve "* ]]; then
        echo "El reporte se abrió mediante el navegador predeterminado (si está disponible)."
    fi
else
    echo "WARNING: Report file was not found: $LATEST_REPORT" >&2
fi
