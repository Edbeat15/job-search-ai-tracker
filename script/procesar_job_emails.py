#!/usr/bin/env python3

import os
import json
import re
import argparse
import html
import time
import hashlib
import difflib
import unicodedata
import threading
import webbrowser
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import plotly.graph_objects as go



# ============================================================
# CONFIGURATION
# ============================================================

# Use the path passed by the launcher when available.
# When run directly, detect the repository from this file.
PROJECT_DIR = Path(__file__).resolve().parent.parent

BASE_DIR = Path(
    os.environ.get("BASE_DIR", str(PROJECT_DIR))
).expanduser().resolve()

DATA_DIR = BASE_DIR / "data"
REPORT_DIR = BASE_DIR / "reportes"

SHEET_FILE = DATA_DIR / "job_emails.xlsx"
RAW_JSON_FILE = DATA_DIR / "job_emails.json"
CLASSIFIED_JSON_FILE = DATA_DIR / "classified_jobs.json"
PROCESSING_STATE_FILE = DATA_DIR / "processing_state.json"
USER_ACTIONS_FILE = DATA_DIR / "user_actions.json"
MANUAL_OVERRIDES_FILE = DATA_DIR / "manual_overrides.json"
REPORT_FILE = REPORT_DIR / "reporte.html"

OLLAMA_URL = os.getenv(
    "OLLAMA_URL",
    "http://localhost:11434/api/chat",
)

OLLAMA_MODEL = os.getenv(
    "OLLAMA_MODEL",
    "granite-job",
)

OLLAMA_TIMEOUT = 600


ALLOWED_STATUSES = {
    "Applied",
    "Action_Required",
    "Rejected",
    "Interview",
    "Job_Offer",
}


# ============================================================
# AI / BATCH CONFIGURATION
# ============================================================

# Maximum email body sent to Ollama.
# The original body remains untouched in job_emails.json.
MAX_BODY_CHARS = 3000

# Maximum estimated prompt size per batch.
#
# Lower values = smaller individual inference workloads.
MAX_BATCH_CHARS = 4500

# Pause between Ollama requests.
BATCH_COOLDOWN_SECONDS = 5


# ============================================================
# UTILITIES
# ============================================================

def log(message):
    print(message, flush=True)


def clean_text(value):
    """Return a safe text representation for scalars and collection values.

    pandas.isna(list_like) returns an array, so using it directly in an if
    statement raises: "truth value of an array is ambiguous". Evidence and
    message IDs can legitimately be lists in the consolidated JSON.
    """
    if value is None:
        return ""

    if isinstance(value, (list, tuple, set)):
        return " ".join(
            text for item in value
            if (text := clean_text(item))
        ).strip()

    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)

    try:
        missing = pd.isna(value)
        # Only evaluate scalar missing-value results. Array-like results must
        # never be coerced to bool.
        if not hasattr(missing, "__len__") and bool(missing):
            return ""
    except (TypeError, ValueError):
        pass

    return str(value).strip()


def normalize_body(text):

    text = clean_text(text)

    text = text.replace(
        "\r\n",
        "\n",
    )

    text = text.replace(
        "\r",
        "\n",
    )

    text = re.sub(
        r"\n{3,}",
        "\n\n",
        text,
    )

    return text.strip()


def truncate_text(
    text,
    max_chars=MAX_BODY_CHARS,
):

    text = clean_text(text)

    if len(text) <= max_chars:
        return text

    return (
        text[:max_chars]
        + "\n\n[EMAIL BODY TRUNCATED FOR AI PROCESSING]"
    )


def safe_filename_date():

    return datetime.now().strftime(
        "%Y-%m-%d"
    )


# ============================================================
# READ EXCEL
# ============================================================

def load_excel():

    log("")
    log("📊 Leyendo Google Sheet exportado...")
    log(f"   Archivo: {SHEET_FILE}")

    if not os.path.exists(SHEET_FILE):

        raise FileNotFoundError(
            f"No existe el archivo:\n{SHEET_FILE}"
        )

    df = pd.read_excel(
        SHEET_FILE,
        sheet_name="Emails",
        engine="openpyxl",
    )

    if df.empty:

        raise ValueError(
            "El Excel está vacío."
        )

    log(
        f"   Filas encontradas: {len(df)}"
    )

    required_columns = [
        "Message_ID",
        "Thread_ID",
        "Date",
        "From",
        "To",
        "Subject",
        "Body",
        "Snippet",
        "Gmail_Link",
        "Retrieved_At",
    ]

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:

        raise ValueError(
            "Faltan columnas requeridas en el Excel:\n"
            + "\n".join(missing)
        )

    for col in [
        "Message_ID",
        "Thread_ID",
        "From",
        "To",
        "Subject",
        "Body",
        "Snippet",
        "Gmail_Link",
        "Retrieved_At",
    ]:

        df[col] = df[col].apply(
            clean_text
        )

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce",
    )

    df = df[
        df["Message_ID"].astype(bool)
    ].copy()

    before = len(df)

    df = df.drop_duplicates(
        subset=["Message_ID"],
        keep="first",
    )

    after = len(df)

    if before != after:

        log(
            f"   Duplicados eliminados: "
            f"{before - after}"
        )

    df = df.sort_values(
        by=[
            "Thread_ID",
            "Date",
        ],
        ascending=[
            True,
            True,
        ],
    )

    log(
        f"   Mensajes válidos: {len(df)}"
    )

    return df


# ============================================================
# RAW JSON
# ============================================================

def dataframe_to_messages(df):

    messages = []

    for _, row in df.iterrows():

        date_value = row["Date"]

        if pd.isna(date_value):

            date_iso = ""

        else:

            date_iso = date_value.isoformat()

        message = {

            "message_id":
                row["Message_ID"],

            "thread_id":
                row["Thread_ID"],

            "date":
                date_iso,

            "from":
                row["From"],

            "to":
                row["To"],

            "subject":
                row["Subject"],

            "body":
                row["Body"],

            "snippet":
                row["Snippet"],

            "gmail_link":
                row["Gmail_Link"],

            "retrieved_at":
                row["Retrieved_At"],
        }

        messages.append(
            message
        )

    return messages


# ============================================================
# GROUP BY THREAD
# ============================================================

def group_messages_by_thread(
    messages
):

    threads = {}

    for message in messages:

        thread_id = message[
            "thread_id"
        ]

        if not thread_id:

            thread_id = message[
                "message_id"
            ]

        if thread_id not in threads:

            threads[thread_id] = {
                "thread_id":
                    thread_id,

                "messages":
                    [],
            }

        threads[
            thread_id
        ][
            "messages"
        ].append(
            message
        )

    for thread in threads.values():

        thread[
            "messages"
        ].sort(
            key=lambda x:
                x.get(
                    "date",
                    "",
                )
        )

    return list(
        threads.values()
    )


def save_raw_json(
    messages,
    threads,
):

    payload = {

        "generated_at":
            datetime.now().isoformat(),

        "source":
            "Google Sheet / Gmail Job_Analyze",

        "message_count":
            len(messages),

        "thread_count":
            len(threads),

        "messages":
            messages,

        "threads":
            threads,
    }

    with open(
        RAW_JSON_FILE,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            payload,
            f,
            ensure_ascii=False,
            indent=2,
        )

    log("")
    log("💾 JSON raw generado:")
    log(
        f"   {RAW_JSON_FILE}"
    )


# ============================================================
# PROCESSING STATE
# ============================================================

def load_processing_state():

    if not os.path.exists(
        PROCESSING_STATE_FILE
    ):

        return {
            "threads": {}
        }

    try:

        with open(
            PROCESSING_STATE_FILE,
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(f)

        if not isinstance(
            data,
            dict,
        ):

            return {
                "threads": {}
            }

        if "threads" not in data:

            data["threads"] = {}

        return data

    except (
        json.JSONDecodeError,
        OSError,
    ):

        log(
            "⚠️ No se pudo leer "
            "processing_state.json."
        )

        log(
            "   Se iniciará un procesamiento "
            "completo."
        )

        return {
            "threads": {}
        }


def save_processing_state(
    state
):

    with open(
        PROCESSING_STATE_FILE,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            state,
            f,
            ensure_ascii=False,
            indent=2,
        )

    log("")
    log(
        "💾 Estado de procesamiento guardado:"
    )

    log(
        f"   {PROCESSING_STATE_FILE}"
    )


def get_thread_message_ids(
    thread
):

    return [
        message["message_id"]
        for message
        in thread["messages"]
        if message.get(
            "message_id"
        )
    ]


def find_pending_threads(
    threads,
    state,
):

    pending = []

    processed_threads = state.get(
        "threads",
        {},
    )

    for thread in threads:

        thread_id = thread[
            "thread_id"
        ]

        current_message_ids = set(
            get_thread_message_ids(
                thread
            )
        )

        previous = processed_threads.get(
            thread_id,
            {},
        )

        processed_message_ids = set(
            previous.get(
                "message_ids",
                [],
            )
        )

        new_message_ids = (
            current_message_ids
            - processed_message_ids
        )

        if new_message_ids:

            pending.append(
                thread
            )

    return pending


def update_processing_state(
    state,
    threads,
):

    if "threads" not in state:

        state["threads"] = {}

    for thread in threads:

        thread_id = thread[
            "thread_id"
        ]

        message_ids = (
            get_thread_message_ids(
                thread
            )
        )

        dates = [
            message.get(
                "date",
                "",
            )
            for message
            in thread["messages"]
            if message.get(
                "date",
                "",
            )
        ]

        dates = [
            date
            for date in dates
            if date
        ]

        last_activity = (
            max(dates)
            if dates
            else ""
        )

        state[
            "threads"
        ][
            thread_id
        ] = {

            "message_ids":
                message_ids,

            "message_count":
                len(message_ids),

            "last_activity":
                last_activity,

            "processed_at":
                datetime.now().isoformat(),
        }


# ============================================================
# LOAD EXISTING CLASSIFICATION
# ============================================================

def load_existing_classification():

    if not os.path.exists(
        CLASSIFIED_JSON_FILE
    ):

        return []

    try:

        with open(
            CLASSIFIED_JSON_FILE,
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(f)

        applications = data.get(
            "applications",
            [],
        )

        if not isinstance(
            applications,
            list,
        ):

            return []

        return applications

    except (
        json.JSONDecodeError,
        OSError,
    ):

        log(
            "⚠️ No se pudo leer "
            "classified_jobs.json."
        )

        log(
            "   Se reconstruirá la clasificación."
        )

        return []


# ============================================================
# PREPARE OLLAMA INPUT
# ============================================================

def prepare_thread_for_ai(
    thread,
    thread_ref,
):

    messages = []

    for message in thread[
        "messages"
    ]:

        body = normalize_body(
            message.get(
                "body",
                "",
            )
        )

        body = truncate_text(
            body
        )

        messages.append({

            "message_id":
                message[
                    "message_id"
                ],

            "date":
                message[
                    "date"
                ],

            "from":
                message[
                    "from"
                ],

            "subject":
                message[
                    "subject"
                ],

            "body":
                body,
        })

    return {

        # IMPORTANT:
        # This short reference is generated by Python.
        # The model does NOT need to reproduce
        # the real Gmail Thread_ID.
        "thread_ref":
            thread_ref,

        "messages":
            messages,
    }


# ============================================================
# SYSTEM PROMPT
# ============================================================

SYSTEM_PROMPT = r"""
You are a highly precise job application tracking assistant.

Analyze the provided email threads and determine whether each
thread represents a job application or recruitment process.

Each thread has a short Python-generated identifier called
"thread_ref".

You MUST return exactly one result for EVERY thread_ref provided.

NEVER invent a thread_ref.

NEVER omit a thread_ref.

Return the exact thread_ref provided in the input.

============================================================
JOB APPLICATION DETECTION
============================================================

Some emails may be unrelated to job searching even though they
were included in the analysis set.

For each thread determine:

is_job_application

Use true when the thread is clearly related to:

- A job application.
- A recruiter.
- A hiring process.
- An interview.
- An assessment.
- Candidate selection.
- Employment.
- A specific job vacancy.

Use false when the thread is clearly unrelated to a job search.

============================================================
STATUS
============================================================

For job applications, status MUST be exactly one of:

Applied
Action_Required
Rejected
Interview
Job_Offer


APPLIED

Use Applied when:

- The application was received.
- The application was submitted.
- The company confirms receipt.
- The application is under review.
- The application is under consideration.
- The company says they are reviewing applications.
- The company says they will contact the candidate if there
  are next steps.
- The company simply acknowledges the application.

"Next steps" by itself does NOT mean Interview.


LINKEDIN JOB ALERTS / JOB RECOMMENDATIONS (IMPORTANT)

LinkedIn job alerts, recommended-job digests, matching vacancies, and
emails listing open roles are NOT applications. Buttons or phrases such
as "Apply", "Solicitar con perfil y CV", "View job", or "See jobs" are
invitations to apply, not proof that an application was submitted and
not an action required by an employer. Mark these threads as
is_job_application=false and status=null unless the message explicitly
confirms a submitted application or concerns an existing hiring process
for a specific role (assessment, interview, rejection, or offer).

ACTION_REQUIRED

Use Action_Required ONLY when the candidate is explicitly required or
requested by an employer/recruiter to perform an action within an existing
application process. A job-alert call-to-action does not qualify.

Examples:

- Complete an assessment.
- Complete a questionnaire.
- Submit documents.
- Provide availability.
- Reply to recruiter.
- Complete coding test.
- Complete technical test.
- Fill out a form.
- Confirm information.

There must be an explicit action requested.


INTERVIEW

Use Interview ONLY when there is explicit evidence of:

- Interview invitation.
- Interview scheduling.
- Recruiter screening call.
- Phone interview.
- Video interview.
- Technical interview.
- Hiring manager interview.
- Explicit request to speak with the candidate about the role.

Generic phrases such as:

- next steps
- we will contact you
- under review
- we are reviewing your application

DO NOT mean Interview.


REJECTED

Use Rejected ONLY when there is explicit evidence that:

- The candidate was not selected.
- The company will not proceed.
- Another candidate was selected.
- The position was filled.
- The application was closed.
- The process will not continue.

If the thread contains:

Interview → Rejection

CURRENT status is Rejected.


JOB_OFFER

Use Job_Offer ONLY when there is explicit evidence of:

- Employment offer.
- Job offer.
- Offer letter.
- Actual offer to employ the candidate.

If:

Interview → Offer

CURRENT status is Job_Offer.


============================================================
CURRENT STATUS
============================================================

Read the thread chronologically.

The latest meaningful event determines the current status.

Examples:

Application received
→ Applied

Application received
→ Assessment requested
→ Action_Required

Application received
→ Interview invitation
→ Interview

Application received
→ Interview
→ Rejection
→ Rejected

Application received
→ Interview
→ Offer
→ Job_Offer

If an interview happened previously and the latest message
is simply "we will be in touch", keep Interview.


============================================================
COMPANY
============================================================

Determine the company primarily from the subject.

Use sender and body as secondary evidence.

Do not invent a company.


============================================================
POSITION
============================================================

Determine the position from subject and body.

Use the most specific job title available.

Do not invent a position.


============================================================
EVIDENCE
============================================================

Provide concise evidence supporting the CURRENT status.

Evidence must be based only on the emails.

Do not invent evidence.


============================================================
CONFIDENCE
============================================================

Return confidence from 0.0 to 1.0.

0.90 - 1.00 = very explicit
0.75 - 0.89 = strong
0.50 - 0.74 = reasonable but ambiguous
below 0.50 = genuinely uncertain


============================================================
OUTPUT
============================================================

Return ONLY valid JSON.

Do not use Markdown.

Do not use code fences.

Return EXACTLY this structure:

{
  "applications": [
    {
      "thread_ref": "T1",
      "is_job_application": true,
      "company": "string",
      "position": "string",
      "status": "Applied",
      "confidence": 0.95,
      "evidence": [
        "string"
      ]
    }
  ]
}

For is_job_application = false:

{
  "thread_ref": "T1",
  "is_job_application": false,
  "company": "",
  "position": "",
  "status": null,
  "confidence": 0.95,
  "evidence": [
    "Thread is unrelated to job searching."
  ]
}

Return exactly one object for every input thread_ref.
"""


def build_user_prompt(
    threads
):

    ai_threads = []

    for index, thread in enumerate(
        threads,
        start=1,
    ):

        thread_ref = (
            f"T{index}"
        )

        ai_threads.append(
            prepare_thread_for_ai(
                thread,
                thread_ref,
            )
        )

    payload = {

        "task":
            "Classify job application threads.",

        "threads":
            ai_threads,
    }

    return json.dumps(
        payload,
        ensure_ascii=False,
        indent=2,
    )


def estimate_prompt_size(
    threads
):

    prompt = build_user_prompt(
        threads
    )

    return len(prompt)


def create_dynamic_batches(
    threads
):

    batches = []

    current_batch = []

    current_size = 0

    for thread in threads:

        thread_size = (
            estimate_prompt_size(
                [thread]
            )
        )

        if (
            thread_size
            > MAX_BATCH_CHARS
        ):

            if current_batch:

                batches.append(
                    current_batch
                )

                current_batch = []
                current_size = 0

            batches.append(
                [thread]
            )

            continue

        if (
            current_batch
            and
            current_size
            + thread_size
            > MAX_BATCH_CHARS
        ):

            batches.append(
                current_batch
            )

            current_batch = []
            current_size = 0

        current_batch.append(
            thread
        )

        current_size += thread_size

    if current_batch:

        batches.append(
            current_batch
        )

    return batches


# ============================================================
# JSON EXTRACTION
# ============================================================

def extract_json(
    text
):

    text = text.strip()

    text = re.sub(
        r"^```json\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"^```\s*",
        "",
        text,
    )

    text = re.sub(
        r"\s*```$",
        "",
        text,
    )

    text = text.strip()

    try:

        return json.loads(
            text
        )

    except json.JSONDecodeError:

        start = text.find(
            "{"
        )

        end = text.rfind(
            "}"
        )

        if (
            start >= 0
            and
            end > start
        ):

            candidate = (
                text[
                    start:end + 1
                ]
            )

            return json.loads(
                candidate
            )

        raise


# ============================================================
# VALIDATE AI RESPONSE
# ============================================================

def validate_classification(
    data,
    source_threads,
):

    if not isinstance(
        data,
        dict,
    ):

        raise ValueError(
            "La respuesta de Ollama "
            "no es un objeto JSON."
        )

    applications = data.get(
        "applications"
    )

    if not isinstance(
        applications,
        list,
    ):

        raise ValueError(
            "La respuesta no contiene "
            "'applications' como lista."
        )

    if len(applications) != len(
        source_threads
    ):

        raise ValueError(
            "Ollama no devolvió exactamente "
            "un resultado por thread. "
            f"Esperados: {len(source_threads)}; "
            f"recibidos: {len(applications)}."
        )

    expected_refs = {
        f"T{i}"
        for i in range(
            1,
            len(source_threads) + 1,
        )
    }

    received_refs = set()

    validated = []

    for app in applications:

        if not isinstance(
            app,
            dict,
        ):

            raise ValueError(
                "Ollama devolvió un elemento "
                "que no es un objeto."
            )

        thread_ref = clean_text(
            app.get(
                "thread_ref"
            )
        )

        if not thread_ref:

            raise ValueError(
                "Una clasificación "
                "no tiene thread_ref."
            )

        if thread_ref not in (
            expected_refs
        ):

            raise ValueError(
                "Ollama devolvió un "
                f"thread_ref inválido: "
                f"{thread_ref}"
            )

        if thread_ref in (
            received_refs
        ):

            raise ValueError(
                "Ollama devolvió "
                "un thread_ref duplicado: "
                f"{thread_ref}"
            )

        received_refs.add(
            thread_ref
        )

        is_job_application = (
            app.get(
                "is_job_application"
            )
        )

        if not isinstance(
            is_job_application,
            bool,
        ):

            raise ValueError(
                f"is_job_application inválido "
                f"para {thread_ref}."
            )

        status = app.get(
            "status"
        )

        if is_job_application:

            status = clean_text(
                status
            )

            if status not in (
                ALLOWED_STATUSES
            ):

                raise ValueError(
                    f"Status inválido para "
                    f"{thread_ref}: {status}"
                )

        else:

            status = None

        confidence = app.get(
            "confidence",
            0.0,
        )

        try:

            confidence = float(
                confidence
            )

        except (
            TypeError,
            ValueError,
        ):

            confidence = 0.0

        confidence = max(
            0.0,
            min(
                1.0,
                confidence,
            ),
        )

        evidence = app.get(
            "evidence",
            [],
        )

        if not isinstance(
            evidence,
            list,
        ):

            evidence = [
                str(evidence)
            ]

        evidence = [
            clean_text(item)
            for item in evidence
            if clean_text(item)
        ]

        validated.append({

            "thread_ref":
                thread_ref,

            "is_job_application":
                is_job_application,

            "company":
                clean_text(
                    app.get(
                        "company"
                    )
                ),

            "position":
                clean_text(
                    app.get(
                        "position"
                    )
                ),

            "status":
                status,

            "confidence":
                confidence,

            "evidence":
                evidence,
        })

    missing = (
        expected_refs
        - received_refs
    )

    if missing:

        raise ValueError(
            "Ollama omitió thread_refs: "
            f"{sorted(missing)}"
        )

    return validated


# ============================================================
# MAP AI RESULTS BACK TO REAL THREADS
# ============================================================

def convert_ai_results_to_applications(
    ai_results,
    source_threads,
):

    applications = []

    for index, result in enumerate(
        ai_results
    ):

        thread = source_threads[
            index
        ]

        thread_id = thread[
            "thread_id"
        ]

        if not result[
            "is_job_application"
        ]:

            log(
                f"   ℹ Thread {thread_id} "
                "marcado como no relacionado "
                "con job search."
            )

            continue

        messages = thread[
            "messages"
        ]

        message_ids = [
            message[
                "message_id"
            ]
            for message
            in messages
            if message.get(
                "message_id"
            )
        ]

        dates = [
            message.get(
                "date",
                "",
            )
            for message
            in messages
            if message.get(
                "date",
                "",
            )
        ]

        dates = [
            date
            for date in dates
            if date
        ]

        last_activity = (
            max(dates)
            if dates
            else ""
        )

        gmail_link = ""

        for message in reversed(
            messages
        ):

            link = clean_text(
                message.get(
                    "gmail_link",
                    "",
                )
            )

            if link:

                gmail_link = link
                break

        applications.append({

            "thread_id":
                thread_id,

            "company":
                result[
                    "company"
                ],

            "position":
                result[
                    "position"
                ],

            "status":
                result[
                    "status"
                ],

            "confidence":
                result[
                    "confidence"
                ],

            "last_activity":
                last_activity,

            "evidence":
                result[
                    "evidence"
                ],

            "message_ids":
                message_ids,

            "gmail_link":
                gmail_link,
        })

    return applications


# ============================================================
# OLLAMA CALL
# ============================================================

def _is_context_overflow(response):
    """Return True when Ollama rejected the prompt because it exceeds num_ctx."""
    if response.status_code not in (400, 413):
        return False
    message = (response.text or "").casefold()
    return (
        "exceeds the available context size" in message
        or "exceed_context_size_error" in message
        or "context length" in message and "exceed" in message
    )


def _shrink_thread_for_retry(thread, total_body_chars=2200, per_message_chars=900):
    """Create a smaller copy of a thread while preserving IDs and message metadata."""
    import copy

    smaller = copy.deepcopy(thread)
    messages = smaller.get("messages", [])
    if not messages:
        return smaller

    # Keep the most recent context preferentially, but preserve chronological order.
    remaining = total_body_chars
    for message in reversed(messages):
        body = normalize_body(message.get("body", ""))
        allowance = min(per_message_chars, remaining)
        if allowance <= 0:
            message["body"] = "[Body omitted to fit Ollama context window]"
            continue
        if len(body) > allowance:
            if allowance >= 160:
                head = int(allowance * 0.65)
                tail = allowance - head
                body = body[:head] + "\n\n[... email content shortened ...]\n\n" + body[-tail:]
            else:
                body = body[:allowance]
        message["body"] = body
        remaining -= min(len(body), allowance)

    return smaller


def _call_ollama_batch(batch_threads, label="batch", depth=0):
    """Call Ollama for one batch; split/reduce prompts automatically on context errors."""
    if depth > 8:
        raise RuntimeError(
            f"No se pudo reducir el prompt de {label} lo suficiente tras varios intentos."
        )

    user = build_user_prompt(batch_threads)
    payload = {
        "model": OLLAMA_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user},
        ],
        "stream": False,
        "format": "json",
        "keep_alive": "10m",
        "options": {"temperature": 0.1},
    }

    try:
        response = requests.post(
            OLLAMA_URL,
            json=payload,
            timeout=OLLAMA_TIMEOUT,
        )
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Error conectando con Ollama en {label}: {e}") from e

    if not response.ok:
        if _is_context_overflow(response):
            log(f"\n⚠️ Contexto excedido en {label}; aplicando recuperación automática.")
            if len(batch_threads) > 1:
                midpoint = max(1, len(batch_threads) // 2)
                left = _call_ollama_batch(
                    batch_threads[:midpoint], f"{label} (parte A)", depth + 1
                )
                time.sleep(BATCH_COOLDOWN_SECONDS)
                right = _call_ollama_batch(
                    batch_threads[midpoint:], f"{label} (parte B)", depth + 1
                )
                return left + right

            original = batch_threads[0]
            reduced = _shrink_thread_for_retry(original)
            original_size = len(build_user_prompt([original]))
            reduced_size = len(build_user_prompt([reduced]))
            if reduced_size >= original_size:
                raise RuntimeError(
                    f"Ollama rechazó {label} por contexto y no fue posible reducir el hilo."
                )
            log(
                f"   ↘ Hilo largo reducido para reintento: "
                f"{original_size:,} → {reduced_size:,} caracteres de prompt."
            )
            return _call_ollama_batch([reduced], f"{label} (hilo reducido)", depth + 1)

        log(f"\n❌ Ollama respondió con error en {label}")
        log(f"   HTTP Status: {response.status_code}")
        log(response.text)
        raise RuntimeError(f"Ollama HTTP {response.status_code}: {response.text}")

    try:
        data = response.json()
    except ValueError as e:
        raise RuntimeError(f"Ollama no devolvió JSON válido en {label}: {e}") from e

    if "message" not in data:
        raise RuntimeError(f"Respuesta de Ollama sin 'message' en {label}.")

    content = data["message"].get("content", "").strip()
    if not content:
        raise RuntimeError(f"Ollama devolvió respuesta vacía en {label}.")

    try:
        classification = extract_json(content)
    except json.JSONDecodeError as e:
        log(f"\n❌ JSON inválido en {label}\n{content}\n")
        raise RuntimeError(f"JSON inválido en {label}: {e}") from e

    ai_results = validate_classification(classification, batch_threads)
    applications = convert_ai_results_to_applications(ai_results, batch_threads)

    log(f"   ✓ Resultados recibidos: {len(ai_results)}")
    log(f"   ✓ Aplicaciones válidas: {len(applications)}")
    return applications


def call_ollama(threads):
    log("")
    log("🧠 Procesando threads con Ollama...")
    log(f"   Modelo: {OLLAMA_MODEL}")
    log(f"   Threads a procesar: {len(threads)}")

    if not threads:
        return {"applications": []}

    batches = create_dynamic_batches(threads)
    log(f"   Batches generados: {len(batches)}")

    all_applications = []
    total_batches = len(batches)

    for batch_number, batch_threads in enumerate(batches, start=1):
        estimated_chars = estimate_prompt_size(batch_threads)
        estimated_tokens = estimated_chars // 4
        log("")
        log(f"🔹 Batch {batch_number}/{total_batches}")
        log(f"   Threads: {len(batch_threads)}")
        log(f"   Prompt estimado: ~{estimated_tokens:,} tokens")

        applications = _call_ollama_batch(
            batch_threads,
            label=f"batch {batch_number}/{total_batches}",
        )
        all_applications.extend(applications)

        if batch_number < total_batches:
            log(
                f"   ⏸ Pausa de {BATCH_COOLDOWN_SECONDS}s "
                "para reducir la carga térmica..."
            )
            time.sleep(BATCH_COOLDOWN_SECONDS)

    log("")
    log(f"✓ Total de aplicaciones nuevas/actualizadas: {len(all_applications)}")
    return {"applications": all_applications}


# ============================================================
# MERGE CLASSIFICATION
# ============================================================

def merge_classification(
    existing_applications,
    updated_applications,
):

    merged = {}

    # Existing results first
    for app in existing_applications:

        thread_id = clean_text(
            app.get(
                "thread_id"
            )
        )

        if thread_id:

            merged[
                thread_id
            ] = app

    # New results replace old results
    for app in updated_applications:

        thread_id = clean_text(
            app.get(
                "thread_id"
            )
        )

        if thread_id:

            merged[
                thread_id
            ] = app

    return list(
        merged.values()
    )


# ============================================================
# SAVE CLASSIFICATION
# ============================================================

def save_classification(
    applications
):

    for app in applications:
        ensure_application_identity(app)

    payload = {

        "generated_at":
            datetime.now().isoformat(),

        "model":
            OLLAMA_MODEL,

        "application_count":
            len(applications),

        "applications":
            applications,
    }

    with open(
        CLASSIFIED_JSON_FILE,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            payload,
            f,
            ensure_ascii=False,
            indent=2,
        )

    log("")
    log(
        "💾 Clasificación guardada:"
    )

    log(
        f"   {CLASSIFIED_JSON_FILE}"
    )


# ============================================================
# STABLE APPLICATION IDENTITY + MANUAL OVERRIDES
# ============================================================

def _source_thread_ids(app):
    values = app.get("thread_ids", []) or []
    if isinstance(values, str):
        values = [values]
    values = list(values) + [app.get("thread_id", "")]
    return sorted({clean_text(value) for value in values if clean_text(value)})


def ensure_application_identity(app):
    """Assign a persistent ID once; edits to company/title must not change it."""
    if not clean_text(app.get("legacy_application_id")):
        legacy_identity = "|".join([
            _normalized_identity_text(app.get("company")),
            _normalized_identity_text(app.get("position")),
        ])
        app["legacy_application_id"] = hashlib.sha256(legacy_identity.encode("utf-8")).hexdigest()[:20]
    if clean_text(app.get("identity_uid")):
        return app
    source_ids = _source_thread_ids(app)
    if source_ids:
        seed = "threads:" + "|".join(source_ids)
    else:
        seed = "legacy:" + "|".join([
            _normalized_identity_text(app.get("company")),
            _normalized_identity_text(app.get("position")),
            clean_text(app.get("last_activity")),
        ])
    app["identity_uid"] = hashlib.sha256(seed.encode("utf-8")).hexdigest()[:24]
    return app


def load_manual_overrides():
    if not MANUAL_OVERRIDES_FILE.exists():
        return {}
    try:
        with open(MANUAL_OVERRIDES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        log("⚠️ No se pudo leer manual_overrides.json; se mantienen los nombres originales.")
        return {}


def save_manual_overrides(payload):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temporary = MANUAL_OVERRIDES_FILE.with_suffix(MANUAL_OVERRIDES_FILE.suffix + ".tmp")
    with open(temporary, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(temporary, MANUAL_OVERRIDES_FILE)


def apply_manual_overrides(applications, overrides=None):
    """Apply user-defined canonical company/title before duplicate grouping."""
    overrides = overrides if overrides is not None else load_manual_overrides()
    for app in applications:
        ensure_application_identity(app)
        # After two records are manually consolidated, their former IDs remain
        # aliases. Continue honoring those overrides on later incremental runs.
        identity_keys = [app["identity_uid"]] + [
            clean_text(value) for value in (app.get("identity_aliases", []) or [])
            if clean_text(value)
        ]
        company = ""
        position = ""
        for identity_key in identity_keys:
            override = overrides.get(identity_key, {}) or {}
            if not isinstance(override, dict):
                continue
            company = clean_text(override.get("company")) or company
            position = clean_text(override.get("position")) or position
        if company:
            app["company"] = company
        if position:
            app["position"] = position
    return applications


def set_manual_override(app, company, position):
    """Persist a canonical identity for one stable application record."""
    ensure_application_identity(app)
    overrides = load_manual_overrides()
    current = overrides.get(app["identity_uid"], {}) or {}
    if not isinstance(current, dict):
        current = {}
    company = clean_text(company)
    position = clean_text(position)
    if company:
        current["company"] = company
    else:
        current.pop("company", None)
    if position:
        current["position"] = position
    else:
        current.pop("position", None)
    if current:
        current["updated_at"] = datetime.now().isoformat()
        overrides[app["identity_uid"]] = current
    else:
        overrides.pop(app["identity_uid"], None)
    save_manual_overrides(overrides)
    return overrides


# ============================================================
# LOCAL DEDUPLICATION / JOB ALERT CLEANUP
# ============================================================

def _norm_vacancy_key(value):
    # Keep seniority terms (Senior/Junior/Lead/etc.): they can identify
    # genuinely different requisitions and must not be silently merged.
    value = clean_text(value).casefold()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _looks_like_linkedin_alert(app):
    company = clean_text(app.get("company", "")).casefold()
    evidence = " ".join(clean_text(x) for x in (app.get("evidence") or [])).casefold()
    position = clean_text(app.get("position", "")).casefold()
    blob = " ".join([company, evidence, position])
    linkedin_related = "linkedin" in company or "linkedin" in blob
    alert_terms = ("job alert", "recommended jobs", "job recommendations", "alerta de empleo",
                   "solicitar con perfil y cv", "matching jobs", "jobs you may be interested",
                   "jobs you might like", "vacantes recomendadas")
    return linkedin_related and any(term in blob for term in alert_terms)


def _normalized_identity_text(value):
    value = unicodedata.normalize("NFKD", clean_text(value).casefold())
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _seniority_tokens(value):
    text = " " + _normalized_identity_text(value) + " "
    terms = {
        "intern", "internship", "junior", "jr", "mid", "middle", "senior", "sr",
        "lead", "staff", "principal", "manager", "director", "head", "vp", "chief",
        "entry", "associate", "expert",
    }
    return {term for term in terms if re.search(r"(?<!\w)" + re.escape(term) + r"(?!\w)", text)}


def _requisition_ids(app):
    blob = " ".join([
        clean_text(app.get("position")),
        clean_text(app.get("evidence")),
        clean_text(app.get("requisition_id")),
    ]).upper()
    patterns = [r"\bREQ(?:UISITION)?[\s:#-]*([A-Z0-9-]{3,})\b", r"\bJOB[\s:#-]*ID[\s:#-]*([A-Z0-9-]{3,})\b", r"\bID[\s:#-]*([A-Z0-9-]{4,})\b"]
    found = set()
    for pattern in patterns:
        found.update(re.findall(pattern, blob))
    return found


def _same_vacancy_candidate(left, right):
    company_left = _normalized_identity_text(left.get("company"))
    company_right = _normalized_identity_text(right.get("company"))
    if not company_left or company_left != company_right:
        return False

    pos_left = _normalized_identity_text(left.get("position"))
    pos_right = _normalized_identity_text(right.get("position"))
    if not pos_left or not pos_right:
        return False

    # Avoid merging distinct seniority levels or explicit requisition IDs,
    # even when the displayed job title is identical.
    seniority_left = _seniority_tokens(pos_left)
    seniority_right = _seniority_tokens(pos_right)
    if seniority_left != seniority_right:
        return False
    req_left = _requisition_ids(left)
    req_right = _requisition_ids(right)
    if req_left and req_right and req_left.isdisjoint(req_right):
        return False
    if pos_left == pos_right:
        return True

    ratio = difflib.SequenceMatcher(None, pos_left, pos_right).ratio()
    # Only merge very close title variants (punctuation/short wording changes).
    return ratio >= 0.93


def _merge_duplicate_group(group):
    group.sort(key=lambda x: clean_text(x.get("last_activity", "")))
    merged = dict(group[-1])
    # Keep the first/oldest group's stable ID when new emails join it.
    for item in group:
        ensure_application_identity(item)
    primary_uid = group[0]["identity_uid"]
    aliases = []
    legacy_ids = []
    for item in group:
        legacy_id = clean_text(item.get("legacy_application_id"))
        if legacy_id and legacy_id not in legacy_ids:
            legacy_ids.append(legacy_id)
        for old_id in item.get("legacy_application_ids", []) or []:
            old_id = clean_text(old_id)
            if old_id and old_id not in legacy_ids:
                legacy_ids.append(old_id)
        uid = clean_text(item.get("identity_uid"))
        if uid and uid not in aliases:
            aliases.append(uid)
        for alias in item.get("identity_aliases", []) or []:
            alias = clean_text(alias)
            if alias and alias not in aliases:
                aliases.append(alias)
    merged["identity_uid"] = primary_uid
    merged["identity_aliases"] = [uid for uid in aliases if uid != primary_uid]
    merged["legacy_application_ids"] = legacy_ids
    evidence, message_ids, thread_ids = [], [], []
    for item in group:
        values = item.get("evidence", []) or []
        if isinstance(values, str):
            values = [values]
        for value in values:
            value = clean_text(value)
            if value and value not in evidence:
                evidence.append(value)
        for value in item.get("message_ids", []) or []:
            value = clean_text(value)
            if value and value not in message_ids:
                message_ids.append(value)
        source_thread_ids = item.get("thread_ids", []) or []
        if isinstance(source_thread_ids, str):
            source_thread_ids = [source_thread_ids]
        source_thread_ids = list(source_thread_ids) + [item.get("thread_id", "")]
        for source_id in source_thread_ids:
            source_id = clean_text(source_id)
            if source_id and source_id not in thread_ids:
                thread_ids.append(source_id)
    merged["evidence"] = evidence[:8]
    merged["message_ids"] = message_ids
    merged["thread_ids"] = thread_ids
    merged["source_thread_count"] = len(thread_ids) or len(group)
    if not clean_text(merged.get("gmail_link")):
        merged["gmail_link"] = next((x.get("gmail_link") for x in reversed(group) if clean_text(x.get("gmail_link"))), "")
    return merged


def deduplicate_applications(applications):
    """Remove alert emails and consolidate exact/near-identical vacancy records conservatively."""
    candidates = []
    for app in applications:
        ensure_application_identity(app)
        if _looks_like_linkedin_alert(app):
            continue
        if not clean_text(app.get("company")) or not clean_text(app.get("position")):
            candidates.append([app])
            continue
        match = None
        for group in candidates:
            representative = group[0]
            if _same_vacancy_candidate(representative, app):
                match = group
                break
        if match is None:
            candidates.append([app])
        else:
            match.append(app)
    result = [_merge_duplicate_group(group) if len(group) > 1 else group[0] for group in candidates]
    return sorted(result, key=lambda x: clean_text(x.get("last_activity", "")), reverse=True)


# ============================================================
# HTML HELPERS
# ============================================================

def esc(value):

    return html.escape(
        clean_text(value)
    )


def status_class(
    status
):

    return re.sub(
        r"[^A-Za-z0-9_-]",
        "",
        status,
    )


# ============================================================
# PLOTLY
# ============================================================

STATUS_CHART_CATEGORIES = ["Pending", "Interview", "Applied", "Rejected", "Offer"]


def get_total_submissions_count(applications, actions=None):
    """Count all tracked submissions, excluding records manually marked unrelated.

    Completed or dismissed follow-ups still count: those applications were submitted
    and remain part of the overall application history.
    """
    actions = actions or {}
    return sum(1 for app in applications if not is_application_unrelated(actions, app))


def get_main_status_counts(applications, actions=None):
    """Return the five dashboard categories after applying persistent user decisions.

    Action_Required contributes to Pending unless completed or dismissed.
    Interview remains Interview unless explicitly dismissed. Completed interviews
    remain in the Interview stage; unrelated applications are excluded entirely.
    """
    actions = actions or {}
    counts = {label: 0 for label in STATUS_CHART_CATEGORIES}
    for app in applications:
        if is_application_unrelated(actions, app):
            continue
        status = clean_text(app.get("status"))
        if status == "Action_Required":
            if not is_task_completed(actions, app, "action") and not is_task_dismissed(actions, app, "action"):
                counts["Pending"] += 1
        elif status == "Interview":
            if not is_task_dismissed(actions, app, "interview"):
                counts["Interview"] += 1
        elif status == "Applied":
            counts["Applied"] += 1
        elif status == "Rejected":
            counts["Rejected"] += 1
        elif status == "Job_Offer":
            counts["Offer"] += 1
    return counts


def create_status_chart(applications, actions=None):
    """Build the consolidated five-category status chart."""
    counts = get_main_status_counts(applications, actions)
    labels = STATUS_CHART_CATEGORIES
    values = [counts[label] for label in labels]
    fig = go.Figure(data=[go.Bar(x=labels, y=values, text=values, textposition="auto")])
    fig.update_layout(
        title="Applications by Current Status",
        xaxis_title="Current status",
        yaxis_title="Applications",
        template="plotly_white",
        height=420,
    )
    return fig


def create_timeline_chart(
    applications
):

    dates = []

    for app in applications:

        date_text = app.get(
            "last_activity",
            "",
        )

        if not date_text:
            continue

        date = pd.to_datetime(
            date_text,
            errors="coerce",
        )

        if not pd.isna(
            date
        ):

            dates.append(
                date
            )

    if not dates:

        fig = go.Figure(data=[go.Scatter(x=[], y=[], mode="lines+markers")])

        fig.update_layout(
            title="Applications Over Time",
            xaxis_title="Date",
            yaxis_title="Applications",
            template="plotly_white",
            height=450,
        )

        return fig

    series = (
        pd.Series(dates)
        .dt.date
        .value_counts()
        .sort_index()
    )

    values = (
        series.values.tolist()
    )

    dates = [
        str(date)
        for date
        in series.index
    ]

    fig = go.Figure(
        data=[
            go.Scatter(
                x=dates,
                y=values,
                mode="lines+markers",
            )
        ]
    )

    fig.update_layout(
        title=
            "Applications Over Time",

        xaxis_title=
            "Date",

        yaxis_title=
            "Applications",

        template=
            "plotly_white",

        height=
            450,
    )

    return fig


# ============================================================
# HTML REPORT
# ============================================================

def application_id(app):
    ensure_application_identity(app)
    return hashlib.sha256(app["identity_uid"].encode("utf-8")).hexdigest()[:20]


def application_action_ids(app):
    ensure_application_identity(app)
    ids = [application_id(app)]
    aliases = app.get("identity_aliases", []) or []
    for uid in aliases:
        alias_id = hashlib.sha256(clean_text(uid).encode("utf-8")).hexdigest()[:20]
        if alias_id not in ids:
            ids.append(alias_id)
    legacy_ids = list(app.get("legacy_application_ids", []) or [])
    legacy_id = clean_text(app.get("legacy_application_id"))
    if legacy_id:
        legacy_ids.append(legacy_id)
    for old_id in legacy_ids:
        old_id = clean_text(old_id)
        if old_id and old_id not in ids:
            ids.append(old_id)
    return ids


def get_application_action_record(actions, app, task_name):
    for app_id in application_action_ids(app):
        record = (actions.get(app_id, {}) or {}).get(task_name, {}) or {}
        if record:
            return record
    return {}


def application_activity_key(app):
    message_ids = app.get("message_ids", []) or []
    if isinstance(message_ids, str):
        message_ids = [message_ids]
    message_ids = sorted({clean_text(value) for value in message_ids if clean_text(value)})
    if message_ids:
        source = "|".join(message_ids)
    else:
        thread_ids = app.get("thread_ids", []) or []
        if isinstance(thread_ids, str):
            thread_ids = [thread_ids]
        source = "|".join(sorted(clean_text(x) for x in thread_ids if clean_text(x)))
        source += "|" + clean_text(app.get("last_activity"))
    return hashlib.sha256(source.encode("utf-8")).hexdigest()[:24]


def load_user_actions():
    if not os.path.exists(USER_ACTIONS_FILE):
        return {}
    try:
        with open(USER_ACTIONS_FILE, "r", encoding="utf-8") as f:
            payload = json.load(f)
        return payload if isinstance(payload, dict) else {}
    except (OSError, json.JSONDecodeError):
        log("⚠️ No se pudo leer user_actions.json; se mostrarán las acciones como pendientes.")
        return {}


def save_user_actions(payload):
    os.makedirs(DATA_DIR, exist_ok=True)
    temporary = USER_ACTIONS_FILE.with_suffix(USER_ACTIONS_FILE.suffix + ".tmp")
    with open(temporary, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(temporary, USER_ACTIONS_FILE)


def is_task_completed(actions, app, task_name):
    record = get_application_action_record(actions, app, task_name)
    return bool(record.get("done")) and record.get("activity_key") == application_activity_key(app)


def is_task_dismissed(actions, app, task_name):
    record = get_application_action_record(actions, app, f"dismiss_{task_name}")
    return bool(record.get("done")) and record.get("activity_key") == application_activity_key(app)


def is_application_unrelated(actions, app):
    record = get_application_action_record(actions, app, "unrelated")
    return bool(record.get("done"))


def create_task_completion_chart(applications, actions):
    """Chart completed, pending and manually dismissed follow-ups."""
    task_labels = ["Actions required", "Interviews"]
    completed_values = [0, 0]
    pending_values = [0, 0]
    dismissed_values = [0, 0]

    for app in applications:
        if is_application_unrelated(actions, app):
            continue
        status = clean_text(app.get("status"))
        if status == "Action_Required":
            index, task = 0, "action"
        elif status == "Interview":
            index, task = 1, "interview"
        else:
            continue

        if is_task_completed(actions, app, task):
            completed_values[index] += 1
        elif is_task_dismissed(actions, app, task):
            dismissed_values[index] += 1
        else:
            pending_values[index] += 1

    fig = go.Figure(data=[
        go.Bar(name="Completed", x=task_labels, y=completed_values, text=completed_values, textposition="auto"),
        go.Bar(name="Pending", x=task_labels, y=pending_values, text=pending_values, textposition="auto"),
        go.Bar(name="Dismissed", x=task_labels, y=dismissed_values, text=dismissed_values, textposition="auto"),
    ])
    fig.update_layout(
        title="Action and Interview Follow-up Progress",
        xaxis_title="Follow-up type",
        yaxis_title="Applications",
        barmode="stack",
        template="plotly_white",
        height=420,
        legend_title_text="Follow-up state",
    )
    return fig, completed_values, pending_values, dismissed_values


def generate_html_report(applications):
    # Generate a single overwriteable HTML report with persistent completion controls.
    os.makedirs(REPORT_DIR, exist_ok=True)
    report_file = REPORT_FILE
    actions = load_user_actions()

    # Remove only reports produced by previous versions of this tracker.
    for filename in os.listdir(REPORT_DIR):
        if re.fullmatch(r"reporte_\d{4}-\d{2}-\d{2}\.html", filename):
            try:
                os.remove(os.path.join(REPORT_DIR, filename))
            except OSError:
                pass

    main_counts = get_main_status_counts(applications, actions)
    total_submissions = get_total_submissions_count(applications, actions)
    status_chart_html = create_status_chart(applications, actions).to_html(full_html=False, include_plotlyjs=True, div_id="status-chart")
    timeline_chart_html = create_timeline_chart([app for app in applications if not is_application_unrelated(actions, app)]).to_html(full_html=False, include_plotlyjs=False)
    sorted_apps = sorted(applications, key=lambda app: clean_text(app.get("last_activity", "")), reverse=True)

    # Prioritize company groups by their most urgent vacancy. Pending Action_Required
    # and pending Interview follow-ups come first, then offers/applied, then rejected,
    # and finally completed/dismissed or otherwise non-pending records.
    def vacancy_display_priority(app):
        status = clean_text(app.get("status"))
        if is_application_unrelated(actions, app):
            return 5
        if status == "Action_Required":
            return 4 if (is_task_completed(actions, app, "action") or is_task_dismissed(actions, app, "action")) else 0
        if status == "Interview":
            return 4 if (is_task_completed(actions, app, "interview") or is_task_dismissed(actions, app, "interview")) else 0
        if status in ("Applied", "Job_Offer"):
            return 1
        if status == "Rejected":
            return 2
        return 3

    company_groups = {}
    for app in sorted_apps:
        company = clean_text(app.get("company")) or "Unknown company"
        company_groups.setdefault(company, []).append(app)

    def company_group_sort(item):
        company, company_apps = item
        return (min(vacancy_display_priority(app) for app in company_apps), company.casefold())

    company_sections = []
    for company, company_apps in sorted(company_groups.items(), key=company_group_sort):
        # sorted_apps is already newest-first; Python's stable sort preserves that
        # order inside each priority tier.
        company_apps = sorted(company_apps, key=vacancy_display_priority)
        vacancy_items = []
        for app in company_apps:
            status = clean_text(app.get("status")) or "Unknown"
            position = clean_text(app.get("position")) or "Position not specified"
            activity = clean_text(app.get("last_activity")) or "—"
            is_unrelated = is_application_unrelated(actions, app)
            task = "action" if status == "Action_Required" else "interview" if status == "Interview" else ""
            task_control = ""
            if is_unrelated:
                task_control = ' <span class="completion-mini">Excluded as unrelated</span>'
            elif task:
                done = is_task_completed(actions, app, task)
                dismissed = is_task_dismissed(actions, app, task)
                state_label = "✓ Completed" if done else "⊘ Dismissed" if dismissed else "○ Pending"
                task_control = ' <span class="completion-mini">' + state_label + "</span>"
            done_state = is_task_completed(actions, app, task) if task else False
            dismissed_state = is_task_dismissed(actions, app, task) if task else False
            vacancy_items.append(
                '<li data-app-id="{}" data-priority="{}" data-status="{}" data-task="{}" data-done="{}" data-dismissed="{}" data-unrelated="{}">'
                '<strong>{}</strong> <span class="status {}">{}</span>{} '
                '<span class="muted">Last activity: {}</span></li>'.format(
                    esc(application_id(app)), vacancy_display_priority(app), esc(status), task,
                    str(done_state).lower(), str(dismissed_state).lower(), str(is_unrelated).lower(),
                    esc(position), status_class(status), esc(status), task_control, esc(activity)
                )
            )
        pending_count = sum(1 for app in company_apps if vacancy_display_priority(app) == 0)
        company_sections.append(
            '<details class="company-group" open data-priority="{}"><summary><strong>{}</strong>'
            '<span class="company-count">{} vacancy/vacancies · {} pending</span></summary><ul>{}</ul></details>'.format(
                min(vacancy_display_priority(app) for app in company_apps),
                esc(company), len(company_apps), pending_count, "".join(vacancy_items)
            )
        )
    company_groups_html = "\n".join(company_sections) or '<p class="muted">No applications to display.</p>'

    rows = []
    for app in sorted_apps:
        evidence = app.get("evidence") or []
        if isinstance(evidence, str):
            evidence = [evidence]
        evidence_html = "".join(f"<li>{esc(item)}</li>" for item in evidence if clean_text(item))
        if not evidence_html:
            evidence_html = '<li class="muted">No evidence provided.</li>'

        gmail_link = clean_text(app.get("gmail_link"))
        gmail_html = f'<a href="{esc(gmail_link)}" target="_blank" rel="noopener noreferrer">Open Gmail</a>' if gmail_link else "—"
        try:
            confidence_value = float(app.get("confidence", 0))
            confidence_text = f"{confidence_value:.0%}"
            confidence_sort = str(confidence_value)
        except (TypeError, ValueError):
            confidence_text, confidence_sort = "—", "-1"

        company = clean_text(app.get("company")) or "Unknown company"
        position = clean_text(app.get("position"))
        status = clean_text(app.get("status"))
        last_activity = clean_text(app.get("last_activity"))
        evidence_text = " ".join(clean_text(x) for x in evidence)
        app_id = application_id(app)
        activity_key = application_activity_key(app)
        is_unrelated = is_application_unrelated(actions, app)
        action_buttons = []
        task = "action" if status == "Action_Required" else "interview" if status == "Interview" else ""
        done = is_task_completed(actions, app, task) if task else False
        dismissed = is_task_dismissed(actions, app, task) if task else False
        if task:
            task_label = ("Action completed" if task == "action" else "Interview completed") if done else ("Mark action done" if task == "action" else "Mark interview done")
            action_buttons.append(f'<button type="button" class="task-button {"done" if done else ""}" data-app-id="{app_id}" data-activity-key="{activity_key}" data-task="{task}" data-done="{str(done).lower()}">{task_label}</button>')
            dismiss_label = "Restore pending" if dismissed else "Dismiss pending"
            action_buttons.append(f'<button type="button" class="dismiss-button {"dismissed" if dismissed else ""}" data-app-id="{app_id}" data-activity-key="{activity_key}" data-task="dismiss_{task}" data-target-task="{task}" data-done="{str(dismissed).lower()}">{dismiss_label}</button>')
        action_buttons.append(f'<button type="button" class="identity-button" data-app-id="{app_id}" data-company="{esc(company)}" data-position="{esc(position)}">Edit company / vacancy</button>')
        unrelated_label = "Restore to tracker" if is_unrelated else "Mark unrelated"
        action_buttons.append(f'<button type="button" class="unrelated-button {"excluded" if is_unrelated else ""}" data-app-id="{app_id}" data-activity-key="{activity_key}" data-task="unrelated" data-done="{str(is_unrelated).lower()}">{unrelated_label}</button>')
        if is_unrelated:
            status_display = "Unrelated · Excluded"
        elif task and done:
            status_display = ("Action_Required" if task == "action" else "Interview") + " · Completed"
        elif task and dismissed:
            status_display = ("Action_Required" if task == "action" else "Interview") + " · Dismissed"
        elif task:
            status_display = ("Action_Required" if task == "action" else "Interview") + " · Pending"
        else:
            status_display = status
        action_control = '<div class="row-actions">' + "".join(action_buttons) + '</div>'
        rows.append(
            f'<tr data-app-id="{app_id}" data-activity-key="{activity_key}" data-base-status="{status_class(status)}" data-unrelated="{str(is_unrelated).lower()}" data-last-activity="{esc(last_activity)}">'
            f'<td data-value="{esc(company).lower()}"><strong>{esc(company)}</strong></td>'
            f'<td data-value="{esc(position).lower()}">{esc(position) or "—"}</td>'
            f'<td data-value="{esc(status_display).lower()}"><span class="status {status_class(status)}">{esc(status_display) or "—"}</span></td>'
            f'<td data-value="{confidence_sort}" data-type="number">{confidence_text}</td>'
            f'<td data-value="{esc(last_activity).lower()}">{esc(last_activity) or "—"}</td>'
            f'<td data-value="{esc(evidence_text).lower()}"><ul>{evidence_html}</ul></td>'
            f'<td data-value="{("open gmail" if gmail_link else "")}">{gmail_html}</td>'
            f'<td data-value="{esc(task_label if status in ("Action_Required", "Interview") else "—").lower()}">{action_control}</td>'
            '</tr>'
        )

    rows_html = "\n".join(rows)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    column_names = ["Company", "Position", "Status / progress", "Confidence", "Last Activity", "Evidence", "Gmail", "Action"]
    filter_cells = "".join(
        '<th><input class="column-filter" type="search" data-column="{}" placeholder="Filter {}" aria-label="Filter {}"></th>'.format(i, esc(name), esc(name))
        for i, name in enumerate(column_names)
    )
    header_cells = "".join(
        '<th><button type="button" class="sort-button" data-column="{}">{} <span aria-hidden="true">↕</span></button></th>'.format(i, esc(name))
        for i, name in enumerate(column_names)
    )

    report_html = f'''<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Job Application Tracker</title><style>
body {{ font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; margin:0; padding:24px; background:#f5f6f8; color:#202124; }}
.container {{ max-width:1800px; margin:auto; }} h1 {{ margin-bottom:5px; }} .subtitle,.muted {{ color:#666; }}
.cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:15px; margin-bottom:25px; }}
.card,.chart,.panel,.table-container {{ background:white; padding:18px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.07); margin-bottom:22px; }}
.card .number {{ font-size:30px; font-weight:700; }} .card .label {{ color:#666; margin-top:5px; }}
.controls {{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin:12px 0 16px; }} .controls input {{ flex:1 1 280px; }}
input {{ box-sizing:border-box; min-width:80px; width:100%; padding:8px; border:1px solid #dadce0; border-radius:6px; font:inherit; background:white; }}
button {{ cursor:pointer; font:inherit; }} .clear-button {{ padding:8px 12px; border:1px solid #dadce0; border-radius:6px; background:#fff; }}
.row-actions {{ display:flex; flex-direction:column; gap:6px; min-width:145px; }} .task-button,.dismiss-button,.unrelated-button,.identity-button {{ padding:7px 9px; border:1px solid #dadce0; border-radius:6px; background:#fff; font-size:12px; text-align:left; }} .task-button.done {{ background:#e6f4ea; border-color:#81c995; }} .dismiss-button.dismissed {{ background:#fef7e0; border-color:#f4c542; }} .unrelated-button.excluded {{ background:#fce8e6; border-color:#e6a09b; }} .completion-mini {{ font-size:12px; margin-left:6px; color:#666; }}
table {{ width:100%; border-collapse:collapse; }} th {{ text-align:left; background:#f1f3f4; padding:9px; vertical-align:top; }} td {{ padding:10px; border-bottom:1px solid #eee; vertical-align:top; }}
thead tr:first-child th {{ white-space:nowrap; }} .sort-button {{ border:0; background:transparent; font-weight:700; padding:2px 0; white-space:nowrap; }}
.filter-row th {{ background:#fafafa; min-width:110px; }} .filter-row input {{ font-size:12px; padding:7px; }}
.status {{ display:inline-block; padding:4px 8px; border-radius:20px; font-size:12px; font-weight:600; white-space:nowrap; }}
.Applied {{ background:#e8f0fe; }} .Action_Required {{ background:#fef7e0; }} .Rejected {{ background:#fce8e6; }} .Interview {{ background:#e6f4ea; }} .Job_Offer {{ background:#e8eaed; }}
ul {{ margin:0; padding-left:18px; }} a {{ text-decoration:none; }} a:hover {{ text-decoration:underline; }}
.company-group {{ background:#fff; border:1px solid #e5e7eb; border-radius:8px; margin:8px 0; padding:12px 14px; }}
.company-group summary {{ cursor:pointer; display:flex; flex-wrap:wrap; justify-content:space-between; gap:8px; }} .company-group ul {{ margin-top:10px; }} .company-group li {{ margin:8px 0; }}
.company-count {{ color:#666; font-size:13px; font-weight:400; }} #visible-count {{ color:#555; font-size:13px; }} .footer {{ margin-top:25px; color:#777; font-size:13px; }}
#save-message {{ font-size:13px; }} .search-button {{ padding:8px 14px; border:1px solid #1a73e8; border-radius:6px; background:#1a73e8; color:#fff; font-weight:600; }}
.visually-hidden {{ position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; border:0; }}
.company-group[data-priority="0"] {{ border-left:5px solid #d97706; }}
.company-group[data-priority="1"] {{ border-left:5px solid #2563eb; }}
@media(max-width:700px) {{ body {{ padding:12px; }} .chart,.panel,.table-container {{ padding:12px; }} .controls {{ align-items:stretch; }} .controls button {{ flex:1 1 auto; }} }}
</style></head><body><div class="container">
<h1>Job Application Tracker</h1><div class="subtitle">Generated locally with Ollama / {esc(OLLAMA_MODEL)}<br>Generated: {esc(generated_at)}</div>
<div class="cards">
<div class="card"><div class="number" id="count-total-submissions">{total_submissions}</div><div class="label">Total Submissions</div></div>
<div class="card"><div class="number" id="count-pending">{main_counts['Pending']}</div><div class="label">Pending</div></div>
<div class="card"><div class="number" id="count-interview">{main_counts['Interview']}</div><div class="label">Interview</div></div>
<div class="card"><div class="number" id="count-applied">{main_counts['Applied']}</div><div class="label">Applied</div></div>
<div class="card"><div class="number" id="count-rejected">{main_counts['Rejected']}</div><div class="label">Rejected</div></div>
<div class="card"><div class="number" id="count-offer">{main_counts['Offer']}</div><div class="label">Offer</div></div></div>
<div class="chart">{status_chart_html}</div><div class="chart">{timeline_chart_html}</div>
<section class="panel"><h2>Vacancies by company</h2><p class="muted">Grouped overview of consolidated vacancies. Expand or collapse each company.</p>{company_groups_html}</section>
<div class="table-container"><h2>All applications</h2><p class="muted">Mark an action or interview as completed. A new email in the same application automatically makes it pending again.</p>
<div class="controls" role="search">
<label class="visually-hidden" for="global-search">Search all applications</label>
<input id="global-search" type="search" placeholder="Search company, vacancy, status, evidence…" aria-label="Search all columns">
<button type="button" class="search-button" id="search-button">Search</button>
<button type="button" class="clear-button" id="clear-filters">Clear search and filters</button>
<span id="visible-count" aria-live="polite"></span><span id="save-message" aria-live="polite"></span></div>
<table id="applications-table"><thead><tr>{header_cells}</tr><tr class="filter-row">{filter_cells}</tr></thead><tbody>{rows_html}</tbody></table></div>
<div class="footer">Source: Google Gmail → Google Sheet → local Python → Ollama<br>Search, column filters and sorting run entirely in this HTML file and work without the local server. Saving completion states or company/vacancy corrections still requires opening the report with --serve. Data: data/user_actions.json and data/manual_overrides.json.</div></div>
<script>
(function() {{
 const table=document.getElementById('applications-table'),tbody=table.querySelector('tbody'),rows=Array.from(tbody.querySelectorAll('tr'));
 const globalSearch=document.getElementById('global-search'),columnFilters=Array.from(document.querySelectorAll('.column-filter'));
 const countLabel=document.getElementById('visible-count'),saveMessage=document.getElementById('save-message');let sortColumn=-1,sortDirection=1;
 function applyFilters() {{
  const term=globalSearch.value.trim().toLocaleLowerCase();
  const filters=columnFilters.map(input=>input.value.trim().toLocaleLowerCase());
  let visible=0;
  rows.forEach(row=>{{
   const cells=Array.from(row.cells);
   const globalMatch=!term||row.textContent.toLocaleLowerCase().includes(term);
   const columnMatch=filters.every((filter,index)=>!filter||(cells[index]&&cells[index].textContent.toLocaleLowerCase().includes(filter)));
   row.hidden=!(globalMatch&&columnMatch);
   if(!row.hidden)visible++;
  }});
  countLabel.textContent='Showing '+visible+' of '+rows.length+' applications';
 }}
 function clearFilters() {{
  globalSearch.value='';
  columnFilters.forEach(input=>input.value='');
  applyFilters();
  globalSearch.focus();
 }}
 function sortTable(column) {{
  if(column===sortColumn)sortDirection*=-1;else{{sortColumn=column;sortDirection=1;}}
  const sorted=rows.slice().sort((a,b)=>{{
   const av=(a.cells[column]?.dataset.value||a.cells[column]?.textContent||'').trim();
   const bv=(b.cells[column]?.dataset.value||b.cells[column]?.textContent||'').trim();
   const an=Number(av),bn=Number(bv);
   const numeric=a.cells[column]?.dataset.type==='number'&&Number.isFinite(an)&&Number.isFinite(bn);
   const comparison=numeric?(an-bn):av.localeCompare(bv,undefined,{{numeric:true,sensitivity:'base'}});
   return comparison*sortDirection;
  }});
  sorted.forEach(row=>tbody.appendChild(row));
  applyFilters();
 }}
 globalSearch.addEventListener('input',applyFilters);
 globalSearch.addEventListener('keydown',event=>{{if(event.key==='Enter'){{event.preventDefault();applyFilters();}}}});
 document.getElementById('search-button').addEventListener('click',applyFilters);
 document.getElementById('clear-filters').addEventListener('click',clearFilters);
 columnFilters.forEach(input=>input.addEventListener('input',applyFilters));
 document.querySelectorAll('.sort-button').forEach(button=>button.addEventListener('click',()=>sortTable(Number(button.dataset.column))));
 function displayPriority(item) {{
  if(item.dataset.unrelated==='true')return 5;
  const status=item.dataset.status||'',task=item.dataset.task||'';
  if(task==='action'||task==='interview')return (item.dataset.done==='true'||item.dataset.dismissed==='true')?4:0;
  if(status==='Applied'||status==='Job_Offer')return 1;
  if(status==='Rejected')return 2;
  return 3;
 }}
 function refreshCompanyGroupOrder() {{
  const panel=Array.from(document.querySelectorAll('.panel')).find(el=>el.querySelector('.company-group'));
  if(!panel)return;
  const groups=Array.from(panel.querySelectorAll('.company-group'));
  groups.forEach(group=>{{
   const list=group.querySelector('ul');if(!list)return;
   const items=Array.from(list.querySelectorAll('li'));
   items.forEach(item=>item.dataset.priority=String(displayPriority(item)));
   items.sort((a,b)=>Number(a.dataset.priority)-Number(b.dataset.priority));
   items.forEach(item=>list.appendChild(item));
   const priorities=items.map(item=>Number(item.dataset.priority));
   const priority=priorities.length?Math.min(...priorities):3;
   group.dataset.priority=String(priority);
   const pending=items.filter(item=>Number(item.dataset.priority)===0).length;
   const count=group.querySelector('.company-count');
   if(count)count.textContent=items.length+' vacancy/vacancies · '+pending+' pending';
  }});
  groups.sort((a,b)=>Number(a.dataset.priority)-Number(b.dataset.priority)||a.querySelector('summary strong').textContent.localeCompare(b.querySelector('summary strong').textContent,undefined,{{sensitivity:'base'}}));
  groups.forEach(group=>panel.appendChild(group));
 }}
 refreshCompanyGroupOrder();
 async function editIdentity(button) {{
  const company=prompt('Canonical company name:',button.dataset.company||'');if(company===null)return;
  const position=prompt('Canonical vacancy / job title:',button.dataset.position||'');if(position===null)return;
  button.disabled=true;saveMessage.textContent='Saving identity correction…';
  try{{const response=await fetch('/api/identity',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{app_id:button.dataset.appId,company:company.trim(),position:position.trim()}})}});
   const result=await response.json();if(!response.ok)throw new Error(result.error||'Unable to save identity');
   saveMessage.textContent='Correction saved. Regrouping applications…';location.reload();
  }}catch(error){{saveMessage.textContent='Could not save correction: '+error.message;button.disabled=false;}}
 }}
 function refreshDashboard() {{
  const allRows=Array.from(tbody.querySelectorAll('tr'));
  const categories=['Pending','Interview','Applied','Rejected','Offer'];
  const counts={{Pending:0,Interview:0,Applied:0,Rejected:0,Offer:0}};
  let totalSubmissions=0;
  const timelineCounts={{}};
  allRows.forEach(row=>{{
   const unrelated=row.dataset.unrelated==='true',base=row.dataset.baseStatus||'',taskButton=row.querySelector('.task-button'),dismissButton=row.querySelector('.dismiss-button');
   if(unrelated)return;
   totalSubmissions++;
   const activityDate=(row.dataset.lastActivity||'').slice(0,10);
   if(/^\\d{{4}}-\\d{{2}}-\\d{{2}}$/.test(activityDate))timelineCounts[activityDate]=(timelineCounts[activityDate]||0)+1;
   if(base==='Applied')counts.Applied++;
   else if(base==='Rejected')counts.Rejected++;
   else if(base==='Job_Offer')counts.Offer++;
   else if(base==='Action_Required'){{
    if(!(taskButton&&taskButton.dataset.done==='true')&&!(dismissButton&&dismissButton.dataset.done==='true'))counts.Pending++;
   }} else if(base==='Interview'){{
    if(!(dismissButton&&dismissButton.dataset.done==='true'))counts.Interview++;
   }}
  }});
  const statusChart=document.getElementById('status-chart');
  if(window.Plotly&&statusChart){{const values=categories.map(label=>counts[label]||0);Plotly.restyle(statusChart,{{x:[categories],y:[values],text:[values]}},[0]);}}
  const timelineChart=document.querySelector('.chart .plotly-graph-div:not(#status-chart)');
  if(window.Plotly&&timelineChart){{const dates=Object.keys(timelineCounts).sort(),values=dates.map(date=>timelineCounts[date]);Plotly.restyle(timelineChart,{{x:[dates],y:[values]}},[0]);}}
  const cardValues={{'count-total-submissions':totalSubmissions,'count-pending':counts.Pending,'count-interview':counts.Interview,'count-applied':counts.Applied,'count-rejected':counts.Rejected,'count-offer':counts.Offer}};
  Object.entries(cardValues).forEach(([id,value])=>{{const el=document.getElementById(id);if(el)el.textContent=value;}});
  allRows.forEach(row=>{{
   const cell=row.cells[2];if(!cell)return;const base=row.dataset.baseStatus||'',unrelated=row.dataset.unrelated==='true',taskButton=row.querySelector('.task-button'),dismissButton=row.querySelector('.dismiss-button');
   let label=base;
   if(unrelated)label='Unrelated · Excluded';
   else if(taskButton){{const prefix=base==='Action_Required'?'Action_Required':'Interview';label=prefix+' · '+(taskButton.dataset.done==='true'?'Completed':dismissButton&&dismissButton.dataset.done==='true'?'Dismissed':'Pending');}}
   cell.innerHTML='<span class="status '+(unrelated?'Rejected':base)+'">'+label+'</span>';cell.dataset.value=label.toLocaleLowerCase();
  }});
 }}
 async function sendAction(button,next) {{
  if(location.protocol!=='http:'&&location.protocol!=='https:'){{saveMessage.textContent='Open this report using ./actualizar_job_tracker.sh --serve to save changes.';return false;}}
  button.disabled=true;saveMessage.textContent='Saving…';
  try{{const response=await fetch('/api/action',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{app_id:button.dataset.appId,activity_key:button.dataset.activityKey,task:button.dataset.task,done:next}})}});
   const result=await response.json();if(!response.ok)throw new Error(result.error||'Unable to save');
   button.dataset.done=String(next);
   if(button.classList.contains('task-button')){{button.classList.toggle('done',next);button.textContent=next?(button.dataset.task==='action'?'Action completed':'Interview completed'):(button.dataset.task==='action'?'Mark action done':'Mark interview done');
    if(next){{const row=button.closest('tr'),dismiss=row.querySelector('.dismiss-button');if(dismiss){{dismiss.dataset.done='false';dismiss.classList.remove('dismissed');dismiss.textContent='Dismiss pending';}}}}
   }} else if(button.classList.contains('dismiss-button')){{button.classList.toggle('dismissed',next);button.textContent=next?'Restore pending':'Dismiss pending';if(next){{const row=button.closest('tr'),taskButton=row.querySelector('.task-button');if(taskButton){{taskButton.dataset.done='false';taskButton.classList.remove('done');taskButton.textContent=taskButton.dataset.task==='action'?'Mark action done':'Mark interview done';}}}}}}
   else if(button.classList.contains('unrelated-button')){{const row=button.closest('tr');row.dataset.unrelated=String(next);button.classList.toggle('excluded',next);button.textContent=next?'Restore to tracker':'Mark unrelated';}}
   const row=button.closest('tr');
   const groupItem=Array.from(document.querySelectorAll('.company-group li')).find(item=>item.dataset.appId===row.dataset.appId);
   if(groupItem){{
    groupItem.dataset.unrelated=row.dataset.unrelated||'false';
    const taskButton=row.querySelector('.task-button'),dismissButton=row.querySelector('.dismiss-button');
    groupItem.dataset.done=taskButton?taskButton.dataset.done:'false';
    groupItem.dataset.dismissed=dismissButton?dismissButton.dataset.done:'false';
   }}
   refreshDashboard();refreshCompanyGroupOrder();saveMessage.textContent=button.dataset.task==='unrelated'?(next?'Saved: excluded as unrelated.':'Saved: restored to tracker.'):button.dataset.task.startsWith('dismiss_')?(next?'Saved: follow-up dismissed.':'Saved: follow-up pending again.'):(next?'Saved: marked completed.':'Saved: marked pending.');applyFilters();return true;
  }}catch(error){{saveMessage.textContent='Could not save: '+error.message;return false;}}finally{{button.disabled=false;}}
 }}
 document.querySelectorAll('.task-button').forEach(button=>button.addEventListener('click',()=>sendAction(button,button.dataset.done!=='true')));
 document.querySelectorAll('.dismiss-button').forEach(button=>button.addEventListener('click',()=>sendAction(button,button.dataset.done!=='true')));
 document.querySelectorAll('.unrelated-button').forEach(button=>button.addEventListener('click',()=>sendAction(button,button.dataset.done!=='true')));
 document.querySelectorAll('.identity-button').forEach(button=>button.addEventListener('click',()=>editIdentity(button)));
 refreshDashboard();applyFilters();
}})();
</script></body></html>'''

    with open(report_file, "w", encoding="utf-8") as f:
        f.write(report_html)
    log("\n📊 Reporte HTML generado (se sobrescribe en cada ejecución):")
    log(f"   {report_file}")
    return report_file


def start_report_server(port=8765):
    # Serve the report and persist completion changes through a local-only JSON API.
    os.chdir(REPORT_DIR)

    class TrackerHandler(SimpleHTTPRequestHandler):
        def log_message(self, fmt, *args):
            log("[UI] " + (fmt % args))

        def do_GET(self):
            if self.path in ("/", "/index.html"):
                self.path = "/reporte.html"
            if self.path == "/api/health":
                body = json.dumps({"ok": True}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            return super().do_GET()

        def do_POST(self):
            if self.path not in ("/api/action", "/api/identity"):
                self.send_error(404)
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if length <= 0 or length > 10000:
                    raise ValueError("Invalid request size")
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                app_id = clean_text(payload.get("app_id"))
                current_apps = load_existing_classification()
                current_app = next((a for a in current_apps if application_id(a) == app_id), None)
                if current_app is None:
                    self._json_response(404, {"error": "Application not found. Regenerate the report."})
                    return
                if self.path == "/api/identity":
                    company = clean_text(payload.get("company"))
                    position = clean_text(payload.get("position"))
                    if not company or not position:
                        raise ValueError("Company and vacancy cannot be empty.")
                    if len(company) > 200 or len(position) > 250:
                        raise ValueError("Company or vacancy is too long.")
                    overrides = set_manual_override(current_app, company, position)
                    apply_manual_overrides(current_apps, overrides)
                    current_apps = deduplicate_applications(current_apps)
                    save_classification(current_apps)
                    generate_html_report(current_apps)
                    self._json_response(200, {"ok": True, "company": company, "position": position, "applications": len(current_apps)})
                    return
                activity_key = clean_text(payload.get("activity_key"))
                task = clean_text(payload.get("task"))
                done = bool(payload.get("done"))
                if not re.fullmatch(r"[a-f0-9]{20}", app_id) or not re.fullmatch(r"[a-f0-9]{24}", activity_key):
                    raise ValueError("Invalid application or activity identifier")
                if task not in ("action", "interview", "dismiss_action", "dismiss_interview", "unrelated"):
                    raise ValueError("Invalid task")

                expected_status = None
                if task in ("action", "dismiss_action"):
                    expected_status = "Action_Required"
                elif task in ("interview", "dismiss_interview"):
                    expected_status = "Interview"
                if expected_status and clean_text(current_app.get("status")) != expected_status:
                    self._json_response(409, {"error": "The application status changed. Refresh the report."})
                    return
                current_activity_key = application_activity_key(current_app)

                # Completing a task requires the exact activity shown in the report.
                # If a newer email arrived, do not allow the old completion to hide it.
                # Dismissal is different: the user explicitly wants to clear the current
                # pending follow-up, so bind the dismissal to the server's current activity.
                # Marking a vacancy unrelated is a manual classification decision and must
                # remain available even when the report's activity key is stale.
                if task in ("action", "interview") and current_activity_key != activity_key:
                    self._json_response(409, {
                        "error": "A newer email exists for this application. The task remains pending; refresh the report before marking it completed.",
                        "current_activity_key": current_activity_key,
                    })
                    return

                state = load_user_actions()
                app_state = state.setdefault(app_id, {})
                now = datetime.now().isoformat()
                if task == "unrelated":
                    if done:
                        app_state[task] = {"done": True, "marked_at": now, "reason": "user_marked_unrelated"}
                    else:
                        app_state.pop(task, None)
                elif task.startswith("dismiss_"):
                    target_task = task[len("dismiss_"):]
                    if done:
                        app_state[task] = {"done": True, "activity_key": current_activity_key, "dismissed_at": now}
                        app_state.pop(target_task, None)
                    else:
                        app_state.pop(task, None)
                else:
                    if done:
                        app_state[task] = {"done": True, "activity_key": current_activity_key, "completed_at": now}
                        app_state.pop(f"dismiss_{task}", None)
                    else:
                        app_state.pop(task, None)
                if not app_state:
                    state.pop(app_id, None)
                save_user_actions(state)
                self._json_response(200, {"ok": True, "done": done, "task": task})
            except (ValueError, TypeError, json.JSONDecodeError) as exc:
                self._json_response(400, {"error": str(exc)})
            except OSError as exc:
                self._json_response(500, {"error": "Could not persist action: " + str(exc)})

        def _json_response(self, status, payload):
            body = json.dumps(payload).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

    server = ThreadingHTTPServer(("127.0.0.1", int(port)), TrackerHandler)
    url = f"http://127.0.0.1:{int(port)}/"
    log("\n🌐 Interfaz interactiva disponible solo en esta computadora:")
    log(f"   {url}")
    log("   Mantén esta terminal abierta mientras uses el reporte. Ctrl+C para detener.")
    try:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
        server.serve_forever()
    except KeyboardInterrupt:
        log("\nServidor del reporte detenido.")
    finally:
        server.server_close()

# ============================================================
# RECLASSIFY EXISTING APPLICATIONS
# ============================================================

def reclassify_existing_applications(existing_applications, threads):
    """Re-run Ollama only for source threads referenced by the saved consolidation."""
    threads_by_id = {
        clean_text(thread.get("thread_id")): thread
        for thread in threads
        if clean_text(thread.get("thread_id"))
    }

    referenced_ids = set()
    for app in existing_applications:
        ids = app.get("thread_ids", []) or []
        if isinstance(ids, str):
            ids = [ids]
        ids = list(ids) + [app.get("thread_id", "")]
        referenced_ids.update(clean_text(value) for value in ids if clean_text(value))

    selected_ids = referenced_ids.intersection(threads_by_id)
    selected_threads = [
        thread for thread in threads
        if clean_text(thread.get("thread_id")) in selected_ids
    ]

    missing_ids = referenced_ids - selected_ids
    log("\n🔄 Reclasificación de aplicaciones existentes")
    log(f"   Aplicaciones consolidadas: {len(existing_applications)}")
    log(f"   Hilos fuente encontrados: {len(selected_threads)}")
    if missing_ids:
        log(f"   ⚠️ Hilos referenciados que no están en el Excel/JSON actual: {len(missing_ids)}")
        log("      Las clasificaciones que dependan de esos hilos se conservarán para evitar pérdida de datos.")

    if not selected_threads:
        raise RuntimeError(
            "No se encontraron hilos originales asociados a las aplicaciones guardadas. "
            "Verifica que job_emails.xlsx incluya los correos originales y vuelve a ejecutar."
        )

    ai_response = call_ollama(selected_threads)
    updated = ai_response.get("applications", [])

    # Remove old records only when all their source threads are available for
    # reclassification. Preserve records with missing source context.
    retained = []
    removed = 0
    for app in existing_applications:
        ids = app.get("thread_ids", []) or []
        if isinstance(ids, str):
            ids = [ids]
        ids = {clean_text(value) for value in list(ids) + [app.get("thread_id", "")] if clean_text(value)}
        if ids and ids.issubset(selected_ids):
            removed += 1
        else:
            retained.append(app)

    applications = apply_manual_overrides(retained + updated)
    applications = deduplicate_applications(applications)
    log(f"   Clasificaciones anteriores reemplazadas: {removed}")
    log(f"   Resultados nuevos de Ollama: {len(updated)}")
    log(f"   Vacantes después de consolidar: {len(applications)}")
    return applications, selected_threads


def print_summary(applications, report_file, threads_processed=0):
    counts = get_main_status_counts(applications, load_user_actions())
    print("\n==========================================")
    print(" RESUMEN")
    print("==========================================")
    total_submissions = get_total_submissions_count(applications, load_user_actions())
    print(f"{'Total Submissions':18} : {total_submissions}")
    for status in STATUS_CHART_CATEGORIES:
        print(f"{status:18} : {counts[status]}")
    print(f"\n📊 Reporte: {report_file}")
    print(f"\n📨 Threads procesados en esta ejecución: {threads_processed}")
    print("\n✓ Proceso terminado.\n")


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Incremental Gmail job application tracker using local Ollama."
    )
    parser.add_argument(
        "--reclassify-existing",
        action="store_true",
        help="Reclassify saved applications with Ollama using their original email threads.",
    )
    parser.add_argument(
        "--serve",
        action="store_true",
        help="After processing, launch the local interactive UI and persist completed actions/interviews.",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Generate the HTML report without opening a browser (for scheduled/headless runs).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="Local UI port (default: 8765).",
    )
    args = parser.parse_args()

    print("")
    print(
        "=========================================="
    )
    print(
        " JOB APPLICATION AI TRACKER"
    )
    print(
        "=========================================="
    )
    print("")
    print(
        f"Model: {OLLAMA_MODEL}"
    )
    print("")

    os.makedirs(
        DATA_DIR,
        exist_ok=True,
    )

    os.makedirs(
        REPORT_DIR,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # 1. Load Excel
    # --------------------------------------------------------

    df = load_excel()

    # --------------------------------------------------------
    # 2. Build RAW JSON
    # --------------------------------------------------------

    messages = (
        dataframe_to_messages(
            df
        )
    )

    threads = (
        group_messages_by_thread(
            messages
        )
    )

    save_raw_json(
        messages,
        threads,
    )

    # --------------------------------------------------------
    # 3. Load processing state
    # --------------------------------------------------------

    state = (
        load_processing_state()
    )

    # --------------------------------------------------------
    # 4. Determine what needs processing
    # --------------------------------------------------------

    pending_threads = (
        find_pending_threads(
            threads,
            state,
        )
    )

    log("")
    log(
        "🔎 Análisis incremental"
    )

    log(
        f"   Threads totales: "
        f"{len(threads)}"
    )

    log(
        f"   Threads pendientes: "
        f"{len(pending_threads)}"
    )

    if pending_threads:

        log(
            "   Motivo: thread nuevo "
            "o contiene Message_ID nuevo."
        )

    else:

        log(
            "   ✓ No hay información nueva "
            "para procesar."
        )

    # --------------------------------------------------------
    # 5. Load previous classification
    # --------------------------------------------------------

    existing_applications = (
        load_existing_classification()
    )

    log(
        f"   Aplicaciones existentes: "
        f"{len(existing_applications)}"
    )

    # --------------------------------------------------------
    # Optional maintenance mode: reclassify the saved consolidation.
    # This deliberately does NOT change processing_state.json.
    # --------------------------------------------------------
    if args.reclassify_existing:
        if not existing_applications:
            raise RuntimeError(
                "No hay aplicaciones guardadas en classified_jobs.json para reclasificar."
            )
        applications, reclassified_threads = reclassify_existing_applications(
            existing_applications, threads
        )
        save_classification(applications)
        report_file = generate_html_report(applications)
        print_summary(applications, report_file, len(reclassified_threads))
        if args.serve:
            start_report_server(args.port)
        elif not args.no_open:
            opened = webbrowser.open(REPORT_FILE.as_uri())
            if not opened:
                log("⚠️ No se pudo abrir automáticamente el navegador. Abre este archivo manualmente:")
                log(f"   {REPORT_FILE}")
        else:
            log("ℹ️ --no-open activo: el reporte se generó sin abrir el navegador.")
        return

    # --------------------------------------------------------
    # 6. Process ONLY pending threads
    # --------------------------------------------------------

    updated_applications = []

    if pending_threads:

        try:

            ai_response = (
                call_ollama(
                    pending_threads
                )
            )

        except requests.exceptions.RequestException as e:

            raise RuntimeError(
                "No fue posible conectarse "
                "a Ollama.\n"
                f"URL: {OLLAMA_URL}\n"
                f"Error: {e}"
            )

        updated_applications = (
            ai_response[
                "applications"
            ]
        )

    else:

        log("")
        log(
            "⏭️ Ollama no será ejecutado."
        )

    # --------------------------------------------------------
    # 7. Merge old + updated
    # --------------------------------------------------------

    applications = (
        merge_classification(
            existing_applications,
            updated_applications,
        )
    )

    applications = apply_manual_overrides(applications)
    before_cleanup = len(applications)
    applications = deduplicate_applications(applications)
    log(f"   Depuración local: {before_cleanup} registros → {len(applications)} vacantes únicas.")
    save_classification(applications)

    if not applications:

        log("")
        log(
            "⚠️ No existen aplicaciones "
            "clasificadas todavía."
        )

        # This can happen if the first execution
        # contains only unrelated threads.

    # --------------------------------------------------------
    # 8. Update processing state
    #
    # IMPORTANT:
    # Only mark pending threads as processed AFTER
    # Ollama successfully returned valid results.
    # --------------------------------------------------------

    if pending_threads:

        update_processing_state(
            state,
            pending_threads,
        )

        save_processing_state(
            state
        )

    # --------------------------------------------------------
    # 9. Save classification
    # --------------------------------------------------------

    save_classification(
        applications
    )

    # --------------------------------------------------------
    # 10. Generate HTML
    # --------------------------------------------------------

    report_file = (
        generate_html_report(
            applications
        )
    )

    # --------------------------------------------------------
    # 11. Summary
    # --------------------------------------------------------
    print_summary(applications, report_file, len(pending_threads))
    if args.serve:
        start_report_server(args.port)
    elif not args.no_open:
        opened = webbrowser.open(REPORT_FILE.as_uri())
        if not opened:
            log("⚠️ No se pudo abrir automáticamente el navegador. Abre este archivo manualmente:")
            log(f"   {REPORT_FILE}")
    else:
        log("ℹ️ --no-open activo: el reporte se generó sin abrir el navegador.")


if __name__ == "__main__":

    main()
