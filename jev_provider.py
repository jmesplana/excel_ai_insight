"""
Jev (typesafe.ai) System One provider.

Where the OpenAI/Azure path asks a model for free text and takes whatever prose
comes back, Jev answers *typed* questions: every answer is constrained to a list
the caller supplied, and comes with a probability per option plus a confidence
statistic. The adapter validates the declared schema before accepting any response.

Confidence (0-1) summarises how concentrated that probability distribution is,
not how likely the answer is to be correct: the API documents it as a
convenience metric rather than a calibrated probability. Choice and Score
report it; Noul does not.

Three question types are supported, mirroring the API's primitives:

  * "choice" -- pick one label from a controlled list (up to 255 options).
  * "score"  -- rate against 2-10 ordered, described levels. The API returns a
                possibly fractional score; we round to the nearest level label.
  * "noul"   -- a yes/no question answered as the probability of "yes".

All questions for one spreadsheet row are sent in a single request: the API
evaluates them in parallel against one shared state, which is both cheaper and
faster than one call per cell.

Configuration is resolved per request with the same precedence as the LLM path:
values the browser sent win over environment variables, so an individual can use
their own key without the server needing one.
"""

import math
import time
import threading
import json
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone
import os
import re
import unicodedata

import requests

TYPESAFE_URL = "https://api.typesafe.ai/v1/systemone"

# "jev-latest" tracks the newest release. Pin an exact version (e.g.
# "jev-1.13.0") in settings if you tune confidence thresholds against it.
DEFAULT_JEV_MODEL = "jev-latest"

# API limits, enforced here so a bad configuration fails with a clear message
# rather than a 400 from the API halfway through a file.
MAX_CHOICE_OPTIONS = 255
MIN_SCORE_LEVELS = 2
MAX_SCORE_LEVELS = 10

QUESTION_TYPES = ("choice", "score", "noul")

# Per-row request timeout, in seconds. Most queries complete in ~100ms; this is
# a generous ceiling that still fails a hung connection before the HTTP layer.
REQUEST_TIMEOUT = 60


class JevConfigError(ValueError):
    """Raised when the resolved Jev configuration is unusable."""


class JevError(RuntimeError):
    """Raised when the Jev API rejects a request or returns something odd."""


def _clean(value):
    """Normalize a request/env value to a non-empty string or None."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def resolve_config(data=None):
    """
    Resolve the effective Jev configuration for one request.

    Args:
        data: the parsed JSON request body (or None). Recognized keys are
            jevApiKey and jevModel.

    Returns:
        A dict with keys: api_key, model.

    Raises:
        JevConfigError: if no API key is available.
    """
    data = data or {}

    api_key = (
        _clean(data.get('jevApiKey'))
        or _clean(os.environ.get('TYPESAFE_API_KEY'))
    )
    if not api_key:
        raise JevConfigError(
            "Missing Jev API key. Add it in Settings or set TYPESAFE_API_KEY."
        )

    model = (
        _clean(data.get('jevModel'))
        or _clean(os.environ.get('TYPESAFE_MODEL'))
        or DEFAULT_JEV_MODEL
    )

    return {"api_key": api_key, "model": model}


def slugify(label, taken=()):
    """
    Turn a human label into an ASCII option key, unique against `taken`.

    Jev keys the criteria dict and echoes the winning key back in `choice`, so
    each key must be stable and unique. Accented and non-Latin labels (the
    French vocabularies this app is often used with) collapse to ASCII here,
    which is why collisions are possible and must be suffixed.
    """
    text = unicodedata.normalize("NFKD", str(label))
    text = text.encode("ascii", "ignore").decode()
    text = re.sub(r"[^\w]+", "_", text.lower()).strip("_")
    base = text or "option"
    slug = base
    counter = 2
    while slug in taken:
        slug = f"{base}_{counter}"
        counter += 1
    return slug


def build_question(spec):
    """
    Build one API question plus the decoder needed to read its answer back.

    Args:
        spec: {"type": "choice"|"score"|"noul", "instructions": str,
               "options": [str, ...]}  -- options are the labels for choice,
               the ordered level descriptions for score, and unused for noul.

    Returns:
        (question, decoder) where decoder is a dict describing how to turn the
        raw answer into a cell value.

    Raises:
        JevConfigError: if the spec cannot be expressed as a Jev question.
    """
    qtype = (spec.get("type") or "choice").lower()
    if qtype not in QUESTION_TYPES:
        raise JevConfigError(
            f"Unknown question type '{qtype}'. Expected choice, score or noul."
        )

    instructions = spec.get("instructions")
    if not isinstance(instructions, (str, dict, list)) or not instructions or (isinstance(instructions, str) and not instructions.strip()):
        raise JevConfigError("Each Jev question needs instructions.")
    options = spec.get("options") or []
    if not isinstance(options, list):
        raise JevConfigError("Options must be an array.")
    labels = [o.get("label") if isinstance(o, dict) else o for o in options]
    if any(not isinstance(label, str) or not label.strip() for label in labels):
        raise JevConfigError("Every option needs a non-empty label.")
    if len(set(labels)) != len(labels):
        raise JevConfigError("Option labels must be unique.")
    descriptions = [o.get("description", o["label"]) if isinstance(o, dict) else o for o in options]

    if any(d is not None and not isinstance(d, (str, dict, list)) for d in descriptions):
        raise JevConfigError("Option descriptions must be text, objects, arrays or null.")

    if qtype == "choice":
        if len(options) < 2:
            raise JevConfigError(
                "A Choice question needs at least 2 options."
            )
        if len(options) > MAX_CHOICE_OPTIONS:
            raise JevConfigError(
                f"A Choice question allows at most {MAX_CHOICE_OPTIONS} options "
                f"({len(options)} given)."
            )
        criteria = {}
        label_map = {}
        for label, description in zip(labels, descriptions):
            slug = slugify(label, criteria)
            # Structured descriptions can explain each label and its boundaries.
            criteria[slug] = description
            label_map[slug] = label
        return (
            {"type": "choice", "instructions": instructions, "criteria": criteria},
            {"type": "choice", "labels": label_map},
        )

    if qtype == "score":
        if not MIN_SCORE_LEVELS <= len(options) <= MAX_SCORE_LEVELS:
            raise JevConfigError(
                f"A Score question needs between {MIN_SCORE_LEVELS} and "
                f"{MAX_SCORE_LEVELS} levels ({len(options)} given)."
            )
        return (
            {"type": "score", "instructions": instructions, "criteria": descriptions},
            {"type": "score", "labels": labels},
        )

    if len(labels) not in (0, 2):
        raise JevConfigError("Noul accepts either no labels or exactly two: yes then no.")
    question = {"type": "noul", "instructions": instructions}
    if "criteria" in spec:
        criteria = spec["criteria"]
        if not isinstance(criteria, dict) or set(criteria) != {"true", "false"}:
            raise JevConfigError("Noul criteria must contain true and false descriptions.")
        question["criteria"] = criteria
    return question, {"type": "noul", "yes": labels[0] if labels else "Yes",
                      "no": labels[1] if labels else "No"}


def number(value, low, high, field):
    if (isinstance(value, bool) or not isinstance(value, (int, float))
            or not math.isfinite(value) or not low <= value <= high):
        raise JevError(f"Invalid {field} in Jev response.")
    return float(value)


def decode_answer(answer, decoder):
    """
    Turn one raw API answer into (value, confidence, detail).

    `detail` is the raw score for a Score question and the yes-probability for
    a Noul, or None for a Choice -- the optional auditing column.

    `confidence` is None for a Noul: the API reports it only for Choice and
    Score. The caller writes an empty cell in that case.
    """
    if not isinstance(answer, dict):
        raise JevError("Malformed answer from Jev.")

    kind = decoder["type"]
    if answer.get("type") != kind:
        raise JevError("Jev answer type does not match the question.")
    if kind == "noul":
        probability = number(answer.get("noul"), 0, 1, "Noul probability")
        return decoder["yes"] if probability >= 0.5 else decoder["no"], None, probability
    confidence = number(answer.get("confidence"), 0, 1, "confidence")
    expected = set(decoder["labels"]) if kind == "choice" else {str(i) for i in range(len(decoder["labels"]))}
    probabilities = answer.get("probabilities")
    if not isinstance(probabilities, dict) or set(probabilities) != expected:
        raise JevError("Jev probability keys do not match the options.")
    total = sum(number(v, 0, 1, "probability") for v in probabilities.values())
    if abs(total - 1) > 0.01:
        raise JevError("Jev probabilities do not sum to one.")
    if kind == "choice":
        key = answer.get("choice")
        if key not in expected:
            raise JevError("Jev returned an unknown choice.")
        return decoder["labels"][key], confidence, None
    raw = number(answer.get("score"), 0, len(decoder["labels"]) - 1, "score")
    return decoder["labels"][math.floor(raw + 0.5)], confidence, raw


# Process-wide pacing shared by workers. Multiple server instances still need
# an account-wide queue for guaranteed quota enforcement.
_pace_lock = threading.Lock()
_next_request = 0.0


def ask(state, questions, *, api_key, model=DEFAULT_JEV_MODEL,
        timeout=10, session=None, details=False, deadline=None,
        requests_per_minute=600, tokens_per_second=100000, max_attempts=3):
    """Bounded transient retries; retain model and usage when details=True."""
    global _next_request
    deadline = deadline or time.monotonic() + 40
    payload = {"model": model, "state": state, "questions": questions}
    # Conservative byte-based estimate for pacing; actual usage is retained.
    estimated_tokens = len(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
    interval = max(60 / requests_per_minute, estimated_tokens / tokens_per_second)
    poster = session.post if session is not None else requests.post
    for attempt in range(max_attempts):
        with _pace_lock:
            now = time.monotonic()
            slot = max(now, _next_request)
            if slot + 1 >= deadline:
                raise JevError("Batch time budget reached. Retry unfinished rows.")
            _next_request = slot + interval
        time.sleep(max(0, slot - time.monotonic()))
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise JevError("Batch time budget reached. Retry unfinished rows.")
        delay = min(2 ** attempt, 8)
        try:
            response = poster(TYPESAFE_URL,
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=payload, timeout=min(timeout, remaining))
        except (requests.Timeout, requests.ConnectionError):
            if attempt + 1 == max_attempts:
                raise JevError("Could not reach Jev after retries. Retry unfinished rows.")
        except requests.RequestException:
            raise JevError("Could not reach Jev.")
        else:
            if response.status_code in (401, 403):
                raise JevError("Invalid Jev API key.")
            if response.status_code in (429, 500, 502, 503, 504, 529):
                try:
                    delay = max(delay, float(response.headers.get("Retry-After", 0)))
                except (ValueError, TypeError):
                    try:
                        until = parsedate_to_datetime(response.headers.get("Retry-After", ""))
                        delay = max(delay, (until - datetime.now(timezone.utc)).total_seconds())
                    except (ValueError, TypeError, OverflowError):
                        pass
                if attempt + 1 == max_attempts:
                    raise JevError("Jev is busy or rate limited. Retry unfinished rows.")
            elif not response.ok:
                # Do not echo provider-controlled bodies which can contain input.
                raise JevError(f"Jev rejected the request (HTTP {response.status_code}).")
            else:
                try:
                    body = response.json()
                except ValueError:
                    raise JevError("Jev returned a response that was not JSON.")
                if not isinstance(body, dict) or not isinstance(body.get("answers"), dict):
                    raise JevError("Jev response did not contain answers.")
                if set(body["answers"]) != set(questions):
                    raise JevError("Jev response has missing or unexpected answers.")
                return body if details else body["answers"]
        if time.monotonic() + delay + 1 >= deadline:
            raise JevError("Retry delay exceeds batch time budget. Retry unfinished rows later.")
        time.sleep(delay)
    raise JevError("Jev request failed.")
