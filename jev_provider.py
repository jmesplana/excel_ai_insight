"""
Jev (typesafe.ai) System One provider.

Where the OpenAI/Azure path asks a model for free text and takes whatever prose
comes back, Jev answers *typed* questions: every answer is constrained to a list
the caller supplied, and comes with a probability per option plus a confidence
statistic. Nothing outside the supplied options can ever be returned, so the
value that lands in a spreadsheet cell needs no parsing or cleanup.

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

    instructions = _clean(spec.get("instructions"))
    if not instructions:
        raise JevConfigError("Each Jev question needs instructions.")

    options = [
        str(option).strip()
        for option in (spec.get("options") or [])
        if str(option).strip()
    ]

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
        labels = {}
        for label in options:
            slug = slugify(label, criteria)
            # The label doubles as the option's description: these lists come
            # from the user's own codebook, where the label *is* the meaning.
            criteria[slug] = label
            labels[slug] = label
        return (
            {"type": "choice", "instructions": instructions, "criteria": criteria},
            {"type": "choice", "labels": labels},
        )

    if qtype == "score":
        if not MIN_SCORE_LEVELS <= len(options) <= MAX_SCORE_LEVELS:
            raise JevConfigError(
                f"A Score question needs between {MIN_SCORE_LEVELS} and "
                f"{MAX_SCORE_LEVELS} levels ({len(options)} given)."
            )
        return (
            {"type": "score", "instructions": instructions, "criteria": list(options)},
            {"type": "score", "labels": list(options)},
        )

    return (
        {"type": "noul", "instructions": instructions},
        {"type": "noul",
         "yes": options[0] if len(options) > 0 else "Yes",
         "no": options[1] if len(options) > 1 else "No"},
    )


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

    confidence = answer.get("confidence")

    if decoder["type"] == "choice":
        key = answer.get("choice")
        # Fall back to the raw key if a label lookup misses, so an unexpected
        # key still lands in the cell rather than blanking it.
        return decoder["labels"].get(key, key), confidence, None

    if decoder["type"] == "score":
        labels = decoder["labels"]
        raw = answer.get("score")
        if raw is None:
            raise JevError("Score answer missing a score.")
        # Round half *up*, not to even: these levels are ordered severities, so
        # an exact midpoint must always resolve to the higher one. Python's
        # built-in round() would send 0.5 down to level 0 but 1.5 up to level 2.
        index = math.floor(float(raw) + 0.5)
        index = max(0, min(len(labels) - 1, index))
        return labels[index], confidence, float(raw)

    probability = answer.get("noul")
    if probability is None:
        raise JevError("Noul answer missing a probability.")
    probability = float(probability)
    label = decoder["yes"] if probability >= 0.5 else decoder["no"]
    return label, confidence, probability


def ask(state, questions, *, api_key, model=DEFAULT_JEV_MODEL,
        timeout=REQUEST_TIMEOUT, session=None):
    """
    Send one System One request and return its `answers` object.

    Args:
        state: the row context -- a string or a JSON-serializable dict.
        questions: {question_key: question} as built by build_question().

    Raises:
        JevError: on an API error or an unreadable response body.
    """
    poster = session.post if session is not None else requests.post
    try:
        response = poster(
            TYPESAFE_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={"model": model, "state": state, "questions": questions},
            timeout=timeout,
        )
    except requests.Timeout:
        raise JevError("Jev request timed out.")
    except requests.RequestException as e:
        raise JevError(f"Could not reach Jev: {type(e).__name__}")

    if response.status_code == 401 or response.status_code == 403:
        raise JevError("Invalid Jev API key.")
    if response.status_code == 429:
        raise JevError("Jev rate limit reached. Wait a moment and resume.")
    if not response.ok:
        raise JevError(_api_error_message(response))

    try:
        body = response.json()
    except ValueError:
        raise JevError("Jev returned a response that was not JSON.")

    answers = body.get("answers")
    if not isinstance(answers, dict):
        raise JevError("Jev response did not contain any answers.")
    return answers


def _api_error_message(response):
    """Best-effort human-readable message from a failed API response."""
    try:
        body = response.json()
    except ValueError:
        return f"Jev error (HTTP {response.status_code})."
    detail = body.get("error") or body.get("message") or body.get("detail")
    if isinstance(detail, dict):
        detail = detail.get("message")
    if detail:
        return f"Jev error: {detail}"
    return f"Jev error (HTTP {response.status_code})."
