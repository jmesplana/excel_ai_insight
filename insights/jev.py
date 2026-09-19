"""Jev (typesafe.ai) analysis routes.

Mirrors the /analyze_batch contract of the LLM path so the browser can reuse
the same BatchRun checkpointing, but the unit of work is a *row* rather than a
cell: every result column for one row is asked in a single System One request,
which the API evaluates in parallel against one shared state.
"""
import logging
from concurrent.futures import ThreadPoolExecutor

import requests
from flask import Blueprint, request, jsonify

from jev_provider import (JevConfigError, JevError, ask, build_question,
                          decode_answer, resolve_config)
from .errors import public_error

logger = logging.getLogger(__name__)

bp = Blueprint("jev", __name__)

# Rows in flight at once. The API allows 1,200 requests/minute; 8 concurrent
# rows keeps a comfortable margin while still saturating a typical batch.
MAX_JEV_WORKERS = 8

MAX_ROWS_PER_BATCH = 100
MAX_QUESTIONS_PER_ROW = 20


@bp.route('/test_jev_connection', methods=['POST'])
def test_jev_connection():
    """Verify a Jev API key with one trivial question before a whole file runs."""
    data = request.json or {}
    try:
        config = resolve_config(data)
    except JevConfigError as e:
        return jsonify({"ok": False, "error": public_error(e)}), 400

    try:
        ask(
            "ping",
            {"_probe": {"type": "noul", "instructions": "Is this a test message?"}},
            api_key=config["api_key"],
            model=config["model"],
            timeout=20,
        )
    except JevError as e:
        status = 401 if "Invalid Jev API key" in str(e) else 400
        return jsonify({"ok": False, "error": public_error(e)}), status

    return jsonify({"ok": True, "provider": "Jev", "model": config["model"]})


def _build_questions(configs):
    """
    Build the shared question set once per batch.

    Every row in a batch asks the same questions, so the API payload and the
    answer decoders are assembled a single time and reused across rows.

    Returns:
        (questions, decoders, output_names) keyed by question key.

    Raises:
        JevConfigError: if any question is not a valid Jev primitive.
    """
    questions = {}
    decoders = {}
    output_names = {}

    for index, config in enumerate(configs):
        # The question key is internal; the output column name is what the user
        # sees, and is not necessarily a valid key (spaces, accents, blanks).
        key = f"q{index}"
        question, decoder = build_question({
            "type": config.get("questionType"),
            "instructions": config.get("instructions"),
            "options": config.get("options"),
        })
        questions[key] = question
        decoders[key] = decoder
        output_names[key] = (config.get('outputColumnName')
                             or f"jev_analysis_{index + 1}")

    return questions, decoders, output_names


@bp.route('/analyze_batch_jev', methods=['POST'])
def analyze_batch_jev():
    """Stateless Jev analysis of a single batch of rows.

    One request per row, carrying every configured question for that row.

    Expected JSON body:
        {
          "jevApiKey": str,
          "jevModel": str,
          "includeConfidence": bool,     # add __confidence / __score columns
          "isFirstBatch": bool,          # validate the key only on the first batch
          "configs": [{"outputColumnName": str, "questionType": str,
                       "instructions": str, "options": [str]}],
          "rows": [{"rowIndex": int, "state": str}]
        }

    Returns:
        {"results": [{"rowIndex": int, "values": {"<outputColumnName>": str}}],
         "errors": int}
    """
    try:
        data = request.json
        if not data:
            raise ValueError("No data received in request.")

        configs = data.get('configs') or []
        rows = data.get('rows') or []
        include_confidence = bool(data.get('includeConfidence'))
        is_first_batch = data.get('isFirstBatch', False)

        if not configs:
            return jsonify({"error": "Missing 'configs' in the request data."}), 400
        if not isinstance(configs, list) or len(configs) > MAX_QUESTIONS_PER_ROW:
            return jsonify(error=f'Use at most {MAX_QUESTIONS_PER_ROW} result columns per batch.'), 400
        if not isinstance(rows, list) or len(rows) > MAX_ROWS_PER_BATCH:
            return jsonify(error=f'Use at most {MAX_ROWS_PER_BATCH} rows per batch.'), 400
        if any(not isinstance(row, dict) or not isinstance(row.get('rowIndex'), int)
               for row in rows):
            return jsonify(error='Rows require an integer rowIndex.'), 400
        if len({row['rowIndex'] for row in rows}) != len(rows):
            return jsonify(error='Each row must have a unique rowIndex.'), 400

        try:
            config = resolve_config(data)
            questions, decoders, output_names = _build_questions(configs)
        except JevConfigError as e:
            return jsonify({"error": public_error(e)}), 400

        # Validate the key once per run to fail fast with one clear message
        # instead of the same auth error on every row.
        if is_first_batch:
            try:
                ask("ping",
                    {"_probe": {"type": "noul", "instructions": "Is this a test message?"}},
                    api_key=config["api_key"], model=config["model"], timeout=20)
            except JevError as e:
                status = 401 if "Invalid Jev API key" in str(e) else 500
                return jsonify({"error": public_error(e)}), status

        error_count = 0
        results_by_row = {}
        blank_rows = []
        tasks = []

        for row in rows:
            row_index = row.get('rowIndex')
            results_by_row[row_index] = {}
            state = row.get('state')
            if state is None or str(state).strip() == '':
                blank_rows.append(row_index)
            else:
                tasks.append((row_index, str(state)))

        # A row with no input is a no-op: fill every column without an API call.
        for row_index in blank_rows:
            for key, name in output_names.items():
                results_by_row[row_index][name] = "No data (empty cell)"
                if include_confidence:
                    results_by_row[row_index][f"{name}__confidence"] = ""
                    if decoders[key]["type"] in ("score", "noul"):
                        results_by_row[row_index][f"{name}__score"] = ""

        def run_row(task):
            """Ask every question for one row in a single API call."""
            row_index, state = task
            values = {}
            try:
                # One session per row keeps the connection pool thread-safe.
                with requests.Session() as session:
                    answers = ask(state, questions, api_key=config["api_key"],
                                  model=config["model"], session=session)
            except Exception as e:
                logger.error("Jev row failed (%s)", type(e).__name__)
                message = f"Error: {public_error(e)}"
                for name in output_names.values():
                    values[name] = message
                return row_index, values, True

            failed = False
            for key, decoder in decoders.items():
                name = output_names[key]
                try:
                    value, confidence, detail = decode_answer(answers.get(key), decoder)
                    values[name] = value
                    if include_confidence:
                        values[f"{name}__confidence"] = (
                            round(confidence, 4) if isinstance(confidence, (int, float)) else "")
                        if decoder["type"] in ("score", "noul"):
                            values[f"{name}__score"] = (
                                round(detail, 4) if isinstance(detail, (int, float)) else "")
                except Exception as e:
                    logger.error("Jev answer failed (%s)", type(e).__name__)
                    values[name] = f"Error: {public_error(e)}"
                    failed = True
            return row_index, values, failed

        if tasks:
            with ThreadPoolExecutor(max_workers=MAX_JEV_WORKERS) as executor:
                for row_index, values, is_error in executor.map(run_row, tasks):
                    results_by_row[row_index].update(values)
                    if is_error:
                        error_count += 1

        results = [{"rowIndex": idx, "values": vals}
                   for idx, vals in results_by_row.items()]
        return jsonify({"results": results, "errors": error_count})

    except Exception as e:
        logger.error("Operation failed (%s)", type(e).__name__)
        return jsonify({"error": public_error(e)}), 500
