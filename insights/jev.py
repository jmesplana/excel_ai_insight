"""Jev routes: concurrent rows, staged questions, explicit decision audits."""
import logging
from concurrent.futures import ThreadPoolExecutor

import requests
from flask import Blueprint, request, jsonify

from jev_provider import JevConfigError, JevError, ask, resolve_config
from .errors import public_error

logger = logging.getLogger(__name__)

bp = Blueprint("jev", __name__)

MAX_ROWS_PER_BATCH = 100


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


@bp.route('/validate_jev_config', methods=['POST'])
def validate_jev_config():
    from .jev_workflow import validate_workflow
    try:
        data = request.get_json() or {}
        validate_workflow(data.get('configs'), data.get('derived'))
        return jsonify(ok=True)
    except (JevConfigError, TypeError, AttributeError) as e:
        return jsonify(error=public_error(e)), 400


@bp.route('/analyze_batch_jev', methods=['POST'])
def analyze_batch_jev():
    from .jev_workflow import validate_workflow, evaluate_row
    import time
    data = request.get_json() or {}
    try:
        configs = data.get('configs')
        derived = data.get('derived', [])
        validate_workflow(configs, derived)
        rows = data.get('rows')
        if not isinstance(rows, list) or not 1 <= len(rows) <= MAX_ROWS_PER_BATCH:
            raise JevConfigError(f'Use 1–{MAX_ROWS_PER_BATCH} rows per batch.')
        if any(not isinstance(r, dict) or type(r.get('rowIndex')) is not int for r in rows):
            raise JevConfigError('Rows require an integer rowIndex.')
        if len({r['rowIndex'] for r in rows}) != len(rows):
            raise JevConfigError('Row indices must be unique.')
        config = resolve_config(data)
        execution = data.get('execution', {})
        limits = {'workers': (4, 1, 8), 'requestsPerMinute': (600, 1, 1200),
                  'tokensPerSecond': (100000, 1, 250000), 'maxAttempts': (3, 1, 4)}
        settings = {}
        for key, (default, low, high) in limits.items():
            value = execution.get(key, default)
            if type(value) is not int or not low <= value <= high:
                raise JevConfigError(f'{key} must be an integer between {low} and {high}.')
            settings[key] = value
    except (JevConfigError, TypeError, AttributeError) as e:
        return jsonify(error=public_error(e)), 400

    deadline = time.monotonic() + 40
    def run_row(row):
        with requests.Session() as session:
            def call(state, questions):
                return ask(state, questions, api_key=config['api_key'], model=config['model'],
                    session=session, details=True, deadline=deadline,
                    requests_per_minute=settings['requestsPerMinute'],
                    tokens_per_second=settings['tokensPerSecond'], max_attempts=settings['maxAttempts'])
            records, calls = evaluate_row(row.get('state'), configs, derived, call, row.get('previous'))
        values = {}
        for name, record in records.items():
            status = record['status']
            values[name] = record.get('value')
            if status == 'error':
                values[name] = 'Error: ' + record['reason']
            elif status == 'empty':
                values[name] = 'No data (empty cell)'
            elif values[name] is None:
                values[name] = 'Needs review'
            if data.get('includeConfidence') and 'raw' in record:
                values[name + '__confidence'] = record.get('confidence') if record.get('confidence') is not None else ''
                if record['raw']['type'] in ('score', 'noul'):
                    values[name + '__score'] = record['detail']
        return {'rowIndex': row['rowIndex'], 'values': values, 'decisions': records, 'calls': calls}

    with ThreadPoolExecutor(max_workers=settings['workers']) as executor:
        results = list(executor.map(run_row, rows))
    errors = sum(any(d['status'] == 'error' for d in r['decisions'].values()) for r in results)
    return jsonify(results=results, errors=errors)
