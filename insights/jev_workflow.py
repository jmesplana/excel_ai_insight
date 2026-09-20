"""Dataset-independent Jev workflow configuration and row evaluation."""
import math
import logging

logger = logging.getLogger(__name__)
from jev_provider import JevConfigError, build_question, decode_answer


def finite(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def validate_workflow(configs, derived=None):
    if not isinstance(configs, list) or not 1 <= len(configs) <= 100:
        raise JevConfigError('Use 1–100 questions per workflow.')
    seen = set()
    for config in configs:
        if not isinstance(config, dict):
            raise JevConfigError('Each question must be an object.')
        unknown = set(config) - {'outputColumnName', 'questionType', 'instructions', 'options', 'review', 'dependsOn', 'branches', 'criteria'}
        if unknown:
            raise JevConfigError('Unknown question settings: ' + ', '.join(sorted(unknown)))
        if config.get('questionType') != 'noul' and 'criteria' in config:
            raise JevConfigError('Use option descriptions for Choice/Score; criteria is for Noul.')
        name = config.get('outputColumnName')
        if not isinstance(name, str) or not name.strip() or name in seen:
            raise JevConfigError('Output names must be non-empty and unique.')
        deps = config.get('dependsOn', [])
        if not isinstance(deps, list) or any(d not in seen for d in deps):
            raise JevConfigError(f'{name}: dependencies must name earlier questions.')
        branches = config.get('branches')
        if branches is not None:
            if not deps or not isinstance(branches, list) or not branches:
                raise JevConfigError(f'{name}: branches need dependencies and a non-empty array.')
            matches = set()
            for branch in branches:
                if isinstance(branch, dict) and set(branch) - {'when', 'options'}:
                    raise JevConfigError(f'{name}: branches only accept when and options.')
                if not isinstance(branch, dict) or not isinstance(branch.get('when'), dict) or set(branch['when']) != set(deps):
                    raise JevConfigError(f'{name}: each branch must match every dependency.')
                match = tuple(str(branch['when'][d]) for d in deps)
                if match in matches:
                    raise JevConfigError(f'{name}: duplicate branch conditions.')
                matches.add(match)
                build_question({**config, 'type': config.get('questionType'), 'options': branch.get('options', [])})
        else:
            build_question({**config, 'type': config.get('questionType')})
        review = config.get('review', {})
        if not isinstance(review, dict) or any(not finite(v) or not 0 <= v <= 1 for v in review.values()):
            raise JevConfigError(f'{name}: review thresholds must be numbers from 0 to 1.')
        allowed_review = {'noMax', 'yesMin'} if config.get('questionType') == 'noul' else {'minConfidence'}
        if set(review) - allowed_review:
            raise JevConfigError(f'{name}: unknown review threshold.')
        if review.get('noMax', .2) >= review.get('yesMin', .8):
            raise JevConfigError(f'{name}: noMax must be below yesMin.')
        seen.add(name)
    if not isinstance(derived or [], list):
        raise JevConfigError('derived must be an array.')
    for item in derived or []:
        name = item.get('outputColumnName') if isinstance(item, dict) else None
        if not isinstance(name, str) or not name.strip() or name in seen:
            raise JevConfigError('Derived output names must be unique.')
        allowed = {'outputColumnName', 'type', 'inputs', 'table'} if item.get('type', 'lookup') == 'lookup' else {'outputColumnName', 'type', 'weights'}
        if set(item) - allowed:
            raise JevConfigError(f'{name}: unknown derived settings.')
        if item.get('type', 'lookup') == 'lookup':
            inputs = item.get('inputs', [])
            if not isinstance(inputs, list) or not inputs or any(not isinstance(i, str) or i not in seen for i in inputs):
                raise JevConfigError(f'{name}: lookup inputs must name earlier outputs.')
            table = item.get('table')
            if not isinstance(table, list) or not table:
                raise JevConfigError(f'{name}: supply a lookup table.')
            keys = set()
            for row in table:
                if not isinstance(row, dict) or not isinstance(row.get('when'), dict) or set(row['when']) != set(inputs) or not isinstance(row.get('value'), str):
                    raise JevConfigError(f'{name}: table entries need when and a string value.')
                key = tuple(str(row['when'][i]) for i in inputs)
                if key in keys:
                    raise JevConfigError(f'{name}: duplicate lookup keys.')
                keys.add(key)
        elif item.get('type') == 'composite':
            weights = item.get('weights')
            scores = {c['outputColumnName']: c for c in configs if c.get('questionType') == 'score'}
            if not isinstance(weights, dict) or not weights or any(k not in scores or not finite(v) or v <= 0 for k, v in weights.items()):
                raise JevConfigError(f'{name}: weights must be positive and reference Score questions.')
        else:
            raise JevConfigError(f'{name}: unknown derived type.')
        seen.add(name)
    return configs


def evaluate_row(state, configs, derived, call, previous=None):
    """Evaluate independent questions together; later stages receive prior decisions.

    Previous successes are validated and reused only within the frozen run.
    Failed, blocked and uncertain answers have distinct statuses.
    """
    records = {}
    remaining = dict(enumerate(configs))
    calls = []
    empty = state is None or (isinstance(state, str) and not state.strip())
    while remaining:
        questions, decoders, active = {}, {}, {}
        for i, config in list(remaining.items()):
            name = config['outputColumnName']
            deps = config.get('dependsOn', [])
            if any(d not in records for d in deps):
                continue
            if empty:
                records[name] = {'status': 'empty', 'value': None}
                del remaining[i]
                continue
            if any(records[d]['status'] != 'ok' for d in deps):
                records[name] = {'status': 'blocked', 'value': None, 'reason': 'Dependency requires review or retry.'}
                del remaining[i]
                continue
            spec = {**config, 'type': config.get('questionType')}
            if 'branches' in config:
                branch = next((b for b in config['branches'] if all(records[d]['value'] == b['when'][d] for d in deps)), None)
                if branch is None:
                    records[name] = {'status': 'blocked', 'value': None, 'reason': 'No configured branch matches.'}
                    del remaining[i]
                    continue
                spec['options'] = branch['options']
            question, decoder = build_question(spec)
            key = f'q{i}'
            old = (previous or {}).get(name, {})
            if old.get('status') in ('ok', 'review') and old.get('raw'):
                try:
                    records[name] = make_record(old['raw'], decoder, config)
                    records[name]['model'] = old.get('model')
                    del remaining[i]
                    continue
                except Exception:
                    pass
            questions[key], decoders[key], active[key] = question, decoder, (i, config)
        if not questions:
            if remaining:
                # Some reused decisions may have unlocked another stage.
                continue
            break
        dependencies = {d for _, c in active.values() for d in c.get('dependsOn', [])}
        context = {'record': state, 'decisions': {d: records[d]['value'] for d in dependencies}} if dependencies else state
        try:
            response = call(context, questions)
            calls.append({'model': response.get('model'), 'usage': response.get('usage', {})})
            answers = response['answers']
            error = None
        except Exception as exc:
            logger.error("Jev request failed (%s)", type(exc).__name__)
            from .errors import public_error
            answers, error = {}, public_error(exc)
        for key, (i, config) in active.items():
            name = config['outputColumnName']
            try:
                if error:
                    from jev_provider import JevError
                    raise JevError(error)
                records[name] = make_record(answers.get(key), decoders[key], config)
                records[name]['model'] = response.get('model')
            except Exception as exc:
                from .errors import public_error
                records[name] = {'status': 'error', 'value': None, 'reason': public_error(exc)}
            del remaining[i]
    for item in derived or []:
        name = item['outputColumnName']
        inputs = item.get('inputs', list(item.get('weights', {})))
        if all(records[k]['status'] == 'empty' for k in inputs):
            records[name] = {'status': 'empty', 'value': None}
        elif any(records[k]['status'] != 'ok' for k in inputs):
            records[name] = {'status': 'blocked', 'value': None, 'reason': 'Inputs require review or retry.'}
        elif item.get('type', 'lookup') == 'lookup':
            found = next((entry for entry in item['table'] if all(records[k]['value'] == entry['when'][k] for k in inputs)), None)
            records[name] = ({'status': 'ok', 'value': found['value']} if found else
                             {'status': 'review', 'value': None, 'reason': 'No configured lookup match.'})
        else:
            weights = item['weights']
            value = sum(records[k]['detail'] / (len(records[k]['labels']) - 1) * w for k, w in weights.items()) / sum(weights.values())
            records[name] = {'status': 'ok', 'value': value}
    return records, calls


def make_record(answer, decoder, config):
    value, confidence, detail = decode_answer(answer, decoder)
    review = config.get('review', {})
    uncertain = confidence is not None and confidence < review.get('minConfidence', .5)
    if decoder['type'] == 'noul':
        uncertain = review.get('noMax', .2) < detail < review.get('yesMin', .8)
        value = decoder['yes'] if detail >= review.get('yesMin', .8) else decoder['no'] if detail <= review.get('noMax', .2) else None
    return {'status': 'review' if uncertain else 'ok', 'value': value,
            'confidence': confidence, 'detail': detail, 'raw': answer,
            'labels': decoder.get('labels'),
            'reason': 'Below configured certainty threshold.' if uncertain else None}
