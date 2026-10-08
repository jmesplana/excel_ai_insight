"""Extractive context selection: candidates are exact spans, never generated text."""
import re

from jev_provider import JevConfigError, MAX_CHOICE_OPTIONS, build_question


def validate_excerpt(config):
    spec = config['sourceExcerpt']
    if (not isinstance(spec, dict) or set(spec) != {'column', 'separators', 'splitBefore'}
            or not isinstance(spec['column'], str) or not spec['column'].strip()):
        raise JevConfigError('sourceExcerpt requires column, separators and splitBefore.')
    for name in ('separators', 'splitBefore'):
        values = spec[name]
        if (not isinstance(values, list) or len(values) > 50
                or any(not isinstance(v, str) or not v or len(v) > 100 for v in values)):
            raise JevConfigError(f'sourceExcerpt.{name} must contain up to 50 non-empty literal strings.')
    build_question({**config, 'type': 'choice', 'options': ['candidate', 'no match']})


def excerpt_candidates(text, spec):
    """Literal boundaries propose candidates; Jev selects the speech act."""
    if not isinstance(text, str):
        return []
    boundaries = {0, len(text)}
    for separator in spec['separators']:
        for match in re.finditer(re.escape(separator), text):
            start, end = match.span()
            if separator == '.' and start > 0 and end < len(text) and text[start - 1].isdigit() and text[end].isdigit():
                continue
            boundaries.add(end)
    for phrase in spec['splitBefore']:
        pattern = r'(?<!\w)' + re.escape(phrase) + r'(?!\w)'
        boundaries.update(m.start() for m in re.finditer(pattern, text, re.IGNORECASE))
    points = sorted(boundaries)
    candidates = []
    for start, end in zip(points, points[1:]):
        while start < end and text[start].isspace():
            start += 1
        while end > start and text[end - 1].isspace():
            end -= 1
        if start < end and any(c.isalnum() for c in text[start:end]):
            candidates.append({'text': text[start:end], 'start': start, 'end': end})
    return candidates


def excerpt_question(state, config):
    spec = config['sourceExcerpt']
    source = state.get(spec['column']) if isinstance(state, dict) else None
    candidates = excerpt_candidates(source, spec)
    if not candidates:
        return None, None, {'status': 'empty', 'value': None, 'reason': 'No source excerpt.'}
    if len(candidates) >= MAX_CHOICE_OPTIONS:
        return None, None, {'status': 'review', 'value': None,
                            'reason': 'Too many candidate excerpts; review source segmentation. No text was truncated.'}
    if len(candidates) == 1:
        span = candidates[0]
        return None, None, {'status': 'ok', 'value': span['text'], 'confidence': None,
                            'sourceColumn': spec['column'], 'sourceSpan': [span['start'], span['end']],
                            'method': 'single_excerpt',
                            'reason': 'Only one candidate; type matching is evaluated by the code question.'}
    options = [{'label': f'excerpt_{i}', 'description': c['text']} for i, c in enumerate(candidates)]
    options.append({'label': 'no_match', 'description':
                    'No single candidate expresses the assigned type unambiguously, or it needs missing context. Require human review.'})
    question, decoder = build_question({**config, 'type': 'choice', 'options': options})
    decoder['labels'] = {f'excerpt_{i}': c['text'] for i, c in enumerate(candidates)}
    decoder['labels']['no_match'] = None
    decoder['spans'] = {f'excerpt_{i}': [c['start'], c['end']] for i, c in enumerate(candidates)}
    return question, decoder, None
