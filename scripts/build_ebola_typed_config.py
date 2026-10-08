"""Build a type-routed workflow from the supplied four-column Ebola workbook.

Run with venv/bin/python scripts/build_ebola_typed_config.py.
Requires openpyxl for reading the source workbook; no model calls are made.
"""
import argparse
import copy
import json
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
FALLBACK = 'À vérifier : aucun code ne correspond au type fourni'


def build_config(workbook, original):
    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        rows = iter(book.active.values)
        if tuple(next(rows)) != ('Type', 'Catégorie', 'Sous-code', 'Code'):
            raise ValueError('Expected Type, Catégorie, Sous-code, Code workbook columns.')
        entries = [tuple(row) for row in rows if any(v is not None for v in row)]
    finally:
        book.close()
    groups = defaultdict(list)
    paths = {}
    for kind, category, subcode, code in entries:
        if any(not isinstance(v, str) or not v.strip() for v in (kind, category, subcode, code)):
            raise ValueError('Every codebook row must contain four non-empty text values.')
        path = ' > '.join((category, subcode, code))
        if path in paths:
            raise ValueError(f'Ambiguous or repeated codebook path: {path}')
        paths[path] = (category, subcode, code)
        groups[kind].append(path)
    old_paths = original['questions'][:2]
    old_union = set().union(*(set(q['options']) for q in old_paths))
    if set(paths) != old_union - {'Non applicable (autre groupe de types)'}:
        raise ValueError('Original path questions and workbook contain different codes.')

    config = copy.deepcopy(original)
    # Keep the carefully tuned clause selection and detailed coding rules, while
    # removing the old two-group routing instructions.
    intro = old_paths[0]['instructions'].split(' This question only covers')[0]
    specific = old_paths[1]['instructions'].split('Origin beliefs:', 1)[1].rsplit('Return one exact configured label.', 1)[0]
    instructions = (intro.replace('record.type_of_feedback', 'state.type_of_feedback').replace('record.Feedback', 'state.Feedback')
        + ' The supplied options have already been restricted to the human-assigned Type. '
        'Choose the most specific matching Catégorie > Sous-code > Code path. '
        'Generic codes repeating their category and Autre(s) codes are fallbacks only when no specific code fits. '
        f'If no interpretable clause matches the assigned Type or no code fits, choose "{FALLBACK}". '
        'Do not choose a code for another clause. Match meaning, not just keywords. '
        'Origin beliefs:' + specific + 'Return one exact configured label.')
    question = {'outputColumnName': 'feedback_path', 'questionType': 'choice',
                'instructions': instructions, 'options': [],
                'sourceBranches': [{'when': {'type_of_feedback': kind}, 'options': options + [FALLBACK]}
                                   for kind, options in groups.items()],
                'review': {'minConfidence': 0.6}}
    config['questions'] = [question] + copy.deepcopy(original['questions'][2:])
    for q in config['questions'][1:]:
        q['instructions'] = q['instructions'].replace('record.type_of_feedback', 'state.type_of_feedback').replace('record.Feedback', 'state.Feedback')
    config['derived'] = []
    for index, name in enumerate(('feedback_category', 'feedback_sous_code', 'feedback_code')):
        config['derived'].append({'outputColumnName': name, 'type': 'lookup', 'inputs': ['feedback_path'],
            'table': [{'when': {'feedback_path': path}, 'value': parts[index]} for path, parts in paths.items()]})
    # Retain the original path columns for existing reports, without two extra
    # predictions. A fallback has no lookup entry, so it produces review/blank.
    for q in old_paths:
        config['derived'].append({'outputColumnName': q['outputColumnName'], 'type': 'lookup', 'inputs': ['feedback_path'],
            'table': [{'when': {'feedback_path': path}, 'value': path if path in q['options']
                       else 'Non applicable (autre groupe de types)'} for path in paths]})
    config['outputOrder'] = ['feedback_category', 'feedback_sous_code', 'feedback_code', 'feedback_path',
                             'feedback_path_1', 'feedback_path_2'] + [q['outputColumnName'] for q in config['questions'][1:]]
    config['execution'].update(batchSize=16, workers=8, reuseIdenticalRows=True, structuredState=True)
    return config


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, default=ROOT / 'sample_jev/Ebola_codes_avec_sous-codes (1).xlsx')
    parser.add_argument('--config', type=Path, default=ROOT / 'sample_jev/ebola-feedback-coding.config_v19 (1).json')
    parser.add_argument('--output', type=Path, default=ROOT / 'sample_jev/ebola-feedback-coding.fast.config.json')
    args = parser.parse_args()
    config = build_config(args.workbook, json.loads(args.config.read_text(encoding='utf-8')))
    args.output.write_text(json.dumps(config, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Wrote {args.output}')
    print('Codes by type:', {b['when']['type_of_feedback']: len(b['options']) - 1 for b in config['questions'][0]['sourceBranches']})


if __name__ == '__main__':
    main()
