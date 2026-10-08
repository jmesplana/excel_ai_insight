import json
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from app import create_app
from insights.jev_workflow import evaluate_row, validate_workflow
from jev_provider import JevConfigError


def answer(question):
    keys = list(question['criteria'])
    return {'type': 'choice', 'choice': keys[0], 'confidence': .95,
            'probabilities': {key: int(i == 0) for i, key in enumerate(keys)}}


class SourceBranchTests(unittest.TestCase):
    def setUp(self):
        self.config = {'outputColumnName': 'Code', 'questionType': 'choice', 'instructions': 'Classify the assigned type.',
            'sourceBranches': [
                {'when': {'type_of_feedback': 'Questions'}, 'options': ['Question A', 'Question B']},
                {'when': {'type_of_feedback': 'Demandes'}, 'options': ['Request A', 'Request B']}]}

    def test_repeated_id_and_text_with_different_types_have_distinct_options(self):
        validate_workflow([self.config])
        states = [{'ID': 1, 'Feedback': 'How does it work? Please send soap.', 'type_of_feedback': t}
                  for t in ['Questions', 'Demandes']]
        def call(state, questions, **kwargs):
            expected = 'question' if state['type_of_feedback'] == 'Questions' else 'request'
            self.assertTrue(all(k.startswith(expected) for k in questions['q0']['criteria']))
            return {'answers': {'q0': answer(questions['q0'])}}
        with patch('insights.jev.ask', side_effect=call) as api:
            result = create_app().test_client().post('/analyze_batch_jev', json={
                'jevApiKey': 'test', 'configs': [self.config],
                'rows': [{'rowIndex': i, 'state': s} for i, s in enumerate(states)]}).get_json()
        self.assertEqual(api.call_count, 2)
        self.assertEqual([r['values']['Code'] for r in result['results']], ['Question A', 'Request A'])

    def test_unknown_or_missing_type_never_falls_back_to_all_options(self):
        for state in [{'Feedback': 'Text'}, {'type_of_feedback': 'Unknown'}, 'type_of_feedback: Questions']:
            call = Mock()
            records, _ = evaluate_row(state, [self.config], [], call)
            call.assert_not_called()
            self.assertEqual(records['Code']['status'], 'review')
            self.assertIsNone(records['Code']['value'])

    def test_invalid_source_branches_fail_before_calls(self):
        for change in [
            {'sourceBranches': []}, {'sourceBranches': {}}, {'sourceBranches': None}, {'questionType': 'score'},
            {'sourceBranches': [self.config['sourceBranches'][0]] * 2},
            {'sourceBranches': [{'when': {'type_of_feedback': 'Q'}, 'options': ['Only']}]},
            {'sourceBranches': [{'when': {'type_of_feedback': ['Q']}, 'options': ['A', 'B']}]},
            {'branches': self.config['sourceBranches']},
        ]:
            with self.subTest(change=change), self.assertRaises(JevConfigError):
                validate_workflow([{**self.config, **change}])

    def test_source_choice_and_other_assessments_share_one_call(self):
        other = {'outputColumnName': 'Sentiment', 'questionType': 'choice', 'instructions': 'Tone', 'options': ['Positive', 'Negative']}
        call = Mock(side_effect=lambda state, qs: {'answers': {k: answer(q) for k, q in qs.items()}})
        records, calls = evaluate_row({'type_of_feedback': 'Questions'}, [self.config, other], [], call)
        self.assertEqual(call.call_count, 1)
        self.assertEqual(calls[0]['questionCount'], 2)
        self.assertEqual(calls[0]['optionCount'], 4)
        self.assertEqual(records['Sentiment']['value'], 'Positive')

    def test_workers_reuse_connections_between_rows(self):
        def call(state, questions, **kwargs):
            return {'answers': {k: answer(q) for k, q in questions.items()}}
        with patch('insights.jev.ask', side_effect=call) as api:
            response = create_app().test_client().post('/analyze_batch_jev', json={
                'jevApiKey': 'test', 'configs': [self.config], 'execution': {'workers': 1},
                'rows': [{'rowIndex': i, 'state': {'type_of_feedback': 'Questions'}} for i in range(3)]})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(len({id(c.kwargs['session']) for c in api.call_args_list}), 1)

    def test_supplied_fast_config_maps_every_path_without_more_model_calls(self):
        path = Path(__file__).resolve().parents[1] / 'sample_jev/ebola-feedback-coding.fast.config.json'
        if not path.exists():
            self.skipTest('Local sample configuration not present')
        config = json.loads(path.read_text())
        validate_workflow(config['questions'], config['derived'])
        path_question = config['questions'][0]
        count = 0
        for branch in path_question['sourceBranches']:
            for selected in branch['options'][:-1]:
                def call(state, questions):
                    q = questions['q0']
                    key = next(k for k, v in q['criteria'].items() if v == selected)
                    return {'answers': {'q0': {'type': 'choice', 'choice': key, 'confidence': .95,
                        'probabilities': {k: int(k == key) for k in q['criteria']}}}}
                records, calls = evaluate_row(branch['when'], [path_question], config['derived'], call)
                self.assertEqual(len(calls), 1)
                self.assertEqual(' > '.join(records[k]['value'] for k in
                    ['feedback_category', 'feedback_sous_code', 'feedback_code']), selected)
                self.assertEqual(sum(records[k]['value'] == selected for k in ['feedback_path_1', 'feedback_path_2']), 1)
                count += 1
        self.assertEqual(count, 441)


if __name__ == '__main__':
    unittest.main()
