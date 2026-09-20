import unittest
from unittest.mock import Mock, patch
from jev_provider import build_question, decode_answer, JevError, JevConfigError, ask
from insights.jev_workflow import validate_workflow, evaluate_row
from app import create_app


def choice(key='a', confidence=.9):
    return {'type': 'choice', 'choice': key, 'confidence': confidence, 'probabilities': {'a': .9, 'b': .1}}


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.parent = {'outputColumnName': 'Category', 'questionType': 'choice', 'instructions': 'Classify', 'options': ['A', 'B']}
        self.child = {'outputColumnName': 'Detail', 'questionType': 'choice', 'instructions': 'Classify detail',
                      'dependsOn': ['Category'], 'branches': [{'when': {'Category': 'A'}, 'options': ['A', 'B']}]}

    def test_dependencies_and_lookups_and_successful_answer_reuse(self):
        derived = [{'outputColumnName': 'Mapped', 'inputs': ['Category', 'Detail'],
                    'table': [{'when': {'Category': 'A', 'Detail': 'A'}, 'value': 'Configured value'}]}]
        validate_workflow([self.parent, self.child], derived)
        call = Mock(side_effect=[{'model': 'pinned', 'answers': {'q0': choice()}}, JevError('Temporary failure')])
        records, calls = evaluate_row('source', [self.parent, self.child], derived, call)
        self.assertEqual(records['Category']['status'], 'ok')
        self.assertEqual(records['Detail']['status'], 'error')
        self.assertEqual(records['Mapped']['status'], 'blocked')
        self.assertEqual(call.call_args.args[0]['decisions'], {'Category': 'A'})
        retry = Mock(return_value={'model': 'pinned', 'answers': {'q1': choice()}})
        records, calls = evaluate_row('source', [self.parent, self.child], derived, retry, records)
        self.assertEqual(retry.call_count, 1)
        self.assertEqual(set(retry.call_args.args[1]), {'q1'})
        self.assertEqual(records['Mapped']['value'], 'Configured value')
        self.assertEqual(records['Category']['model'], 'pinned')

    def test_uncertain_parent_blocks_child_and_noul_uses_configured_thresholds(self):
        call = Mock(return_value={'answers': {'q0': choice(confidence=.3)}})
        records, _ = evaluate_row('source', [self.parent, self.child], [], call)
        self.assertEqual(call.call_count, 1)
        self.assertEqual(records['Category']['status'], 'review')
        self.assertEqual(records['Detail']['status'], 'blocked')
        noul = {'outputColumnName': 'Flag', 'questionType': 'noul', 'instructions': 'Present?', 'review': {'noMax': .1, 'yesMin': .9}}
        records, _ = evaluate_row('source', [noul], [], Mock(return_value={'answers': {'q0': {'type': 'noul', 'noul': .6}}}))
        self.assertEqual(records['Flag']['status'], 'review')
        self.assertIsNone(records['Flag']['value'])
        self.assertEqual(records['Flag']['detail'], .6)

    def test_invalid_graphs_fail_before_api(self):
        with self.assertRaises(JevConfigError): validate_workflow([self.child, self.parent])
        with self.assertRaises(JevConfigError): validate_workflow([self.parent, self.parent])
        with self.assertRaises(JevConfigError): validate_workflow([{**self.parent, 'review': {'minConfidence': float('nan')}}])

    def test_unknown_choice_invalid_probabilities_and_noul_are_rejected(self):
        _, decoder = build_question({**self.parent, 'type': 'choice'})
        for raw in [{}, {**choice(), 'choice': 'unknown'}, {**choice(), 'probabilities': {'a': 2, 'b': -1}}, {**choice(), 'type': 'score'}]:
            with self.assertRaises(JevError): decode_answer(raw, decoder)
        _, decoder = build_question({'type': 'noul', 'instructions': 'True?'})
        for value in [-.1, 1.1, float('nan'), True, '0.5']:
            with self.assertRaises(JevError): decode_answer({'type': 'noul', 'noul': value}, decoder)

    def test_structured_criteria_preserved(self):
        q, decoder = build_question({'type': 'choice', 'instructions': {'question': 'Which?'},
            'options': [{'label': 'A', 'description': {'includes': ['alpha']}}, 'B']})
        self.assertEqual(q['criteria']['a'], {'includes': ['alpha']})
        self.assertEqual(decoder['labels']['a'], 'A')

    def test_empty_rows_cost_nothing_and_derived_is_blocked(self):
        call = Mock()
        records, _ = evaluate_row(None, [self.parent, self.child], [], call)
        call.assert_not_called()
        self.assertEqual({d['status'] for d in records.values()}, {'empty'})

    def test_retry_429_and_preserve_usage(self):
        busy = Mock(status_code=429, headers={'Retry-After': '2'})
        ok = Mock(status_code=200, ok=True)
        ok.json.return_value = {'model': 'pinned', 'answers': {'q': {'type': 'noul', 'noul': 1}}, 'usage': {'input_tokens': 42}}
        session = Mock()
        session.post.side_effect = [busy, ok]
        with patch('jev_provider.time.sleep') as sleep, patch('jev_provider._next_request', 0):
            response = ask('state', {'q': {'type': 'noul', 'instructions': 'True?'}}, api_key='secret', session=session, details=True)
        self.assertEqual(session.post.call_count, 2)
        self.assertEqual(response['usage']['input_tokens'], 42)
        self.assertTrue(any(c.args[0] >= 2 for c in sleep.call_args_list))

    def test_report_uses_evidence_without_row_sampling_and_rejects_oversize(self):
        client = create_app().test_client()
        packet = {'coverage': {'total': 10001, 'processed': 10001}, 'distributions': [{'name': 'Category', 'counts': [{'value': 'Tail', 'count': 5001}]}]}
        mock = Mock()
        mock.chat.completions.create.return_value = []
        with patch('insights.jev_report.get_client_and_model', return_value=(mock, 'test')):
            response = client.post('/jev_report', json={'report': packet})
            response.get_data()
        self.assertIn('10001', mock.chat.completions.create.call_args.kwargs['messages'][1]['content'])
        packet['huge'] = 'x' * 180001
        self.assertEqual(client.post('/jev_report', json={'report': packet}).status_code, 400)

    def test_composite_uses_normalized_scores_and_configured_weights(self):
        configs = [{'outputColumnName': 'Short', 'questionType': 'score', 'instructions': 'Rate', 'options': ['Low','High']},
                   {'outputColumnName': 'Long', 'questionType': 'score', 'instructions': 'Rate', 'options': ['Low','Mid','High']}]
        derived = [{'outputColumnName': 'Index', 'type': 'composite', 'weights': {'Short': 2, 'Long': 1}}]
        validate_workflow(configs, derived)
        call = Mock(return_value={'answers': {
            'q0': {'type':'score','score':1,'confidence':1,'probabilities':{'0':0,'1':1}},
            'q1': {'type':'score','score':1,'confidence':1,'probabilities':{'0':0,'1':1,'2':0}}}})
        records, _ = evaluate_row('input', configs, derived, call)
        self.assertAlmostEqual(records['Index']['value'], 2.5 / 3)
        records, _ = evaluate_row(None, configs, derived, call)
        self.assertEqual(records['Index']['status'], 'empty')

    def test_raw_decisions_are_retained_without_visible_confidence_columns(self):
        client = create_app().test_client()
        with patch('insights.jev.ask', return_value={'model':'pinned','answers':{'q0':choice()},'usage':{'input_tokens':17}}):
            response = client.post('/analyze_batch_jev', json={'jevApiKey':'key','configs':[self.parent],
                'rows':[{'rowIndex':8,'state':'source'}], 'includeConfidence':False})
        result = response.get_json()['results'][0]
        self.assertNotIn('Category__confidence', result['values'])
        self.assertEqual(result['decisions']['Category']['raw']['probabilities']['a'], .9)
        self.assertEqual(result['calls'][0]['usage']['input_tokens'],17)
