import unittest
from unittest.mock import patch, MagicMock
from app import create_app
from insights import icd, analysis, jev
from insights.errors import AnalysisOutputError, public_error
import jev_provider

class AppTests(unittest.TestCase):
    def setUp(self):
        self.app = create_app()
        self.client = self.app.test_client()

    def test_pages_and_modules(self):
        for path in ['/', '/about', '/how-it-works', '/static/js/app.js', '/static/js/results.js', '/static/vendor/purify.es.mjs']:
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200, path)

    def test_batch_preserves_row_indices_and_errors(self):
        with patch('insights.analysis.get_client_and_model', return_value=(MagicMock(), 'test')), patch('insights.analysis.analyze_text', side_effect=['positive', RuntimeError('secret-key private-row')]):
            response = self.client.post('/analyze_batch', json={'configs': [{'id':'a','outputColumnName':'Result'}], 'rows':[
                {'rowIndex':7,'inputs':{'a':'good'}}, {'rowIndex':9,'inputs':{'a':'bad'}}, {'rowIndex':10,'inputs':{'a':None}}]})
        data = response.get_json()
        self.assertEqual(data['errors'], 1)
        self.assertEqual([r['rowIndex'] for r in data['results']], [7,9,10])
        self.assertEqual(data['results'][0]['values']['Result'], 'positive')
        self.assertNotIn('secret-key', response.get_data(as_text=True))

    def test_chat_does_not_log_payload(self):
        with self.assertLogs(self.app.logger, level='ERROR') as logs, patch('insights.chat.get_client_and_model', side_effect=RuntimeError('secret-key')):
            # Route reports a safe message without logging request content.
            self.client.post('/chat_with_data', json={'rows':[{'diagnosis':'private-diagnosis'}], 'question':'summary', 'apiKey':'secret-key'})
            self.app.logger.error('test marker')
        self.assertNotIn('secret-key', '\n'.join(logs.output))
        self.assertNotIn('private-diagnosis', '\n'.join(logs.output))

    def test_icd_mixed_provenance(self):
        with self.app.app_context(), patch.object(icd, 'ai_full_icd', return_value={'icd11Code':'A','icd10Code':'B'}), patch.object(icd, 'icd_codeinfo', return_value='https://id.who.int/icd/release/11/2026-01/mms/123'), patch.object(icd, 'icd_entity_title', return_value='term'), patch.object(icd,'load_icd_maps'), patch.object(icd,'derive_icd10',return_value={'code':''}):
            result = icd._icd_llm_fallback('token',0,'input','en','en',['icd11','icd10'],{'model':'test'})
        self.assertEqual(result['outputs']['icd11']['codeSource'], 'WHO')
        self.assertEqual(result['outputs']['icd10']['codeSource'], 'LLM suggestion')
        self.assertIn('LLM', result['source'])

    def test_jev_batch_asks_every_question_in_one_call_per_row(self):
        # One request per row, carrying all questions -- not one per cell.
        answers = {'q0': {'choice': 'plainte', 'confidence': 0.91},
                   'q1': {'score': 1.6, 'confidence': 0.55}}
        with patch.object(jev, 'ask', return_value=answers) as ask:
            response = self.client.post('/analyze_batch_jev', json={
                'jevApiKey': 'secret-key', 'includeConfidence': True,
                'configs': [
                    {'outputColumnName': 'Type', 'questionType': 'choice',
                     'instructions': 'Classer', 'options': ['Plainte', 'Refus']},
                    {'outputColumnName': 'Urgence', 'questionType': 'score',
                     'instructions': 'Evaluer', 'options': ['Faible', 'Moyenne', 'Élevée']}],
                'rows': [{'rowIndex': 4, 'state': 'Feedback: a'},
                         {'rowIndex': 5, 'state': '   '}]})
        data = response.get_json()
        self.assertEqual(data['errors'], 0)
        # Two rows, but only the non-blank one costs a call.
        self.assertEqual(ask.call_count, 1)
        self.assertEqual(len(ask.call_args.args[1]), 2)
        values = {r['rowIndex']: r['values'] for r in data['results']}
        # Labels come back decoded, not as the internal option slugs.
        self.assertEqual(values[4]['Type'], 'Plainte')
        self.assertEqual(values[4]['Urgence'], 'Élevée')
        self.assertEqual(values[4]['Urgence__score'], 1.6)
        self.assertEqual(values[4]['Type__confidence'], 0.91)
        self.assertEqual(values[5]['Type'], 'No data (empty cell)')
        self.assertNotIn('secret-key', response.get_data(as_text=True))

    def test_jev_row_failure_is_isolated_and_key_never_leaks(self):
        with self.assertLogs(level='ERROR') as logs, patch.object(
                jev, 'ask', side_effect=[{'q0': {'choice': 'a', 'confidence': 0.9}},
                                         RuntimeError('secret-key leaked')]):
            response = self.client.post('/analyze_batch_jev', json={
                'jevApiKey': 'secret-key',
                'configs': [{'outputColumnName': 'R', 'questionType': 'choice',
                             'instructions': 'i', 'options': ['A', 'B']}],
                'rows': [{'rowIndex': 0, 'state': 'x'}, {'rowIndex': 1, 'state': 'y'}]})
        data = response.get_json()
        # One bad row does not abandon the batch.
        self.assertEqual(data['errors'], 1)
        self.assertEqual(len(data['results']), 2)
        self.assertNotIn('secret-key', response.get_data(as_text=True))
        self.assertNotIn('secret-key', '\n'.join(logs.output))

    def test_jev_invalid_question_is_rejected_before_any_call(self):
        with patch.object(jev, 'ask') as ask:
            response = self.client.post('/analyze_batch_jev', json={
                'jevApiKey': 'k',
                'configs': [{'outputColumnName': 'R', 'questionType': 'choice',
                             'instructions': 'i', 'options': ['OnlyOne']}],
                'rows': [{'rowIndex': 0, 'state': 'x'}]})
        self.assertEqual(response.status_code, 400)
        ask.assert_not_called()
        # The message names the actual problem so it can be fixed.
        self.assertIn('at least 2 options', response.get_json()['error'])

    def test_jev_slugs_are_ascii_and_unique(self):
        question, decoder = jev_provider.build_question({
            'type': 'choice', 'instructions': 'i',
            # Two labels that collapse to the same ASCII slug.
            'options': ['Appréciation', 'Appreciation', 'Refus']})
        slugs = list(question['criteria'])
        self.assertEqual(len(slugs), len(set(slugs)))
        self.assertTrue(all(s.isascii() for s in slugs))
        # Every slug still decodes back to the label the user typed.
        self.assertEqual(sorted(decoder['labels'].values()),
                         sorted(['Appréciation', 'Appreciation', 'Refus']))

    def test_jev_score_rounds_to_nearest_level_and_clamps(self):
        _, decoder = jev_provider.build_question({
            'type': 'score', 'instructions': 'i', 'options': ['Low', 'Mid', 'High']})
        self.assertEqual(jev_provider.decode_answer({'score': 0.4}, decoder)[0], 'Low')
        self.assertEqual(jev_provider.decode_answer({'score': 1.4}, decoder)[0], 'Mid')
        # Midpoints round *up* consistently: on an ordered severity scale
        # built-in round() would send 0.5 down but 1.5 up (banker's rounding).
        self.assertEqual(jev_provider.decode_answer({'score': 0.5}, decoder)[0], 'Mid')
        self.assertEqual(jev_provider.decode_answer({'score': 1.5}, decoder)[0], 'High')
        # An out-of-range score is clamped rather than raising an IndexError.
        self.assertEqual(jev_provider.decode_answer({'score': 9.0}, decoder)[0], 'High')
        self.assertEqual(jev_provider.decode_answer({'score': -2.0}, decoder)[0], 'Low')

    def test_jev_config_errors_are_shown_but_api_errors_stay_generic(self):
        # The user's own setup mistake must be readable...
        self.assertIn('at least 2 options', public_error(
            jev_provider.JevConfigError('A Choice question needs at least 2 options.')))
        # ...while an unexpected exception still reveals nothing.
        self.assertNotIn('boom', public_error(RuntimeError('boom')))

    def test_analysis_rejects_truncated_output(self):
        client = MagicMock()
        client.chat.completions.create.return_value.choices[0].finish_reason = 'length'
        with self.assertRaises(AnalysisOutputError):
            analysis.analyze_text(client, 'input', 'translate', 'test', 4096)
        self.assertEqual(client.chat.completions.create.call_args.kwargs['max_tokens'], 4096)

    def test_long_input_never_silently_truncates(self):
        client = MagicMock()
        with self.assertRaises(AnalysisOutputError):
            analysis.analyze_text(client, 'x' * 8001, 'translate')
        client.chat.completions.create.assert_not_called()

    def test_duplicate_row_indices_rejected_before_provider_call(self):
        with patch('insights.analysis.get_client_and_model') as provider:
            response = self.client.post('/analyze_batch', json={'configs':[{'id':'a'}], 'rows':[
                {'rowIndex':1,'inputs':{}}, {'rowIndex':1,'inputs':{}}]})
        self.assertEqual(response.status_code,400)
        provider.assert_not_called()

if __name__ == '__main__': unittest.main()
