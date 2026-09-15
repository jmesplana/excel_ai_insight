import unittest
from unittest.mock import patch, MagicMock
from app import create_app
from insights import icd, analysis
from insights.errors import AnalysisOutputError

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
