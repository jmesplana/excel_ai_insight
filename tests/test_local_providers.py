import os
import unittest
from unittest.mock import patch, MagicMock

from app import create_app
from llm_provider import resolve_config, build_client, LLMConfigError
from jev_provider import resolve_config as classifier_config, JevConfigError
from local_services import local_url
from laya_provider import validate_configs


class LocalProviderTests(unittest.TestCase):
    def setUp(self):
        self.client = create_app().test_client()

    def test_ollama_never_uses_cloud_keys_or_models(self):
        with patch.dict(os.environ, {'OPENAI_API_KEY': 'cloud-secret', 'OPENAI_MODEL': 'cloud-model'}):
            config = resolve_config({'provider': 'ollama', 'apiKey': 'azure-secret', 'ollamaModel': 'gemma3:1b'})
        self.assertEqual(config['api_key'], 'ollama')
        self.assertEqual(config['model'], 'gemma3:1b')
        with patch('llm_provider.OpenAI') as sdk:
            build_client(config)
            self.assertEqual(sdk.call_args.kwargs['base_url'], 'http://127.0.0.1:11434/v1')
            self.assertEqual(sdk.call_args.kwargs['api_key'], 'ollama')
            sdk.call_args.kwargs['http_client'].close()

    def test_loopback_only_and_no_url_credentials(self):
        for value in ['https://evil.example', 'http://127.0.0.1.evil.example',
                      'http://secret@localhost:8000', 'http://localhost:8000/v1',
                      'http://localhost:8000?token=secret', 'file:///tmp/model', 'http://localhost:bad']:
            with self.subTest(value=value), self.assertRaises(ValueError):
                local_url(value)
        self.assertEqual(local_url('http://[::1]:8000/'), 'http://[::1]:8000')

    def test_laya_independent_of_jev_credentials(self):
        with patch.dict(os.environ, {'TYPESAFE_API_KEY': 'cloud-secret', 'TYPESAFE_MODEL': 'jev-pinned'}, clear=True):
            config = classifier_config({'classificationProvider': 'laya', 'jevApiKey': 'also-secret'})
        self.assertEqual(config['api_key'], '')
        self.assertEqual(config['model'], 'multilingual')

    def test_laya_branch_limit_before_inference(self):
        with self.assertRaises(JevConfigError):
            validate_configs([{'branches': [{'options': list(range(101))}]}])

    def test_laya_runs_existing_workflow_and_preserves_audit(self):
        response = MagicMock(status_code=200)
        response.json.return_value = {'answers': {'q0': {'type': 'choice', 'choice': 'a',
            'confidence': .9, 'probabilities': {'a': .95, 'b': .05}}}, 'usage': {'input_tokens': 20}}
        with patch('laya_provider.requests.Session') as session, patch('insights.jev.ask') as cloud:
            post = session.return_value.__enter__.return_value.post
            post.return_value = response
            result = self.client.post('/analyze_batch_jev', json={
                'classificationProvider': 'laya', 'jevApiKey': 'cloud-secret',
                'configs': [{'outputColumnName': 'Category', 'questionType': 'choice', 'instructions': 'Choose', 'options': ['A', 'B']}],
                'rows': [{'rowIndex': 0, 'state': 'Example'}]}).get_json()
            self.assertEqual(result['errors'], 0)
            record = result['results'][0]
            self.assertEqual(record['values']['Category'], 'A')
            self.assertEqual(record['decisions']['Category']['model'], 'laya:multilingual')
            self.assertEqual(post.call_args.args[0], 'http://127.0.0.1:8000/v1/systemone')
            self.assertNotIn('cloud-secret', str(post.call_args))
            self.assertFalse(post.call_args.kwargs['allow_redirects'])
            cloud.assert_not_called()

    def test_laya_failure_does_not_fall_back_to_cloud(self):
        import requests
        with patch('laya_provider.requests.Session') as session, patch('insights.jev.ask') as cloud:
            session.return_value.__enter__.return_value.post.side_effect = requests.ConnectionError('private secret')
            response = self.client.post('/test_jev_connection', json={'classificationProvider': 'laya'})
            self.assertEqual(response.status_code, 400)
            self.assertNotIn('private secret', response.get_data(as_text=True))
            cloud.assert_not_called()

    def test_installed_ollama_models(self):
        with patch('requests.Session') as session:
            session.return_value.__enter__.return_value.get.return_value.json.return_value = {'models': [{'name': 'gemma3:1b'}]}
            response = self.client.post('/ollama_models', json={})
            self.assertEqual(response.get_json(), {'models': ['gemma3:1b']})
