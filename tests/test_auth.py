import os
import re
import unittest
from unittest.mock import patch

from app import create_app
from insights import auth

ENV = {'AUTH_ALLOWED_EMAILS': 'Person@Example.org', 'AUTH_ALLOWED_DOMAINS': 'ifrc.org',
       'SECRET_KEY': 'test-secret', 'APP_BASE_URL': 'https://insights.example'}
NODE = {'id': 'a', 'name': 'Translate', 'type': 'analyze', 'inputs': ['text'], 'provider': 'azure',
        'instructions': 'Translate.', 'fields': [{'name': 'english', 'options': []}]}


class AllowListTest(unittest.TestCase):
    def test_emails_and_exact_domains(self):
        with patch.dict(os.environ, ENV):
            self.assertTrue(auth.is_allowed(' person@example.org '))
            self.assertTrue(auth.is_allowed('Someone@IFRC.org'))
            for email in ('other@example.org', 'x@sub.ifrc.org', 'x@ifrc.org.evil.com', 'x@evilifrc.org', 'ifrc.org', ''):
                self.assertFalse(auth.is_allowed(email), email)

    def test_disabled_without_a_list(self):
        with patch.dict(os.environ, {'AUTH_ALLOWED_EMAILS': '', 'AUTH_ALLOWED_DOMAINS': ''}):
            self.assertFalse(auth.auth_enabled())

    def test_safe_next_stays_on_site(self):
        for target in ('https://evil.com', '//evil.com', '/\\evil.com', 'javascript:alert(1)', ''):
            self.assertEqual(auth.safe_next(target), '/pipeline', target)
        self.assertEqual(auth.safe_next('/pipeline?x=1'), '/pipeline?x=1')


class SignInFlowTest(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, ENV)
        self.env.start()
        self.addCleanup(self.env.stop)
        self.client = create_app().test_client()

    def link_for(self, email):
        with patch('insights.auth.send_link') as send:
            response = self.client.post('/login', data={'email': email, 'next': '/pipeline'})
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'Check your email', response.data)
        url = send.call_args.args[1]
        self.assertTrue(url.startswith('https://insights.example/auth/verify?token='))
        return url[len('https://insights.example'):]

    def test_studio_requires_sign_in(self):
        response = self.client.get('/pipeline')
        self.assertEqual(response.status_code, 302)
        self.assertIn('/login', response.headers['Location'])
        api = self.client.post('/api/pipeline/step', json={'node': NODE, 'record': {'text': 'x'}})
        self.assertEqual(api.status_code, 401)
        self.assertEqual(self.client.get('/api/pipeline/presets').status_code, 401)
        self.assertEqual(self.client.get('/').status_code, 200)

    def test_unapproved_email_is_told_and_gets_no_email(self):
        with patch('insights.auth.send_link') as send:
            response = self.client.post('/login', data={'email': 'someone@gmail.com'})
        self.assertEqual(response.status_code, 403)
        self.assertIn(b'not on the approved list', response.data)
        send.assert_not_called()

    def test_link_signs_in_and_reaches_studio(self):
        path = self.link_for('Someone@ifrc.org')
        response = self.client.get(path)
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers['Location'].endswith('/pipeline'))
        page = self.client.get('/pipeline')
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"organization's AI connection", page.data)
        self.assertIn(b'someone@ifrc.org', page.data)
        self.client.post('/logout')
        self.assertEqual(self.client.get('/pipeline').status_code, 302)

    def test_tampered_and_expired_links_fail(self):
        path = self.link_for('someone@ifrc.org')
        bad = self.client.get(path[:-3] + 'abc')
        self.assertIn('error=invalid', bad.headers['Location'])
        with patch('itsdangerous.timed.time.time', return_value=10**10):
            expired = self.client.get(path)
        self.assertIn('error=expired', expired.headers['Location'])
        self.assertEqual(self.client.get('/pipeline').status_code, 302)

    def test_removed_from_list_is_signed_out(self):
        self.client.get(self.link_for('person@example.org'))
        self.assertEqual(self.client.get('/pipeline').status_code, 200)
        with patch.dict(os.environ, {'AUTH_ALLOWED_EMAILS': 'other@example.org'}):
            self.assertEqual(self.client.get('/pipeline').status_code, 302)

    def test_steps_pass_browser_credentials_through(self):
        self.client.get(self.link_for('someone@ifrc.org'))
        with patch('insights.pipeline.execute_step', return_value={'status': 'ok', 'items': [{'english': 'hi'}]}) as run:
            response = self.client.post('/api/pipeline/step', json={'node': NODE, 'record': {'text': 'x'},
                                                                    'credentials': {'apiKey': 'browser-key'}})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(run.call_args.args[2], {'apiKey': 'browser-key'})

    def test_missing_secret_key_fails_closed(self):
        with patch.dict(os.environ, {'SECRET_KEY': ''}):
            client = create_app().test_client()
            self.assertEqual(client.get('/pipeline').status_code, 503)


SERVER = {'AZURE_OPENAI_API_KEY': 'server-key', 'AZURE_OPENAI_ENDPOINT': 'https://server.openai.azure.com',
          'AZURE_OPENAI_DEPLOYMENT': 'dep', 'OPENAI_API_KEY': 'server-openai', 'TYPESAFE_API_KEY': 'server-jev'}


class ServerCredentialTest(unittest.TestCase):
    """Server keys go only to signed-in users, and never to an endpoint a visitor chose."""

    def setUp(self):
        self.env = patch.dict(os.environ, {**ENV, **SERVER})
        self.env.start()
        self.addCleanup(self.env.stop)
        self.app = create_app()

    def resolve(self, data, email=None):
        from llm_provider import resolve_config
        with self.app.test_request_context('/'):
            if email:
                from flask import session
                session['email'] = email
            return resolve_config(data)

    def test_signed_out_visitors_get_no_server_keys(self):
        from llm_provider import LLMConfigError
        from jev_provider import JevConfigError, resolve_config as jev_config
        for data in ({'provider': 'azure'}, {'provider': 'openai'}):
            with self.assertRaises(LLMConfigError):
                self.resolve(data)
        with self.app.test_request_context('/'), self.assertRaises(JevConfigError):
            jev_config({'classificationProvider': 'jev'})

    def test_signed_in_users_fall_back_to_server_keys(self):
        config = self.resolve({'provider': 'azure'}, 'someone@ifrc.org')
        self.assertEqual((config['api_key'], config['endpoint'], config['model']), ('server-key', 'https://server.openai.azure.com', 'dep'))
        self.assertEqual(self.resolve({'provider': 'openai'}, 'someone@ifrc.org')['api_key'], 'server-openai')

    def test_browser_keys_take_precedence(self):
        config = self.resolve({'provider': 'azure', 'apiKey': 'mine', 'azureEndpoint': 'https://mine.openai.azure.com'})
        self.assertEqual((config['api_key'], config['endpoint']), ('mine', 'https://mine.openai.azure.com'))

    def test_server_key_never_goes_to_a_browser_endpoint(self):
        config = self.resolve({'provider': 'azure', 'azureEndpoint': 'https://attacker.example'}, 'someone@ifrc.org')
        self.assertEqual((config['api_key'], config['endpoint']), ('server-key', 'https://server.openai.azure.com'))

    def test_removed_user_loses_server_keys(self):
        from llm_provider import LLMConfigError
        with self.assertRaises(LLMConfigError):
            self.resolve({'provider': 'azure'}, 'someone@gmail.com')

    def test_without_sign_in_server_keys_work_as_before(self):
        with patch.dict(os.environ, {'AUTH_ALLOWED_EMAILS': '', 'AUTH_ALLOWED_DOMAINS': ''}):
            self.assertEqual(self.resolve({'provider': 'azure'})['api_key'], 'server-key')


if __name__ == '__main__':
    unittest.main()
