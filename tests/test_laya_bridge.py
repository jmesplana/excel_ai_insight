import unittest
from unittest.mock import patch
from local_laya_bridge import create_bridge, allowed_origins

ORIGIN = 'https://aidstack-test.vercel.app'

class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.client = create_bridge(origins=ORIGIN).test_client()

    def test_preflight_and_private_network_headers(self):
        response = self.client.options('/analyze_batch_jev', headers={
            'Origin':ORIGIN, 'Access-Control-Request-Method':'POST',
            'Access-Control-Request-Headers':'content-type', 'Access-Control-Request-Private-Network':'true'})
        self.assertEqual(response.status_code, 204)
        self.assertEqual(response.headers['Access-Control-Allow-Origin'], ORIGIN)
        self.assertEqual(response.headers['Access-Control-Allow-Private-Network'], 'true')
        self.assertNotIn('Access-Control-Allow-Credentials', response.headers)

    def test_untrusted_origins_cannot_trigger_work(self):
        for origin in ['https://evil.example', 'https://aidstack-test.vercel.app.evil.example', 'null', None]:
            with self.subTest(origin=origin), patch('laya_provider.ask') as ask:
                response = self.client.post('/test_jev_connection', json={}, headers={'Origin':origin} if origin else {})
                self.assertEqual(response.status_code, 403)
                self.assertNotIn('Access-Control-Allow-Origin', response.headers)
                ask.assert_not_called()

    def test_connector_forces_local_laya_even_if_browser_requests_cloud(self):
        with patch('laya_provider.ask', return_value={}) as ask, patch('insights.jev.ask') as cloud:
            response = self.client.post('/test_jev_connection', headers={'Origin':ORIGIN}, json={
                'classificationProvider':'jev', 'layaEndpoint':'https://evil.example', 'jevApiKey':'secret'})
            self.assertEqual(response.status_code, 200)
            self.assertEqual(ask.call_args.kwargs['config']['endpoint'], 'http://127.0.0.1:8000')
            self.assertEqual(response.json['provider'], 'Laya')
            cloud.assert_not_called()

    def test_rejects_dns_rebinding_form_posts_and_other_routes(self):
        self.assertEqual(self.client.post('/test_jev_connection', headers={'Origin':ORIGIN,'Host':'evil.example'}, json={}).status_code, 403)
        self.assertEqual(self.client.post('/test_jev_connection', headers={'Origin':ORIGIN}, data='x').status_code, 415)
        self.assertEqual(self.client.post('/analyze_batch', headers={'Origin':ORIGIN}, json={}).status_code, 404)
        self.assertEqual(self.client.post('/test_jev_connection', headers={'Origin':ORIGIN}, json=[]).status_code, 400)

    def test_health_identifies_connector(self):
        response = self.client.get('/health', headers={'Origin':ORIGIN})
        self.assertEqual(response.json['service'], 'aidstack-laya-connector')

    def test_origin_configuration_refuses_wildcards_paths_and_remote_http(self):
        for origin in ['*', 'https://*.vercel.app', 'https://site.example/path', 'http://remote.example']:
            with self.assertRaises(ValueError): allowed_origins(origin)
