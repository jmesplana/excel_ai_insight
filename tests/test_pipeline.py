import os
import json
import unittest
import httpx
from openai import RateLimitError
from insights.pipeline_errors import retry_delay
from unittest.mock import MagicMock, patch
from app import create_app
from insights.pipeline import validate_node, execute_step, AnalysisOutputError


def extract():
    return {'id': 'a', 'name': 'Extract', 'type': 'extract', 'inputs': ['Feedback'], 'provider': 'azure', 'instructions': 'Extract paired messages.', 'fields': [{'name': 'excerpt', 'options': []}, {'name': 'type', 'options': ['Thanks', 'Question']}], 'verbatimInput': 'Feedback', 'verbatimField': 'excerpt'}


def completion(value, finish='stop'):
    response = MagicMock()
    response.choices[0].finish_reason = finish
    response.choices[0].message.refusal = None
    response.choices[0].message.content = json.dumps(value)
    return response


class PipelineTests(unittest.TestCase):
    def test_drc_translation_review_rereads_original_without_previous_gloss(self):
        with self.client.get('/api/pipeline/presets/drc-ebola-framework') as template:
            node = template.get_json()['nodes'][0]
        record = {'Feedback': 'Tondimi maso bolobi'}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.side_effect = [
            completion({'status': 'review', 'reason': 'Unclear word', 'values': {'source_translation_en': 'We accept eyes spoke'}}),
            completion({'status': 'ok', 'reason': '', 'values': {'source_translation_en': 'We accept what you said.'}}),
        ]
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            result = execute_step(node, record, {})
        self.assertEqual(result['attempts'], 2)
        self.assertEqual(result['items'], [{'source_translation_en': 'We accept what you said.'}])
        for call in api.call_args_list:
            system, user = call.kwargs['messages']
            self.assertIn('collection setting', system['content'])
            self.assertIn('Ebola response in the DRC', json.loads(user['content'])['instructions'])
        retry = json.loads(api.call_args.kwargs['messages'][1]['content'])
        self.assertEqual(retry['record'], record)
        self.assertIn('reviewCheck', retry)
        self.assertNotIn('previousReviewReason', retry)
        self.assertNotIn('eyes spoke', json.dumps(retry))

    def test_partial_translation_is_returned_only_as_a_review_draft(self):
        node = {**extract(), 'type': 'analyze', 'fields': [{'name': 'translation_en', 'options': []}]}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.return_value = completion({'status': 'review', 'reason': 'Unclear reference',
                                       'values': {'translation_en': 'What is it called?'}})
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            result = execute_step(node, {'Feedback': 'Maladie Yangon kombo nayango nini'}, {})
        self.assertEqual(result['status'], 'review')
        self.assertEqual(result['drafts'], {'translation_en': 'What is it called?'})
        self.assertNotIn('items', result)

    def test_translation_review_retains_first_draft_when_recheck_returns_null(self):
        with self.client.get('/api/pipeline/presets/drc-ebola-framework') as template:
            node = template.get_json()['nodes'][0]
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        for placeholder in (None, 'null'):
            with self.subTest(placeholder=placeholder):
                api.side_effect = [completion({'status': 'review', 'reason': 'Uncertain nuance',
                                               'values': {'source_translation_en': value}})
                                   for value in ('We accept what you said.', placeholder)]
                with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
                    result = execute_step(node, {'Feedback': 'Tondimi maso bolobi'}, {})
                self.assertEqual(result['status'], 'review')
                self.assertEqual(result['drafts'], {'source_translation_en': 'We accept what you said.'})
                self.assertTrue(result['tentative'])
                schema = api.call_args.kwargs['response_format']['json_schema']['schema']
                self.assertEqual(next(iter(schema['properties'])), 'values')

    def setUp(self):
        # Keep a local .env sign-in list from changing these tests.
        env = patch.dict(os.environ, {'AUTH_ALLOWED_EMAILS': '', 'AUTH_ALLOWED_DOMAINS': ''})
        env.start()
        self.addCleanup(env.stop)
        self.client = create_app().test_client()

    def test_page_and_validation(self):
        with self.client.get('/pipeline') as response:
            self.assertEqual(response.status_code, 200)
        for node in [None, {}, {**extract(), 'fields': [None]}, {**extract(), 'provider': 'jev'}, {**extract(), 'maxOutputTokens': 'bad'}]:
            self.assertEqual(self.client.post('/api/pipeline/validate', json={'nodes': [node]}).status_code, 400)
        self.assertEqual(self.client.post('/api/pipeline/validate', json={'nodes': [extract()]}).status_code, 200)

    def test_builtin_workflow_templates(self):
        response = self.client.get('/api/pipeline/presets/drc-ebola-framework')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()['format'], 'aidstack-insights-pipeline')
        self.assertEqual(self.client.get('/api/pipeline/presets/unknown').status_code, 404)

    def test_numbers_are_review_without_api(self):
        with patch('insights.pipeline.get_client_and_model') as api:
            for value in [123, '', None, '123 ??']:
                out = self.client.post('/api/pipeline/step', json={'node': extract(), 'record': {'Feedback': value}}).get_json()
                self.assertEqual(out['status'], 'review')
            api.assert_not_called()

    def test_no_space_pairs_and_scoped_input(self):
        text = 'Merci pour votre aide.Où est le centre ?Quand ouvre-t-il ?'
        items = [{'excerpt': 'Merci pour votre aide.', 'type': 'Thanks'}, {'excerpt': 'Où est le centre ?', 'type': 'Question'}, {'excerpt': 'Quand ouvre-t-il ?', 'type': 'Question'}]
        client = MagicMock()
        client.with_options.return_value.chat.completions.create.return_value = completion({'status':'ok','reason':'','items':items})
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'deployment')) as resolver:
            result = execute_step({**extract(), 'model': 'chosen', 'inputs': ['Feedback', 'Context', 'Location']}, {'Feedback': text, 'Context': 'Centre opening hours', 'Location': 'Kabondo', 'private': 'not sent'}, {'apiKey': 'secret'})
        self.assertEqual(result['items'], items)
        call = client.with_options.return_value.chat.completions.create.call_args.kwargs
        self.assertEqual(call['response_format']['type'], 'json_schema')
        self.assertNotIn('not sent', json.dumps(call))
        self.assertEqual(json.loads(call['messages'][1]['content'])['record'], {'Feedback': text, 'Context': 'Centre opening hours', 'Location': 'Kabondo'})
        self.assertNotIn('schema', json.loads(call['messages'][1]['content']))
        self.assertEqual(resolver.call_args.args[0]['azureDeployment'], 'chosen')
        client.close.assert_called_once()

    def test_invalid_pairs_retry_then_review_not_accepted(self):
        for items in [[{'excerpt':'made up','type':'Question'}], [{'excerpt':'hello'}], [{'excerpt':'hello','type':'Bad'}], []]:
            client = MagicMock()
            client.with_options.return_value.chat.completions.create.return_value = completion({'status':'ok','reason':'','items':items})
            with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
                response = self.client.post('/api/pipeline/step', json={'node':extract(),'record':{'Feedback':'hello'}}).get_json()
            self.assertEqual(response['status'], 'review')
            self.assertEqual(client.with_options.return_value.chat.completions.create.call_count, 2)

    def test_truncation_and_refusal(self):
        client = MagicMock()
        client.with_options.return_value.chat.completions.create.return_value = completion({}, 'length')
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            out = self.client.post('/api/pipeline/step', json={'node':extract(),'record':{'Feedback':'hello'}}).get_json()
        self.assertEqual(out['status'], 'review')
        response = completion({})
        response.choices[0].message.refusal = 'No'
        client.with_options.return_value.chat.completions.create.return_value = response
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            self.assertEqual(execute_step(extract(), {'Feedback':'hello'}, {})['status'], 'review')

    def test_errors_do_not_expose_provider_secrets(self):
        with patch('insights.pipeline.get_client_and_model', side_effect=RuntimeError('private-feedback secret-key')):
            response = self.client.post('/api/pipeline/step', json={'node':extract(),'record':{'Feedback':'hello'}})
        self.assertEqual(response.get_json()['status'], 'error')
        self.assertNotIn('secret-key', response.get_data(as_text=True))

    def test_split_refuses_mismatch_and_keeps_repeats(self):
        node = {'type':'split','inputs':['excerpt','type'],'delimiter':' || '}
        out = execute_step(node, {'excerpt':'a || b || c','type':'Question || Question || Question'}, {})
        self.assertEqual(len(out['items']), 3)
        for value in ['Question || Question', 'Question ||  || Question']:
            with self.assertRaises(AnalysisOutputError):
                execute_step(node, {'excerpt':'a || b || c','type':value}, {})

    def test_jev_branch_and_confidence(self):
        node = {'type':'classify','inputs':['excerpt','type'], 'provider':'jev', 'instructions':'Classify', 'fields':[{'name':'code','options':[]}], 'sourceBranches':[{'when':{'type':'Question'},'options':['Centre','Hours']}], 'minConfidence':.6}
        validate_node(node)
        answer = {'model':'test', 'answers': {'q0': {'type':'choice','choice':'hours','confidence':.9,'probabilities':{'centre':.1,'hours':.9}}}}
        with patch('insights.jev.classify', return_value=answer) as classifier:
            out = execute_step(node, {'excerpt':'When?','type':'Question'}, {'jevApiKey':'secret'})
        self.assertEqual(out['status'], 'ok')
        self.assertEqual(out['items'], [{'code':'Hours'}])
        self.assertEqual(set(classifier.call_args.args[1]['q0']['criteria']), {'centre','hours'})
        with self.assertRaises(AnalysisOutputError):
            execute_step(node, {'excerpt':'When?','type':'Unknown'}, {'jevApiKey':'secret'})

    def test_lookup_and_required_validation(self):
        lookup = {'type':'lookup','inputs':['code'],'output':'category','table':[{'when':{'code':'C'},'value':'Centre'}]}
        self.assertEqual(execute_step(lookup, {'code':'C'}, {})['items'][0]['category'], 'Centre')
        with self.assertRaises(AnalysisOutputError): execute_step(lookup, {'code':'X'}, {})
        validate = {'type':'validate','inputs':['type'],'rules':[{'column':'type','allowed':['Question']}]}
        for value in ['', 'Thanks', None]:
            with self.assertRaises(AnalysisOutputError): execute_step(validate, {'type':value}, {})

    def test_explicit_review_outcome(self):
        node = {**extract(), 'reviewValues':['Question']}
        with patch('insights.pipeline.execute_step', return_value={'status':'ok','items':[{'excerpt':'Why?','type':'Question'}]}):
            result = self.client.post('/api/pipeline/step', json={'node':node,'record':{'Feedback':'Why?'}}).get_json()
        self.assertEqual(result['status'], 'review')

    def test_full_example_validates_and_azure_branch_is_constrained(self):
        with open('presets/drc-ebola-framework.pipeline.json', encoding='utf-8') as f:
            workflow = json.load(f)
        self.assertEqual(self.client.post('/api/pipeline/validate', json={'nodes': workflow['nodes']}).status_code, 200)
        node = {'type':'classify','inputs':['excerpt','type','prior'], 'provider':'azure', 'instructions':'Classify using decisions.prior', 'fields':[{'name':'code','options':[]}], 'sourceBranches':[{'when':{'type':'Question'},'options':['Centre','Hours']}]}
        client = MagicMock()
        client.with_options.return_value.chat.completions.create.return_value = completion({'status':'ok','reason':'','values':{'code':'Hours'}})
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'azure-test')):
            out = execute_step(node, {'excerpt':'When?','type':'Question','prior':'Some code'}, {})
        self.assertEqual(out['items'], [{'code':'Hours'}])
        kwargs = client.with_options.return_value.chat.completions.create.call_args.kwargs
        schema = kwargs['response_format']['json_schema']['schema']
        self.assertEqual(schema['properties']['values']['properties']['code']['enum'], ['Centre','Hours',None])
        self.assertEqual(json.loads(kwargs['messages'][1]['content'])['decisions']['prior'], 'Some code')


    def test_rate_limit_returns_cooldown_without_immediate_provider_retry(self):
        response = httpx.Response(429, headers={'retry-after': '30', 'retry-after-ms': '45000'}, request=httpx.Request('POST', 'https://example.test'))
        exc = RateLimitError('private provider response', response=response, body={'code':'rate_limit_exceeded','message':'secret'})
        client = MagicMock()
        client.with_options.return_value.chat.completions.create.side_effect = exc
        with patch('insights.pipeline.get_client_and_model', return_value=(client,'test')):
            out = self.client.post('/api/pipeline/step', json={'node':extract(),'record':{'Feedback':'Why?'}}).get_json()
        self.assertEqual(out['errorCode'], 'rate_limit')
        self.assertEqual(out['retryAfterSeconds'], 45)
        self.assertTrue(out['retryable'])
        self.assertEqual(client.with_options.return_value.chat.completions.create.call_count, 1)
        self.assertNotIn('secret', json.dumps(out))

    def test_quota_is_not_treated_as_transient_rate_limit(self):
        for body in [{'code':'insufficient_quota'}, {'error':{'type':'insufficient_quota'}}]:
            response = httpx.Response(429, headers={'retry-after':'1'}, request=httpx.Request('POST','https://example.test'))
            exc = RateLimitError('secret',response=response,body=body)
            with patch('insights.pipeline.execute_step',side_effect=exc):
                out = self.client.post('/api/pipeline/step',json={'node':extract(),'record':{'Feedback':'Why?'}}).get_json()
            self.assertEqual(out['errorCode'],'quota_exhausted')
            self.assertFalse(out['retryable'])
            self.assertNotIn('retryAfterSeconds',out)
            self.assertNotIn('secret',json.dumps(out))

    def test_retry_after_dates_milliseconds_and_invalid_headers(self):
        self.assertEqual(retry_delay({'retry-after':'Thu, 01 Jan 1970 00:02:00 GMT'}, now=60), 60)
        self.assertEqual(retry_delay({'x-ms-retry-after-ms':'1200'}), 1.2)
        self.assertIsNone(retry_delay({'retry-after':'NaN'}))
        self.assertIsNone(retry_delay({'retry-after':'-4'}))


    def test_mixed_language_review_gets_one_evidence_based_recheck(self):
        node = extract()
        text = 'Merci. Aksanti.'
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.side_effect = [completion({'status':'review','reason':'The feedback spans multiple languages and lacks context.','items':[]}),
                           completion({'status':'ok','reason':'','items':[{'excerpt':'Merci.','type':'Thanks'},{'excerpt':'Aksanti.','type':'Thanks'}]})]
        with patch('insights.pipeline.get_client_and_model',return_value=(client,'test')):
            result = execute_step(node, {'Feedback':text}, {})
        self.assertEqual(result['status'],'ok')
        self.assertEqual(api.call_count,2)
        prompt = json.loads(api.call_args_list[1].kwargs['messages'][1]['content'])
        self.assertIn('reviewCheck',prompt)
        self.assertNotIn('previousReviewReason', prompt)
        self.assertEqual(prompt['record']['Feedback'],text)
        self.assertIn('NEVER grounds for review',api.call_args_list[0].kwargs['messages'][0]['content'])

    def test_multilingual_recheck_does_not_force_acceptance(self):
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.side_effect = [completion({'status':'review','reason':'Mixed-language feedback.','items':[]}),
                           completion({'status':'review','reason':'The phrase “xyz” cannot be interpreted.','items':[]})]
        with patch('insights.pipeline.get_client_and_model',return_value=(client,'test')):
            result = execute_step(extract(),{'Feedback':'Merci xyz'}, {})
        self.assertEqual(result['status'],'review')
        self.assertEqual(api.call_count,2)
        self.assertNotIn('items',result)

    def test_ambiguous_referent_can_still_yield_a_verbatim_question(self):
        text = 'Maladie Yangon kombo nayango nini'
        with self.client.get('/api/pipeline/presets/drc-ebola-framework') as template:
            node = template.get_json()['nodes'][1]
        gloss = 'What is this disease called?'
        item = {'translation_en': gloss, 'feedback_excerpt': text, 'type_of_feedback': 'Questions'}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.side_effect = [
            completion({'status': 'review', 'reason': 'The reference is ambiguous.', 'items': []}),
            completion({'status': 'ok', 'reason': '', 'items': [item]}),
        ]
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            response = self.client.post('/api/pipeline/step', json={'node': node, 'record': {'Feedback': text, 'source_translation_en': gloss}})
        self.assertEqual(response.get_json()['items'], [item])
        self.assertEqual(response.get_json()['attempts'], 2)

    def test_drc_action_review_reconsiders_operational_consequence(self):
        with self.client.get('/api/pipeline/presets/drc-ebola-framework') as template:
            node = next(n for n in template.get_json()['nodes'] if n['name'] == 'feedback_action')
        record = {'feedback_excerpt': 'Kama una cikia ile utapata dawa',
                  'type_of_feedback': 'Observations, perceptions ou croyances',
                  'translation_en': 'If you hear that, you will get medicine.',
                  'feedback_path': 'Croyances concernant la disponibilité des traitements approuvés'}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        common = {'plainMeaning': 'If you hear that, you will get medicine.'}
        api.side_effect = [
            completion({**common, 'status': 'review', 'reason': 'Belief or reassurance?', 'values': {'feedback_action': None}}),
            completion({**common, 'status': 'ok', 'reason': '', 'values': {'feedback_action': 'Document and monitor trend'}}),
        ]
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            result = self.client.post('/api/pipeline/step', json={'node': node, 'record': record}).get_json()
        self.assertEqual(result['items'], [{'feedback_action': 'Document and monitor trend'}])
        schema = api.call_args.kwargs['response_format']['json_schema']['schema']
        self.assertEqual(schema['required'][0], 'plainMeaning')
        retry = json.loads(api.call_args.kwargs['messages'][1]['content'])
        self.assertIn('same first action', retry['reviewCheck'])

    def test_no_code_path_rechecks_even_if_model_said_ok_and_keeps_uncertain_candidate_in_review(self):
        with self.client.get('/api/pipeline/presets/drc-ebola-framework') as template:
            node = next(n for n in template.get_json()['nodes'] if n['name'] == 'feedback_path')
        record = {'feedback_excerpt': 'Tondimi maso bolobi', 'translation_en': 'We accepted what you said.',
                  'type_of_feedback': 'Appréciations'}
        no_code = node['reviewValues'][0]
        path = 'Appréciation concernant la réponse > Appréciation concernant la réponse > Appréciations des informations reçues'
        for candidate in (path, None):
            with self.subTest(candidate=candidate):
                client = MagicMock()
                api = client.with_options.return_value.chat.completions.create
                api.side_effect = [
                    completion({'status': 'ok', 'reason': '', 'values': {'feedback_path': no_code}}),
                    completion({'status': 'review', 'reason': 'Possible acknowledgment', 'values': {'feedback_path': candidate}})]
                with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
                    result = self.client.post('/api/pipeline/step', json={'node': node, 'record': record}).get_json()
                self.assertEqual(result['status'], 'review')
                self.assertEqual(result.get('drafts', {}).get('feedback_path'), candidate)
                schema = api.call_args.kwargs['response_format']['json_schema']['schema']
                choices = schema['properties']['values']['properties']['feedback_path']['enum']
                self.assertIn(None, choices)
                self.assertIn(path, choices)
                self.assertNotIn(no_code, choices)

    def test_drc_path_receives_translation_and_rechecks_within_assigned_type(self):
        with self.client.get('/api/pipeline/presets/drc-ebola-framework') as template:
            node = next(n for n in template.get_json()['nodes'] if n['name'] == 'feedback_path')
        for excerpt, gloss, kind, path in [
            ('Maladie Yangon kombo nayango nini', 'What is this disease called?', 'Questions',
             'Questions sur la maladie > Autres > Autres questions sur la maladie'),
            ("Tuna sikiya basi n'a maneno yenu", 'We just hear your words', 'Observations, perceptions ou croyances',
             'Observations sur la réponse > Information > Observations concernant l’information sur la réponse'),
        ]:
            with self.subTest(excerpt=excerpt):
                record = {'feedback_excerpt': excerpt, 'translation_en': gloss, 'type_of_feedback': kind}
                client = MagicMock()
                api = client.with_options.return_value.chat.completions.create
                api.side_effect = [
                    completion({'plainMeaning': gloss, 'status': 'review', 'reason': 'Unclear tone.', 'values': {'feedback_path': None}}),
                    completion({'plainMeaning': gloss, 'status': 'ok', 'reason': '', 'values': {'feedback_path': path}}),
                ]
                with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
                    result = self.client.post('/api/pipeline/step', json={'node': node, 'record': record}).get_json()
                self.assertEqual(result['items'], [{'feedback_path': path}])
                retry = json.loads(api.call_args.kwargs['messages'][1]['content'])
                self.assertEqual(retry['record'], record)
                self.assertIn('available paths', retry['reviewCheck'])

    def test_presets_folder_is_listed_and_served_by_id(self):
        listing = self.client.get('/api/pipeline/presets').get_json()['presets']
        ids = [p['id'] for p in listing]
        self.assertIn('drc-ebola-framework', ids)
        self.assertIn('drc-feedback-messages', ids)
        self.assertTrue(all(p['name'] and p['description'] and p['steps'] for p in listing))
        for bad in ('unknown', '..%2Fapp', 'drc-ebola-framework.pipeline'):
            self.assertEqual(self.client.get('/api/pipeline/presets/' + bad).status_code, 404)

    def test_drc_wording_alone_no_longer_switches_on_hidden_behaviour(self):
        # Version 1 matched these words and field names in code. Behaviour now comes only from the workflow file.
        node = {**extract(), 'instructions': 'Feedback from the DRC in Congolese Swahili.',
                'fields': [{'name': 'excerpt', 'options': []}, {'name': 'type', 'options': ['Thanks']}]}
        path = {'id': 'p', 'name': 'path', 'type': 'classify', 'inputs': ['excerpt', 'translation_en'], 'provider': 'azure',
                'instructions': 'In the DRC.', 'fields': [{'name': 'feedback_path', 'options': ['A', 'B']}]}
        for step, reply in ((node, {'status': 'ok', 'reason': '', 'items': [{'excerpt': 'Merci', 'type': 'Thanks'}]}),
                            (path, {'status': 'ok', 'reason': '', 'values': {'feedback_path': 'A'}})):
            client = MagicMock()
            api = client.with_options.return_value.chat.completions.create
            api.return_value = completion(reply)
            with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
                execute_step(step, {k: 'Merci' for k in step['inputs']}, {})
            system = api.call_args.kwargs['messages'][0]['content']
            schema = api.call_args.kwargs['response_format']['json_schema']['schema']
            self.assertNotIn('oral community feedback', system)
            self.assertNotIn('plainMeaning', schema['properties'])

    def test_workflow_context_and_step_guidance_reach_the_model(self):
        node = {**extract(), 'systemNotes': 'Step note.', 'interpretFirst': True}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.return_value = completion({'plainMeaning': 'Thanks', 'status': 'ok', 'reason': '', 'items': [{'excerpt': 'Merci', 'type': 'Thanks'}]})
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            self.client.post('/api/pipeline/step', json={'node': node, 'record': {'Feedback': 'Merci'}, 'context': 'Setting note.'})
            system = api.call_args.kwargs['messages'][0]['content']
            self.assertTrue(system.endswith('Setting note. Step note. '))
            self.assertEqual(api.call_args.kwargs['response_format']['json_schema']['schema']['required'][0], 'plainMeaning')
            self.client.post('/api/pipeline/step', json={'node': {**node, 'useContext': False}, 'record': {'Feedback': 'Merci'}, 'context': 'Setting note.'})
            self.assertNotIn('Setting note.', api.call_args.kwargs['messages'][0]['content'])
        self.assertEqual(self.client.post('/api/pipeline/step', json={'node': node, 'record': {}, 'context': 5}).status_code, 400)
        for bad in ({'useContext': 'yes'}, {'systemNotes': 3}, {'fallbackOption': 'X'}):
            self.assertEqual(self.client.post('/api/pipeline/validate', json={'nodes': [{**node, **bad}]}).status_code, 400)

    def test_fallback_option_works_for_any_column_name(self):
        node = {'id': 'c', 'name': 'Topic', 'type': 'classify', 'inputs': ['Feedback'], 'provider': 'azure', 'instructions': 'Pick a topic.',
                'fields': [{'name': 'topic', 'options': ['Water', 'Health', 'None fits']}], 'fallbackOption': 'None fits',
                'reviewValues': ['None fits'], 'recheck': 'Look again.'}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.side_effect = [completion({'status': 'ok', 'reason': '', 'values': {'topic': 'None fits'}}),
                           completion({'status': 'ok', 'reason': '', 'values': {'topic': 'Water'}})]
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            result = execute_step(node, {'Feedback': 'No tap'}, {})
        self.assertEqual(result['items'], [{'topic': 'Water'}])
        retry = api.call_args.kwargs
        self.assertEqual(json.loads(retry['messages'][1]['content'])['reviewCheck'], 'Look again.')
        self.assertNotIn('None fits', retry['response_format']['json_schema']['schema']['properties']['values']['properties']['topic']['enum'])

    def test_reviewer_note_is_sent_apart_from_the_feedback(self):
        node = {**extract(), 'type': 'analyze', 'fields': [{'name': 'english', 'options': []}]}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.return_value = completion({'status': 'ok', 'reason': '', 'values': {'english': 'We accept your words.'}})
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            self.client.post('/api/pipeline/step', json={'node': node, 'record': {'Feedback': 'Tondimi maso bolobi'}, 'note': ' maso = words '})
            system, user = api.call_args.kwargs['messages']
            payload = json.loads(user['content'])
            self.assertEqual(payload['reviewerNote'], 'maso = words')
            self.assertEqual(payload['record'], {'Feedback': 'Tondimi maso bolobi'})
            self.assertIn('not as part of the feedback', system['content'])
            self.client.post('/api/pipeline/step', json={'node': node, 'record': {'Feedback': 'Tondimi maso bolobi'}})
            system, user = api.call_args.kwargs['messages']
            self.assertNotIn('reviewerNote', json.loads(user['content']))
            self.assertNotIn('reviewer', system['content'])
        self.assertEqual(self.client.post('/api/pipeline/step', json={'node': node, 'record': {}, 'note': 'x' * 5001}).status_code, 400)

    def test_model_may_skip_only_when_the_step_allows_it(self):
        node = {**extract(), 'type': 'analyze', 'fields': [{'name': 'english', 'options': []}]}
        client = MagicMock()
        api = client.with_options.return_value.chat.completions.create
        api.return_value = completion({'status': 'skip', 'reason': 'Random keystrokes', 'values': {'english': None}})
        with patch('insights.pipeline.get_client_and_model', return_value=(client, 'test')):
            result = execute_step({**node, 'allowSkip': True}, {'Feedback': 'kdjfa'}, {})
            self.assertEqual(result['status'], 'skipped')
            self.assertEqual(api.call_count, 1)
            schema = api.call_args.kwargs['response_format']['json_schema']['schema']
            self.assertEqual(schema['properties']['status']['enum'], ['ok', 'review', 'skip'])
            self.assertIn('never skip it', api.call_args.kwargs['messages'][0]['content'])
            api.reset_mock()
            # Without allowSkip, "skip" is not a valid answer: the record goes to review instead.
            result = self.client.post('/api/pipeline/step', json={'node': node, 'record': {'Feedback': 'kdjfa'}}).get_json()
            self.assertEqual(result['status'], 'review')
            schema = api.call_args.kwargs['response_format']['json_schema']['schema']
            self.assertEqual(schema['properties']['status']['enum'], ['ok', 'review'])
            self.assertNotIn('skip', api.call_args.kwargs['messages'][0]['content'])
        self.assertEqual(self.client.post('/api/pipeline/validate', json={'nodes': [{**node, 'allowSkip': 'yes'}]}).status_code, 400)

    def test_pipeline_assets_revalidate_and_entry_module_is_versioned(self):
        page = self.client.get('/pipeline')
        self.assertIn('must-revalidate',page.headers['Cache-Control'])
        self.assertRegex(page.get_data(as_text=True),r'pipeline\.js\?v=[a-f0-9]{12}')
        asset = self.client.get('/static/js/pipeline-core.js?v=4')
        self.assertEqual(asset.status_code,200)
        self.assertIn('no-cache',asset.headers['Cache-Control'])
        asset.close()
        page.close()
