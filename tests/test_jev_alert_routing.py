"""Offline checks for the configured alert-routing rule, not model accuracy."""
import copy
import json
import unittest
from pathlib import Path
from unittest.mock import Mock

from insights.jev_workflow import evaluate_row, validate_workflow
from jev_provider import build_question


class AlertRoutingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = Path(__file__).resolve().parents[1] / 'sample_jev/ebola-feedback-coding.excerpt.config.json'
        if not path.exists():
            raise unittest.SkipTest('Local Ebola sample configuration not present')
        cls.config = json.loads(path.read_text())
        cls.routing = next(d for d in cls.config['derived'] if d['outputColumnName'] == 'feedback_sector')
        cls.questions = [copy.deepcopy(next(q for q in cls.config['questions'] if q['outputColumnName'] == name))
                         for name in cls.routing['inputs']]
        for question in cls.questions:
            question.pop('dependsOn')
            question.pop('allowReviewDependencies')
        cls.decoders = [build_question({**q, 'type': q['questionType']})[1] for q in cls.questions]

    def provider(self, action, sector, confidences=(.95, .95), omit=None):
        def call(state, questions):
            answers = {}
            for i, value in enumerate((action, sector)):
                if i == omit:
                    continue
                key = next(k for k, label in self.decoders[i]['labels'].items() if label == value)
                answers[f'q{i}'] = {'type': 'choice', 'choice': key, 'confidence': confidences[i],
                                   'probabilities': {k: int(k == key) for k in questions[f'q{i}']['criteria']}}
            return {'answers': answers}
        return Mock(side_effect=call)

    def test_every_alert_routes_to_surveillance_and_other_actions_keep_suggestion(self):
        validate_workflow(self.config['questions'], self.config['derived'])
        validate_workflow(self.questions, [self.routing])
        for action in self.decoders[0]['labels'].values():
            for sector in self.decoders[1]['labels'].values():
                with self.subTest(action=action, sector=sector):
                    call = self.provider(action, sector)
                    records, calls = evaluate_row({'Feedback': 'Example'}, self.questions, [self.routing], call)
                    self.assertEqual(records['feedback_sector']['value'],
                                     'Surveillance' if action == 'Trigger public health alert' else sector)
                    self.assertEqual(records['feedback_sector_suggested']['value'], sector)
                    self.assertEqual(records['feedback_sector']['status'], 'ok')
                    self.assertNotIn('confidence', records['feedback_sector'])
                    self.assertEqual(len(calls), 1)

    def test_r03_routing_retains_uncertainty_raw_suggestion_and_retry_behavior(self):
        for confidences in ((.43, .36), (.43, .95), (.95, .36)):
            call = self.provider('Trigger public health alert', 'Clinical care', confidences)
            records, _ = evaluate_row({'Feedback': 'Mon enfant souffre d’hémorragie'},
                                      self.questions, [self.routing], call)
            self.assertEqual(records['feedback_sector']['value'], 'Surveillance')
            self.assertEqual(records['feedback_sector']['status'], 'review')
            suggested = records['feedback_sector_suggested']
            self.assertEqual(suggested['value'], 'Clinical care')
            self.assertEqual(suggested['confidence'], confidences[1])
            self.assertIn('raw', suggested)
            retry = Mock()
            retried, _ = evaluate_row({'Feedback': 'Mon enfant souffre d’hémorragie'},
                                     self.questions, [self.routing], retry, records)
            retry.assert_not_called()
            self.assertEqual(retried['feedback_sector'], records['feedback_sector'])

    def test_missing_input_cannot_produce_a_certified_sector(self):
        for omit in (0, 1):
            records, _ = evaluate_row({'Feedback': 'Example'}, self.questions, [self.routing],
                                      self.provider('Trigger public health alert', 'Clinical care', omit=omit))
            self.assertEqual(records['feedback_sector']['status'], 'blocked')
            self.assertIsNone(records['feedback_sector']['value'])


if __name__ == '__main__':
    unittest.main()
