import copy
import unittest
from unittest.mock import Mock

from insights.jev_excerpt import excerpt_candidates
from insights.jev_workflow import evaluate_row, validate_workflow
from jev_provider import JevConfigError, JevError


SPEC = {'column': 'Feedback', 'separators': [';', '?', '!', '\n', '.'],
        'splitBefore': ['et vous', 'et nous', 'mais nous']}


def choice(question, key=None, confidence=.95):
    keys = list(question['criteria'])
    key = key or keys[0]
    return {'type': 'choice', 'choice': key, 'confidence': confidence,
            'probabilities': {k: int(k == key) for k in keys}}


class ExcerptTests(unittest.TestCase):
    def setUp(self):
        self.configs = [
            {'outputColumnName': 'Excerpt', 'questionType': 'choice', 'instructions': 'Select the assigned type.',
             'options': [], 'sourceExcerpt': SPEC, 'review': {'minConfidence': .7}},
            {'outputColumnName': 'Code', 'questionType': 'choice', 'instructions': 'Code the excerpt.',
             'dependsOn': ['Excerpt'], 'allowReviewDependencies': True,
             'sourceBranches': [{'when': {'type_of_feedback': t}, 'options': ['A', 'B']}
                                for t in ['Demandes', 'Appréciations', 'Signalements et préoccupations']]},
            {'outputColumnName': 'Action', 'questionType': 'choice', 'instructions': 'Assess the coded excerpt.',
             'dependsOn': ['Code'], 'allowReviewDependencies': True, 'options': ['Act', 'Wait']},
        ]
        validate_workflow(self.configs)

    def test_exact_spans_preserve_accents_punctuation_and_decimal(self):
        text = "  Fièvre à 38.5 degrés;Ebola n'existe pas et vous avez du courage pour la sensibilisation  "
        spans = excerpt_candidates(text, SPEC)
        self.assertEqual([s['text'] for s in spans], ['Fièvre à 38.5 degrés;', "Ebola n'existe pas", 'et vous avez du courage pour la sensibilisation'])
        for span in spans:
            self.assertEqual(text[span['start']:span['end']], span['text'])
        self.assertEqual(excerpt_candidates('De l’eau et du savon', SPEC)[0]['text'], 'De l’eau et du savon')

    def test_unrelated_symptoms_and_denial_are_absent_from_all_later_calls(self):
        cases = [
            ('Mon enfant souffre d’hémorragie; cette maladie n’existe pas; nous demandons aux autorités de faire respecter les mesures barrières aux chauffeurs de bus',
             'Demandes', 'nous demandons aux autorités de faire respecter les mesures barrières aux chauffeurs de bus', ['hémorragie', 'n’existe pas']),
            ("Ebola n'existe pas et vous avez du courage pour la sensibilisation", 'Appréciations',
             'et vous avez du courage pour la sensibilisation', ["n'existe pas"]),
        ]
        for text, kind, selected, forbidden in cases:
            with self.subTest(kind=kind):
                states = []
                def call(state, questions):
                    states.append(copy.deepcopy(state))
                    if 'q0' in questions:
                        key = next(k for k, v in questions['q0']['criteria'].items() if v == selected)
                        return {'answers': {'q0': choice(questions['q0'], key, confidence=.4)}}
                    self.assertEqual(state['record']['Feedback'], selected)
                    self.assertEqual(state['record']['type_of_feedback'], kind)
                    for word in forbidden:
                        self.assertNotIn(word, str(state))
                    return {'answers': {k: choice(q) for k, q in questions.items()}}
                source = {'Feedback': text, 'type_of_feedback': kind}
                records, calls = evaluate_row(source, self.configs, [], call)
                self.assertEqual(len(calls), 3)
                self.assertEqual(source['Feedback'], text)
                self.assertEqual(records['Excerpt']['value'], selected)
                a, b = records['Excerpt']['sourceSpan']
                self.assertEqual(text[a:b], selected)
                self.assertEqual(records['Code']['status'], 'review')
                self.assertEqual(records['Action']['status'], 'review')
                retry = Mock()
                reused, _ = evaluate_row(source, self.configs, [], retry, records)
                retry.assert_not_called()
                self.assertEqual(reused['Excerpt']['sourceSpan'], [a, b])
                self.assertEqual(reused['Action']['status'], 'review')

    def test_single_excerpt_skips_selection_call_without_claiming_model_confidence(self):
        call = Mock(side_effect=lambda state, qs: {'answers': {k: choice(q) for k, q in qs.items()}})
        records, calls = evaluate_row({'Feedback': 'Merci pour la sensibilisation', 'type_of_feedback': 'Appréciations'}, self.configs, [], call)
        self.assertEqual(len(calls), 2)
        self.assertNotIn('q0', call.call_args_list[0].args[1])
        self.assertEqual(records['Excerpt']['method'], 'single_excerpt')
        self.assertIsNone(records['Excerpt']['confidence'])
        self.assertNotIn('raw', records['Excerpt'])

    def test_selected_actual_medical_report_is_retained_not_suppressed(self):
        text = 'Mon enfant souffre d’hémorragie; veuillez nous donner du savon'
        selected = 'Mon enfant souffre d’hémorragie;'
        def call(state, qs):
            if 'q0' in qs:
                key = next(k for k, v in qs['q0']['criteria'].items() if v == selected)
                return {'answers': {'q0': choice(qs['q0'], key)}}
            self.assertEqual(state['record']['Feedback'], selected)
            self.assertNotIn('savon', str(state))
            return {'answers': {k: choice(q) for k, q in qs.items()}}
        records, calls = evaluate_row({'Feedback': text, 'type_of_feedback': 'Signalements et préoccupations'}, self.configs, [], call)
        self.assertEqual(len(calls), 3)
        self.assertEqual(records['Excerpt']['value'], selected)

    def test_no_match_and_provider_failure_block_every_later_question(self):
        source = {'Feedback': 'Clause one; clause two', 'type_of_feedback': 'Demandes'}
        call = Mock(side_effect=lambda state, qs: {'answers': {'q0': choice(qs['q0'], 'no_match')}})
        records, _ = evaluate_row(source, self.configs, [], call)
        self.assertEqual(call.call_count, 1)
        self.assertEqual(records['Excerpt']['status'], 'review')
        self.assertIsNone(records['Excerpt']['value'])
        self.assertEqual(records['Code']['status'], 'blocked')
        self.assertEqual(records['Action']['status'], 'blocked')
        call = Mock(side_effect=JevError('Unavailable'))
        records, _ = evaluate_row(source, self.configs, [], call)
        self.assertEqual(call.call_count, 1)
        self.assertEqual(records['Action']['status'], 'blocked')

    def test_empty_and_oversize_sources_do_not_trigger_unfocused_classification(self):
        for text in ['', ' ; ? ', ';'.join(f'Clause {i}' for i in range(255))]:
            call = Mock()
            records, _ = evaluate_row({'Feedback': text, 'type_of_feedback': 'Demandes'}, self.configs, [], call)
            call.assert_not_called()
            self.assertIsNone(records['Excerpt']['value'])
            self.assertIsNone(records['Action']['value'])

    def test_graph_cannot_bypass_the_excerpt(self):
        bad = copy.deepcopy(self.configs)
        bad[2].pop('dependsOn')
        bad[2].pop('allowReviewDependencies')
        with self.assertRaisesRegex(JevConfigError, 'depend on the source excerpt'):
            validate_workflow(bad)
        for spec in [None, {**SPEC, 'splitBefore': ['']}, {**SPEC, 'unexpected': True}]:
            bad = copy.deepcopy(self.configs)
            bad[0]['sourceExcerpt'] = spec
            with self.assertRaises(JevConfigError):
                validate_workflow(bad)


if __name__ == '__main__':
    unittest.main()
