"""Build an Ebola workflow that isolates one source excerpt before coding."""
import json

from build_ebola_typed_config import ROOT, build_config
from refine_ebola_typed_config import refine


def build():
    original = json.loads((ROOT / 'sample_jev/ebola-feedback-coding.config_v19 (1).json').read_text(encoding='utf-8'))
    config = refine(build_config(ROOT / 'sample_jev/Ebola_codes_avec_sous-codes (1).xlsx', original))
    excerpt = {
        'outputColumnName': 'feedback_excerpt', 'questionType': 'choice', 'options': [],
        'instructions': (
            'Select exactly one supplied excerpt expressing state.type_of_feedback, the authoritative human-assigned '
            'speech act. Questions: the question being asked; Demandes: an explicit or implied request; '
            'Suggestions: proposed improvement; Appréciations: thanks, blessing or encouragement; '
            'Observations, perceptions ou croyances: the observation, belief or fear; '
            'Signalements et préoccupations: the reported problem; Allégations et Signalement d’incident: the allegation/incident. '
            'A rhetorical question can imply a request or concern. Choose the first matching independent speech act '
            'if several match. Do not select an urgent report for a row assigned to a request, question or appreciation '
            'merely because that report is urgent. A causal explanation within a candidate belongs to that candidate. '
            'Select no_match if none matches unambiguously or essential context is missing. '
            'The chosen excerpt alone will be coded; do not invent, combine or rewrite text.'),
        'sourceExcerpt': {'column': 'Feedback', 'separators': [';', '?', '!', '\n', '.'],
                          'splitBefore': ['et vous', 'et nous', 'et je', 'et on', 'et cette',
                                          'mais vous', 'mais nous', 'mais je', 'mais on', 'mais cette',
                                          'et merci', 'mais merci', 'et courage', 'mais courage',
                                          'et veuillez', 'et prière', 'et pourquoi', 'et comment']},
        'review': {'minConfidence': 0.7},
    }
    path = config['questions'][0]
    path['dependsOn'] = ['feedback_excerpt']
    path['allowReviewDependencies'] = True
    path['instructions'] = (
        'record.Feedback is the selected verbatim excerpt, not the full original row. Code only this evidence for '
        'record.type_of_feedback. Do not infer unseen clauses or actors. ' +
        path['instructions'].replace('state.Feedback', 'record.Feedback').replace('state.type_of_feedback', 'record.type_of_feedback'))
    notes = {
        'Questions sur le rôle de la Croix-Rouge':
            'The question is about the actor\'s role or responsibility itself. Mentioning Croix-Rouge while asking why '
            'aquatabs or other prevention items are no longer distributed is an assistance/item-supply question, not an actor-role question.',
        'Questions sur la fourniture d’autres articles de prévention des maladies':
            'Questions about providing or stopping provision of prevention items, including aquatabs, soap or disinfectant, '
            'even when the provider such as Croix-Rouge is named.',
        'Questions sur l’évolution de l’épidémie':
            'Current case numbers, confirmed cases, geographical spread and changes in the outbreak. '
            'Combien de cas confirmés à Goma is a case-count question, not an Autre question.',
        'Questions sur les voyages internationaux':
            'Explicit cross-border or international travel. A destination alone does not establish a border crossing; '
            'travel between Goma and Bunia is domestic within DRC, not international.',
        'Croyances selon lesquelles l’épidémie a été apportée par des étrangers':
            'An explicit belief that outsiders brought or created the outbreak. A foreign doctor giving a vaccine '
            'associated with adverse effects is a vaccine-effects belief, not an assertion that foreigners brought the epidemic.',
        'Observations ou croyances concernant les effets secondaires du vaccin':
            'Reported problems or adverse effects attributed to vaccination, including when the vaccinator\'s nationality is mentioned.',
        'Demande de renforcer les MHP (mesures d’hygiène et de prévention)':
            'Explicit installation of hand-washing stations or strengthening/enforcing hygiene measures. '
            'Simply asking to be brought a hand-washing device is equipment supply. Water supply or public standpipes alone '
            'are not explicitly hygiene enforcement and use the workbook\'s generic request code.',
        'Demande d’autres équipements de prévention':
            'Supply soap, disinfectant, hand-washing kits/devices or prevention equipment. '
            'Amener des lave-mains asks for equipment; placer/installer stations in public places asks for installation.',
    }
    for branch in path['sourceBranches']:
        for i, option in enumerate(branch['options']):
            label = option['label'] if isinstance(option, dict) else option
            code = label.split(' > ')[-1]
            if code in notes:
                branch['options'][i] = {'label': label, 'description': label + '. ' + notes[code]}
    for question in config['questions'][1:]:
        # These output fields are categories, not numerical averages. A Score
        # response can round an uncertain mixture into an intermediate class.
        if question['outputColumnName'] in ('feedback_sensitivity', 'feedback_criticality'):
            question['questionType'] = 'choice'
        question['instructions'] = (
            'record.Feedback contains ONLY the selected excerpt. decisions.feedback_path is its provisional code. '
            'Use only the evidence in this excerpt. ' + question['instructions'])
        if question['outputColumnName'] == 'feedback_sensitivity':
            descriptions = [
                'Routine information, disease prevention, ordinary supply requests, praise, impersonal beliefs or hypothetical '
                'health questions without a concrete confidentiality, misconduct or personal-safety concern. '
                'Se protéger against infection or matériels de protection does not by itself mean humanitarian Protection or sensitive content.',
                'A concrete allegation of misconduct, discrimination, stigma, coercion, violence, conflict or personal-safety '
                'concern in this excerpt. General disease prevention is not this kind of protection concern.',
                'An actual identifiable person\'s medical information, targeted threat, safeguarding or SEA, violence or '
                'serious identifiable allegation in this excerpt. A hypothetical child/person in a general question is not an actual patient.',
            ]
            for option, description in zip(question['options'], descriptions):
                option['description'] = description
        if question['outputColumnName'] == 'feedback_sector':
            descriptions = {
                'Protection': 'Humanitarian protection from violence, exploitation, abuse, discrimination or rights violations. '
                              'Not general protection against infection; se protéger d\'Ebola alone is health prevention.',
                'WASH': 'Water, sanitation, hygiene, aquatabs, soap, disinfectant and hand-washing equipment. '
                        'If the selected path codes medicines within a mixed request, use Health for that coded subject.',
                'Health': 'Medical treatment, medicines, health services and clinical information. '
                          'An approved-treatment supply code belongs here even if the same excerpt also mentions water.',
                'RCCE/CEA': 'Community information, prevention advice, rumours and risk communication, including how '
                            'drivers can avoid infection. Not humanitarian Protection just because the text says se protéger.',
            }
            question['options'] = [{'label': label, 'description': descriptions[label]} if label in descriptions else label
                                   for label in question['options']]
    config['questions'].insert(0, excerpt)
    config['outputOrder'].insert(0, 'feedback_excerpt')
    # Enforce the rubric's alert routing after inference, keeping the original
    # suggestion and its confidence visible. The lookup propagates input review
    # flags and does not invent a confidence for the rule-derived final sector.
    sector = next(q for q in config['questions'] if q['outputColumnName'] == 'feedback_sector')
    action = next(q for q in config['questions'] if q['outputColumnName'] == 'feedback_action')
    sector['outputColumnName'] = 'feedback_sector_suggested'
    labels = lambda q: [o['label'] if isinstance(o, dict) else o for o in q['options']]
    config['derived'].append({
        'outputColumnName': 'feedback_sector', 'type': 'lookup',
        'inputs': ['feedback_action', 'feedback_sector_suggested'],
        'table': [
            {'when': {'feedback_action': a, 'feedback_sector_suggested': s},
             'value': 'Surveillance' if a == 'Trigger public health alert' else s}
            for a in labels(action) for s in labels(sector)
        ],
    })
    config['outputOrder'].insert(config['outputOrder'].index('feedback_sector') + 1,
                                 'feedback_sector_suggested')
    config['report']['evidenceColumns'] = ['Feedback', 'type_of_feedback', 'feedback_excerpt']
    return config


if __name__ == '__main__':
    output = ROOT / 'sample_jev/ebola-feedback-coding.excerpt.config.json'
    output.write_text(json.dumps(build(), ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Wrote {output}')
