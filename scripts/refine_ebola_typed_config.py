"""Build the second revision after reviewing the first type-routed sample.

No API calls. Preserve all workbook labels and the previous fast configuration.
"""
import json
from pathlib import Path

from build_ebola_typed_config import ROOT, FALLBACK, build_config


# Clarify boundaries without changing any label from the workbook.
CODE_NOTES = {
    'Appréciation concernant la réponse actors':
        'General thanks, blessings or encouragement addressed to responders, with no named profession or activity. '
        'Includes Nous vous remercions, Que Dieu vous bénisse, Merci et courage, Courage pour le travail. '
        'Do not assume the recipient is a health worker.',
    'Appréciations ou encouragements envers les agents de santé':
        'Thanks or encouragement explicitly addressed to health workers. Generic merci or courage alone does not identify health workers.',
    'Appréciations des informations reçues':
        'Thanks for sensitisation, messages, explanations or information, including encouragement specifically for sensitisation.',
    'Appréciations de l’engagement communautaire':
        'Praise specifically for community participation, consultation or involvement; ordinary thanks for information belongs to informations reçues.',
    'Croyances concernant la non-existence de la maladie':
        'Bare denial of the disease. When denial is explained by profit, a plot, foreign origin or supernatural origin, prefer that specific explanatory belief.',
    'Croyances concernant l’existence de la maladie':
        'An assertion that the disease exists. Do not convert doubt, denial, reported rumours or a supernatural explanation into simple affirmation.',
    'Croyances selon lesquelles l’épidémie fait partie d’un complot':
        'A plot, fabricated scheme or cooperation of authorities/humanitarians. Prefer the money-making code if funding or profit is the explanation.',
    'Croyances selon lesquelles certaines personnes/organisations gagnent de l’argent à cause de l’épidémie':
        'A claim that the outbreak was invented or used to obtain money, funding or profit. Includes denial justified by financial gain.',
    'Croyances selon lesquelles l’épidémie a été apportée par des étrangers':
        'Disease attributed to foreigners or outsiders, including claims it was created by white people. Prefer this explanation over a bare denial when both occur together.',
    'Croyances selon lesquelles l’épidémie est d’origine surnaturelle':
        'Religious, scriptural, spiritual or supernatural explanations of the outbreak, including fulfilment of scripture.',
    'Croyances concernant la gravité de la maladie':
        'Severity or lethality of the disease itself. Fear of being killed or harmed at a treatment centre concerns centre safety instead.',
    'Observations ou croyances concernant la sécurité des centres de traitement':
        'Fear that entering a CTE is dangerous or a death sentence. The danger is attributed to the centre, not just disease severity.',
    'Croyances selon lesquelles les structures de santé ne sont pas sûres':
        'Fear of harm, unsafe care, forced quarantine or being misdiagnosed in a hospital. Distinguish this from cost, distance or service availability.',
    'Croyances concernant l’accès aux soins de santé':
        'Access barriers such as cost, distance, admission or availability. Prefer unsafe-facility beliefs when fear of harm or forced quarantine explains avoidance.',
    'Croyances concernant la disponibilité des traitements approuvés':
        'A claim that treatment or medicines are or are not available, including a belief clause embedded in a question about cure.',
    'Demande d’autres équipements de prévention':
        'Supply soap, disinfectant, hand-washing kits, masks or protective equipment. In a request for water AND soap, soap has this specific code; do not default to Autre demande.',
    'Demande de fournir un traitement approuvé':
        'Supply medicines or approved treatment. In a request for water AND medicines, medicines have this specific code.',
    'Demande de renforcer les MHP (mesures d’hygiène et de prévention)':
        'Install hand-washing stations in public places or strengthen/enforce preventive measures; distinguish installation/enforcement from simply supplying kits.',
    'Autre demande':
        'A request with no more specific code in this workbook, such as water supply alone. If another requested item has a specific code, prefer that code.',
    'Questions sur la disponibilité des traitements approuvés':
        'Whether approved medicines or treatments exist or are available. Do not confuse availability with questions about efficacy.',
    'Plaintes concernant l’information reçue':
        'A complaint about information received. Difficulty delivering messages because of distrust is not automatically a complaint about received information; review if no reporting code fits.',
}

PATH_INSTRUCTIONS = (
    'Code only the clause or implied speech act in state.Feedback matching the authoritative human-assigned '
    'state.type_of_feedback. A question can imply a request; do not reject the assigned Type just because of grammar. '
    'Options already belong to that Type. Never use a different Type\'s clause. '
    'Choose the most specific matching code, keeping its exact full path. For several same-Type subjects, '
    'prefer a specific code over a generic catch-all; if equally specific, use the first expressed subject. '
    'For a belief plus its explanation, code the explanatory belief: financial gain before conspiracy, '
    'conspiracy before unspecified foreign origin, and explicit foreign or supernatural origin before bare existence/denial. '
    'An assertion of existence followed by a religious explanation is not merely an existence claim. '
    'General thanks, blessings and encouragement assigned to Appréciations are meaningful appreciation: '
    'use the general responder appreciation code unless a specific actor/activity is named. '
    'Do not assume an unnamed recipient is a health worker. '
    'Use an existing generic code when it accurately represents meaningful feedback outside the specific codes. '
    f'Choose "{FALLBACK}" only for absent/uninterpretable matching content or when even the generic codes do not fit. '
    'A reported belief is not a verified fact. Read the option descriptions to distinguish close codes.'
)

ASSESSMENT_SCOPE = (
    'record.type_of_feedback is authoritative. decisions.feedback_path is the assigned codebook path. '
    'In record.Feedback, assess ONLY the clause or implied speech act supporting that path and that Type. '
    'If several clauses support it, use the first. Do not select another subject or import another clause\'s symptoms, '
    'allegation, emotion or request. Use surrounding words only to resolve references. '
    'The assigned path may be uncertain: treat it as provisional evidence, not proof. '
)

ASSESSMENTS = {
    'feedback_sensitivity': (
        'Assess confidentiality/handling sensitivity, not urgency. General or hypothetical health questions are Not sensitive, '
        'even if they say your child, pregnancy or bleeding. An identifiable real person\'s current medical situation is '
        'Highly sensitive / confidential only when it belongs to this selected clause. '
        'Routine supply requests, generic thanks and impersonal disease beliefs are Not sensitive unless the selected '
        'clause also contains a concrete protection, misconduct, stigma or safety concern.'),
    'feedback_criticality': (
        'Assess operational urgency. Praise and routine suggestions: Low. Clear questions, requests, denial or '
        'conspiracy beliefs needing a response: at least Medium. Serious refusal, healthcare avoidance, disruption '
        'or harmful mistrust: High. A specific real person currently symptomatic, an unreported death, safeguarding '
        'or imminent danger in this clause: Critical / Urgent. A general/hypothetical question is not a current case; '
        'si une personne saigne and si ton enfant est atteint alone do not describe an actual patient. '
        'Never choose Low merely because a conspiracy claim is phrased calmly.'),
    'feedback_status': (
        'This is content/handling status, not model confidence. New means an interpretable, relevant clause has a '
        'usable specific code, including a short request for soap or an interpretable blessing. '
        'Needs review means no matching clause, uninterpretable content, interrupted/withdrawn/blocked interaction, '
        'unrelated topic/language, the no-code fallback, or a codebook gap represented by an Autre/Autres catch-all. '
        'Inspect decisions.feedback_path: a water-only request assigned Autre demande needs review for a codebook gap; '
        'a clear soap request with a prevention-equipment code is New. Generic responder appreciation is a valid code, '
        'not a no-code fallback. Uncertainty between specific codes alone does not change New to Needs review.'),
    'feedback_action': (
        'Choose the first operational action, not an action already completed. Pure appreciation: No action. '
        'A straightforward or hypothetical question: Provide information / response. '
        'Denial, conspiracy, false-cure or harmful mistrust beliefs: Adapt RCCE messages (including a calmly phrased '
        'cooperation des autorités claim). Supply requests: Refer to responsible sector. '
        'For a real current symptomatic person or unreported death in this selected clause: Trigger public health alert. '
        'If actuality is ambiguous: Individual follow-up. General/hypothetical symptoms are not grounds for an alert. '
        'Threats, violence, SEA or safeguarding in this clause: Immediate escalation. '
        'Do not alert or escalate because of another clause.'),
    'feedback_sector': (
        'Choose the sector responsible for the selected coded issue, not the delivery method or strongest unrelated '
        'clause. Water, soap, aquatabs, disinfectant and hand-washing equipment/stations: WASH. Medicines, treatment '
        'availability and access to clinical care: Health or Clinical care. Vaccines: Vaccination. '
        'General disease information, rumours, beliefs and unspecified appreciation in this community-feedback '
        'workflow: RCCE/CEA. Actor roles/governance: Coordination. MHPSS requires explicit psychosocial support, '
        'not merely courage or thanks. Logistics requires an explicit transport/storage/supply-chain issue; '
        'amener or fournir alone does not make a health/WASH request Logistics. '
        'A real suspected case or death requiring an alert in this selected clause: Surveillance.'),
    'feedback_sentiment': (
        'Positive: thanks, blessing, appreciation or encouragement. Negative: fear, dissatisfaction, mistrust, '
        'explicit denial, conspiracy or accusation, even calmly phrased. Neutral: ordinary information questions, '
        'requests or factual observations without emotion. Do not label simple affirmation of disease existence or '
        'religious interpretation Negative solely because it is a belief. Mixed only if both positive and negative '
        'tones occur within this selected clause; Unclear if insufficient evidence. Never borrow another Type\'s tone.'),
}


def refine(config):
    path = config['questions'][0]
    path['instructions'] = PATH_INSTRUCTIONS
    for branch in path['sourceBranches']:
        branch['options'] = [
            {'label': label, 'description': label + '. ' + CODE_NOTES[label.split(' > ')[-1]]}
            if label.split(' > ')[-1] in CODE_NOTES else label for label in branch['options']]
    for question in config['questions'][1:]:
        question['dependsOn'] = ['feedback_path']
        question['allowReviewDependencies'] = True
        question['instructions'] = ASSESSMENT_SCOPE + ASSESSMENTS[question['outputColumnName']]
    config['exportedAt'] = '2026-10-08T00:00:00Z'
    return config


def main():
    original = json.loads((ROOT / 'sample_jev/ebola-feedback-coding.config_v19 (1).json').read_text(encoding='utf-8'))
    config = refine(build_config(ROOT / 'sample_jev/Ebola_codes_avec_sous-codes (1).xlsx', original))
    output = ROOT / 'sample_jev/ebola-feedback-coding.refined.config.json'
    output.write_text(json.dumps(config, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(f'Wrote {output}')


if __name__ == '__main__':
    main()
