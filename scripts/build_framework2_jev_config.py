"""Emit the Type x Topic Jev config."""
import json, datetime
from collections import defaultdict

import pathlib

# Paths are resolved against the repo root so this runs from any directory.
ROOT = pathlib.Path(__file__).resolve().parent.parent

MAX_CHOICE, MIN_SCORE, MAX_SCORE = 255, 2, 10
G = json.load(open(ROOT / 'framework2_topic_grid.json', encoding='utf-8'))
prev = {q['outputColumnName']: q for q in json.load(
    open(ROOT / 'examples/ebola-feedback-coding.jev-config.json', encoding='utf-8'))['questions']}

types, topics = G['types'], G['topics']
valid = defaultdict(list)
for k in G['grid']:
    t, tp = k.split('||')
    valid[t].append(tp)

# Spell out which topics exist for each type: 23 of the 84 cells have no
# sous-dimension in the framework, and the flat format cannot block them.
per_type = "\n".join(f"- {t} : {' | '.join(valid[t])}" for t in types)

CARRIED = ["feedback_sensibilite", "feedback_criticalite", "feedback_action",
           "feedback_secteur", "feedback_referent"]

questions = [
    {
        "outputColumnName": "framework2_type",
        "questionType": "choice",
        "instructions": (
            "Choisis le type de feedback dominant selon le Framework 2 (Ebola) — "
            "l'acte de langage, indépendamment du sujet. Un seul type. Si "
            "plusieurs sujets sont présents, choisis celui qui nécessite "
            "l'action la plus importante."),
        "options": types,
    },
    {
        "outputColumnName": "framework2_topic",
        "questionType": "choice",
        "instructions": (
            "Choisis le sujet dominant du feedback, indépendamment du type. "
            "Un seul sujet.\n\n"
            "Le sujet doit exister pour le type choisi. Sujets valides par "
            "type :\n\n" + per_type),
        "options": topics,
    },
] + [prev[n] for n in CARRIED]

cfg = {
    "format": "aidstack-insights-jev-config",
    "version": 1,
    "exportedAt": datetime.datetime.now(datetime.timezone.utc)
                    .isoformat(timespec='milliseconds').replace('+00:00', 'Z'),
    "sheetName": "Sheet1",
    "sourceColumns": ["Feedback"],
    "includeConfidence": True,
    "questions": questions,
}
for q in cfg["questions"]:
    n = len(q["options"])
    lo, hi = (MIN_SCORE, MAX_SCORE) if q["questionType"] == "score" else (2, MAX_CHOICE)
    assert lo <= n <= hi, (q["outputColumnName"], n)
    assert len(set(q["options"])) == n

out = ROOT / 'examples/ebola-framework2.jev-config.json'
json.dump(cfg, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
open(out, 'a').write('\n')
print(f'wrote {out}')
for q in cfg["questions"]:
    print(f'  {q["outputColumnName"]:28s} {q["questionType"]:7s} {len(q["options"]):3d} opts')

# Optional v2 hierarchy: all taxonomy content lives in the imported JSON.
# The runtime workflow engine contains no Ebola/Framework 2 special cases.
import copy
hierarchical = copy.deepcopy(cfg)
hierarchical['version'] = 2
hierarchical['questions'][1] = {
    'outputColumnName': 'framework2_topic', 'questionType': 'choice',
    'instructions': 'Choisis le sujet dominant dans record pour le type accepté dans decisions.framework2_type.',
    'dependsOn': ['framework2_type'],
    'branches': [{'when': {'framework2_type': kind}, 'options': options} for kind, options in valid.items()],
    'review': {'minConfidence': 0.6},
}
taxonomy = json.load(open(ROOT / 'framework2_taxonomy.json', encoding='utf-8'))['taxonomy']
leaf_branches = []
lookup_table = []
for pair, subdimension in G['grid'].items():
    kind, topic = pair.split('||')
    when = {'framework2_type': kind, 'framework2_topic': topic}
    lookup_table.append({'when': when, 'value': subdimension})
    options = list(dict.fromkeys(taxonomy[kind][subdimension]))
    # Explicit catch-all also gives singleton branches a genuine alternative.
    options.append('Aucun code ne correspond / information insuffisante')
    leaf_branches.append({'when': when, 'options': options})
hierarchical['questions'].append({
    'outputColumnName': 'framework2_code', 'questionType': 'choice',
    'instructions': 'Choisis le code précis correspondant au feedback dans record, pour le type et sujet acceptés dans decisions. Choisis la catégorie de repli si aucun code ne convient.',
    'dependsOn': ['framework2_type', 'framework2_topic'],
    'branches': leaf_branches, 'review': {'minConfidence': 0.6},
})
hierarchical['derived'] = [{'outputColumnName': 'framework2_sous_dimension', 'type': 'lookup',
    'inputs': ['framework2_type', 'framework2_topic'], 'table': lookup_table}]
hierarchical['report'] = {'title': 'Synthèse du feedback — Framework 2',
    'instructions': 'Rédige une synthèse en français avec les effectifs, les thèmes dominants, la criticité, les limites et les suites suggérées. Distingue les croyances rapportées des faits vérifiés.',
    'crossTabs': [['framework2_type', 'framework2_topic']], 'groupBy': [], 'evidenceColumns': ['Feedback'], 'maxExamples': 12}
hierarchical['execution'] = {'batchSize': 4, 'workers': 4}
out_v2 = ROOT / 'examples/ebola-framework2-hierarchical.jev-config.json'
with open(out_v2, 'w', encoding='utf-8') as target:
    json.dump(hierarchical, target, ensure_ascii=False, indent=2)
    target.write('\n')
print(f'wrote {out_v2}')
