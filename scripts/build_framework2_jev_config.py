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
