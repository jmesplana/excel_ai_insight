"""Derive the (Type, Topic) -> Sous-dimension grid from the parsed taxonomy."""
import json, re, unicodedata
from collections import defaultdict

import pathlib

# Paths are resolved against the repo root so this runs from any directory.
ROOT = pathlib.Path(__file__).resolve().parent.parent

tax = json.load(open(ROOT / 'framework2_taxonomy.json', encoding='utf-8'))['taxonomy']

def norm(s):
    s = unicodedata.normalize('NFKD', s.replace('’', "'"))
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip().lower()

PREFIX = re.compile(
    r'^(observations?(\s+ou\s+croyances?)?|croyances?|suggestions?|'
    r'appreciations?|appreciation|demandes?|demande|questions?|'
    r'signalements?|signalement|allegations?|allegation)\b'
    r'(\s+(sur|concernant|relative[s]?|relatif[s]?|liee?[s]?|de|du|des|a|aux|au|d))*\b',
    re.I)

def strip(s):
    t = PREFIX.sub('', norm(s)).strip()
    return re.sub(r"^(l'|la |le |les )", '', t).strip() or norm(s)

# Explicit merges: the same topic spelled differently across types. Each maps a
# stripped label to the canonical topic. Listed here rather than inferred, so
# every judgment call is reviewable.
MERGE = {
    "traitement": "Traitement",
    "traitements": "Traitement",
    "traitement de la maladie": "Traitement",
    "vaccins": "Vaccins",
    "vaccination": "Vaccins",
    "services de sante": "Services de santé",
    "services de soins de sante": "Services de santé",
    "surete et la securite": "Sûreté et sécurité",
    "surete et securite": "Sûreté et sécurité",
    "securite et la surete": "Sûreté et sécurité",
    "maladie": "Maladie",
    "epidemie": "Épidémie",
    # Content is about the disease (symptoms, transmission, recovery), not the
    # outbreak trend — so it lands under Maladie, filling the Demandes/Maladie gap.
    "plus d'informations sur l'epidemie": "Maladie",
    "impact de l'epidemie": "Impact de l'épidémie",
    "comportements": "Comportements",
    "'aide liee aux comportements": "Comportements",
    "msp": "MSP",
    "reponse": "Réponse",
    "acteurs de la reponse": "Acteurs de la réponse",
    "autres observations, perceptions ou croyances": "Autres",
    "autres questions": "Autres",
    "autre demande": "Autres",
}

grid = {}           # (type, topic) -> sous-dimension label
unmapped = []
for t, subs in tax.items():
    for s in subs:
        key = strip(s)
        topic = MERGE.get(key)
        if topic is None:
            unmapped.append((t, s, key)); continue
        if (t, topic) in grid:
            print(f'!! COLLISION {t} / {topic}: {grid[(t,topic)]}  vs  {s}')
        grid[(t, topic)] = s

TOPICS = ["Maladie", "Épidémie", "Impact de l'épidémie", "Comportements", "MSP",
          "Traitement", "Services de santé", "Vaccins", "Réponse",
          "Acteurs de la réponse", "Sûreté et sécurité", "Autres"]
TYPES = list(tax)

print(f'sous-dimensions mapped: {len(grid)} / 61')
print(f'unmapped: {unmapped or "none"}')
print(f'topics: {len(TOPICS)}')
print()
print('GRID  (. = no such sous-dimension in the framework)')
print(f'{"":26s}' + ''.join(f'{t[:9]:11s}' for t in TYPES))
for topic in TOPICS:
    row = ''.join(f'{("YES" if (t,topic) in grid else "."):11s}' for t in TYPES)
    print(f'{topic:26s}{row}')
gaps = [(t, tp) for t in TYPES for tp in TOPICS if (t, tp) not in grid]
print(f'\ncells: {len(TYPES)*len(TOPICS)}  filled: {len(grid)}  gaps: {len(gaps)}')

json.dump({'types': TYPES, 'topics': TOPICS,
           'grid': {f'{t}||{tp}': v for (t, tp), v in grid.items()},
           'gaps': [{'type': t, 'topic': tp} for t, tp in gaps]},
          open(ROOT / 'framework2_topic_grid.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=2)
