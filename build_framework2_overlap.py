"""Measure how much leaf vocabulary the types actually share within a topic.

The sous-dimension level is Type x Topic, but the LEAF level is not: what
people observe about a subject and what they ask about it differ. This
quantifies that, so the doc cites a measurement rather than an impression.
"""
import json, re, unicodedata

STRIP = [
    r'^observations?( ou croyances?)?( concernant| sur)?\s*',
    r'^declarations? (indiquant|concernant)\s*',
    r'^croyances?( selon lesquelles| indiquant| concernant| relatives? aux?)?\s*',
    r'^questions?( sur| indiquant| relative[s]? aux?| concernant)?\s*',
    r"^suggestions?( sur| concernant| de| d'| relatives? aux?| pour)?\s*",
    r'^appreciations?( sur| concernant| des| du| de)?\s*',
    r"^demandes?( de| d'| relative[s]? aux?| concernant| liee[s]? aux?| relatif)?\s*",
    r"^signalements?( sur| lie[s]? aux?| concernant| relatif aux?| de| d')?\s*",
    r'^plaintes? concernant\s*',
    r"^allegations? (de|d'|concernant)\s*",
]

def key(s):
    """Reduce a leaf label to its subject, dropping the type prefix and articles."""
    s = unicodedata.normalize('NFKD', s.replace('’', "'"))
    s = ''.join(c for c in s if not unicodedata.combining(c)).lower()
    for pattern in STRIP:
        s = re.sub(pattern, '', s)
    s = re.sub(r"^(l'|la |le |les |d'|du |des |une |un |aux |au |a )", '', s).strip()
    return ' '.join(re.sub(r'\b(du|des|au|aux|le|la|les|de|d|l)\b', '', s).split())

def main():
    tax = json.load(open('framework2_taxonomy.json', encoding='utf-8'))['taxonomy']
    grid_raw = json.load(open('framework2_topic_grid.json', encoding='utf-8'))['grid']

    by_topic = {}
    for k, sub in grid_raw.items():
        t, tp = k.split('||')
        by_topic.setdefault(tp, {})[t] = sub

    rows = []
    for tp, members in by_topic.items():
        if len(members) < 2:
            continue
        sets = {t: {key(c) for c in tax[t][s]} for t, s in members.items()}
        names = list(sets)
        pairs = ratio = 0
        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                a, b = sets[names[i]], sets[names[j]]
                pairs += 1
                # Jaccard would punish size differences; this asks "of the
                # smaller list, how much is echoed in the larger?"
                ratio += len(a & b) / max(1, min(len(a), len(b)))
        rows.append({'topic': tp, 'types': len(members), 'pairs': pairs,
                     'overlap': round(100 * ratio / pairs)})

    rows.sort(key=lambda r: -r['overlap'])
    mean = round(sum(r['overlap'] for r in rows) / len(rows))
    json.dump({'per_topic': rows, 'mean_overlap_pct': mean},
              open('framework2_overlap.json', 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)

    print(f'{"TOPIC":26s} {"types":>5s} {"overlap":>8s}')
    for r in rows:
        print(f'{r["topic"]:26s} {r["types"]:5d} {r["overlap"]:7d}%')
    print(f'\nmean leaf overlap across topic columns: {mean}%')

if __name__ == '__main__':
    main()
