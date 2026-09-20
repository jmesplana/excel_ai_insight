"""Parse the Framework 2 sheet into a nested Type -> Sous-dimension -> Code dict."""
import zipfile, re, json, unicodedata
from xml.etree import ElementTree as ET
from collections import Counter

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
XLSX = 'examples/sample_framework_bvd.xlsx'
HEADER_ROW, FIRST_DATA_ROW, LAST_ROW, LAST_COL = 2, 3, 34, 69

# Source-sheet header typos: the header cell does not match the label used in
# the parent column. Bound by position + verified contents, not name similarity.
HEADER_OVERRIDES = {
    'Questions sur les MSP': 'Demande relative aux MSP2',
    'Allégations concernant  les acteurs de la réponse':
        'Allegations sur les acteurs de la réponse',
}

z = zipfile.ZipFile(XLSX)
sst = [''.join(t.text or '' for t in si.iter(f'{{{NS["m"]}}}t'))
       for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS)]

def colnum(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n

cells = {}
for row in ET.fromstring(z.read('xl/worksheets/sheet1.xml')).iter(f'{{{NS["m"]}}}row'):
    ri = int(row.get('r'))
    for c in row.findall('m:c', NS):
        v = c.find('m:v', NS)
        if v is None or v.text is None:
            continue
        val = (sst[int(v.text)] if c.get('t') == 's' else v.text).strip()
        if val:
            cells[(ri, colnum(c.get('r')))] = val

SEP = re.compile(r'^_{2,}.*_{2,}$')
def is_sep(s):
    return bool(SEP.match(s)) or set(s) <= set('_ ')

def norm(s):
    s = unicodedata.normalize('NFKD', s.replace('’', "'"))
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r'\s+', ' ', s).strip().lower()

headers = {c: cells[(HEADER_ROW, c)] for c in range(1, LAST_COL + 1)
           if (HEADER_ROW, c) in cells}
by_norm = {norm(h): c for c, h in headers.items()}
for label, target in HEADER_OVERRIDES.items():
    by_norm[norm(label)] = by_norm[norm(target)]

def column(c):
    return [cells[(r, c)] for r in range(FIRST_DATA_ROW, LAST_ROW + 1)
            if (r, c) in cells and not is_sep(cells[(r, c)])]

report = {'unresolved': [], 'empty_sous_dimensions': [], 'orphan_columns': [],
          'separators_dropped': sum(1 for v in cells.values() if is_sep(v)),
          'header_overrides_applied': HEADER_OVERRIDES}

taxonomy, used = {}, {1}
for typ in column(1):
    tc = by_norm[norm(typ)]
    used.add(tc)
    taxonomy[typ] = {}
    for sub in column(tc):
        sc = by_norm.get(norm(sub))
        if sc is None:
            report['unresolved'].append({'type': typ, 'sous_dimension': sub})
            taxonomy[typ][sub] = []
            continue
        used.add(sc)
        codes = column(sc)
        if not codes:
            report['empty_sous_dimensions'].append({'type': typ, 'sous_dimension': sub})
        taxonomy[typ][sub] = codes

report['orphan_columns'] = [headers[c] for c in sorted(headers) if c not in used]
leaves = [x for subs in taxonomy.values() for cs in subs.values() for x in cs]
report['duplicate_code_labels'] = sorted(k for k, n in Counter(leaves).items() if n > 1)
report['counts'] = {'types': len(taxonomy),
                    'sous_dimensions': sum(len(s) for s in taxonomy.values()),
                    'codes': len(leaves), 'unique_codes': len(set(leaves))}

with open('framework2_taxonomy.json', 'w', encoding='utf-8') as f:
    json.dump({'framework': 'Framework 2', 'disease': 'Ebola',
               'source': XLSX, 'taxonomy': taxonomy, 'report': report},
              f, ensure_ascii=False, indent=2)
print(json.dumps(report, ensure_ascii=False, indent=2))
