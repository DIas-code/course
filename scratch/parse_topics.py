# -*- coding: utf-8 -*-
import docx, json, re
d = docx.Document('ПМ 10 РУПД 2025-2026.docx')
tbl = d.tables[2]
data = {'10.1': [], '10.2': [], '10.3': []}
for r in tbl.rows:
    c = [x.text.strip().replace(chr(10), ' ') for x in r.cells]
    topic = c[3] if len(c) > 3 else ''
    typ = c[-1].strip()
    m = re.match(r'Тема\s*(\d+)\.(\d+)', topic)
    if not m:
        continue
    key = '10.' + m.group(1)
    if key not in data:
        continue
    data[key].append({'title_ru': topic, 'type_ru': typ,
                       'theor': c[5].strip(), 'lab': c[6].strip()})
lines = ['%s rows=%d' % (k, len(v)) for k, v in data.items()]
json.dump(data, open('scratch/topics_raw.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
open('scratch/counts.txt', 'w', encoding='utf-8').write('\n'.join(lines))
