# -*- coding: utf-8 -*-
"""Build canonical per-LO ordered topic lists (merge duplicate 1.14 in LO 10.1),
map РУПД activity type -> (type_en, kind), and emit subagent briefing chunks."""
import json, collections

raw = json.load(open('scratch/topics_raw.json', encoding='utf-8'))

TYPE_MAP = {
    'Изучение нового материала': ('Study of new material', 'lecture'),
    'Комбинированный урок': ('Combined lesson', 'lecture'),
    'Урок совершенствования ЗУН': ('Skills-improvement lesson', 'practical'),
    'Урок применения и закрепления знаний': ('Application and consolidation of knowledge', 'practical'),
    'Урок комплексного применения знаний': ('Comprehensive application of knowledge', 'practical'),
    'Урок комплексной оценки знаний': ('Comprehensive assessment of knowledge', 'practical'),
    'Урок обобщения и оценки знаний': ('Generalization and assessment of knowledge', 'practical'),
}

# distinct types report
types = collections.Counter()
for lo, items in raw.items():
    for it in items:
        types[it['type_ru']] += 1

# merge duplicate 1.14 in LO 10.1: indexes 13 & 14 (0-based) are both "Тема 1.14"
lo1 = raw['10.1']
assert lo1[13]['title_ru'].startswith('Тема 1.14') and lo1[14]['title_ru'].startswith('Тема 1.14'), \
    (lo1[13]['title_ru'], lo1[14]['title_ru'])
merged = {
    'title_ru': 'Тема 1.14 Сетевой анализ и раскраска графов: алгоритмы Форда-Фалкерсона, Эдмондса-Карпа и жадного раскрашивания',
    'type_ru': 'Изучение нового материала',
}
raw['10.1'] = lo1[:13] + [merged] + lo1[15:]

canon = {}
for lo, items in raw.items():
    lst = []
    for i, it in enumerate(items, 1):
        te, kind = TYPE_MAP[it['type_ru']]
        # strip leading "Тема X.Y " from russian title
        t = it['title_ru']
        parts = t.split(' ', 1)
        title_ru_clean = parts[1] if len(parts) > 1 and parts[0] == 'Тема' else t
        # remove leading code like "1.14 " if still present
        import re
        title_ru_clean = re.sub(r'^\d+\.\d+\s*', '', title_ru_clean).strip()
        lst.append({'pos': i, 'title_ru': title_ru_clean, 'type_en': te, 'kind': kind})
    canon[lo] = lst

json.dump(canon, open('scratch/topics_canon.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

report = ['DISTINCT TYPES:']
for t, n in types.items():
    report.append('  %3d  %s' % (n, t))
report.append('')
for lo, lst in canon.items():
    lec = sum(1 for x in lst if x['kind'] == 'lecture')
    prc = len(lst) - lec
    report.append('%s: %d lessons (%d lecture / %d practical) = %d h' % (lo, len(lst), lec, prc, len(lst)*2))
open('scratch/report.txt', 'w', encoding='utf-8').write('\n'.join(report))
