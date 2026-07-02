# -*- coding: utf-8 -*-
import json, os
canon = json.load(open('scratch/topics_canon.json', encoding='utf-8'))
os.makedirs('scratch/chunks', exist_ok=True)

PLAN = {
    '10.1': [(1, 12), (13, 24), (25, 36), (37, 48)],
    '10.2': [(1, 12), (13, 24), (25, 36)],
    '10.3': [(1, 14), (15, 28), (29, 42), (43, 55), (56, 66)],
}
manifest = []
for lo, ranges in PLAN.items():
    lokey = lo.replace('.', '_')  # 10_1
    items = canon[lo]
    for ci, (a, b) in enumerate(ranges, 1):
        slice_ = [x for x in items if a <= x['pos'] <= b]
        name = 'frag_%s_c%d' % (lokey, ci)
        chunk = {
            'lo_id': lo,
            'chunk': ci,
            'out_file': 'umkd_gen/%s.json' % name,
            'topics': slice_,
        }
        json.dump(chunk, open('scratch/chunks/%s.json' % name, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1)
        manifest.append((name, lo, ci, len(slice_), a, b))

lines = ['%s  LO %s chunk %d  %d topics (pos %d-%d)' % (n, lo, c, k, a, b)
         for (n, lo, c, k, a, b) in manifest]
open('scratch/chunks/manifest.txt', 'w', encoding='utf-8').write('\n'.join(lines))
print('\n'.join(lines))
print('TOTAL chunks:', len(manifest), 'lessons:', sum(m[3] for m in manifest))
