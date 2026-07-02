# -*- coding: utf-8 -*-
"""Assemble subagent fragment JSONs into lo_10_1/2/3.json and validate."""
import os, json, sys

HERE = os.path.dirname(os.path.abspath(__file__))

PLAN = {
    '10_1': (['frag_10_1_c1', 'frag_10_1_c2', 'frag_10_1_c3', 'frag_10_1_c4'], 48),
    '10_2': (['frag_10_2_c1', 'frag_10_2_c2', 'frag_10_2_c3'], 36),
    '10_3': (['frag_10_3_c1', 'frag_10_3_c2', 'frag_10_3_c3', 'frag_10_3_c4', 'frag_10_3_c5'], 66),
}


def load(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return json.load(f)


def validate_lesson(les, where):
    errs = []
    for k in ('title', 'type_en', 'kind'):
        if not les.get(k):
            errs.append(f'{where}: missing {k}')
    kind = les.get('kind')
    if kind == 'lecture':
        if len(les.get('outline', [])) < 3:
            errs.append(f'{where} ({les.get("title")}): outline<3')
        if len(les.get('body', [])) < 10:
            errs.append(f'{where} ({les.get("title")}): body<10 (={len(les.get("body", []))})')
        if len(les.get('questions', [])) != 5:
            errs.append(f'{where} ({les.get("title")}): questions!=5 (={len(les.get("questions", []))})')
    else:
        for k in ('objective', 'theory', 'materials', 'procedure', 'tasks', 'conclusion'):
            if not les.get(k):
                errs.append(f'{where} ({les.get("title")}): missing {k}')
        if len(les.get('questions', [])) != 5:
            errs.append(f'{where} ({les.get("title")}): questions!=5 (={len(les.get("questions", []))})')
    return errs


def main():
    all_errs = []
    summary = []
    for lo_key, (frags, expected) in PLAN.items():
        lessons = []
        for fr in frags:
            path = os.path.join(HERE, fr + '.json')
            if not os.path.exists(path):
                all_errs.append(f'MISSING fragment: {fr}.json')
                continue
            data = load(fr + '.json')
            if not isinstance(data, list):
                all_errs.append(f'{fr}.json is not a JSON array')
                continue
            for i, les in enumerate(data):
                all_errs += validate_lesson(les, f'{fr}[{i}]')
            lessons += data
        lo_file = f'lo_{lo_key}.json'
        lo = load(lo_file)
        lec = sum(1 for x in lessons if x.get('kind') == 'lecture')
        prc = len(lessons) - lec
        summary.append(f'{lo_file}: {len(lessons)} lessons (exp {expected}) — {lec} lecture / {prc} practical')
        if len(lessons) != expected:
            all_errs.append(f'{lo_file}: COUNT {len(lessons)} != expected {expected}')
        lo['lessons'] = lessons
        with open(os.path.join(HERE, lo_file), 'w', encoding='utf-8') as f:
            json.dump(lo, f, ensure_ascii=False, indent=1)
    print('\n'.join(summary))
    if all_errs:
        print('\n--- ISSUES (%d) ---' % len(all_errs))
        print('\n'.join(all_errs))
        sys.exit(2)
    print('\nAll fragments assembled and validated OK.')


if __name__ == '__main__':
    main()
