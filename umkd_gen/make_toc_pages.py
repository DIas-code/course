# -*- coding: utf-8 -*-
"""Map each CONTENT entry of УМКД_ПМ01_Готовый.docx to its real page number.

Reads the headings + page numbers extracted from Word (headings.txt), aligns
them with the reconstructed table-of-contents labels, and writes a separate
file with the page numbers. Does NOT touch the source .docx.
"""
import io, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = r'C:\Users\Dias\AppData\Local\Temp\claude\C--Users-Dias-Desktop-course\767528e1-1ebd-40dc-8f49-105500f0c51f\scratchpad'
HEADINGS = os.path.join(SCRATCH, 'headings.txt')
OUT = os.path.join(HERE, '..', 'Содержание_страницы.txt')

ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']

los = []
for k in ['1_1', '1_2', '1_3', '1_4', '1_5', '1_6', '1_7']:
    los.append(json.load(open(os.path.join(HERE, f'lo_{k}.json'), encoding='utf-8')))

# reconstruct (toc_label, expected_real_heading) pairs in document order
pairs = []
pairs.append(('INTRODUCTION', 'INTRODUCTION'))
pairs.append(('I. ACADEMIC CALENDAR OF MODULE', 'I. ACADEMIC CALENDAR OF MODULE'))
SECTION_MAT = {'practical': 'PRACTICAL CLASS MATERIALS',
               'lab': 'LABORATORY WORK MATERIALS',
               'industrial': 'INDUSTRIAL TRAINING MATERIALS'}
for idx, lo in enumerate(los):
    rn = ROMAN[idx + 1]
    lid = lo['id']
    has_lec = any(l['kind'] == 'lecture' for l in lo['lessons'])
    others = [l for l in lo['lessons'] if l['kind'] != 'lecture']
    kind = others[0]['kind'] if others else 'practical'
    pairs.append((f'{rn}. LO {lid} {lo["title"]}', f'{rn}. LO {lid} {lo["title"]}'))
    pairs.append((f'1. LO {lid} Academic calendar', '1. ACADEMIC CALENDAR'))
    pairs.append((f'2. LO {lid} Glossary', '2. GLOSSARY'))
    if has_lec:
        pairs.append((f'3. LO {lid} Lecture materials', '3. LECTURE MATERIALS'))
    pairs.append((f'4. LO {lid} Practical / training materials', f'4. {SECTION_MAT[kind]}'))
    pairs.append((f'5. LO {lid} List of software and multimedia support', '5. LIST OF SOFTWARE AND MULTIMEDIA SUPPORT FOR TRAINING SESSIONS'))
    pairs.append((f'6. LO {lid} List of literature', '6. LIST OF LITERATURE'))
n = len(los) + 2
pairs.append(('IX. Self-Study Topics for the Module', 'IX. SELF-STUDY TOPICS FOR THE MODULE'))
pairs.append(('X. Exam Questions', 'X. EXAM QUESTIONS'))

# read extracted headings
rows = []
for line in io.open(HEADINGS, encoding='utf-8'):
    line = line.rstrip('\n')
    if not line.strip():
        continue
    pg, txt = line.split('\t', 1)
    rows.append((int(pg), txt.strip()))

# real content begins at the 2nd 'INTRODUCTION'
intro_idx = [i for i, (_, t) in enumerate(rows) if t == 'INTRODUCTION']
real = rows[intro_idx[1]:]

assert len(real) == len(pairs), f'count mismatch: real={len(real)} pairs={len(pairs)}'

out = []
out.append('CONTENT — page numbers (from УМКД_ПМ01_Готовый.docx as rendered in Word)')
out.append('=' * 78)
out.append('')
mism = 0
for (toc_label, expected), (pg, got) in zip(pairs, real):
    if got != expected:
        mism += 1
    dots = '.' * max(3, 74 - len(toc_label) - len(str(pg)))
    indent = '    ' if toc_label[0].isdigit() else ''
    out.append(f'{indent}{toc_label} {dots} {pg}')

io.open(OUT, 'w', encoding='utf-8').write('\n'.join(out) + '\n')
print('alignment mismatches:', mism)
print('entries:', len(pairs))
print('written:', os.path.abspath(OUT))
