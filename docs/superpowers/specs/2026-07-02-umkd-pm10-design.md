# УМКД ПМ 10 — design / spec

Build the full Educational-Methodological Complex (EMC / УМКД) for module
**PM 10 «Development of software program code»**, matching the structure,
depth and formatting of the existing `УМКД_ПМ01_Готовый.docx`, following
`ПРАВИЛА_УМКД.md` and the TIPO-2025 methodical recommendations. All content in
English. Source of truth: `ПМ 10 РУПД 2025-2026.docx`.

## Module facts (from РУПД)

- Specialty 06130100 «Software (by types)», qualification 4S06130105
  «Information Systems Technician». Total volume stated as **336 h / 12.5
  credits** (title + intro), matching the РУПД header.
- Developers: K.D. Rakhym, A.A. Daulbay, G. Aibaruly. SCMC «Software» chair:
  A.E. Beket. Prerequisites PM 08, PM 09; post-requisites: pre-diploma
  practice and diploma design. Instructor: K.D. Rakhym.
- **Three learning outcomes** (vs 7 in PM 01):
  - LO 10.1 — Formalization and algorithmization of tasks — 96 h — 48 lessons
    (36 lectures + 12 practical).
  - LO 10.2 — Develop program code from ready requirement specifications —
    72 h — 36 lessons (28 lectures + 8 practical).
  - LO 10.3 — Debug program code at the level of program modules — 132 h —
    66 lessons (48 lectures + 18 practical).
  - Directed study totals **150 lessons / 300 h**; weights 19 % / 15 % / 26 %
    (+40 % exam = 100 %).

## Decisions

- **LO 10.1 duplicate «Тема 1.14».** The РУПД numbers two distinct lessons as
  1.14 (network-flow analysis; graph coloring), giving 49 rows while the hours
  table says 96 h = 48 lessons. **Resolution (user-approved):** merge the two
  into one lecture "Network flow and graph coloring: Ford–Fulkerson,
  Edmonds–Karp and greedy coloring algorithms" → 48 lessons, matches 96 h.
- **Activity-type → kind mapping:** «Изучение нового материала» → lecture
  (Study of new material); «Урок совершенствования ЗУН» → practical
  (Skills-improvement lesson); «Урок комплексной оценки знаний» → practical
  (Comprehensive assessment of knowledge).
- **Module hours in text:** 336 h (РУПД header), as in PM 01.

## Pipeline (reuse `umkd_gen/`, do not disturb PM 01)

- `build_umkd.py` generalized: reads a module-config file named on argv
  (default `module.json`), taking `lo_files` and `out_name` from it. PM 01
  still builds via the defaults.
- New `module_pm10.json` (title page, approval sheet, introduction, weights,
  software, literature, self-study, exam) — hand-authored.
- New `lo_10_1.json`, `lo_10_2.json`, `lo_10_3.json` (glossary + ordered
  lessons). Glossaries (8–13 terms each) authored centrally; lesson content
  produced by parallel subagents.

## Content generation (parallel subagents)

Twelve subagents, each writing a JSON array of lesson objects for an ordered
slice of topics (12–14 lessons each), in the exact schema and depth of
`ПРАВИЛА_УМКД.md` §8 and the PM 01 эталон:

- Lecture: `outline` (4–5 items) + `body` (12–18 substantive paragraphs, last
  begins "Thus, …") + exactly 5 `questions`.
- Practical: `objective` + `theory` (2–3 paras) + `materials` + `procedure`
  (~6 steps) + `tasks` (3) + `conclusion` + 5 `questions`.

Each agent receives its topic slice (position, faithful English title to use as
`title`, `type_en`, `kind`), the schema, and the PM 01 exemplar. Agents write
to `umkd_gen/frag_<lo>_<chunk>.json`. The coordinator assembles fragments into
`lo_10_*.json` (validating counts and kinds), then renders the docx and fills
the CONTENT page numbers via Word COM (read-only on the source), per
ПРАВИЛА §12.

## Verification

- Calendar rows per LO = hours ÷ 2 (48 / 36 / 66); module 150 lessons / 300 h.
- Weights 19+15+26 + 40 = 100.
- All РУПД topics present, order preserved; LO-without-lecture rule N/A (all
  three LOs have lectures).
- Formatting: TNR, margins 25/20/15/20, centered footer page numbers — inherited
  from the shared builder.
