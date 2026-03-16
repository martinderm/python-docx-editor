# python-docx Quick Reference

## Setup

- Standard:
  - `py -3 -m pip install python-docx`
- Optional (Entwicklung gegen lokalen Clone):
  - `py -3 -m pip install -e vendor/python-docx`

## Common operations

- Load file: `Document("file.docx")`
- Read paragraphs: `for p in doc.paragraphs: p.text`
- Read tables: `for t in doc.tables ...`
- Add paragraph: `doc.add_paragraph("Text")`
- Save: `doc.save("out.docx")`

## LLM extraction helper

- Script: `scripts/extract_docx_for_llm.py`
- Default-Output: `docx-structure.v2` (hierarchisch: `sections -> content -> children`)
- Inhalte werden typisiert als `paragraph`, `list_item`, `table` (+ `headers`, `footers`)
- Stabile `block_id`s bleiben erhalten (`p_*`, `t_*`, `h_*`, `f_*`)
- Beispiele:
  - Struktur (v2): `py -3 scripts/extract_docx_for_llm.py --in input.docx --out structure.v2.json`
  - optional RAG (v1): `py -3 scripts/extract_docx_for_llm.py --in input.docx --out structure.v2.json --rag-output rag.v1.json`

## Minimal writeback helper

- Script: `scripts/apply_docx_patch.py`
- Zweck: gezielte Änderungen mit `block_id` zurückschreiben (ohne Full-Rewrite)
- Unterstützte Operation aktuell:
  - `replace_text` mit `block_id`, `find`, `replace`, optional `expected_matches` (Default 1)
- Sicherheitsverhalten:
  - Fehler bei Mehrfachtreffern (wenn `expected_matches` nicht passt)
  - Fehler bei Treffern über Run-Grenzen (kein stilles Layout-Risiko)
- Beispiel:
  - `py -3 scripts/apply_docx_patch.py --in in.docx --out out.docx --patch patch.json`

## Project self-test

- `py -3 scripts/selftest.py`
- Erzeugt temporäre Fixture-DOCX, testet v2-Extraktion + Patch-Writeback und hinterlässt keine Repo-Artefakte.

## Caveats

- `.docx` is OOXML (zip + xml). Complex formatting can be lost with naive `paragraph.text = ...` writes.
- For safe edits, prefer targeted run-level manipulation.
- Tracked changes/comments are limited in python-docx and may require low-level OOXML edits.
