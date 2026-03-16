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

Default-Workflow bei inhaltlicher Überarbeitung:
1. v2 extrahieren
2. bei Projektbezug passende Datei aus `memory/references/projects/<project>/...` lesen
3. Range-basierten Patch-Plan generieren (`replace_paragraph_range`)
4. `removed_preview` prüfen und erst dann final anwenden

- Script: `scripts/extract_docx_for_llm.py`
- Default-Output: `docx-structure.v2` (hierarchisch: `sections -> content -> children`)
- Inhalte werden typisiert als `paragraph`, `list_item`, `table` (+ `headers`, `footers`)
- Stabile `block_id`s bleiben erhalten (`p_*`, `t_*`, `h_*`, `f_*`)
- Beispiele:
  - Struktur (v2): `py -3 scripts/extract_docx_for_llm.py --in input.docx --out structure.v2.json`
  - optional RAG (v1): `py -3 scripts/extract_docx_for_llm.py --in input.docx --out structure.v2.json --rag-output rag.v1.json`

## Preview helper

- Script: `scripts/docx_preview.py`
- v2-Section aus JSON ansehen:
  - `py -3 scripts/docx_preview.py v2-section --json structure.v2.json --title "3. Governance and Roles"`
- DOCX-Output um einen Heading-Text herum prüfen:
  - `py -3 scripts/docx_preview.py docx-around --in out.docx --contains "3. Governance and Roles" --lines 25`

## Minimal writeback helper

- Script: `scripts/apply_docx_patch.py`
- Zweck: gezielte Änderungen mit `block_id` zurückschreiben (ohne Full-Rewrite)
- Unterstützte Operationen:
  - `replace_text` mit `block_id`, `find`, `replace`, optional `expected_matches` (Default 1)
  - `set_paragraph` mit `block_id`, optional `text` oder `runs[]`, optional `style`, optional `expected_contains`, optional `markdown:true`
  - `delete_paragraph` mit `block_id`, optional `expected_contains`
  - `replace_paragraph_range` mit `start_block_id`, `end_block_id`, `new_paragraphs[]` (optional `allow_headings`)
  - `replace_paragraph_range_markdown` mit `start_block_id`, `end_block_id`, `markdown` (optional `allow_headings`)
- Markdown→Word (neu):
  - Inline: `*kursiv*`, `**fett**`, `***fett+kursiv***`, `` `code` ``, `[Text](https://...)`
  - Block-Level via `replace_paragraph_range_markdown`: Heading (`#`), Listen (`-` / `1.`), Quote (`>`), Trennlinie (`---`), Tabellen (`|...|`)
- Sicherheitsverhalten:
  - Fehler bei Mehrfachtreffern (wenn `expected_matches` nicht passt)
  - Fehler bei Treffern über Run-Grenzen (kein stilles Layout-Risiko)
  - Bei `replace_paragraph_range`: Bereich wird vollständig ersetzt (keine leeren Listenreste)
- Reihenfolge-Tipp bei mehreren Range-Edits:
  - von unten nach oben patchen (größere `p_*` zuerst), damit Block-IDs konsistent bleiben
- Beispiel:
  - `py -3 scripts/apply_docx_patch.py --in in.docx --out out.docx --patch patch.json`

## Project self-test

- `py -3 scripts/selftest.py`
- Erzeugt temporäre Fixture-DOCX, testet v2-Extraktion + Patch-Writeback und hinterlässt keine Repo-Artefakte.

## Caveats

- `.docx` is OOXML (zip + xml). Complex formatting can be lost with naive `paragraph.text = ...` writes.
- For safe edits, prefer targeted run-level manipulation.
- Tracked changes/comments are limited in python-docx and may require low-level OOXML edits.
