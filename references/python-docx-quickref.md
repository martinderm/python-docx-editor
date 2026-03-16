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

## Caveats

- `.docx` is OOXML (zip + xml). Complex formatting can be lost with naive `paragraph.text = ...` writes.
- For safe edits, prefer targeted run-level manipulation.
- Tracked changes/comments are limited in python-docx and may require low-level OOXML edits.
