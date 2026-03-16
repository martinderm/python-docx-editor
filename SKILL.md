---
name: python-docx-skill
description: "Arbeiten mit Word-Dokumenten (.docx) über die Python-Bibliothek python-docx: Text extrahieren, Dokumente erstellen/ändern, Tabellen lesen, einfache Ersetzungen und Struktur-Checks. Use when ein Agent DOCX-Dateien automatisiert verarbeiten soll oder reproduzierbare DOCX-Operationen per Script braucht."
---

# Python DOCX Skill

Nutze diesen Skill für wiederholbare DOCX-Aufgaben mit `python-docx`.

## Workflow

1. Stelle sicher, dass `python-docx` verfügbar ist.
   - Standard-Installation:
     - `py -3 -m pip install python-docx`
   - Optional für Entwicklung gegen lokalen Clone:
     - `py -3 -m pip install -e vendor/python-docx`
2. Nutze für Standardaufgaben `scripts/docx_ops.py`.
3. Für Spezialfälle lade `references/python-docx-quickref.md` und implementiere gezielt.

## Standardbefehle

- Text extrahieren:
  - `py -3 scripts/docx_ops.py text --in <datei.docx>`
- Statistik (JSON):
  - `py -3 scripts/docx_ops.py stats --in <datei.docx>`
- Einfache Ersetzung in Absätzen + Tabellen:
  - `py -3 scripts/docx_ops.py replace --in <in.docx> --out <out.docx> --find "Alt" --replace "Neu"`

## Grenzen

- `replace` ist bewusst einfach und kann Formatierung beeinflussen.
- Für formatkritische Änderungen auf Run-Ebene arbeiten statt `paragraph.text` komplett neu zu setzen.
