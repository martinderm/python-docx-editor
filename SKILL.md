---
name: python-docx-editor
description: "Arbeiten mit Word-Dokumenten (.docx) über die Python-Bibliothek python-docx: Text extrahieren, Dokumente erstellen/ändern, Tabellen lesen, einfache Ersetzungen und Struktur-Checks. Use when ein Agent DOCX-Dateien automatisiert verarbeiten soll oder reproduzierbare DOCX-Operationen per Script braucht."
---

# Python DOCX Editor

Nutze diesen Skill für wiederholbare DOCX-Aufgaben mit `python-docx`.

## Workflow

1. Stelle sicher, dass `python-docx` verfügbar ist.
   - Standard-Installation:
     - `py -3 -m pip install python-docx`
   - Optional für Entwicklung gegen lokalen Clone:
     - `py -3 -m pip install -e vendor/python-docx`
2. Prüfe, ob das Dokument Projektbezug hat (Titel, Projektnamen, Akronyme, Partnerlisten, WP-Struktur).
3. Wenn Projektbezug vorliegt: lade vor Text-Transformationen die passende Projektreferenz aus
   - `memory/references/projects/<project>/...` (falls vorhanden)
   - und nutze sie als Stil-/Terminologie-Quelle (Begriffe, Rollen, formale Benennungen, Tonalität).
4. Nutze für Standardaufgaben `scripts/docx_ops.py`.
5. Für Spezialfälle lade `references/python-docx-quickref.md` und implementiere gezielt.

## Standardbefehle

- Text extrahieren:
  - `py -3 scripts/docx_ops.py text --in <datei.docx>`
- Statistik (JSON):
  - `py -3 scripts/docx_ops.py stats --in <datei.docx>`
- Einfache Ersetzung in Absätzen + Tabellen:
  - `py -3 scripts/docx_ops.py replace --in <in.docx> --out <out.docx> --find "Alt" --replace "Neu"`
- LLM-freundliche Extraktion (v2, hierarchisch nach Sections/Content):
  - `py -3 scripts/extract_docx_for_llm.py --in <in.docx> --out <structure.v2.json>`
- Optionaler RAG-Output (v1, flat blocks+chunks):
  - `py -3 scripts/extract_docx_for_llm.py --in <in.docx> --out <structure.v2.json> --rag-output <rag.v1.json>`
- Gezieltes Minimal-Writeback via Patch:
  - `py -3 scripts/apply_docx_patch.py --in <in.docx> --out <out.docx> --patch <patch.json>`
- Section-/DOCX-Preview für Review:
  - `py -3 scripts/docx_preview.py v2-section --json <structure.v2.json> --title "<Section Title>"`
  - `py -3 scripts/docx_preview.py docx-around --in <out.docx> --contains "<Heading Text>" --lines 25`

## LLM-Default Editing Procedure (verbindlich)

Nutze für Überarbeitungen standardmäßig diesen Ablauf (kein manuelles Range-Basteln im Normalfall):

1. **Extrahieren (v2)**
   - `py -3 scripts/extract_docx_for_llm.py --in <in.docx> --out <structure.v2.json>`
2. **Kontext laden**
   - Bei Projektbezug zusätzlich relevante Infos aus `memory/references/projects/<project>/...` laden (Terminologie, Rollen, Zielgruppen, Ton).
3. **Edit-Plan erzeugen (LLM)**
   - LLM erstellt einen Patch-Plan mit `ops[]` für `apply_docx_patch.py`.
   - Für größere Umformungen immer `replace_paragraph_range` verwenden.
   - `replace_text` nur für kleine, lokale Korrekturen.
4. **Semantik-Check vor Apply**
   - Bei Listen/Abschnitts-Rewrites: Original-Claims erfassen und im Zieltext wiederfinden.
   - Keine inhaltstragenden Punkte stillschweigend entfernen.
5. **Apply + Review**
   - `py -3 scripts/apply_docx_patch.py --in <in.docx> --out <out.docx> --patch <patch.json>`
   - `results[].removed_preview` prüfen (bei Range-Edits).
6. **Feinschliff**
   - Bei Bedarf zweiten, kleinen Patchlauf für Stil-/Rhythmuskorrektur.

### Geltungsbereich der Überarbeitung

Der Workflow gilt nicht nur für Listen, sondern allgemein für:
- Abschnittsverdichtung (zu lang/zu redundant)
- Abschnittsausbau (zu knapp)
- Umstrukturierung (Listen ↔ Fließtext)
- Terminologie-/Stilkonsistenz über das Dokument

## Rewrite-Strategie (Abschnitte allgemein)

Bei Abschnittsverbesserungen:
- Identifiziere den Zielabschnitt in v2 (`Heading` + zugehöriger Content-Bereich).
- Entscheide Modus explizit:
  - `narrative` (Fließtext)
  - `hybrid` (Fließtext + wenige kompakte Punkte)
  - `preserve-list` (Liste bleibt, sprachlich verbessert)
- Erzeuge einen zusammenhängenden Patch pro semantischem Block (Range-basiert statt atomisiert).
- Halte Claim-Abdeckung hoch: Kerninhalte bleiben erhalten, nur Redundanz wird reduziert.

## JSON Schemas & Contracts

### 1) Input-Schema: `extract_docx_for_llm.py`

Pflichtparameter:
- `--in` (Pfad zu bestehender `.docx`)
- `--out` (Zielpfad)

Optionale Parameter:
- `--table-row-batch` (Integer, Default: `25`; bei `<= 0` wird die ganze Tabelle als ein Block/Part behandelt)
- `--rag-output` (Pfad für optionalen v1-RAG-Output)
- `--max-chars` (Integer, Default: `4000`, nur für `--rag-output`, intern mindestens `500`)
- `--overlap-blocks` (Integer, Default: `1`, nur für `--rag-output`, intern mindestens `0`)

Validierung:
- `--in` muss existieren
- `--in` muss `.docx`-Endung haben
- Zielordner von `--out` wird bei Bedarf erstellt

Formales JSON-Schema (CLI-Äquivalent):

```json
{
  "type": "object",
  "required": ["in", "out"],
  "properties": {
    "in": {
      "type": "string",
      "description": "Pfad zu bestehender .docx-Datei"
    },
    "out": {
      "type": "string",
      "description": "Ausgabepfad für v2-JSON"
    },
    "rag_output": {
      "type": "string",
      "description": "Optionaler Ausgabepfad für v1 (RAG blocks+chunks)"
    },
    "max_chars": {
      "type": "integer",
      "default": 4000,
      "minimum": 500,
      "description": "Script erzwingt intern mindestens 500"
    },
    "overlap_blocks": {
      "type": "integer",
      "default": 1,
      "minimum": 0,
      "description": "Script erzwingt intern mindestens 0"
    },
    "table_row_batch": {
      "type": "integer",
      "default": 25,
      "description": "<=0 bedeutet ganze Tabelle als ein Block"
    }
  },
  "additionalProperties": false
}
```

---

### 2) Output-Schema (Default): `docx-structure.v2`

Der Standard-Output ist hierarchisch und gruppiert Inhalte nach Sections (Heading-basiert).

```json
{
  "type": "object",
  "required": ["schema", "source", "stats", "document"],
  "properties": {
    "schema": { "const": "docx-structure.v2" },
    "source": { "type": "string" },
    "stats": {
      "type": "object",
      "required": ["sections", "content", "headers", "footers", "pre_heading_nodes"],
      "properties": {
        "sections": { "type": "integer", "minimum": 0 },
        "content": {
          "type": "object",
          "required": ["paragraph", "list_item", "table"],
          "properties": {
            "paragraph": { "type": "integer", "minimum": 0 },
            "list_item": { "type": "integer", "minimum": 0 },
            "table": { "type": "integer", "minimum": 0 }
          }
        },
        "headers": { "type": "integer", "minimum": 0 },
        "footers": { "type": "integer", "minimum": 0 },
        "pre_heading_nodes": { "type": "integer", "minimum": 0 }
      }
    },
    "document": {
      "type": "object",
      "required": ["pre_heading", "sections", "headers", "footers"],
      "properties": {
        "pre_heading": { "type": "array" },
        "sections": { "type": "array" },
        "headers": { "type": "array" },
        "footers": { "type": "array" }
      }
    }
  }
}
```

`document.sections[]` Knotenstruktur:
- `level`, `title`, `block_id`, `style`
- `content[]` mit typisierten Nodes:
  - `paragraph`
  - `list_item` (optional mit `list: {num_id, level}`)
  - `table` (mit `parts[]`, je Part eigene `block_id`)
- `children[]` für Unter-Sections

Hinweise zu IDs:
- Absätze/Headings: `p_<n>`
- Tabellenbereiche: `t_<tableIndex>_r<start>_<end>`
- Header: `h_<n>`
- Footer: `f_<n>`

### 2b) Optionales RAG-Schema: `docx-llm-chunks.v1` (`--rag-output`)

Wenn `--rag-output <pfad>` gesetzt ist, wird zusätzlich ein v1-JSON mit `blocks[]` und `chunks[]` erzeugt (kompatibel für Retrieval-Pipelines).

---

### 3) Input-Schema: `apply_docx_patch.py`

Pflichtparameter:
- `--in` (Pfad zu bestehender `.docx`)
- `--out` (Zielpfad für neue `.docx`)
- `--patch` (Pfad zu Patch-JSON)

Aktuell unterstützte Operationen:
- `replace_text`
- `set_paragraph`
- `delete_paragraph`
- `replace_paragraph_range`

Patch-Datei-Schema:

```json
{
  "type": "object",
  "required": ["ops"],
  "properties": {
    "ops": {
      "type": "array",
      "items": {
        "oneOf": [
          {
            "type": "object",
            "required": ["op", "block_id", "find", "replace"],
            "properties": {
              "op": { "const": "replace_text" },
              "block_id": { "type": "string" },
              "find": { "type": "string", "minLength": 1 },
              "replace": { "type": "string" },
              "expected_matches": { "type": "integer", "minimum": 0, "default": 1 }
            },
            "additionalProperties": false
          },
          {
            "type": "object",
            "required": ["op", "block_id", "text"],
            "properties": {
              "op": { "const": "set_paragraph" },
              "block_id": { "type": "string" },
              "text": { "type": "string" },
              "style": { "type": "string" },
              "expected_contains": { "type": "string" }
            },
            "additionalProperties": false
          },
          {
            "type": "object",
            "required": ["op", "block_id"],
            "properties": {
              "op": { "const": "delete_paragraph" },
              "block_id": { "type": "string" },
              "expected_contains": { "type": "string" }
            },
            "additionalProperties": false
          },
          {
            "type": "object",
            "required": ["op", "start_block_id", "end_block_id", "new_paragraphs"],
            "properties": {
              "op": { "const": "replace_paragraph_range" },
              "start_block_id": { "type": "string", "pattern": "^p_[0-9]+$" },
              "end_block_id": { "type": "string", "pattern": "^p_[0-9]+$" },
              "expected_start_contains": { "type": "string" },
              "expected_end_contains": { "type": "string" },
              "allow_headings": { "type": "boolean", "default": false },
              "new_paragraphs": {
                "type": "array",
                "minItems": 1,
                "items": {
                  "type": "object",
                  "required": ["text"],
                  "properties": {
                    "text": { "type": "string" },
                    "style": { "type": "string" }
                  },
                  "additionalProperties": false
                }
              }
            },
            "additionalProperties": false
          }
        ]
      }
    }
  },
  "additionalProperties": false
}
```

Zusatzschema für Listen→Fließtext-Refactoring (`replace_paragraph_range`):

```json
{
  "op": "replace_paragraph_range",
  "start_block_id": "p_27",
  "end_block_id": "p_30",
  "expected_start_contains": "Develops and maintains",
  "expected_end_contains": "Define QM-Indicators",
  "new_paragraphs": [
    { "text": "Neuer Fließtext ...", "style": "Normal" },
    { "text": "Optionaler kompakter Punkt ...", "style": "Compact" }
  ]
}
```

Regeln für `replace_paragraph_range`:
- Ersetzt den gesamten Absatzbereich (`p_*` von start bis end, inkl.) atomar.
- Fügt `new_paragraphs` ein und entfernt den alten Bereich vollständig (keine leeren Rest-Bullets).
- Schützt Headings standardmäßig: wenn der Bereich Heading-Absätze enthält, bricht die Operation ab (nur mit `allow_headings: true` überschreibbar).
- Bei mehreren Range-Operationen im selben Patch **von unten nach oben** (höhere `p_*` zuerst) arbeiten, damit IDs stabil bleiben.

Qualitätsregeln für Listen→Fließtext:
- Behandle zusammenhängende Listen als **eine semantische Einheit** (nicht Punkt-für-Punkt umformulieren).
- Führe vor dem Patch eine Claim-Liste der Originalpunkte (1 Claim je Bullet).
- Stelle sicher, dass jeder Claim im neuen Text wieder auftaucht (Inhaltserhalt statt freier Verdichtung).
- Vermeide „Lonely bullet remnants": wenn nur 1 Punkt übrig bleibt, in Fließtext integrieren.
- Bevorzuge bei operativen Abschnitten aktive Verben und konkrete Zuständigkeiten; keine unnötige Abstraktion.
- Nutze `removed_preview` im Patch-Resultat als Pflicht-Review vor finaler Freigabe.

---

### 4) Failure / Validation Contract (Writeback)

`apply_docx_patch.py` bricht hart mit Fehler ab, wenn:
- `block_id` nicht gefunden wird
- `find` leer ist
- `expected_matches` nicht der tatsächlichen Trefferanzahl entspricht
- ein Treffer über Run-Grenzen geht (kein stilles Full-Rewrite)
- eine unbekannte `op` verwendet wird

Erfolgsoutput (stdout) ist JSON:
- `in`, `out`, `ops`, `results[]`
- pro Operation: `op`, `block_id`, `matches`, `changes`, `status`

## Prompt Contract für robuste Überarbeitungen

Wenn der User „Abschnitt verbessern/überarbeiten“ sagt, soll der Agent intern diesen Plan erzwingen:
- `target_section`: exakter Heading-Titel
- `mode`: `narrative | hybrid | preserve-list`
- `intent`: `shorten | expand | clarify | harmonize-style`
- `must_keep`: Liste nicht verlierbarer Kernpunkte/Claims
- `project_context_used`: welche Datei aus `memory/references/projects/...` verwendet wurde (falls vorhanden)
- `style_profile`: gewünschter Stil (explizit vom User oder aus Kontext abgeleitet)

Stil-Klärung (verbindlich):
- Wenn Stil/Tonalität aus Prompt + Dokumentkontext **nicht klar** ableitbar ist, **vor dem Patch** den User kurz nach dem gewünschten Stil fragen.
- Erst nach Stil-Klärung den finalen Patch erzeugen.

Default-Style (wenn User keinen eigenen Stil vorgibt und Kontext konsistent ist):
- Write in a natural, fluent, and precise style.
- Prefer active voice, clear structure, and concise phrasing.
- Vary sentence length and rhythm without sacrificing coherence.
- Use less common words only when they add technical or conceptual precision, not decorative flair.
- Maintain grammatical accuracy, logical clarity, and stylistic consistency.
- Increase linguistic diversity (“perplexity” and “burstiness”) only insofar as it improves readability and nuance.

Ausgabeformat vor Patch-Anwendung (kurz):
- **Claims alt** (stichwortartig)
- **Claims neu** (Mapping alt→neu)
- **Style profile** (1-2 Zeilen)
- **Ops-Plan** (`replace_paragraph_range` bevorzugt)

Patch nur anwenden, wenn kein Kernclaim ungemappt bleibt.

## Quickstart (maintainer)

1. Umgebung prüfen:
   - `py -3 -m pip install python-docx`
2. Syntax prüfen:
   - `py -3 -m py_compile scripts/extract_docx_for_llm.py scripts/apply_docx_patch.py scripts/selftest.py`
3. End-to-End-Selbsttest (ohne Repo-Artefakte):
   - `py -3 scripts/selftest.py`

## Release Readiness Checklist

Vor Veröffentlichung sicherstellen:
- [ ] `scripts/selftest.py` läuft lokal grün
- [ ] v2-Output (`docx-structure.v2`) bleibt default
- [ ] Optionaler v1-RAG-Output funktioniert via `--rag-output`
- [ ] `apply_docx_patch.py` validiert `expected_matches` und blockiert Run-Grenzkonflikte
- [ ] `SKILL.md`-Schemas entsprechen dem tatsächlichen CLI-Verhalten
- [ ] Keine Testartefakte versehentlich versioniert (lokale Dateien bleiben in `Tests/`)

## Versioning Notes

- **v2 default**: `extract_docx_for_llm.py --out ...` erzeugt `docx-structure.v2`.
- **v1 optional**: Nur bei gesetztem `--rag-output` wird zusätzlich `docx-llm-chunks.v1` erzeugt.
- **Patch-Kompatibilität**: `apply_docx_patch.py` nutzt die gleiche Paragraph-ID-Logik wie der Extractor (nur nicht-leere Body-Paragraphen erhalten `p_<n>`).

## Grenzen

- `replace` in `docx_ops.py` ist bewusst einfach und kann Formatierung beeinflussen.
- `apply_docx_patch.py` ist minimal-sicher: bei Run-Grenzkonflikten wird verweigert statt riskant umgeschrieben.
- Für formatkritische Änderungen auf Run-Ebene arbeiten statt `paragraph.text` komplett neu zu setzen.
