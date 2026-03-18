# python-docx – Listenverhalten (Kurznotizen)

Stand: 2026-03-18

## Was die offizielle python-docx-Doku klar zeigt

1. Listen werden im üblichen python-docx-Workflow über **Paragraph-Styles** gesetzt:
   - `style='List Bullet'`
   - `style='List Number'`

2. Styles werden per `paragraph.style` oder direkt bei `add_paragraph(..., style=...)` gesetzt.

3. Wichtig für internationale Dokumente:
   - Built-in Style-Namen müssen in python-docx über den **englischen Namen** angesprochen werden (z. B. `Heading 1`, `List Bullet`, `List Number`).

## Relevanz für unseren Patcher

- Nur Style setzen reicht in manchen Vorlagen nicht für robustes Listenverhalten über Kapitelgrenzen hinweg.
- Deshalb kombinieren wir:
  - python-docx Style-Setzung (`List Bullet`/`List Number`) **plus**
  - direkte OOXML-Nummerierungssteuerung (`w:numPr`, eigene `w:num`-Instanz, ggf. Override).

So vermeiden wir:
- fehlende Bullet-Symbole in UL,
- ungewollte Fortsetzung von OL-Nummern in späteren Kapiteln.

## Quellen (offiziell)

1. python-docx Startseite (Beispiel mit `List Bullet` / `List Number`):
   - https://python-docx.readthedocs.io/en/latest/

2. Working with Styles (Style-Zugriff, Anwendung, englische Built-in-Namen):
   - https://python-docx.readthedocs.io/en/latest/user/styles-using.html

Hinweis: Für die tieferen Nummerierungsdetails gilt weiter die OOXML-Quelle in
`references/ooxml-numbering-notes.md` (Microsoft Learn / WordprocessingML-Mapping).