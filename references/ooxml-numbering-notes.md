# OOXML Numbering – Kurznotizen für python-docx-editor

Stand: 2026-03-18

## Relevantes Modell (WordprocessingML)

- `w:abstractNum` = abstrakte Listendefinition (Format/Level, z. B. bullet oder decimal).
- `w:num` = konkrete Instanz einer abstrakten Definition (`w:abstractNumId`).
- Absatz-seitig wird Liste über `w:numPr` gesetzt:
  - `w:numId` (welche Instanz)
  - `w:ilvl` (welcher Listenlevel)

Für unser Problem wichtig:
- Unterschiedliche Kapitel sollten bei geordneten Listen (OL) nicht dieselbe laufende Instanz teilen, sonst zählt Word weiter.
- Für Neustart einer OL ist pro neuem Listenblock eine frische `w:num`-Instanz sinnvoll; alternativ/ergänzend über `w:lvlOverride` + `w:startOverride` explizit Startwert setzen.

## Override/Restart

- `w:lvlOverride` ist ein Override pro Level innerhalb einer `w:num`-Instanz.
- `w:startOverride` setzt den Startwert für diesen Level (z. B. 1).

Praktische Konsequenz für den Patcher:
1. Für UL immer eine Bullet-Definition verwenden (`numFmt=bullet`).
2. Für OL pro Markdown-Listenblock eigene Nummerierungsinstanz erzeugen (kein ungewolltes Kapitel-Weiterzählen).
3. Optional streng: pro OL-Block zusätzlich `lvlOverride(ilvl=0) + startOverride(1)` setzen.

## Quelle (Microsoft Learn)

1. NumberingInstance (`w:num`)
   - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.numberinginstance?view=openxml-3.0.1
   - Kernaussage laut Seite: „Numbering Definition Instance“, serialisiert als `w:num`.

2. AbstractNum (`w:abstractNum`)
   - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.abstractnum?view=openxml-3.0.1
   - Kernaussage laut Seite: „Abstract Numbering Definition“, serialisiert als `w:abstractNum`.

3. LevelOverride (`w:lvlOverride`)
   - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.leveloverride?view=openxml-3.0.1
   - Kernaussage laut Seite: `w:lvlOverride`, enthält u. a. `StartOverrideNumberingValue`.

4. StartOverrideNumberingValue (`w:startOverride`)
   - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.startoverridenumberingvalue?view=openxml-3.0.1
   - Kernaussage laut Seite: „Numbering Level Starting Value Override“, serialisiert als `w:startOverride`.

## Hinweis zur Norm

Die eigentliche normative Spezifikation liegt in ECMA-376 / ISO/IEC 29500 (WordprocessingML). Microsoft Learn beschreibt die OpenXML-SDK-Objektklassen und deren XML-Mapping, was für unsere Implementierung im Patch-Skript praxisnah und ausreichend präzise ist.