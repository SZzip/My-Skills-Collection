---
name: word-md-roundtrip
description: >-
  Roundtrip zwischen Microsoft Word (.docx) und Markdown. Verwende diesen Skill,
  wenn ein Word-Dokument verlustarm in eine hierarchische, KI-kontextschonende
  Markdown-Struktur zerlegt werden soll ("Word in Markdown zerlegen", "docx
  aufteilen/splitten", "Word importieren"), oder wenn eine zuvor erzeugte
  Markdown-Struktur wieder zu einem Word-Dokument zusammengefuehrt werden soll
  ("Markdown zurueck nach Word", "Roundtrip", "wieder zu docx zusammenfuehren").
  Ebenfalls verwenden bei XML-gebundenen Platzhaltern/Eingabefeldern in Word
  (Content Controls, Custom XML) im Zusammenhang mit Markdown-Bearbeitung.
  Nur Windows mit installiertem Microsoft Word.
---

# Word ⇆ Markdown Roundtrip

Zerlegt ein Word-Dokument in eine rekursive Ordner-/Markdown-Struktur (Kapitel
werden nie mitten im Sinn zerschnitten, Zielgroesse token-basiert) und fuehrt sie
verlustarm wieder zu Word zusammen. Skripte: PowerShell + Word.Application-COM +
Pandoc. Laeuft unter Windows PowerShell 5.1 **und** PowerShell 7 (Word-COM-STA
wird automatisch sichergestellt).

Skriptverzeichnis: `~/.claude/skills/word-md-roundtrip/scripts`
(absolut: `C:\Users\SecuSolveSebastianSc\.claude\skills\word-md-roundtrip\scripts`).

## Voraussetzungen

- Microsoft Word installiert (COM-Zugriff).
- Pandoc — wird bei Fehlen automatisch via `winget` installiert (ohne Rueckfrage).
- Aufruf bevorzugt mit `pwsh -NoProfile -File <skript>`; `powershell.exe` geht auch.

## Schritt 1 — Import: Word → Markdown

```powershell
pwsh -NoProfile -File "C:\Users\SecuSolveSebastianSc\.claude\skills\word-md-roundtrip\scripts\Convert-WordToMarkdown.ps1" -Path "<PFAD\zur\Datei.docx>"
```

Parameter:

| Parameter | Default | Bedeutung |
|---|---|---|
| `-Path` | (Pflicht) | Quelle (.docx) |
| `-OutDir` | Ordner der Quelle | Zielordner; darin entsteht `01 - <Titel>\` |
| `-MaxSplitLevel` | `3` | Maximale Tiefe, bis zu der bei zu grossen Kapiteln gesplittet wird (1=nur H1) |
| `-TargetTokens` | `6000` | Ziel-Tokenfenster je Datei (Heuristik) |
| `-MaxTokens` | `8000` | Harte Obergrenze; danach Absatz-Split in `(Teil k)` |

Ergebnis-Struktur:

```
01 - <Titel>/
  00 - <Titel> Hub.md        # Navigation (WikiLinks) + zentrale fields:-Registry
  01 - Vorspann.md           # nur falls Inhalt vor erster Ueberschrift 1
  02 - <Kapitel>.md          # Einzeldatei, wenn <= Token-Ziel
  03 - <Kapitel>/            # Ordner, wenn aufgeteilt
    00 - <Kapitel>.md        #   Eigener Text + ![[..]]-Embeds der Unterkapitel
    01 - <Unterkapitel>.md
  assets/                    # extrahierte Bilder (relative Links)
  .roundtrip/state.json      # internes Statefile (Rueckweg)
  .backup/                   # Backups beim Rueckweg
```

### Pflicht-Nachbearbeitung durch dich (KI)

Nach erfolgreichem Import enthaelt die Hub-Datei je Kapitel den Platzhalter
`<!-- summary:NN -->`. **Lies jedes Kapitel** und ersetze jeden Platzhalter
durch eine praegnante deutsche 1–2-Satz-Zusammenfassung (worum geht es, was
findet die KI dort). Nur den Platzhalter ersetzen, WikiLink und Struktur
unveraendert lassen. Das macht spaetere Ladevorgaenge kontextschonend.

## Schritt 2 — Rueckweg: Markdown → Word

```powershell
pwsh -NoProfile -File "C:\Users\SecuSolveSebastianSc\.claude\skills\word-md-roundtrip\scripts\Convert-MarkdownToWord.ps1" -DocDir "<...\01 - Titel>"
```

| Parameter | Bedeutung |
|---|---|
| `-DocDir` | Der beim Import erzeugte `01 - <Titel>`-Ordner (Pflicht) |
| `-NoOverwrite` | Schreibt `"<Titel> (roundtrip).docx"` neben das Original statt es zu ueberschreiben |
| `-OutFile <Pfad>` | Schreibt das Ergebnis genau dorthin, Original unangetastet (z. B. um nicht in OneDrive zu schreiben) |

**Achtung — destruktiv:** Ohne `-NoOverwrite` wird das Original-DOCX
**ueberschrieben**. Vorher legt der Skript automatisch ein zeitgestempeltes
Backup in `<DocDir>\.backup\` an. Weise den Nutzer beim ersten Mal kurz darauf
hin. Gerendert wird im Originalstil (Pandoc `--reference-doc`).

## Platzhalter / XML-gebundene Eingabefelder

- Im Markdown wird ein Platzhalter als `{{id}}` im Fliesstext geschrieben.
- Die Definition jedes Feldes steht zentral in der **Frontmatter der Hub-Datei**
  unter `fields:` (id, label, type, default, value, required, choices, format).
  `type` ∈ `text | multiline | rich | date | choice | checkbox`.
- Beim Rueckweg wird daraus ein Custom XML Part erzeugt und jedes `{{id}}` durch
  ein gebundenes Word-Inhaltssteuerelement ersetzt (Werte aus `value`/`default`).
- Bidirektional: bereits gebundene Controls im Quell-DOCX werden beim Import als
  `fields:` + aktuelle Werte zurueckgewonnen und der Text durch `{{id}}` ersetzt.
- Feldwert aendern = `value` in der Hub-Frontmatter anpassen, dann Rueckweg.
- `{{ }}` ist reserviert; literale doppelte geschweifte Klammern im Fachtext
  vermeiden bzw. anders schreiben.

## Selbsttest

```powershell
pwsh -NoProfile -File "C:\Users\SecuSolveSebastianSc\.claude\skills\word-md-roundtrip\scripts\Test-Roundtrip.ps1"
```

Erzeugt ein Beispiel-DOCX (Ueberschriften, Tabelle, Liste, Fussnote, gebundenes
Feld), macht Import + Rueckweg und gibt Token-Statistik sowie einen
Plaintext-Vergleich aus. Mit `-Path <docx>` gegen ein echtes Dokument testbar.

## Grenzen (bewusst, dem Nutzer bei Bedarf nennen)

- Nicht round-getrippt: Kommentare, Aenderungsverfolgung, Kopf-/Fusszeilen,
  Word-Felder, automatische Inhaltsverzeichnisse als lebende Felder, sehr
  komplexe Stile. Erhalten bleiben: Text, Ueberschriftenstruktur, Tabellen,
  Listen, Bilder, Fuss-/Endnoten, die Markdown-Formatierung und (separat) die
  XML-gebundenen Platzhalter.
- Formeln/Gleichungen: Best-Effort ueber Pandoc, ggf. Qualitaetsverlust.
- `[[WikiLinks]]`/`![[Embeds]]` sind Nicht-Standard-Markdown; nur dieser Skill
  loest sie beim Rueckweg auf. Bestehende Struktur/Numerierung nicht manuell
  umbenennen, sonst bricht die Embed-Aufloesung.
- Rueckweg ohne `-NoOverwrite` ueberschreibt das Original (Backup in `.backup\`).
- Platzhalter-Injektion braucht Word-COM; faellt Word aus, entfaellt nur dieser
  Schritt, der Rest des Rueckwegs funktioniert.
