# My Skills Collection

A personal collection of [Claude Code](https://claude.com/claude-code) skills,
versioned and shareable. Each skill lives in its own folder under `skills/` and
can be installed by copying it into your local Claude skills directory.

## Repository structure

```
My-Skills-Collection/
├── LICENSE
├── README.md
└── skills/
    └── <skill-name>/
        ├── SKILL.md          # skill manifest (name + description frontmatter)
        └── scripts/          # optional helper scripts
```

## Skills

| Skill | Description |
|-------|-------------|
| [`word-md-roundtrip`](skills/word-md-roundtrip/SKILL.md) | Round-trip between Microsoft Word (`.docx`) and a chapter-split, AI-context-friendly Markdown structure. PowerShell + Word COM + Pandoc; hierarchical folder split by token budget, navigation hub with WikiLinks, image extraction, and XML-bound placeholders (content controls). **Not lossless** — Markdown-representable content is preserved; see [Limitations](#limitations--not-lossless). |

## Installation

Claude Code discovers skills placed in your user skills directory. To install a
skill from this collection, copy its folder there:

**Windows (PowerShell):**

```powershell
Copy-Item -Recurse -Force `
  ".\skills\word-md-roundtrip" `
  "$env:USERPROFILE\.claude\skills\"
```

**macOS / Linux:**

```bash
cp -r ./skills/word-md-roundtrip ~/.claude/skills/
```

Once installed, the skill is invoked via `/word-md-roundtrip` or picked up
automatically by Claude when a matching request is made.

> Note: scripts are stored as UTF-8 **with BOM** so they parse correctly under
> both Windows PowerShell 5.1 and PowerShell 7. Preserve the byte content when
> copying (a plain file copy does this; avoid re-encoding round-trips).

## Skill: word-md-roundtrip

Splits a Word document into a recursive Markdown folder structure (chapters are
never cut mid-meaning, sized to a token budget), and merges it back into Word.

The round trip is **not lossless**: it faithfully preserves text, heading
structure, tables, lists, images, foot-/endnotes and Markdown-representable
formatting (plus separately tracked XML-bound placeholders), but everything that
Markdown/Pandoc cannot represent is dropped or approximated — see
[Limitations](#limitations--not-lossless).

**Requirements:** Windows · Microsoft Word (COM automation) · PowerShell 5.1 or 7
· Pandoc (auto-installed via `winget` if missing).

**Quick usage** (run from the installed skill's `scripts/` folder):

```powershell
# Word -> Markdown (non-destructive; creates "01 - <Title>/" next to the source)
pwsh -NoProfile -File .\Convert-WordToMarkdown.ps1 -Path "C:\Docs\Manual.docx"

# Markdown -> Word (overwrites the original; auto-backup in .backup\)
pwsh -NoProfile -File .\Convert-MarkdownToWord.ps1 -DocDir ".\01 - Manual"

# ...or write the result elsewhere, leaving the original untouched
pwsh -NoProfile -File .\Convert-MarkdownToWord.ps1 -DocDir ".\01 - Manual" -OutFile "C:\out\Manual.docx"

# End-to-end self-test (generates a sample document)
pwsh -NoProfile -File .\Test-Roundtrip.ps1
```

### Limitations — not lossless

By deliberate scope, the round trip does **not** preserve:

- **Word-only constructs:** comments, tracked changes, headers/footers, Word
  fields, and live tables of contents. An auto-generated TOC chapter is
  intentionally discarded (it is regenerable in Word and only adds noise).
- **Complex/custom styling:** formatting is reduced to what Markdown + Pandoc can
  express. The original document is used as a Pandoc *reference doc* so the
  result resembles it, but exact layout, spacing, themes and advanced styles are
  approximated, not byte-preserved.
- **Equations / formulas:** best-effort via Pandoc; may degrade or be lost.
- **Placeholders outside the body flow:** content controls on cover pages, in
  text boxes, or in headers/footers are recovered into the hub `fields:` registry
  as metadata, but produce no `{{token}}` in the body and are therefore **not
  re-inserted** on the reverse pass (Pandoc does not export those regions).
- **Title & names:** the document title is taken from Word metadata — any typo in
  the source carries through. File/folder/hub names are stripped of Markdown
  formatting and newlines (the raw heading line is kept verbatim in the body, so
  formatting survives the round trip there).
- **Token sizing is heuristic:** chapter split sizes use an approximate token
  estimate, not an exact tokenizer.

Operational notes:

- **Markdown → Word overwrites the original by default** (a timestamped copy is
  written to `.backup\`). Use `-NoOverwrite` or `-OutFile <path>` to avoid this.
- `[[WikiLinks]]` / `![[embeds]]` are non-standard Markdown understood only by
  this skill's reverse pass. Do **not** manually rename the generated folders or
  numbering, or the reverse merge will break.
- `{{ }}` is reserved for placeholders; avoid literal double curly braces in body
  text.

See [`skills/word-md-roundtrip/SKILL.md`](skills/word-md-roundtrip/SKILL.md) for
the full workflow, parameters, placeholder/field model, and the complete list of
documented scope limits.

## Adding a new skill

Create `skills/<skill-name>/` containing a `SKILL.md` with YAML frontmatter
(`name` and `description`) and any supporting `scripts/`. Add a row to the
**Skills** table above, then commit and push.

## License

[MIT](LICENSE) © 2026 SZzip
