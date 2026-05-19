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
| [`word-md-roundtrip`](skills/word-md-roundtrip/SKILL.md) | Lossless round-trip between Microsoft Word (`.docx`) and a chapter-split, AI-context-friendly Markdown structure. PowerShell + Word COM + Pandoc; hierarchical folder split by token budget, navigation hub with WikiLinks, image extraction, and XML-bound placeholders (content controls). |

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

See [`skills/word-md-roundtrip/SKILL.md`](skills/word-md-roundtrip/SKILL.md) for
the full workflow, parameters, placeholder/field model, and documented scope
limits.

## Adding a new skill

Create `skills/<skill-name>/` containing a `SKILL.md` with YAML frontmatter
(`name` and `description`) and any supporting `scripts/`. Add a row to the
**Skills** table above, then commit and push.

## License

[MIT](LICENSE) © 2026 SZzip
