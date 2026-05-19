<#
.SYNOPSIS
  Fuehrt die Markdown-Struktur wieder zu einem Word-Dokument zusammen (Rueckweg).
.DESCRIPTION
  Markdown -> Word. Loest die ![[...]]-Embeds rekursiv auf, fuehrt alles in der
  urspruenglichen Lesereihenfolge zusammen, rendert per Pandoc im Originalstil
  (reference.docx) und injiziert die in der Hub-Frontmatter (fields:) definierten
  Platzhalter als XML-gebundene Inhaltssteuerelemente.

  ACHTUNG: Ueberschreibt standardmaessig das Original-DOCX. Vorher wird automatisch
  eine zeitgestempelte Kopie in <DocDir>\.backup\ abgelegt. Mit -NoOverwrite wird
  stattdessen "<Titel> (roundtrip).docx" neben dem Original geschrieben. Mit
  -OutFile <Pfad> wird genau dorthin geschrieben (Original unangetastet) — ideal,
  um nicht in einen synchronisierten Ordner (OneDrive) zu schreiben.
.EXAMPLE
  pwsh -File Convert-MarkdownToWord.ps1 -DocDir "C:\Doku\01 - Handbuch"
.EXAMPLE
  pwsh -File Convert-MarkdownToWord.ps1 -DocDir ".\01 - Vertrag" -NoOverwrite
.EXAMPLE
  pwsh -File Convert-MarkdownToWord.ps1 -DocDir ".\01 - Vertrag" -OutFile C:\temp\Vertrag.docx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DocDir,
    [switch]$NoOverwrite,
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\_Common.ps1"
. "$PSScriptRoot\_Export.ps1"

Invoke-WmrInSta `
    -LoadFiles @("$PSScriptRoot\_Common.ps1", "$PSScriptRoot\_Export.ps1") `
    -EntryFunction 'Invoke-WmrMarkdownToWord' `
    -BoundParameters $PSBoundParameters
