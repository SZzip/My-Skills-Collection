<#
.SYNOPSIS
  Zerlegt ein Word-Dokument verlustarm in eine hierarchische Markdown-Struktur.
.DESCRIPTION
  Word -> Markdown. Spiegelt die Ueberschriftenhierarchie als Ordnerstruktur,
  splittet nach Token-Budget (Kapitel werden nie mitten im Sinn zerschnitten),
  extrahiert Bilder nach assets/, erzeugt eine Hub-Navigationsdatei und gewinnt
  vorhandene XML-gebundene Platzhalter (Content Controls) zurueck.
  Word-COM benoetigt STA; laeuft unter Windows PowerShell 5.1 und PowerShell 7.
.EXAMPLE
  pwsh -File Convert-WordToMarkdown.ps1 -Path "C:\Doku\Handbuch.docx"
.EXAMPLE
  pwsh -File Convert-WordToMarkdown.ps1 -Path .\Vertrag.docx -MaxSplitLevel 4 -TargetTokens 5000
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$OutDir,
    [ValidateRange(1,6)][int]$MaxSplitLevel = 3,
    [ValidateRange(500,200000)][int]$TargetTokens = 6000,
    [ValidateRange(500,200000)][int]$MaxTokens = 8000
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\_Common.ps1"
. "$PSScriptRoot\_Import.ps1"

Invoke-WmrInSta `
    -LoadFiles @("$PSScriptRoot\_Common.ps1", "$PSScriptRoot\_Import.ps1") `
    -EntryFunction 'Invoke-WmrWordToMarkdown' `
    -BoundParameters $PSBoundParameters
