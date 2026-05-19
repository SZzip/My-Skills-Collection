<#
.SYNOPSIS
  End-to-End-Selbsttest fuer den word-md-roundtrip-Skill.
.DESCRIPTION
  Erzeugt (oder verwendet) ein Test-DOCX, fuehrt Import (Word->MD) und Rueckweg
  (MD->Word, -NoOverwrite) aus und vergleicht Original vs. Ergebnis als Pandoc-
  Plaintext. Gibt Token-Statistik je erzeugter Datei und einen Kurz-Diff aus.
.EXAMPLE
  pwsh -File Test-Roundtrip.ps1
.EXAMPLE
  pwsh -File Test-Roundtrip.ps1 -Path C:\eigenes\dokument.docx -WorkDir C:\temp\wmrtest
#>
[CmdletBinding()]
param(
    [string]$Path,
    [string]$WorkDir,
    [int]$TargetTokens = 1200,
    [int]$MaxSplitLevel = 4
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\_Common.ps1"
. "$PSScriptRoot\_Import.ps1"
. "$PSScriptRoot\_Export.ps1"
. "$PSScriptRoot\_SelfTest.ps1"

Invoke-WmrInSta `
    -LoadFiles @("$PSScriptRoot\_Common.ps1", "$PSScriptRoot\_Import.ps1", "$PSScriptRoot\_Export.ps1", "$PSScriptRoot\_SelfTest.ps1") `
    -EntryFunction 'Invoke-WmrSelfTest' `
    -BoundParameters $PSBoundParameters
