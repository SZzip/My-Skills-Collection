<#
  _Common.ps1 — Gemeinsame Helfer fuer den Skill "word-md-roundtrip".

  Enthaelt:
   - STA-Bootstrap (Word-COM benoetigt STA; PowerShell 7 ist standardmaessig MTA)
   - Word.Application-COM-Helfer
   - Pandoc-Autocheck + automatische Installation via winget
   - Token-Heuristik
   - Slug/Sanitize fuer Dateinamen
   - Minimaler YAML-Frontmatter-Leser/-Schreiber (auf das vom Skill erzeugte Format zugeschnitten)

  Wird von den Entry-Skripten dot-gesourct. Definiert nur Funktionen/Variablen,
  fuehrt selbst keine Aktion aus.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Konstanten -------------------------------------------------------------

$script:WmrSchemaVersion = 1
$script:WmrXmlNamespace  = 'urn:word-md-roundtrip:fields'
$script:WmrXmlPrefix     = 'f'

# WdContentControlType
$script:WdCC = @{
    RichText     = 0
    Text         = 1
    ComboBox     = 3
    DropdownList = 4
    Date         = 6
    CheckBox     = 8
}
# WdSaveFormat / Alerts / Find
$script:WdFormatDocx       = 16   # wdFormatDocumentDefault (.docx)
$script:WdAlertsNone       = 0
$script:WdDoNotSaveChanges = 0
$script:WdFindStop         = 0

# --- Logging ----------------------------------------------------------------

function Write-WmrInfo { param([string]$Message) Write-Host "[wmr] $Message" }
function Write-WmrWarn { param([string]$Message) Write-Warning "[wmr] $Message" }
function Write-WmrStep { param([string]$Message) Write-Host "[wmr] >> $Message" -ForegroundColor Cyan }

# --- Datei-I/O (UTF-8 ohne BOM) --------------------------------------------

function Get-WmrText {
    param([Parameter(Mandatory)][string]$Path)
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Set-WmrText {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

# --- Token-Heuristik --------------------------------------------------------

function Get-WmrTokenEstimate {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $chars = $Text.Length
    $words = ([regex]::Matches($Text, '\S+')).Count
    $byChars = [Math]::Ceiling($chars / 4.0)
    $byWords = [Math]::Ceiling($words * 1.33)
    return [int][Math]::Max($byChars, $byWords)
}

function Get-WmrWordCount {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    return ([regex]::Matches($Text, '\S+')).Count
}

# --- Slug / sichere Dateinamen ---------------------------------------------

function ConvertTo-WmrPlainTitle {
    <# Liefert den Kapitelnamen als Klartext OHNE Markdown-Formatierung und ohne
       Zeilenumbrueche. Basis fuer Dateiname/Slug/Hub-Anzeige. Die rohe
       Ueberschriftenzeile im Inhalt bleibt davon unberuehrt (Roundtrip-Treue).
       Einzelnes '_' bleibt erhalten (Bezeichner wie IEC_62443 nicht zerstoeren). #>
    param([AllowEmptyString()][string]$Title)
    if ([string]::IsNullOrWhiteSpace($Title)) { return '' }
    $t = $Title
    $t = [regex]::Replace($t, '\{[^}]*\}', '')                       # Pandoc-Attribute {#id .cls}
    $t = [regex]::Replace($t, '!\[([^\]]*)\]\([^)]*\)', '$1')        # Bild -> alt
    $t = [regex]::Replace($t, '\[([^\]]*)\]\([^)]*\)', '$1')         # Link -> Text
    $t = [regex]::Replace($t, '!?\[\[(?:[^\]\|]*\|)?([^\]]*)\]\]', '$1')  # [[a|b]] -> b
    $t = [regex]::Replace($t, '<[^>]+>', '')                         # HTML-Tags
    $t = $t -replace '\*\*', '' -replace '~~', '' -replace '__', '' -replace '`+', '' -replace '\*', ''
    $t = [regex]::Replace($t, '\\([\\`*_{}\[\]()#+.!~>|-])', '$1')   # Pandoc-Escapes \x -> x
    $t = ($t -replace '[\r\n\t]+', ' ') -replace '\s{2,}', ' '       # Zeilenumbrueche raus
    return $t.Trim()
}

function ConvertTo-WmrSafeName {
    <# Macht aus einem Ueberschriftstext einen dateisystem-sicheren Namen.
       Umlaute bleiben erhalten; nur \ / : * ? " < > | und Steuerzeichen werden ersetzt. #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [int]$MaxLength = 80
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = 'Ohne Titel' }
    $clean = $Name -replace '[\\/:\*\?"<>\|]', '-'
    $clean = $clean -replace '[\x00-\x1F]', ' '
    $clean = ($clean -replace '\s+', ' ').Trim().TrimEnd('.', ' ')
    if ($clean.Length -gt $MaxLength) { $clean = $clean.Substring(0, $MaxLength).TrimEnd('.', ' ') }
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = 'Ohne Titel' }
    return $clean
}

function ConvertTo-WmrSlug {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)
    # Zeichencodes statt Literale (encoding-unabhaengig); .Replace ist case-sensitive
    $s = $Name
    $s = $s.Replace(([char]0x00E4).ToString(), 'ae').Replace(([char]0x00F6).ToString(), 'oe').Replace(([char]0x00FC).ToString(), 'ue')
    $s = $s.Replace(([char]0x00C4).ToString(), 'Ae').Replace(([char]0x00D6).ToString(), 'Oe').Replace(([char]0x00DC).ToString(), 'Ue')
    $s = $s.Replace(([char]0x00DF).ToString(), 'ss')
    $s = $s.ToLowerInvariant()
    $s = $s -replace '[^a-z0-9]+', '-'
    $s = $s.Trim('-')
    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'abschnitt' }
    return $s
}

# --- Minimaler YAML-Frontmatter-Reader/-Writer ------------------------------
# Bewusst auf das vom Skill selbst erzeugte Format beschraenkt:
#   - skalare key: value
#   - key: [a, b] (Inline-Liste)
#   - fields: als Liste von Maps (- id: ...\n    key: value ...)

function Format-WmrYamlScalar {
    param($Value)
    if ($null -eq $Value) { return '""' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return "$Value" }
    $s = [string]$Value
    if ($s -eq '') { return '""' }
    if ($s -match '^[\w \-./äöüÄÖÜß]+$' -and $s -notmatch '^\s' -and $s -notmatch '\s$') {
        return $s
    }
    return '"' + ($s -replace '\\', '\\' -replace '"', '\"') + '"'
}

function ConvertTo-WmrFrontmatter {
    <# $Data : [ordered]/hashtable. Schluessel 'fields' wird als Liste von Maps
       behandelt (Array von Hashtables), 'children'/'choices' als Inline-Liste. #>
    param([Parameter(Mandatory)]$Data)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('---')
    foreach ($key in $Data.Keys) {
        $val = $Data[$key]
        if ($key -eq 'fields') {
            [void]$sb.AppendLine('fields:')
            foreach ($f in @($val)) {
                $first = $true
                foreach ($fk in $f.Keys) {
                    $fv = $f[$fk]
                    if ($fk -eq 'choices') {
                        $items = (@($fv) | ForEach-Object { Format-WmrYamlScalar $_ }) -join ', '
                        $line = "$fk`: [$items]"
                    } else {
                        $line = "$fk`: $(Format-WmrYamlScalar $fv)"
                    }
                    if ($first) { [void]$sb.AppendLine("  - $line"); $first = $false }
                    else        { [void]$sb.AppendLine("    $line") }
                }
            }
        }
        elseif ($val -is [System.Array] -or ($val -is [System.Collections.IEnumerable] -and $val -isnot [string])) {
            $items = (@($val) | ForEach-Object { Format-WmrYamlScalar $_ }) -join ', '
            [void]$sb.AppendLine("$key`: [$items]")
        }
        else {
            [void]$sb.AppendLine("$key`: $(Format-WmrYamlScalar $val)")
        }
    }
    [void]$sb.AppendLine('---')
    return $sb.ToString()
}

function Split-WmrFrontmatter {
    <# Trennt Frontmatter und Body. Rueckgabe: PSCustomObject @{ Data; Body; HasFrontmatter } #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Content)
    $nl = "`n"
    $norm = $Content -replace "`r`n", "`n"
    if ($norm -notmatch '^﻿?---\r?\n') {
        return [pscustomobject]@{ Data = @{}; Body = $Content; HasFrontmatter = $false }
    }
    $lines = $norm.Split("`n")
    if ($lines.Count -lt 2 -or ($lines[0].TrimStart([char]0xFEFF)) -ne '---') {
        return [pscustomobject]@{ Data = @{}; Body = $Content; HasFrontmatter = $false }
    }
    $end = -1
    for ($i = 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -eq '---') { $end = $i; break }
    }
    if ($end -lt 0) {
        return [pscustomobject]@{ Data = @{}; Body = $Content; HasFrontmatter = $false }
    }
    $fmLines = $lines[1..($end-1)]
    $body = if ($end + 1 -le $lines.Count - 1) { ($lines[($end+1)..($lines.Count-1)] -join $nl) } else { '' }
    $data = ConvertFrom-WmrYamlBlock -Lines $fmLines
    return [pscustomobject]@{ Data = $data; Body = $body; HasFrontmatter = $true }
}

function ConvertFrom-WmrYamlScalar {
    param([AllowEmptyString()][string]$Raw)
    $v = $Raw.Trim()
    if ($v -eq '' -or $v -eq '""' -or $v -eq "''") { return '' }
    if ($v -eq 'true')  { return $true }
    if ($v -eq 'false') { return $false }
    if ($v -match '^-?\d+$') { return [int]$v }
    if (($v.StartsWith('"') -and $v.EndsWith('"')) -or ($v.StartsWith("'") -and $v.EndsWith("'"))) {
        $inner = $v.Substring(1, $v.Length - 2)
        return ($inner -replace '\\"', '"' -replace '\\\\', '\')
    }
    return $v
}

function ConvertFrom-WmrInlineList {
    param([string]$Raw)
    $inner = $Raw.Trim()
    if ($inner.StartsWith('[')) { $inner = $inner.Substring(1) }
    if ($inner.EndsWith(']'))   { $inner = $inner.Substring(0, $inner.Length - 1) }
    if ([string]::IsNullOrWhiteSpace($inner)) { return @() }
    return @($inner.Split(',') | ForEach-Object { ConvertFrom-WmrYamlScalar $_ })
}

function ConvertFrom-WmrYamlBlock {
    param([string[]]$Lines)
    $data = [ordered]@{}
    $i = 0
    while ($i -lt $Lines.Count) {
        $line = $Lines[$i]
        if ($line.Trim() -eq '' -or $line.TrimStart().StartsWith('#')) { $i++; continue }

        if ($line -match '^([A-Za-z0-9_]+):\s*$') {
            $key = $Matches[1]
            # Block: entweder Liste von Maps (fields) oder Liste von Skalaren
            $items = New-Object System.Collections.ArrayList
            $i++
            $current = $null
            while ($i -lt $Lines.Count) {
                $l = $Lines[$i]
                if ($l -match '^\S') { break }                       # zurueck auf Top-Level
                if ($l.Trim() -eq '') { $i++; continue }
                if ($l -match '^\s*-\s*(.+)$') {
                    $rest = $Matches[1]
                    if ($rest -match '^([A-Za-z0-9_]+):\s*(.*)$') {    # Map-Eintrag
                        if ($null -ne $current) { [void]$items.Add($current) }
                        $current = [ordered]@{}
                        $mk = $Matches[1]; $mv = $Matches[2]
                        $current[$mk] = if ($mv.Trim().StartsWith('[')) { ConvertFrom-WmrInlineList $mv } else { ConvertFrom-WmrYamlScalar $mv }
                    } else {                                          # Skalar-Listeneintrag
                        [void]$items.Add((ConvertFrom-WmrYamlScalar $rest))
                    }
                }
                elseif ($l -match '^\s+([A-Za-z0-9_]+):\s*(.*)$' -and $null -ne $current) {
                    $mk = $Matches[1]; $mv = $Matches[2]
                    $current[$mk] = if ($mv.Trim().StartsWith('[')) { ConvertFrom-WmrInlineList $mv } else { ConvertFrom-WmrYamlScalar $mv }
                }
                $i++
            }
            if ($null -ne $current) { [void]$items.Add($current) }
            $data[$key] = $items.ToArray()
            continue
        }
        if ($line -match '^([A-Za-z0-9_]+):\s*(.+)$') {
            $key = $Matches[1]; $raw = $Matches[2]
            if ($raw.Trim().StartsWith('[')) { $data[$key] = ConvertFrom-WmrInlineList $raw }
            else { $data[$key] = ConvertFrom-WmrYamlScalar $raw }
        }
        $i++
    }
    return $data
}

function Remove-WmrFrontmatter {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Content)
    return (Split-WmrFrontmatter -Content $Content).Body
}

# --- JSON-State -------------------------------------------------------------

function Save-WmrState {
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][string]$Path)
    $json = $State | ConvertTo-Json -Depth 12
    Set-WmrText -Path $Path -Content $json
}
function Read-WmrState {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-WmrText -Path $Path | ConvertFrom-Json)
}

# --- Pandoc -----------------------------------------------------------------

function Assert-WmrPandoc {
    <# Stellt sicher, dass pandoc verfuegbar ist. Installiert es bei Bedarf
       automatisch via winget (ohne Rueckfrage, gemaess Skill-Entscheidung). #>
    $cmd = Get-Command pandoc -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    Write-WmrWarn 'Pandoc nicht gefunden — automatische Installation via winget.'
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw 'Pandoc fehlt und winget ist nicht verfuegbar. Bitte Pandoc manuell installieren: https://pandoc.org/installing.html'
    }
    & winget install --id JohnMacFarlane.Pandoc -e --silent `
        --accept-package-agreements --accept-source-agreements | Out-Host

    # PATH der laufenden Sitzung aktualisieren (winget aktualisiert nur kuenftige Sitzungen)
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user    = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = (@($machine, $user) | Where-Object { $_ }) -join ';'

    $cmd = Get-Command pandoc -ErrorAction SilentlyContinue
    if (-not $cmd) {
        $guess = Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe'
        if (Test-Path -LiteralPath $guess) { return $guess }
        throw 'Pandoc-Installation fehlgeschlagen oder nicht im PATH. Bitte Terminal neu starten und erneut versuchen.'
    }
    return $cmd.Source
}

function Invoke-WmrPandoc {
    param([Parameter(Mandatory)][string]$PandocPath, [Parameter(Mandatory)][string[]]$Arguments)
    Write-WmrInfo "pandoc $($Arguments -join ' ')"
    & $PandocPath @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Pandoc-Aufruf fehlgeschlagen (Exitcode $LASTEXITCODE)." }
}

# --- Word-COM ---------------------------------------------------------------

function New-WmrWord {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try { $word.DisplayAlerts = $script:WdAlertsNone } catch {}
    try { $word.ScreenUpdating = $false } catch {}
    return $word
}

function Stop-WmrWord {
    param($Word)
    if ($null -eq $Word) { return }
    try { $Word.Quit($script:WdDoNotSaveChanges) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function Open-WmrDoc {
    param($Word, [Parameter(Mandatory)][string]$Path, [bool]$ReadOnly = $true)
    $full = (Resolve-Path -LiteralPath $Path).Path
    return $Word.Documents.Open($full, $false, $ReadOnly)
}

function Test-WmrGenericHeading {
    <# True, wenn der Text eine generische TOC-/Gliederungs-Ueberschrift ist
       (kein sinnvoller Dokumenttitel, oft das geflachte Word-Inhaltsverzeichnis). #>
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $t = ($Text -replace '\{[^}]*\}', '' -replace '[*_~`#]', '')
    $t = ($t -replace '\s+', ' ').Trim().ToLowerInvariant()
    $generic = @('content', 'contents', 'inhalt', 'inhaltsverzeichnis',
                 'table of contents', 'agenda', 'index', 'toc', 'gliederung')
    return ($generic -contains $t)
}

function Get-WmrCleanFileTitle {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $stripped = [regex]::Replace($base, '^\s*\d+[\s._)\-]+', '').Trim()
    if ([string]::IsNullOrWhiteSpace($stripped)) { return $base }
    return $stripped
}

function Get-WmrDocTitle {
    param($Doc, [Parameter(Mandatory)][string]$FallbackFileName)
    # 1. Dokumenteigenschaft "Title"
    try {
        $t = $Doc.BuiltInDocumentProperties.Item('Title').Value
        if ($t -and "$t".Trim() -and -not (Test-WmrGenericHeading "$t")) { return "$t".Trim() }
    } catch {}
    # 2. Erster Absatz mit Stil "Titel"/"Title" (generische uebergehen)
    try {
        foreach ($p in $Doc.Paragraphs) {
            $sn = ''
            try { $sn = $p.Style.NameLocal } catch {}
            if ($sn -match '^(Titel|Title)$') {
                $txt = ($p.Range.Text -replace '[\r\n\a]', '').Trim()
                if ($txt -and -not (Test-WmrGenericHeading $txt)) { return $txt }
            }
            break  # nur der erste Absatz ist als Titel plausibel
        }
    } catch {}
    # 3. Erste sinnvolle Heading-1 (generische TOC-Ueberschriften ueberspringen)
    try {
        foreach ($p in $Doc.Paragraphs) {
            $sn = ''
            try { $sn = $p.Style.NameLocal } catch {}
            if ($sn -match '^(Überschrift 1|Heading 1)$') {
                $txt = ($p.Range.Text -replace '[\r\n\a]', '').Trim()
                if ($txt -and -not (Test-WmrGenericHeading $txt)) { return $txt }
            }
        }
    } catch {}
    # 4. Dateiname (fuehrende Nummerierung entfernen)
    return (Get-WmrCleanFileTitle -FileName $FallbackFileName)
}

function Get-WmrHeadingStyleMap {
    <# Liefert Map Ebene(1..9) -> lokaler Word-Stilname (informativ fuer Frontmatter). #>
    param($Doc)
    $map = [ordered]@{}
    for ($lvl = 1; $lvl -le 9; $lvl++) {
        try {
            $st = $Doc.Styles.Item("Heading $lvl")
            $map["$lvl"] = $st.NameLocal
        } catch {
            $map["$lvl"] = "Heading $lvl"
        }
    }
    return $map
}

# --- STA-Bootstrap ----------------------------------------------------------

function Test-WmrSta {
    return ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA)
}

function Invoke-WmrInSta {
    <# Fuehrt eine Entry-Funktion garantiert in einem STA-Thread aus.
       Ist die aktuelle Sitzung bereits STA (z. B. Windows PowerShell 5.1),
       wird direkt aufgerufen. Sonst (PowerShell 7 = MTA) in einem STA-Runspace.
       -LoadFiles : Dateien, die im Runspace (in Reihenfolge) dot-gesourct werden,
                    damit die Entry-Funktion + Helfer dort verfuegbar sind. #>
    param(
        [Parameter(Mandatory)][string[]]$LoadFiles,
        [Parameter(Mandatory)][string]$EntryFunction,
        [Parameter(Mandatory)][hashtable]$BoundParameters
    )

    if (Test-WmrSta) {
        return & $EntryFunction @BoundParameters
    }

    Write-WmrInfo 'PowerShell laeuft im MTA-Modus — starte STA-Runspace fuer Word-COM.'
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    foreach ($f in $LoadFiles) { [void]$ps.AddScript(". `"$f`"").AddStatement() }
    [void]$ps.AddCommand($EntryFunction).AddParameters($BoundParameters)

    $async  = $ps.BeginInvoke()
    $iInfo = 0; $iWarn = 0; $iErr = 0
    while (-not $async.IsCompleted) {
        Start-Sleep -Milliseconds 200
        while ($iInfo -lt $ps.Streams.Information.Count) { Write-Host $ps.Streams.Information[$iInfo].MessageData; $iInfo++ }
        while ($iWarn -lt $ps.Streams.Warning.Count)     { Write-Warning $ps.Streams.Warning[$iWarn].Message;       $iWarn++ }
        while ($iErr  -lt $ps.Streams.Error.Count)       { Write-Host ("ERR: " + $ps.Streams.Error[$iErr]); $iErr++ }
    }
    try {
        $result = $ps.EndInvoke($async)
    } catch {
        throw
    } finally {
        while ($iInfo -lt $ps.Streams.Information.Count) { Write-Host $ps.Streams.Information[$iInfo].MessageData; $iInfo++ }
        while ($iWarn -lt $ps.Streams.Warning.Count)     { Write-Warning $ps.Streams.Warning[$iWarn].Message;       $iWarn++ }
        $ps.Dispose(); $rs.Dispose()
    }
    if ($ps.HadErrors -and $ps.Streams.Error.Count -gt 0) {
        throw $ps.Streams.Error[0]
    }
    return $result
}
