<#
  _Export.ps1 — Markdown -> Word (Rueckweg / Zusammenfuehrung).
  Definiert nur Funktionen. Setzt _Common.ps1 als bereits geladen voraus.
#>

function Remove-WmrFootnoteAppendix {
    param([string]$Text)
    return [regex]::Replace($Text, '\r?\n?<!-- wmr:footnotes -->.*?<!-- /wmr:footnotes -->', '', 'Singleline')
}

function ConvertTo-WmrRootRelativeAssets {
    param([string]$Text)
    $t = [regex]::Replace($Text, '(\!\[[^\]]*\]\()\s*(?:\.\./)+assets/', { param($m) $m.Groups[1].Value + 'assets/' })
    $t = [regex]::Replace($t, '(<img[^>]*\ssrc=")(?:\.\./)+assets/', { param($m) $m.Groups[1].Value + 'assets/' })
    return $t
}

function ConvertTo-WmrMarkdownImages {
    <# Pandocs DOCX-Writer ignoriert rohes HTML-<img>. Daher <img>-Tags in
       Markdown-Bildsyntax ![alt](src){width=.. height=..} umwandeln, damit das
       Bild beim Rueckweg eingebettet wird (Groesse via Pandoc-Attribut erhalten). #>
    param([string]$Text)
    return [regex]::Replace($Text, '<img\b[^>]*?/?>', {
        param($m)
        $tag = $m.Value
        $src = ''; $alt = ''; $w = ''; $h = ''
        $x = [regex]::Match($tag, 'src\s*=\s*"([^"]*)"');  if ($x.Success) { $src = $x.Groups[1].Value }
        if (-not $src) { $x = [regex]::Match($tag, "src\s*=\s*'([^']*)'"); if ($x.Success) { $src = $x.Groups[1].Value } }
        $x = [regex]::Match($tag, 'alt\s*=\s*"([^"]*)"');  if ($x.Success) { $alt = $x.Groups[1].Value }
        $st = [regex]::Match($tag, 'style\s*=\s*"([^"]*)"')
        if ($st.Success) {
            $wm = [regex]::Match($st.Groups[1].Value, 'width\s*:\s*([^;]+)');  if ($wm.Success) { $w = $wm.Groups[1].Value.Trim() }
            $hm = [regex]::Match($st.Groups[1].Value, 'height\s*:\s*([^;]+)'); if ($hm.Success) { $h = $hm.Groups[1].Value.Trim() }
        }
        if (-not $w) { $x = [regex]::Match($tag, '\bwidth\s*=\s*"([^"]*)"');  if ($x.Success) { $w = $x.Groups[1].Value } }
        if (-not $h) { $x = [regex]::Match($tag, '\bheight\s*=\s*"([^"]*)"'); if ($x.Success) { $h = $x.Groups[1].Value } }
        if (-not $src) { return '' }
        $parts = @()
        if ($w) { $parts += "width=`"$w`"" }
        if ($h) { $parts += "height=`"$h`"" }
        $attr = if ($parts.Count) { '{' + ($parts -join ' ') + '}' } else { '' }
        $altEsc = ($alt -replace '[\[\]]', ' ').Trim()
        return "![$altEsc]($src)$attr"
    }, 'Singleline')
}

function Resolve-WmrEntry {
    <# Liest eine .md (relativer Pfad ohne .md), entfernt Frontmatter +
       Fussnoten-Anhang und loest ![[...]]-Embeds rekursiv auf. #>
    param([string]$DocRootAbs, [string]$RelNoExt, [System.Collections.Generic.HashSet[string]]$Seen)
    $file = Join-Path $DocRootAbs ($RelNoExt -replace '/', '\')
    $file = "$file.md"
    if (-not (Test-Path -LiteralPath $file)) {
        Write-WmrWarn "Embed-Ziel fehlt, uebersprungen: $RelNoExt"
        return ''
    }
    $key = $RelNoExt.ToLowerInvariant()
    if ($Seen.Contains($key)) {
        Write-WmrWarn "Zyklischer Embed erkannt, uebersprungen: $RelNoExt"
        return ''
    }
    [void]$Seen.Add($key)

    $raw  = Get-WmrText -Path $file
    $body = Remove-WmrFrontmatter -Content $raw
    $body = Remove-WmrFootnoteAppendix -Text $body
    $body = ConvertTo-WmrRootRelativeAssets -Text $body

    $lines = ($body -replace "`r`n", "`n").Split("`n")
    $out = New-Object System.Collections.ArrayList
    foreach ($ln in $lines) {
        $m = [regex]::Match($ln, '^\s*!\[\[([^\]\|]+?)(?:\|[^\]]*)?\]\]\s*$')
        if ($m.Success) {
            $target = $m.Groups[1].Value.Trim()
            $resolved = Resolve-WmrEntry -DocRootAbs $DocRootAbs -RelNoExt $target -Seen $Seen
            [void]$out.Add($resolved)
        } else {
            [void]$out.Add($ln)
        }
    }
    [void]$Seen.Remove($key)
    # "## Unterkapitel"-Ueberschrift ist reine Navigationshilfe -> entfernen
    $joined = ($out -join "`n")
    $joined = [regex]::Replace($joined, '(?m)^\#\#\s+Unterkapitel\s*$', '')
    return $joined.Trim("`n")
}

function Get-WmrHubFields {
    param([string]$DocRootAbs)
    $hub = Get-ChildItem -LiteralPath $DocRootAbs -Filter '00 - * Hub.md' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $hub) { return @() }
    $parsed = Split-WmrFrontmatter -Content (Get-WmrText -Path $hub.FullName)
    if ($parsed.Data.Contains('fields')) { return @($parsed.Data['fields']) }
    return @()
}

function Add-WmrContentControls {
    <# Word-COM-Nachlauf: ersetzt {{id}} durch XML-gebundene Inhaltssteuerelemente. #>
    param($Doc, $Fields)
    if (-not $Fields -or @($Fields).Count -eq 0) { return }

    $ns = $script:WmrXmlNamespace
    $pfx = $script:WmrXmlPrefix
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<$pfx`:fields xmlns:$pfx=`"$ns`">")
    foreach ($f in @($Fields)) {
        $val = ''
        if ($f.Contains('value') -and "$($f.value)" -ne '') { $val = "$($f.value)" }
        elseif ($f.Contains('default')) { $val = "$($f.default)" }
        if ($f.type -eq 'checkbox') {
            $val = ($(if ("$val" -match '^(true|1|ja|yes|x)$') { 'true' } else { 'false' }))
        }
        $esc = [System.Security.SecurityElement]::Escape($val)
        [void]$sb.Append("<$pfx`:$($f.id)>$esc</$pfx`:$($f.id)>")
    }
    [void]$sb.Append("</$pfx`:fields>")
    $part = $Doc.CustomXMLParts.Add($sb.ToString())
    $prefixMapping = "xmlns:$pfx='$ns'"

    foreach ($f in @($Fields)) {
        $ccType = switch ($f.type) {
            'rich'      { $script:WdCC.RichText }
            'multiline' { $script:WdCC.Text }
            'choice'    { $script:WdCC.DropdownList }
            'date'      { $script:WdCC.Date }
            'checkbox'  { $script:WdCC.CheckBox }
            default     { $script:WdCC.Text }
        }
        $xpath = "/$pfx`:fields[1]/$pfx`:$($f.id)[1]"
        $token = "{{$($f.id)}}"
        $guard = 0
        $rng = $Doc.Content
        $find = $rng.Find
        $find.ClearFormatting()
        $find.Text = $token
        $find.Forward = $true
        $find.MatchWildcards = $false
        $find.Wrap = $script:WdFindStop
        while ($find.Execute() -and $guard -lt 500) {
            $guard++
            $cc = $Doc.ContentControls.Add($ccType, $rng)
            try { $cc.Tag = $f.id } catch {}
            try { $cc.Title = $(if ($f.Contains('label') -and "$($f.label)" -ne '') { "$($f.label)" } else { $f.id }) } catch {}
            if ($f.type -eq 'multiline') { try { $cc.MultiLine = $true } catch {} }
            if ($f.type -eq 'date' -and $f.Contains('format') -and "$($f.format)" -ne '') {
                try { $cc.DateDisplayFormat = "$($f.format)" } catch {}
            }
            if ($f.type -eq 'choice' -and $f.Contains('choices')) {
                try {
                    $cc.DropdownListEntries.Clear()
                    foreach ($ch in @($f.choices)) { [void]$cc.DropdownListEntries.Add("$ch", "$ch") }
                } catch {}
            }
            try {
                if ($f.type -eq 'checkbox') {
                    $checked = ("$($f.value)" -match '^(true|1|ja|yes|x)$')
                    try { $cc.Checked = [bool]$checked } catch {}
                } else {
                    $disp = ''
                    if ($f.Contains('value') -and "$($f.value)" -ne '') { $disp = "$($f.value)" }
                    elseif ($f.Contains('default')) { $disp = "$($f.default)" }
                    if ($disp -ne '') { $cc.Range.Text = $disp }
                }
            } catch {}
            try { $null = $cc.XMLMapping.SetMapping($xpath, $prefixMapping, $part) } catch {
                Write-WmrWarn "XML-Bindung fuer '$($f.id)' fehlgeschlagen: $($_.Exception.Message)"
            }
            $rng.Collapse(0)  # wdCollapseEnd
            $rng = $Doc.Content
            $rng.Start = $cc.Range.End
            $find = $rng.Find
            $find.ClearFormatting()
            $find.Text = $token
            $find.Forward = $true
            $find.MatchWildcards = $false
            $find.Wrap = $script:WdFindStop
        }
    }
}

function Invoke-WmrMarkdownToWord {
    param(
        [Parameter(Mandatory)][string]$DocDir,
        [switch]$NoOverwrite,
        [string]$OutFile
    )

    if (-not (Test-Path -LiteralPath $DocDir)) { throw "Dokumentordner nicht gefunden: $DocDir" }
    $docRootAbs = (Resolve-Path -LiteralPath $DocDir).Path
    $statePath  = Join-Path $docRootAbs '.roundtrip\state.json'
    if (-not (Test-Path -LiteralPath $statePath)) {
        throw "Kein .roundtrip\state.json in '$docRootAbs'. Wurde der Ordner mit Convert-WordToMarkdown erzeugt?"
    }
    $state  = Read-WmrState -Path $statePath
    $pandoc = Assert-WmrPandoc

    $rtDir   = Join-Path $docRootAbs '.roundtrip'
    $refDocx = Join-Path $docRootAbs ($state.referenceDocx -replace '/', '\')
    if (-not (Test-Path -LiteralPath $refDocx)) { throw "reference.docx fehlt: $refDocx" }

    Write-WmrStep 'Loese Embeds auf und fuehre Markdown zusammen.'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $sb = [System.Text.StringBuilder]::new()
    foreach ($e in $state.topEntries) {
        $resolved = Resolve-WmrEntry -DocRootAbs $docRootAbs -RelNoExt $e.target -Seen $seen
        if ($resolved.Trim()) {
            if ($sb.Length -gt 0) { [void]$sb.AppendLine(''); [void]$sb.AppendLine('') }
            [void]$sb.Append($resolved.Trim("`n"))
        }
    }
    # Globale Fussnoten-Definitionen einmalig anhaengen
    if ($state.PSObject.Properties.Name -contains 'footnotes' -and @($state.footnotes).Count -gt 0) {
        [void]$sb.AppendLine(''); [void]$sb.AppendLine('')
        foreach ($fn in $state.footnotes) { [void]$sb.AppendLine($fn.text) }
    }
    $combined = Join-Path $rtDir 'combined.md'
    Set-WmrText -Path $combined -Content (ConvertTo-WmrMarkdownImages -Text $sb.ToString())

    Write-WmrStep 'Rendere DOCX (Pandoc, Originalstil via reference-doc).'
    $rebuilt = Join-Path $rtDir 'rebuilt.docx'
    Invoke-WmrPandoc -PandocPath $pandoc -Arguments @(
        $combined, '-f', 'markdown+pipe_tables', '-t', 'docx',
        "--reference-doc=$refDocx", "--resource-path=$docRootAbs",
        '-o', $rebuilt
    )

    # Platzhalter-Injektion via Word-COM
    $fields = Get-WmrHubFields -DocRootAbs $docRootAbs
    if (@($fields).Count -gt 0) {
        Write-WmrStep "Injiziere $(@($fields).Count) XML-gebundene Eingabefelder (Word-COM)."
        $word = $null
        try {
            $word = New-WmrWord
            $doc  = Open-WmrDoc -Word $word -Path $rebuilt -ReadOnly $false
            Add-WmrContentControls -Doc $doc -Fields $fields | Out-Null
            $doc.Save()
            $doc.Close($script:WdDoNotSaveChanges)
        } catch {
            Write-WmrWarn "Platzhalter-Injektion fehlgeschlagen (Schritt wird uebersprungen): $($_.Exception.Message)"
        } finally {
            Stop-WmrWord -Word $word
        }
    }

    if ($OutFile) {
        $target = $OutFile
        if (-not [System.IO.Path]::GetExtension($target)) { $target = "$target.docx" }
        $od = Split-Path -Parent $target
        if ($od -and -not (Test-Path -LiteralPath $od)) { New-Item -ItemType Directory -Path $od -Force | Out-Null }
        Copy-Item -LiteralPath $rebuilt -Destination $target -Force
        Write-WmrInfo "Ausgabe (Original unangetastet): $target"
        return [pscustomobject]@{ Output = $target; Overwritten = $false }
    }

    if ($NoOverwrite) {
        $alt = Join-Path (Split-Path -Parent $state.sourceDocx) ("$([System.IO.Path]::GetFileNameWithoutExtension($state.sourceDocx)) (roundtrip).docx")
        Copy-Item -LiteralPath $rebuilt -Destination $alt -Force
        Write-WmrInfo "Ausgabe (Original unangetastet): $alt"
        return [pscustomobject]@{ Output = $alt; Overwritten = $false }
    }

    # Backup + Original ueberschreiben
    $backupDir = Join-Path $docRootAbs '.backup'
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $origName = [System.IO.Path]::GetFileNameWithoutExtension($state.sourceDocx)
    if (Test-Path -LiteralPath $state.sourceDocx) {
        $backup = Join-Path $backupDir "$origName-$stamp.docx"
        Copy-Item -LiteralPath $state.sourceDocx -Destination $backup -Force
        Write-WmrInfo "Backup angelegt: $backup"
    } else {
        Write-WmrWarn "Original nicht gefunden ($($state.sourceDocx)) — es wird neu angelegt."
    }
    Copy-Item -LiteralPath $rebuilt -Destination $state.sourceDocx -Force
    Write-WmrInfo "Original ueberschrieben: $($state.sourceDocx)"
    return [pscustomobject]@{ Output = $state.sourceDocx; Overwritten = $true }
}
