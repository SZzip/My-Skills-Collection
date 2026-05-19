<#
  _Import.ps1 — Word -> Markdown (Zerlegung).
  Definiert nur Funktionen. Wird von Convert-WordToMarkdown.ps1 (und im
  STA-Runspace) dot-gesourct. Setzt _Common.ps1 als bereits geladen voraus.
#>

# --- Markdown-Heading-Baum --------------------------------------------------

function ConvertTo-WmrHeadingTree {
    <# Parst Markdown in einen Baum. Rueckgabe: PSCustomObject @{ Preamble; Children; Footnotes }
       Knoten: @{ Level; Title; HeadingLine; Body; Children } #>
    param([Parameter(Mandatory)][string]$Markdown)

    $text = $Markdown -replace "`r`n", "`n"

    # 1. Fussnoten-Definitionen herausloesen (Block = [^x]: Zeile + eingerueckte Folgezeilen)
    $allLines = $text.Split("`n")
    $footnotes = New-Object System.Collections.ArrayList
    $kept = New-Object System.Collections.ArrayList
    $i = 0
    while ($i -lt $allLines.Count) {
        $l = $allLines[$i]
        if ($l -match '^\[\^([^\]]+)\]:\s?(.*)$') {
            $label = $Matches[1]
            $block = New-Object System.Collections.ArrayList
            [void]$block.Add($l)
            $i++
            while ($i -lt $allLines.Count) {
                $n = $allLines[$i]
                if ($n -match '^(\s{4,}|\t)' -or $n.Trim() -eq '') {
                    # blank-Zeile nur uebernehmen, wenn danach noch Einrueckung folgt
                    if ($n.Trim() -eq '') {
                        if ($i + 1 -lt $allLines.Count -and $allLines[$i+1] -match '^(\s{4,}|\t)') {
                            [void]$block.Add($n); $i++; continue
                        } else { break }
                    }
                    [void]$block.Add($n); $i++
                } else { break }
            }
            [void]$footnotes.Add([pscustomobject]@{ Label = $label; Text = ($block -join "`n") })
            continue
        }
        [void]$kept.Add($l); $i++
    }

    # 2. Heading-Baum aus den verbliebenen Zeilen
    $root = [pscustomobject]@{ Level = 0; Title = ''; HeadingLine = ''; Body = ''; Children = (New-Object System.Collections.ArrayList) }
    $preamble = New-Object System.Collections.ArrayList
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    $current = $null
    $inFence = $false
    $fenceTok = ''
    $bodyBuf = New-Object System.Collections.ArrayList

    function Flush-Body {
        param($node, $buf)
        if ($null -ne $node) { $node.Body = ($buf -join "`n").Trim("`n") }
    }

    foreach ($line in $kept) {
        $trim = $line.TrimStart()
        if (-not $inFence -and ($trim -match '^(```+|~~~+)')) {
            $inFence = $true; $fenceTok = $Matches[1].Substring(0,3)
        }
        elseif ($inFence -and ($trim -match '^(```+|~~~+)\s*$')) {
            $inFence = $false
        }

        if (-not $inFence -and $line -match '^(#{1,6})\s+(.+?)\s*#*\s*$') {
            $level = $Matches[1].Length
            $title = $Matches[2].Trim()
            if ($null -ne $current) { Flush-Body $current $bodyBuf } else { $preamble.AddRange(@($bodyBuf)) | Out-Null }
            $bodyBuf = New-Object System.Collections.ArrayList

            while ($stack.Count -gt 1 -and $stack.Peek().Level -ge $level) { [void]$stack.Pop() }
            $parent = $stack.Peek()
            $node = [pscustomobject]@{
                Level = $level; Title = $title; HeadingLine = $line
                Body = ''; Children = (New-Object System.Collections.ArrayList)
            }
            [void]$parent.Children.Add($node)
            $stack.Push($node)
            $current = $node
        }
        else {
            [void]$bodyBuf.Add($line)
        }
    }
    if ($null -ne $current) { Flush-Body $current $bodyBuf } else { $preamble.AddRange(@($bodyBuf)) | Out-Null }

    Remove-WmrEmptyNodes -Children $root.Children

    return [pscustomobject]@{
        Preamble  = ($preamble -join "`n").Trim("`n")
        Children  = $root.Children
        Footnotes = $footnotes
    }
}

function Remove-WmrEmptyNodes {
    <# Entfernt rekursiv (bottom-up) Knoten ohne Titeltext UND ohne Inhalt
       UND ohne Kinder (z. B. leere Word-Heading-1-Absaetze). #>
    param($Children)
    $keep = New-Object System.Collections.ArrayList
    foreach ($n in $Children) {
        Remove-WmrEmptyNodes -Children $n.Children
        $titleClean = ($n.Title -replace '\{[^}]*\}', '' -replace '[*_~`]', '').Trim()
        $titleEmpty = [string]::IsNullOrWhiteSpace($titleClean)
        $bodyEmpty  = [string]::IsNullOrWhiteSpace($n.Body)
        if ($titleEmpty -and $bodyEmpty -and $n.Children.Count -eq 0) { continue }
        [void]$keep.Add($n)
    }
    $Children.Clear()
    foreach ($k in $keep) { [void]$Children.Add($k) }
}

function Join-WmrRel {
    <# Verkettet relative Pfadsegmente mit '/'. Leere Basis -> nur Child
       (Join-Path verbietet leeres -Path). #>
    param([AllowEmptyString()][string]$Base, [Parameter(Mandatory)][string]$Child)
    if ([string]::IsNullOrEmpty($Base)) { return $Child }
    return ($Base.TrimEnd('/', '\') + '/' + $Child)
}

function Get-WmrSubtreeMarkdown {
    param($Node)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine($Node.HeadingLine)
    if ($Node.Body -and $Node.Body.Trim()) { [void]$sb.AppendLine(''); [void]$sb.AppendLine($Node.Body) }
    foreach ($c in $Node.Children) {
        [void]$sb.AppendLine('')
        [void]$sb.Append((Get-WmrSubtreeMarkdown -Node $c))
    }
    return $sb.ToString()
}

# --- Fussnoten / Bilder -----------------------------------------------------

function Add-WmrUsedFootnotes {
    <# Haengt die im Text referenzierten Fussnoten-Definitionen zur Lesbarkeit an. #>
    param([string]$Body, $Footnotes)
    if (-not $Footnotes -or $Footnotes.Count -eq 0) { return $Body }
    $used = [ordered]@{}
    foreach ($m in [regex]::Matches($Body, '\[\^([^\]]+)\]')) {
        $lbl = $m.Groups[1].Value
        if (-not $used.Contains($lbl)) { $used[$lbl] = $true }
    }
    if ($used.Count -eq 0) { return $Body }
    $defs = New-Object System.Collections.ArrayList
    foreach ($lbl in $used.Keys) {
        $fn = $Footnotes | Where-Object { $_.Label -eq $lbl } | Select-Object -First 1
        if ($fn) { [void]$defs.Add($fn.Text) }
    }
    if ($defs.Count -eq 0) { return $Body }
    return ($Body.TrimEnd("`n") + "`n`n<!-- wmr:footnotes -->`n`n" + ($defs -join "`n") + "`n`n<!-- /wmr:footnotes -->")
}

function Set-WmrAssetDepth {
    <# Passt assets/-Bildpfade an die Ordnertiefe der Zieldatei an. #>
    param([string]$Body, [int]$Depth)
    $prefix = ('../' * $Depth) + 'assets/'
    # Bildlinks und HTML-img normalisieren -> <prefix>assets/<rest>
    $b = [regex]::Replace($Body, '(\!\[[^\]]*\]\()\s*[^)]*?assets[\\/]+', { param($m) $m.Groups[1].Value + $prefix })
    $b = [regex]::Replace($b, '(<img[^>]*\ssrc=")[^"]*?assets[\\/]+', { param($m) $m.Groups[1].Value + $prefix })
    $b = $b -replace ([regex]::Escape($prefix) + '([^\s\)"]*)'), ($prefix + '$1')
    return ($b -replace '(assets/[^\s\)"]*)\\', '$1/')
}

# --- Datei-/Ordner-Emission -------------------------------------------------

function Get-WmrFrontmatterData {
    param($Title, $Doc, $Source, $Type, $Level, $Order, $Body, $StyleMap)
    $tokens = Get-WmrTokenEstimate -Text $Body
    $words  = Get-WmrWordCount -Text $Body
    $hs = ''
    if ($Level -ge 1 -and $StyleMap -and $StyleMap.Contains("$Level")) { $hs = $StyleMap["$Level"] }
    return [ordered]@{
        title         = $Title
        doc           = $Doc
        source        = $Source
        type          = $Type
        level         = $Level
        order         = $Order
        slug          = (ConvertTo-WmrSlug $Title)
        heading_style = $hs
        tokens_est    = $tokens
        words         = $words
        generated     = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        schema        = $script:WmrSchemaVersion
    }
}

function Write-WmrNode {
    <# Schreibt einen Knoten rekursiv. Rueckgabe: relativer Pfad (ohne .md)
       zur erzeugten Einstiegsdatei (fuer den Eltern-Embed-Link). #>
    param(
        $Node, [int]$Number, [int]$Depth,
        [string]$DirAbs, [string]$RelDir,
        [string]$DocTitle, [string]$Source, $StyleMap, $Footnotes,
        [int]$TargetTokens, [int]$MaxTokens, [int]$MaxSplitLevel
    )
    $nn = '{0:00}' -f $Number
    $plain = ConvertTo-WmrPlainTitle $Node.Title   # ohne Markdown/Newlines
    $safe = ConvertTo-WmrSafeName $plain
    $subtree = Get-WmrSubtreeMarkdown -Node $Node
    $subtreeTokens = Get-WmrTokenEstimate -Text $subtree
    $canDescend = ($Node.Children.Count -gt 0) -and ($Depth -lt $MaxSplitLevel)
    $order = if ($RelDir) { ($RelDir -replace '[^0-9]', '') + '.' + $nn } else { $nn }

    if (($subtreeTokens -le $TargetTokens) -or (-not $canDescend)) {
        # Ggf. zu grosser Blattknoten -> Absatz-Split in Teil-Dateien
        if (($subtreeTokens -gt $MaxTokens) -and (-not $canDescend)) {
            $folderRel = Join-WmrRel $RelDir "$nn - $safe"
            $folderAbs = Join-Path $DirAbs "$nn - $safe"
            New-Item -ItemType Directory -Path $folderAbs -Force | Out-Null
            $parts = Split-WmrByParagraph -Text $subtree -MaxTokens $TargetTokens
            $depth1 = ($folderRel -split '/').Count
            $embeds = New-Object System.Collections.ArrayList
            for ($p = 0; $p -lt $parts.Count; $p++) {
                $pn = '{0:00}' -f ($p + 1)
                $ptitle = "$plain (Teil $($p+1))"
                $pbody = Set-WmrAssetDepth -Body $parts[$p] -Depth $depth1
                $pbody = Add-WmrUsedFootnotes -Body $pbody -Footnotes $Footnotes
                $fm = Get-WmrFrontmatterData -Title $ptitle -Doc $DocTitle -Source $Source -Type 'chapter' -Level $Node.Level -Order "$order.$pn" -Body $pbody -StyleMap $StyleMap
                Set-WmrText -Path (Join-Path $folderAbs "$pn - $(ConvertTo-WmrSafeName $ptitle).md") -Content ((ConvertTo-WmrFrontmatter $fm) + "`n" + $pbody + "`n")
                [void]$embeds.Add("![[${folderRel}/$pn - $(ConvertTo-WmrSafeName $ptitle)]]")
            }
            $introBody = ($embeds -join "`n")
            $fm0 = Get-WmrFrontmatterData -Title $plain -Doc $DocTitle -Source $Source -Type 'chapter' -Level $Node.Level -Order $order -Body '' -StyleMap $StyleMap
            Set-WmrText -Path (Join-Path $folderAbs "00 - $safe.md") -Content ((ConvertTo-WmrFrontmatter $fm0) + "`n" + $Node.HeadingLine + "`n`n" + $introBody + "`n")
            return "$folderRel/00 - $safe"
        }

        $depthF = if ($RelDir) { ($RelDir -split '/').Count } else { 0 }
        $body = Set-WmrAssetDepth -Body $subtree -Depth $depthF
        $body = Add-WmrUsedFootnotes -Body $body -Footnotes $Footnotes
        $fm = Get-WmrFrontmatterData -Title $plain -Doc $DocTitle -Source $Source -Type 'chapter' -Level $Node.Level -Order $order -Body $body -StyleMap $StyleMap
        $rel = Join-WmrRel $RelDir "$nn - $safe"
        Set-WmrText -Path (Join-Path $DirAbs "$nn - $safe.md") -Content ((ConvertTo-WmrFrontmatter $fm) + "`n" + $body + "`n")
        return $rel
    }

    # Ordnerknoten: 00 - Title.md + Kinder
    $folderRel = Join-WmrRel $RelDir "$nn - $safe"
    $folderAbs = Join-Path $DirAbs "$nn - $safe"
    New-Item -ItemType Directory -Path $folderAbs -Force | Out-Null
    $depthC = ($folderRel -split '/').Count
    $childEmbeds = New-Object System.Collections.ArrayList
    $cn = 1
    foreach ($child in $Node.Children) {
        $childRel = Write-WmrNode -Node $child -Number $cn -Depth ($Depth + 1) `
            -DirAbs $folderAbs -RelDir $folderRel -DocTitle $DocTitle -Source $Source `
            -StyleMap $StyleMap -Footnotes $Footnotes -TargetTokens $TargetTokens `
            -MaxTokens $MaxTokens -MaxSplitLevel $MaxSplitLevel
        [void]$childEmbeds.Add("![[$childRel]]")
        $cn++
    }
    $ownBody = Set-WmrAssetDepth -Body $Node.Body -Depth $depthC
    $ownBody = Add-WmrUsedFootnotes -Body $ownBody -Footnotes $Footnotes
    $introParts = @($Node.HeadingLine)
    if ($ownBody -and $ownBody.Trim()) { $introParts += ''; $introParts += $ownBody }
    $introParts += ''; $introParts += '## Unterkapitel'; $introParts += ''
    $introParts += ($childEmbeds -join "`n")
    $intro = ($introParts -join "`n")
    $fm0 = Get-WmrFrontmatterData -Title $plain -Doc $DocTitle -Source $Source -Type 'chapter' -Level $Node.Level -Order $order -Body $Node.Body -StyleMap $StyleMap
    Set-WmrText -Path (Join-Path $folderAbs "00 - $safe.md") -Content ((ConvertTo-WmrFrontmatter $fm0) + "`n" + $intro + "`n")
    return "$folderRel/00 - $safe"
}

function Split-WmrByParagraph {
    <# Teilt Text an Absatzgrenzen (Leerzeilen) in Bloecke <= MaxTokens.
       Kein Block wird mitten im Absatz zerschnitten. #>
    param([string]$Text, [int]$MaxTokens)
    $paras = [regex]::Split($Text, "`n[ \t]*`n")
    $parts = New-Object System.Collections.ArrayList
    $cur = New-Object System.Collections.ArrayList
    $curTok = 0
    foreach ($p in $paras) {
        $pt = Get-WmrTokenEstimate -Text $p
        if ($cur.Count -gt 0 -and ($curTok + $pt) -gt $MaxTokens) {
            [void]$parts.Add(($cur -join "`n`n")); $cur = New-Object System.Collections.ArrayList; $curTok = 0
        }
        [void]$cur.Add($p); $curTok += $pt
    }
    if ($cur.Count -gt 0) { [void]$parts.Add(($cur -join "`n`n")) }
    if ($parts.Count -eq 0) { [void]$parts.Add($Text) }
    return $parts
}

# --- Platzhalter-Rueckgewinnung (Content Controls) --------------------------

function Get-WmrBoundFields {
    <# Erfasst gebundene Content Controls und ersetzt sie auf der Arbeitskopie
       durch literales {{id}}. Rueckgabe: Array von Feld-Hashtables. #>
    param($Doc)
    $fields = New-Object System.Collections.ArrayList
    $ccs = @()
    try { $ccs = @($Doc.ContentControls) } catch { return @() }
    foreach ($cc in $ccs) {
        $mapped = $false
        try { $mapped = [bool]$cc.XMLMapping.IsMapped } catch { $mapped = $false }
        $id = ''
        try { if ($cc.Tag)   { $id = "$($cc.Tag)" } } catch {}
        if (-not $id) { try { if ($cc.Title) { $id = "$($cc.Title)" } } catch {} }
        if (-not $mapped -and -not $id) { continue }
        if (-not $id) { $id = 'feld_' + ([guid]::NewGuid().ToString('N').Substring(0,6)) }

        $typeNum = 1; try { $typeNum = [int]$cc.Type } catch {}
        $type = switch ($typeNum) {
            $script:WdCC.RichText     { 'rich' }
            $script:WdCC.Text         { 'text' }
            $script:WdCC.ComboBox     { 'choice' }
            $script:WdCC.DropdownList { 'choice' }
            $script:WdCC.Date         { 'date' }
            $script:WdCC.CheckBox     { 'checkbox' }
            default                   { 'text' }
        }
        $multi = $false; try { $multi = [bool]$cc.MultiLine } catch {}
        if ($type -eq 'text' -and $multi) { $type = 'multiline' }

        $value = ''
        if ($type -eq 'checkbox') {
            try { $value = [bool]$cc.Checked } catch { $value = $false }
        } else {
            try { $value = ($cc.Range.Text -replace '[\r\n\a\x07]', ' ').Trim() } catch {}
            try { if ($cc.ShowingPlaceholderText) { $value = '' } } catch {}
        }
        $choices = @()
        if ($type -eq 'choice') {
            try { foreach ($e in $cc.DropdownListEntries) { if ($e.Value) { $choices += "$($e.Value)" } } } catch {}
        }
        $fmt = ''
        if ($type -eq 'date') { try { $fmt = "$($cc.DateDisplayFormat)" } catch {} }

        $field = [ordered]@{ id = $id; label = ($(try { "$($cc.Title)" } catch { $id })); type = $type }
        if ($type -eq 'date' -and $fmt) { $field['format'] = $fmt }
        if ($choices.Count -gt 0) { $field['choices'] = $choices }
        $field['default'] = ''
        $field['value']   = $value
        if (-not ($fields | Where-Object { $_.id -eq $id })) { [void]$fields.Add($field) }

        # Control auf der Arbeitskopie durch Token ersetzen
        try {
            $rng = $cc.Range
            $rng.Text = "{{$id}}"
            $cc.Delete($false)
        } catch {
            try { $cc.Delete($true) } catch {}
        }
    }
    return $fields.ToArray()
}

# --- Hub --------------------------------------------------------------------

function Write-WmrHub {
    param([string]$DocRootAbs, [string]$DocTitle, [string]$Source, $TopEntries, $Fields)
    $fm = [ordered]@{
        title     = "$DocTitle – Hub"
        doc       = $DocTitle
        source    = $Source
        type      = 'hub'
        generated = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        schema    = $script:WmrSchemaVersion
    }
    if ($Fields -and @($Fields).Count -gt 0) { $fm['fields'] = @($Fields) }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append((ConvertTo-WmrFrontmatter $fm))
    [void]$sb.AppendLine("`n# $DocTitle — Übersicht")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('> Navigations-Hub. Lade gezielt einzelne Kapitel, um den KI-Kontext zu schonen.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Kapitel')
    [void]$sb.AppendLine('')
    foreach ($e in $TopEntries) {
        [void]$sb.AppendLine("- [[$($e.Target)|$($e.Title)]] — <!-- summary:$($e.Key) -->")
    }
    if ($Fields -and @($Fields).Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('## Platzhalter / Felder')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Die Feldwerte werden zentral in der Frontmatter dieser Datei (`fields:`) gepflegt und beim Rueckweg als XML-gebundene Eingabefelder in Word eingesetzt.')
        [void]$sb.AppendLine('')
        foreach ($f in @($Fields)) { [void]$sb.AppendLine("- ``{{$($f.id)}}`` — $($f.label) ($($f.type))") }
    }
    Set-WmrText -Path (Join-Path $DocRootAbs "00 - $(ConvertTo-WmrSafeName $DocTitle) Hub.md") -Content $sb.ToString()
}

# --- Haupt-Entry ------------------------------------------------------------

function Invoke-WmrWordToMarkdown {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$OutDir,
        [int]$MaxSplitLevel = 3,
        [int]$TargetTokens  = 6000,
        [int]$MaxTokens     = 8000
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "Word-Datei nicht gefunden: $Path" }
    $srcFull = (Resolve-Path -LiteralPath $Path).Path
    if (-not $OutDir) { $OutDir = Split-Path -Parent $srcFull }
    if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

    $pandoc = Assert-WmrPandoc

    Write-WmrStep "Oeffne Word-Dokument: $srcFull"
    $word = $null; $doc = $null; $workCopy = $null
    try {
        $word = New-WmrWord
        $doc  = Open-WmrDoc -Word $word -Path $srcFull -ReadOnly $true
        $docTitle = Get-WmrDocTitle -Doc $doc -FallbackFileName $srcFull
        $styleMap = Get-WmrHeadingStyleMap -Doc $doc

        $docRootName = "01 - $(ConvertTo-WmrSafeName $docTitle)"
        $docRootAbs  = Join-Path $OutDir $docRootName
        $rtDir       = Join-Path $docRootAbs '.roundtrip'
        $assetsAbs   = Join-Path $docRootAbs 'assets'
        New-Item -ItemType Directory -Path $rtDir -Force | Out-Null

        # reference.docx = Kopie des Originals (fuer Pandoc-Stilvorlage beim Rueckweg)
        $refDocx = Join-Path $rtDir 'reference.docx'
        Copy-Item -LiteralPath $srcFull -Destination $refDocx -Force

        # Platzhalter-Rueckgewinnung auf Arbeitskopie
        Write-WmrStep 'Pruefe gebundene Inhaltssteuerelemente (Platzhalter).'
        $workCopy = Join-Path $rtDir 'work.docx'
        Copy-Item -LiteralPath $srcFull -Destination $workCopy -Force
        $doc.Close($script:WdDoNotSaveChanges); $doc = $null
        $wdoc = Open-WmrDoc -Word $word -Path $workCopy -ReadOnly $false
        $fields = Get-WmrBoundFields -Doc $wdoc
        if (@($fields).Count -gt 0) {
            Write-WmrInfo "$(@($fields).Count) gebundene Felder erkannt: $((@($fields) | ForEach-Object { $_.id }) -join ', ')"
            $wdoc.Save()
        }
        $wdoc.Close($script:WdDoNotSaveChanges); $wdoc = $null
        $pandocInput = if (@($fields).Count -gt 0) { $workCopy } else { $srcFull }
    }
    finally {
        if ($doc)  { try { $doc.Close($script:WdDoNotSaveChanges) } catch {} }
        Stop-WmrWord -Word $word
    }

    # DOCX -> Markdown
    Write-WmrStep 'Konvertiere DOCX -> Markdown (Pandoc).'
    $fullMd = Join-Path $rtDir 'full.md'
    Invoke-WmrPandoc -PandocPath $pandoc -Arguments @(
        $pandocInput, '-f', 'docx', '-t', 'gfm+footnotes',
        '--wrap=none', "--extract-media=$assetsAbs", '-o', $fullMd
    )
    $md = Get-WmrText -Path $fullMd

    # Bildpfade -> root-relativ "assets/..."
    $md = [regex]::Replace($md, '(\!\[[^\]]*\]\()\s*[^)]*?assets[\\/]+', { param($m) $m.Groups[1].Value + 'assets/' })
    $md = [regex]::Replace($md, '(<img[^>]*\ssrc=")[^"]*?assets[\\/]+', { param($m) $m.Groups[1].Value + 'assets/' })
    $md = $md -replace '(assets/[^\s\)"]*?)\\', '$1/'

    Write-WmrStep 'Parse Heading-Struktur und zerlege nach Token-Budget.'
    $tree = ConvertTo-WmrHeadingTree -Markdown $md

    $topEntries = New-Object System.Collections.ArrayList
    $num = 1

    # Vorspann (Inhalt vor erster Ueberschrift)
    if ($tree.Preamble -and $tree.Preamble.Trim()) {
        $num = 1
        $pre = Set-WmrAssetDepth -Body $tree.Preamble -Depth 0
        $pre = Add-WmrUsedFootnotes -Body $pre -Footnotes $tree.Footnotes
        $fm = Get-WmrFrontmatterData -Title 'Vorspann' -Doc $docTitle -Source $srcFull -Type 'vorspann' -Level 0 -Order '01' -Body $pre -StyleMap $styleMap
        Set-WmrText -Path (Join-Path $docRootAbs '01 - Vorspann.md') -Content ((ConvertTo-WmrFrontmatter $fm) + "`n" + $pre + "`n")
        [void]$topEntries.Add([pscustomobject]@{ Key = '01'; Title = 'Vorspann'; Target = '01 - Vorspann' })
        $num = 2
    } else {
        $num = 2
    }

    foreach ($chapter in $tree.Children) {
        # Auto-Inhaltsverzeichnis (geflachtes Word-Feld) verwerfen: generischer
        # Titel + viele Anker-Links bzw. kein eigener Inhalt.
        if (Test-WmrGenericHeading $chapter.Title) {
            $sub = Get-WmrSubtreeMarkdown -Node $chapter
            $anchorLinks = ([regex]::Matches($sub, '\]\(#')).Count
            if ($anchorLinks -ge 3 -or [string]::IsNullOrWhiteSpace($chapter.Body)) {
                Write-WmrInfo "TOC-Kapitel verworfen: '$($chapter.Title)' ($anchorLinks Anker-Links)"
                continue
            }
        }
        $rel = Write-WmrNode -Node $chapter -Number $num -Depth 1 `
            -DirAbs $docRootAbs -RelDir '' -DocTitle $docTitle -Source $srcFull `
            -StyleMap $styleMap -Footnotes $tree.Footnotes `
            -TargetTokens $TargetTokens -MaxTokens $MaxTokens -MaxSplitLevel $MaxSplitLevel
        [void]$topEntries.Add([pscustomobject]@{ Key = ('{0:00}' -f $num); Title = (ConvertTo-WmrPlainTitle $chapter.Title); Target = $rel })
        $num++
    }

    # Hub
    Write-WmrStep 'Schreibe Hub-Datei.'
    Write-WmrHub -DocRootAbs $docRootAbs -DocTitle $docTitle -Source $srcFull -TopEntries $topEntries -Fields $fields

    # State
    $state = [ordered]@{
        schema        = $script:WmrSchemaVersion
        generated     = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        sourceDocx    = $srcFull
        referenceDocx = '.roundtrip/reference.docx'
        docTitle      = $docTitle
        docRoot       = $docRootName
        assetsDir     = 'assets'
        maxSplitLevel = $MaxSplitLevel
        targetTokens  = $TargetTokens
        maxTokens     = $MaxTokens
        xml           = [ordered]@{ namespace = $script:WmrXmlNamespace; prefix = $script:WmrXmlPrefix }
        styleMap      = $styleMap
        fields        = @($fields)
        footnotes     = @($tree.Footnotes | ForEach-Object { [ordered]@{ label = $_.Label; text = $_.Text } })
        topEntries    = @($topEntries | ForEach-Object { [ordered]@{ key = $_.Key; title = $_.Title; target = $_.Target } })
    }
    Save-WmrState -State $state -Path (Join-Path $rtDir 'state.json')

    Write-WmrInfo "Fertig. Struktur unter: $docRootAbs"
    Write-WmrInfo "Naechster Schritt (KI): Kapitel lesen und im Hub die <!-- summary:NN --> Platzhalter durch 1-2-Satz-Zusammenfassungen ersetzen."
    return [pscustomobject]@{ DocRoot = $docRootAbs; Hub = (Join-Path $docRootAbs "00 - $(ConvertTo-WmrSafeName $docTitle) Hub.md"); Fields = @($fields).Count }
}
