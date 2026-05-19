<#
  _SelfTest.ps1 — Worker fuer den End-to-End-Selbsttest.
  Definiert nur Funktionen. Setzt _Common.ps1/_Import.ps1/_Export.ps1 als geladen voraus.
#>

function New-WmrSampleDocx {
    param([string]$OutPath)
    $word = $null
    try {
        $word = New-WmrWord
        $doc  = $word.Documents.Add()
        try { $doc.BuiltInDocumentProperties.Item('Title').Value = 'Testdokument Roundtrip' } catch {}
        $sel = $word.Selection
        $lorem = (1..6 | ForEach-Object { "Dies ist Absatz $_ mit ausreichend Text, damit das Token-Budget je Kapitel ueberschritten wird und der rekursive Split greift. " * 4 }) -join "`n"

        function Add-Para($s, $t) { $s.Style = 'Standard'; $s.TypeText($t); $s.TypeParagraph() }
        function Add-Head($s, $d, $t, $lvl) {
            try { $s.Style = $d.Styles.Item("Heading $lvl") } catch { $s.Style = $d.Styles.Item("Überschrift $lvl") }
            $s.TypeText($t); $s.TypeParagraph(); $s.Style = 'Standard'
        }

        Add-Head $sel $doc 'Einleitung' 1
        Add-Para $sel "Sehr geehrte/r {{kundenname}}, dieses Dokument dient dem Roundtrip-Test. $lorem"
        $fnRange = $sel.Range; $fnRange.Collapse(0)
        [void]$doc.Footnotes.Add($sel.Range, '', 'Dies ist eine Beispiel-Fussnote.')
        $sel.TypeParagraph()

        Add-Head $sel $doc 'Hauptteil' 1
        Add-Para $sel "Der Hauptteil enthaelt Unterkapitel. $lorem"
        Add-Head $sel $doc 'Unterkapitel A' 2
        Add-Para $sel "Inhalt A. $lorem"
        Add-Head $sel $doc 'Unterkapitel B' 2
        Add-Para $sel "Inhalt B mit Liste:"
        $sel.Range.ListFormat.ApplyBulletDefault()
        $sel.TypeText('Erster Punkt'); $sel.TypeParagraph()
        $sel.TypeText('Zweiter Punkt'); $sel.TypeParagraph()
        $sel.Range.ListFormat.RemoveNumbers()
        Add-Para $sel $lorem

        Add-Head $sel $doc 'Anhang' 1
        Add-Para $sel 'Tabelle:'
        $tblRange = $sel.Range
        $tbl = $doc.Tables.Add($tblRange, 2, 2)
        $tbl.Cell(1,1).Range.Text = 'Spalte 1'
        $tbl.Cell(1,2).Range.Text = 'Spalte 2'
        $tbl.Cell(2,1).Range.Text = 'Wert A'
        $tbl.Cell(2,2).Range.Text = 'Wert B'
        $sel.EndKey(6) | Out-Null  # wdStory
        $sel.TypeParagraph()
        Add-Para $sel "Abschliessender Text. $lorem"

        # Gebundenes Content Control (Platzhalter-Rueckgewinnungstest)
        try {
            $ns = $script:WmrXmlNamespace; $pfx = $script:WmrXmlPrefix
            $xml = "<$pfx`:fields xmlns:$pfx=`"$ns`"><$pfx`:kundenname>Mustermann GmbH</$pfx`:kundenname></$pfx`:fields>"
            $part = $doc.CustomXMLParts.Add($xml)
            $sel.EndKey(6) | Out-Null
            $sel.TypeParagraph()
            $sel.TypeText('Kunde: ')
            $cc = $doc.ContentControls.Add($script:WdCC.Text, $sel.Range)
            $cc.Tag = 'kundenname'; $cc.Title = 'Kundenname'
            $cc.XMLMapping.SetMapping("/$pfx`:fields[1]/$pfx`:kundenname[1]", "xmlns:$pfx='$ns'", $part)
        } catch { Write-WmrWarn "Konnte gebundenes Control nicht erzeugen: $($_.Exception.Message)" }

        if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }
        $doc.SaveAs([ref]$OutPath, [ref]$script:WdFormatDocx)
        $doc.Close($script:WdDoNotSaveChanges)
    } finally {
        Stop-WmrWord -Word $word
    }
}

function Get-WmrPlainText {
    param([string]$Pandoc, [string]$Docx)
    $tmp = [System.IO.Path]::GetTempFileName() + '.txt'
    Invoke-WmrPandoc -PandocPath $Pandoc -Arguments @($Docx, '-f', 'docx', '-t', 'plain', '--wrap=none', '-o', $tmp)
    $t = Get-WmrText -Path $tmp
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    return (($t -replace '\s+', ' ').Trim())
}

function Invoke-WmrSelfTest {
    param([string]$Path, [string]$WorkDir, [int]$TargetTokens = 1200, [int]$MaxSplitLevel = 4)

    if (-not $WorkDir) { $WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) ("wmrtest-" + (Get-Date).ToString('yyyyMMddHHmmss')) }
    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
    $pandoc = Assert-WmrPandoc

    if (-not $Path) {
        $Path = Join-Path $WorkDir 'Testdokument.docx'
        Write-WmrStep "Erzeuge Beispiel-DOCX: $Path"
        New-WmrSampleDocx -OutPath $Path
    }
    $srcCopy = Join-Path $WorkDir ([System.IO.Path]::GetFileName($Path))
    if ((Resolve-Path -LiteralPath $Path).Path -ne (Join-Path $WorkDir ([System.IO.Path]::GetFileName($Path)))) {
        Copy-Item -LiteralPath $Path -Destination $srcCopy -Force
    }

    Write-WmrStep '=== IMPORT (Word -> Markdown) ==='
    $imp = Invoke-WmrWordToMarkdown -Path $srcCopy -OutDir $WorkDir -TargetTokens $TargetTokens -MaxTokens ([int]($TargetTokens*1.34)) -MaxSplitLevel $MaxSplitLevel
    $docRoot = $imp.DocRoot

    Write-Host ''
    Write-Host '--- Token-Statistik je Datei ---'
    $mdFiles = Get-ChildItem -LiteralPath $docRoot -Recurse -Filter '*.md' -File
    foreach ($f in $mdFiles) {
        $body = Remove-WmrFrontmatter -Content (Get-WmrText -Path $f.FullName)
        $tok  = Get-WmrTokenEstimate -Text $body
        $flag = if ($tok -gt [int]($TargetTokens*1.34)) { '  <-- > MaxTokens!' } else { '' }
        Write-Host ("  {0,6} tok  {1}{2}" -f $tok, $f.FullName.Substring($docRoot.Length+1), $flag)
    }

    Write-WmrStep '=== RUECKWEG (Markdown -> Word, -NoOverwrite) ==='
    $exp = Invoke-WmrMarkdownToWord -DocDir $docRoot -NoOverwrite

    Write-WmrStep '=== VERGLEICH (Pandoc-Plaintext) ==='
    $a = Get-WmrPlainText -Pandoc $pandoc -Docx $srcCopy
    $b = Get-WmrPlainText -Pandoc $pandoc -Docx $exp.Output
    # Platzhalter-Tokens vs. injizierte Werte angleichen, damit der Vergleich
    # die (korrekte) Feld-Aufloesung nicht als Abweichung wertet.
    $stateFile = Join-Path $docRoot '.roundtrip\state.json'
    $stFields = @()
    if (Test-Path -LiteralPath $stateFile) { $stFields = @((Read-WmrState -Path $stateFile).fields) }
    foreach ($fld in $stFields) {
        $a = $a -replace [regex]::Escape("{{$($fld.id)}}"), ''
        $v = ''
        if ($fld.PSObject.Properties.Name -contains 'value') { $v = "$($fld.value)" }
        if ($v) { $b = $b.Replace($v, '') }
    }
    $wordsA = @($a -split '\s+' | Where-Object { $_ })
    $wordsB = @($b -split '\s+' | Where-Object { $_ })
    # Reihenfolge-unabhaengiger Multiset-Vergleich (Dice-Koeffizient)
    $freqA = @{}; foreach ($w in $wordsA) { $freqA[$w] = 1 + ($(if ($freqA.ContainsKey($w)) { $freqA[$w] } else { 0 })) }
    $freqB = @{}; foreach ($w in $wordsB) { $freqB[$w] = 1 + ($(if ($freqB.ContainsKey($w)) { $freqB[$w] } else { 0 })) }
    $inter = 0
    foreach ($k in $freqA.Keys) { if ($freqB.ContainsKey($k)) { $inter += [Math]::Min($freqA[$k], $freqB[$k]) } }
    $tot = $wordsA.Count + $wordsB.Count
    $pct = if ($tot -gt 0) { [Math]::Round(200.0 * $inter / $tot, 1) } else { 100 }

    Write-Host ''
    Write-Host ("Original-Woerter : {0}" -f $wordsA.Count)
    Write-Host ("Rebuilt-Woerter  : {0}" -f $wordsB.Count)
    Write-Host ("Inhaltsgleichheit: {0}% (Multiset, Platzhalter-normalisiert)" -f $pct)
    $missing = @()
    foreach ($k in $freqA.Keys) { if (-not $freqB.ContainsKey($k)) { $missing += $k } }
    if ($missing.Count -gt 0) {
        Write-WmrWarn ("Nur im Original (max. 15): " + (($missing | Select-Object -First 15) -join ' '))
        Write-Host '(Kommentare/Felder/Kopf-Fusszeilen sind erwartungsgemaess nicht Teil des Roundtrips.)'
    } else {
        Write-WmrInfo 'Alle Original-Woerter im Ergebnis vorhanden.'
    }
    if ($pct -ge 97) { Write-WmrInfo 'Roundtrip-Treue OK.' }
    else { Write-WmrWarn 'Inhaltsgleichheit unter 97% — bitte Artefakte pruefen.' }

    Write-WmrInfo "Artefakte unter: $WorkDir"
    return [pscustomobject]@{ WorkDir = $WorkDir; DocRoot = $docRoot; Rebuilt = $exp.Output; MatchPercent = $pct }
}
