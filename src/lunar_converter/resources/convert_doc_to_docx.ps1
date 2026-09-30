<#
.SYNOPSIS
    Wandelt .doc-Dateien mit Microsoft Word in .docx um.

.DESCRIPTION
    Wird vom LUNAR-Testfallkonverter automatisch je .doc-Datei aufgerufen, kann aber
    auch eigenstaendig genutzt werden.

    Verarbeitet alle .doc-Dateien direkt im Eingabeordner (alphabetisch, ohne Unterordner,
    ohne temporaere Word-Dateien "~$..."). Jede Datei wird schreibgeschuetzt geoeffnet und als
    .docx im Ausgabeordner gespeichert. Vorhandene .docx-Dateien des Eingabeordners werden
    unveraendert mitkopiert, sodass der Ausgabeordner direkt als --input-dir des Konverters
    dienen kann.

    Sicherheit und Robustheit:
    - Es wird nichts ueberschrieben; existiert die Zieldatei, schlaegt nur diese Datei fehl.
    - Makros werden beim Oeffnen deaktiviert.
    - Kennwortgeschuetzte Dokumente schlagen fehl, statt einen Dialog zu oeffnen.
    - Die Quelldateien werden nicht veraendert.

    Exit-Codes: 0 = alle Dateien erfolgreich, 1 = mindestens eine Datei fehlgeschlagen,
    2 = ungueltige Parameter oder Word nicht verfuegbar.

.PARAMETER InputDir
    Ordner mit .doc- (und optional .docx-) Dateien.

.PARAMETER OutputDir
    Zielordner fuer die .docx-Dateien. Wird angelegt, falls er fehlt. Muss sich vom
    Eingabeordner unterscheiden.

.EXAMPLE
    .\src\lunar_converter\resources\convert_doc_to_docx.ps1 -InputDir C:\temp\input -OutputDir C:\temp\input_docx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Nur beim Aufruf aus Python gesetzt; im interaktiven Einsatz bleibt die Konsolenkodierung unveraendert.
if ($env:LUNAR_UTF8_OUTPUT -eq '1') {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$wdFormatXMLDocument = 12
$wdDoNotSaveChanges = 0
$wdAlertsNone = 0
$msoAutomationSecurityForceDisable = 3
# Absichtlich falsches Kennwort: geschuetzte Dokumente scheitern sofort statt einen Dialog zu oeffnen.
$dummyPassword = 'lunar-kein-kennwort-' + [guid]::NewGuid().ToString('N')

function Write-Info([string]$Message) { Write-Host ("{0} INFO    {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) }
function Write-Fehler([string]$Message) { [Console]::Error.WriteLine(("{0} ERROR   {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)) }

if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) {
    Write-Fehler "Eingabeordner '$InputDir' existiert nicht."
    exit 2
}
$inputPath = (Resolve-Path -LiteralPath $InputDir).ProviderPath
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
elseif (-not (Test-Path -LiteralPath $OutputDir -PathType Container)) {
    Write-Fehler "Ausgabepfad '$OutputDir' existiert, ist aber kein Ordner."
    exit 2
}
$outputPath = (Resolve-Path -LiteralPath $OutputDir).ProviderPath
if ($inputPath.TrimEnd('\') -ieq $outputPath.TrimEnd('\')) {
    Write-Fehler 'Ausgabeordner muss sich vom Eingabeordner unterscheiden.'
    exit 2
}

$files = @(
    Get-ChildItem -LiteralPath $inputPath -File |
        Where-Object { ($_.Extension -ieq '.doc' -or $_.Extension -ieq '.docx') -and -not $_.Name.StartsWith('~$') } |
        Sort-Object -Property @{ Expression = { $_.Name.ToLowerInvariant() } }, Name
)
$docCount = @($files | Where-Object { $_.Extension -ieq '.doc' }).Count
Write-Info ("Start: {0} Datei(en) in '{1}', davon {2} .doc." -f $files.Count, $inputPath, $docCount)

$converted = 0
$copied = 0
$failed = 0
$word = $null
try {
    if ($docCount -gt 0) {
        try {
            $word = New-Object -ComObject Word.Application
        }
        catch {
            Write-Fehler "Microsoft Word konnte nicht gestartet werden: $($_.Exception.Message)"
            exit 2
        }
        $word.Visible = $false
        $word.DisplayAlerts = $wdAlertsNone
        $word.AutomationSecurity = $msoAutomationSecurityForceDisable
    }

    foreach ($file in $files) {
        $target = Join-Path $outputPath ($file.BaseName + '.docx')
        if (Test-Path -LiteralPath $target) {
            Write-Fehler "$($file.Name): Zieldatei '$target' existiert bereits; es wird nichts ueberschrieben."
            $failed++
            continue
        }
        if ($file.Extension -ieq '.docx') {
            Copy-Item -LiteralPath $file.FullName -Destination $target
            Write-Info "$($file.Name): .docx unveraendert kopiert."
            $copied++
            continue
        }

        $doc = $null
        try {
            # Parameter: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument
            $doc = $word.Documents.Open($file.FullName, $false, $true, $false, $dummyPassword)
            $doc.SaveAs2($target, $wdFormatXMLDocument)
            Write-Info "$($file.Name): nach '$target' umgewandelt."
            $converted++
        }
        catch {
            $failed++
            Write-Fehler "$($file.Name): Umwandlung fehlgeschlagen: $($_.Exception.Message)"
            if (Test-Path -LiteralPath $target) {
                Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
            }
        }
        finally {
            if ($null -ne $doc) {
                try { $doc.Close($wdDoNotSaveChanges) } catch { Write-Fehler "$($file.Name): Dokument konnte nicht geschlossen werden." }
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }
}
finally {
    if ($null -ne $word) {
        # Word nur beenden, wenn keine fremden Dokumente offen sind (Instanz koennte wiederverwendet sein).
        try {
            if ($word.Documents.Count -eq 0) { $word.Quit() }
        }
        catch {
            Write-Fehler "Word konnte nicht sauber beendet werden: $($_.Exception.Message)"
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Write-Info ("Ende: {0} umgewandelt, {1} kopiert, {2} fehlgeschlagen. Ausgabe: '{3}'" -f $converted, $copied, $failed, $outputPath)
if ($failed -gt 0) { exit 1 }
exit 0
