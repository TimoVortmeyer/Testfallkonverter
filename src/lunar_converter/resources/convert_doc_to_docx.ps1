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

    Mit -Recurse wird die Unterordnerstruktur gespiegelt. Im Rekursionsmodus werden
    VBA-Makros aus .doc und makrohaltigen .docx beim Speichern als normales .docx entfernt.
    -ShowProgress zeigt Dateistatus sowie Laufzeit-, Gesamtzeit- und Restzeitschaetzung.

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

.PARAMETER Recurse
    Verarbeitet Unterordner rekursiv und spiegelt deren Struktur im Ausgabeordner.

.PARAMETER ShowProgress
    Zeigt einen Fortschrittsbalken mit Dateistatus und Zeitschaetzung.

.EXAMPLE
    .\src\lunar_converter\resources\convert_doc_to_docx.ps1 -InputDir C:\temp\input -OutputDir C:\temp\input_docx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [switch]$Recurse,
    [switch]$ShowProgress
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
Add-Type -AssemblyName System.IO.Compression.FileSystem
# Absichtlich falsches Kennwort: geschuetzte Dokumente scheitern sofort statt einen Dialog zu oeffnen.
$dummyPassword = 'lunar-kein-kennwort-' + [guid]::NewGuid().ToString('N')

function Write-Info([string]$Message) { Write-Host ("{0} INFO    {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) }
function Write-Fehler([string]$Message) { [Console]::Error.WriteLine(("{0} ERROR   {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)) }

function Test-HasVbaProject([string]$Path) {
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -ieq 'word/vbaProject.bin') { return $true }
        }
        return $false
    }
    finally { $archive.Dispose() }
}

function Format-PreparationDuration([double]$Seconds) {
    $wholeSeconds = [int][Math]::Max(0, [Math]::Round($Seconds))
    $hours = [Math]::Floor($wholeSeconds / 3600)
    $minutes = [Math]::Floor(($wholeSeconds % 3600) / 60)
    $seconds = $wholeSeconds % 60
    if ($hours -gt 0) { return ("{0:00}:{1:00}:{2:00}" -f $hours, $minutes, $seconds) }
    return ("{0:00}:{1:00}" -f $minutes, $seconds)
}

function Ensure-Word {
    if ($null -eq $script:word) {
        $script:word = New-Object -ComObject Word.Application
        $script:word.Visible = $false
        $script:word.DisplayAlerts = $wdAlertsNone
        $script:word.AutomationSecurity = $msoAutomationSecurityForceDisable
    }
    return $script:word
}

if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) {
    Write-Fehler "Eingabeordner '$InputDir' existiert nicht."
    exit 2
}
$inputPath = (Resolve-Path -LiteralPath $InputDir).ProviderPath
$outputCandidate = [System.IO.Path]::GetFullPath($OutputDir)
if ($Recurse) {
    $trimChars = [char[]]@('\', '/')
    $normalizedInput = $inputPath.TrimEnd($trimChars)
    $normalizedOutput = $outputCandidate.TrimEnd($trimChars)
    $inputPrefix = $normalizedInput + [System.IO.Path]::DirectorySeparatorChar
    $outputPrefix = $normalizedOutput + [System.IO.Path]::DirectorySeparatorChar
    if (
        $normalizedInput.Equals($normalizedOutput, [StringComparison]::OrdinalIgnoreCase) -or
        $normalizedOutput.StartsWith($inputPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        $normalizedInput.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase)
    ) {
        Write-Fehler 'Eingabe- und Ausgabeordner duerfen im Rekursionsmodus nicht identisch sein oder ineinander liegen.'
        exit 2
    }
}
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

$foundFiles = if ($Recurse) {
    Get-ChildItem -LiteralPath $inputPath -File -Recurse
} else {
    Get-ChildItem -LiteralPath $inputPath -File
}
$relativeRoot = $inputPath.TrimEnd([char[]]@('\', '/'))
$files = @(
    foreach ($file in $foundFiles | Where-Object {
        ($_.Extension -ieq '.doc' -or $_.Extension -ieq '.docx') -and -not $_.Name.StartsWith('~$')
    }) {
        $relativePath = if ($Recurse) {
            $file.FullName.Substring($relativeRoot.Length).TrimStart([char[]]@('\', '/'))
        } else {
            $file.Name
        }
        [PSCustomObject]@{ Source = $file; RelativePath = $relativePath }
    }
) | Sort-Object -Property @{ Expression = { $_.RelativePath.ToLowerInvariant() } }, RelativePath
$docCount = @($files | Where-Object { $_.Source.Extension -ieq '.doc' }).Count
Write-Info ("Start: {0} Datei(en) in '{1}', davon {2} .doc." -f $files.Count, $inputPath, $docCount)

$converted = 0
$copied = 0
$macrosRemoved = 0
$failed = 0
$script:word = $null
$preparationWatch = [System.Diagnostics.Stopwatch]::StartNew()

function Update-PreparationProgress([int]$Completed, [string]$RelativePath, [string]$Action) {
    if (-not $ShowProgress -or $files.Count -eq 0) { return }
    $elapsed = $preparationWatch.Elapsed.TotalSeconds
    $estimateText = 'Gesamt --:-- | Rest --:--'
    if ($Completed -gt 0) {
        $estimatedTotal = $elapsed / $Completed * $files.Count
        $secondsRemaining = [int][Math]::Max(0, [Math]::Ceiling($estimatedTotal - $elapsed))
        $estimateText = "Gesamt ~$(Format-PreparationDuration $estimatedTotal) | Rest ~$(Format-PreparationDuration $secondsRemaining)"
    }
    $percent = [int][Math]::Floor(100 * $Completed / $files.Count)
    $status = "$Completed/$($files.Count) | Laufzeit $(Format-PreparationDuration $elapsed) | $estimateText"
    $progressParameters = @{
        Activity = 'DOCX-Vorbereitung'
        Status = $status
        CurrentOperation = "$Action - $RelativePath"
        PercentComplete = $percent
    }
    if ($Completed -gt 0) { $progressParameters.SecondsRemaining = $secondsRemaining }
    Write-Progress @progressParameters
}

try {
    $position = 0
    foreach ($entry in $files) {
        $file = $entry.Source
        $targetRelativePath = [System.IO.Path]::ChangeExtension($entry.RelativePath, '.docx')
        $target = Join-Path $outputPath $targetRelativePath
        $targetDirectory = Split-Path -Parent $target
        $action = if ($file.Extension -ieq '.doc') { 'Konvertiere DOC nach DOCX' } else { 'Kopiere und prüfe Makros' }
        Update-PreparationProgress $position $entry.RelativePath $action

        if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
        }
        if (Test-Path -LiteralPath $target) {
            Write-Fehler "$($file.Name): Zieldatei '$target' existiert bereits; es wird nichts ueberschrieben."
            $failed++
            $position++
            Update-PreparationProgress $position $entry.RelativePath 'Fehler'
            continue
        }

        $doc = $null
        try {
            $hasMacros = $false
            if ($file.Extension -ieq '.docx') {
                $hasMacros = Test-HasVbaProject $file.FullName
            }
            if ($file.Extension -ieq '.doc' -or $hasMacros) {
                $word = Ensure-Word
                # Parameter: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument
                $doc = $word.Documents.Open($file.FullName, $false, $true, $false, $dummyPassword)
                $doc.SaveAs2($target, $wdFormatXMLDocument)
                if ($hasMacros) {
                    Write-Info "$($file.Name): als makrofreies DOCX gespeichert; VBA-Makros entfernt."
                    $macrosRemoved++
                } else {
                    Write-Info "$($file.Name): nach '$target' umgewandelt."
                    $converted++
                }
            } else {
                Copy-Item -LiteralPath $file.FullName -Destination $target
                Write-Info "$($file.Name): DOCX unveraendert kopiert."
                $copied++
            }
        }
        catch {
            $failed++
            Write-Fehler "$($file.Name): Vorbereitung fehlgeschlagen: $($_.Exception.Message)"
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
        $position++
        $finishedAction = if ($failed -gt 0 -and -not (Test-Path -LiteralPath $target)) { 'Fehler' } else { 'fertig' }
        Update-PreparationProgress $position $entry.RelativePath $finishedAction
    }
}
finally {
    if ($ShowProgress) { Write-Progress -Activity 'DOCX-Vorbereitung' -Completed }
    if ($null -ne $script:word) {
        # Word nur beenden, wenn keine fremden Dokumente offen sind (Instanz koennte wiederverwendet sein).
        try {
            if ($script:word.Documents.Count -eq 0) { $script:word.Quit() }
        }
        catch {
            Write-Fehler "Word konnte nicht sauber beendet werden: $($_.Exception.Message)"
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:word)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Write-Info ("Ende: {0} umgewandelt, {1} kopiert, {2} Dateien von VBA-Makros bereinigt, {3} fehlgeschlagen. Ausgabe: '{4}'" -f $converted, $copied, $macrosRemoved, $failed, $outputPath)
if ($failed -gt 0) { exit 1 }
exit 0
