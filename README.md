# LUNAR-Testfallkonverter (.docx/.doc → Jira/Xray-Import-JSON)

## 1. Zweck und Abgrenzung

Der Konverter liest LUNAR-Testdokumente im Word-Format, erkennt das passende
Dokumentprofil, extrahiert je Dokument genau einen Testfall und schreibt ein
`testcase.json` samt Screenshots für den **bereits vorhandenen, separaten**
Jira/Xray-Importprozess.

Nicht Bestandteil:

- keine Jira- oder Xray-API-Aufrufe, keine Zugangsdaten,
- kein Import, kein Upload – das Ergebnis ist ausschließlich ein Dateiexport.

## 2. Vom POC zum Batch-Konverter

Der POC verarbeitete genau eine DOCX-Datei (`--input-file`). Die aktuelle
Version ist ein Batch-Konverter mit klar getrennten Schichten:

```text
Word-Dokument
  -> DOC-Konvertierung über Microsoft Word        (doc_converter.py, resources/convert_doc_to_docx.ps1)
  -> technische DOCX-/OOXML-Extraktion           (docx_reader.py, ooxml_traversal.py, models.py)
  -> automatische Profilerkennung                (profile_loader.py, profile_detector.py, step_table.py)
  -> semantisches, toolneutrales Testfallmodell  (lunar_parser.py, semantic_model.py)
  -> Bildextraktion und Bildzuordnung            (image_extractor.py, image_normalizer.py, image_assignment.py)
  -> Ziel-Renderer für das Import-JSON           (target_renderer.py)
  -> JSON- und Referenzvalidierung               (validator.py)
  -> atomarer Export je Testfall                 (filesystem.py)
  -> Reporting und Logging                       (reporting.py, logging_setup.py)
```

Orchestriert wird der Ablauf in `conversion_service.py`, die CLI liegt in `cli.py`.

| Ebene | Inhalt | Module |
|---|---|---|
| Technisches Quellmodell | Absätze, Tabellen, Zellen, Textsegmente, Bildreferenzen in Dokumentreihenfolge | `models.py`, `ooxml_traversal.py`, `docx_reader.py` |
| Semantisches Modell | Testfallfelder, Testschritte, Bildpositionen (`ImageMarker`) | `semantic_model.py`, `lunar_parser.py` |
| Zielmodell | `testcase.json` | ausschließlich `target_renderer.py` |

Neu hinzugekommene Module: `semantic_model.py`, `step_table.py`, `reporting.py`,
`source_discovery.py`, `image_normalizer.py`, `doc_converter.py` samt
PowerShell-Ressource `resources/convert_doc_to_docx.ps1`.

## 3. Unterstützte Eingabeformate

- `.docx` – direkte Verarbeitung.
- `.doc` – wird automatisch mit Microsoft Word nach `.docx` umgewandelt
  (Abschnitt 4) und danach wie `.docx` verarbeitet.

Verarbeitet werden alle Dateien mit diesen Endungen (Groß-/Kleinschreibung egal)
im Eingabeordner **und in allen Unterordnern**, sortiert nach relativem Pfad
(ohne Groß-/Kleinschreibung, z. B. `a.docx`, `Bereich B/b.docx`, `Bereich B/Tief/c.doc`).
Temporäre Word-Dateien (`~$…`) werden ignoriert, Verzeichnisverknüpfungen
(Symlinks/Junctions) nicht verfolgt. Liegt der Output-Ordner innerhalb des
Eingabeordners, wird er nicht durchsucht. Im Log steht je Datei der Pfad relativ
zum Eingabeordner. Ein Dokument entspricht genau einem Testfall.

## 4. `.doc`-Konvertierung mit Microsoft Word

`.doc`-Dateien werden während des Laufs automatisch über das lokal installierte
Microsoft Word (COM-Automatisierung) umgewandelt. Dazu ruft der Konverter je
Datei das mitgelieferte Skript `src/lunar_converter/resources/convert_doc_to_docx.ps1`
auf – ohne Shell, mit `pwsh` (bevorzugt) oder `powershell` aus dem `PATH`:

```text
pwsh -NoProfile -NonInteractive -File convert_doc_to_docx.ps1 -InputDir <temp>\eingabe -OutputDir <temp>\ausgabe
```

- Die Quelle wird in einen temporären Arbeitsordner kopiert und dort
  schreibgeschützt geöffnet; die Originaldatei bleibt unverändert.
- Makros werden beim Öffnen deaktiviert, Warndialoge unterdrückt;
  kennwortgeschützte Dokumente schlagen fehl statt einen Dialog zu öffnen.
- Word wird je Datei gestartet und am Ende nur beendet, wenn keine anderen
  Dokumente geöffnet sind (Dauer etwa 5–10 Sekunden je `.doc`).
- Schlägt die Umwandlung fehl (kein PowerShell, Word nicht startbar, Fehler beim
  Öffnen/Speichern, Zeitlimit 300 Sekunden), wird **nur diese Datei** mit
  `doc_conversion_failed` abgelehnt; `stdout`/`stderr` stehen im Log und Report.
- Die Ausführungsrichtlinie von PowerShell wird nicht umgangen; sie muss lokale
  Skripte erlauben (z. B. `RemoteSigned`).
- Voraussetzungen: Windows mit installiertem Word und angemeldetem Benutzer.
  Word-Automatisierung ist für interaktive Arbeitsplätze gedacht, nicht für
  unbeaufsichtigte Server-/Dienstausführung.

Das Skript kann auch eigenständig genutzt werden, z. B. um `.doc`-Dateien nur
umzuwandeln:

```powershell
.\src\lunar_converter\resources\convert_doc_to_docx.ps1 -InputDir C:\temp\input -OutputDir C:\temp\input_docx
```

## 5. Installation und virtuelle Umgebung

Unter Windows kann der Konverter mit `start_testfallkonverter.cmd` gestartet
werden. Beim ersten Start erstellt der Launcher `.venv` und installiert die
Abhängigkeiten aus `requirements.txt`; dafür sind Python 3.11+ und
Internetverbindung erforderlich. CLI-Argumente werden weitergereicht:

```powershell
.\start_testfallkonverter.cmd convert --input-dir .\input --output-dir .\output\lauf1
```

Für `.doc`-Dateien wird weiterhin Microsoft Word benötigt.

```powershell
cd Testfallkonverter
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -e .
```

`pip install -e .` ist optional; ohne Installation vorher `$env:PYTHONPATH = "src"` setzen.
Voraussetzung: Python 3.11 oder neuer.

## 6. CLI

### Profile vorab prüfen

Der Preflight prüft rekursiv alle `.docx`- und `.doc`-Dateien gegen alle Profile
und schreibt pro Datei Status, Treffer, Prüfgründe und Fehler in eine CSV:

Es wird empfohlen, den Preflight vor der eigentlichen Konvertierung auszuführen.
So lassen sich fehlende oder mehrdeutige Profilzuordnungen früh erkennen und
klären.

```powershell
python -m lunar_converter preflight `
  --input-dir .\input `
  --csv-path .\output\profile-preflight.csv
```

Eine Vorabkonvertierung ist nicht nötig. `.docx` wird direkt gelesen; `.doc`
wird temporär mit Microsoft Word nach `.docx` konvertiert. Der Preflight ändert
die Quelldateien nicht und erzeugt keine Testfall-Exports. Exit-Code `0` bedeutet,
dass für jede Datei genau ein Profil passt; `1` bedeutet mindestens einen
fehlenden/mehrdeutigen Treffer oder Dateifehler. Die CSV ist UTF-8 mit BOM und
Semikolon als Trennzeichen. Zusätzlich entsteht daneben eine Logdatei mit dem
gleichen Format wie bei der Konvertierung, standardmäßig `<csv-name>.preflight.log`.
Mit `--log-level DEBUG` werden auch die Einzelgründe jeder Profilprüfung protokolliert.

```powershell
python -m lunar_converter convert `
  --input-dir .\input `
  --output-dir .\output\lauf_2026-09-30
```

| Parameter | Pflicht | Bedeutung |
|---|---|---|
| `--input-dir` | ja | Eingabeordner mit `.docx`/`.doc`, inklusive Unterordner |
| `--output-dir` | ja | Output-Basisordner (muss leer sein oder wird angelegt) |
| `--profile <id>` | nein | Nur dieses Profil prüfen (Debugging/Migration), kein Fallback |
| `--log-level` | nein | `DEBUG`, `INFO` (Standard), `WARNING`, `ERROR` |
| `--dry-run` | nein | Analysieren und validieren, keine Exportdaten schreiben |
| `--fail-fast` | nein | Beim ersten Dateifehler abbrechen |
| `--schema-path` | nein | JSON-Schema, Standard `schema/testcase.schema.json` |
| `--config-dir` | nein | Konfigurationsordner mit `profiles/`, Standard `config/` |

Beispiele:

```powershell
# Ordner mit DOC- und DOCX-Dateien (.doc wird automatisch mit Word umgewandelt)
python -m lunar_converter convert --input-dir .\input --output-dir .\output\lauf1

# Debug-Ausführung
python -m lunar_converter convert --input-dir .\input --output-dir .\output\debug --log-level DEBUG

# Dry Run
python -m lunar_converter convert --input-dir .\input --output-dir .\output\probe --dry-run

# Nur ein bestimmtes Profil prüfen
python -m lunar_converter convert --input-dir .\input --output-dir .\output\lauf2 --profile lunar_standard_v1
```

## 7. Regeln für den Output-Basisordner

Vor jeder Dateiverarbeitung (Preflight):

1. Ordner existiert nicht → wird angelegt.
2. Ordner existiert und ist leer → wird verwendet.
3. Ordner enthält Dateien oder Unterordner → **gesamter Lauf bricht ab**
   (Exit-Code 2), es wird nichts geschrieben oder überschrieben.

Nach erfolgreichem Preflight entstehen zusätzlich `conversion.log` und
`conversion-report.json` im Output-Ordner. Das gilt auch für `--dry-run`; für
einen anschließenden echten Lauf daher einen neuen Output-Ordner verwenden.

Ausgabe je erfolgreicher Quelldatei:

```text
<output-dir>/
  conversion.log
  conversion-report.json
  <dateiname-ohne-erweiterung>/
    testcase.json
    screenshots/
      0001.png
      0002.png
```

Der Output ist auch bei Unterordnern im Eingabeordner **flach** (eine Ebene
Testfallordner), weil der Importer `<ordner>/*/testcase.json` liest.

Der Ordnername wird ausschließlich aus dem Dateinamen gebildet (nie aus der
Summary). Auch die `summary` in `testcase.json` wird aus dem Dateinamen ohne
`.doc`- oder `.docx`-Endung gebildet. Ungültige Zeichen (`<>:"/\|?*`, Steuerzeichen)
und – wie im POC – Leerraum werden durch `_` ersetzt; reservierte Windows-Namen (`CON`, `NUL`, …)
erhalten ein `_`. Ergeben zwei Quelldateien denselben Ordnernamen (z. B.
`Test Fall.docx` und `Test_Fall.docx` oder gleichnamige Dateien in verschiedenen
Unterordnern), wird die in der Sortierung spätere mit `output_name_conflict`
abgelehnt; die Meldung nennt die Datei, die den Namen bereits belegt.

## 8. Atomarer Export pro Testfall

1. Temporären Arbeitsordner anlegen (`tempfile`),
2. extrahieren, Bilder als PNG ablegen,
3. `testcase.json` rendern,
4. Pflichtfelder, JSON-Schema und Bildreferenzen validieren,
5. erst dann den vollständigen Ordner in einen versteckten Staging-Ordner im
   Output-Verzeichnis kopieren und per `os.rename` in einem Schritt freigeben.

Bei Fehlern wird aufgeräumt; es bleibt kein unvollständiger Testfallordner zurück.

## 9. Automatische Profilerkennung

Für jedes Dokument werden alle Profile aus `config/profiles/*.json` geprüft –
**nie** anhand des Dateinamens. Ein Profil passt genau dann, wenn

1. alle `required_markers` im normalisierten Dokumenttext vorkommen,
2. eine Tabelle alle `step_table.required_columns` abdeckt (Kopfzeile in den
   ersten drei Tabellenzeilen; Spaltennamen-Varianten über `field_aliases`),
3. mindestens eine fachlich befüllte Schrittzeile vorhanden ist (mindestens eine
   der Spalten System, Beschreibung, Erwartetes/Tatsächliches Ergebnis enthält
   Text oder Bild).

| Ergebnis | Verhalten |
|---|---|
| genau ein Profil passt | Profil wird verwendet, geloggt und im Report gespeichert |
| kein Profil passt | Status `no_matching_profile`, kein Export, weiter mit nächster Datei |
| mehrere Profile passen | Status `ambiguous_profile`, kein Export, **keine** willkürliche Auswahl |

Mit `--profile <id>` wird ausschließlich dieses Profil geprüft (ohne Fallback).
Ein unbekanntes Profil ist ein globaler Fehler (Exit-Code 2).

Textnormalisierung (zentral in `text_normalizer.py`): Groß-/Kleinschreibung,
Zeilenumbrüche, Tabs, Mehrfach-Leerzeichen, Bindestrichvarianten, Leerzeichen um
`/` und `-`, führende/nachgestellte Leerzeichen sowie Satzzeichen am Ende von
Überschriften (`.:;,!?`) werden ignoriert.

## 10. Profilformat

Profile sind klein und deklarativ. Erlaubt sind **ausschließlich**:

```json
{
  "id": "lunar_standard_v1",
  "name": "Standard-Testfall mit Systemspalte (LUNAR und EDDI)",
  "required_markers": ["Testablauf", "Kurzbeschreibung"],
  "step_table": {
    "required_columns": ["System / Komponente", "Beschreibung des Testschritts", "Erwartete Ergebnisse"]
  },
  "info_table": {
    "required_labels": ["Kurzbeschreibung"]
  },
  "field_aliases": {
    "testfallname": ["Testfall", "Testfallname"],
    "system": ["System / Komponente", "System / Transaktion"]
  }
}
```

`info_table.required_labels` ist optional: Die Tabelle mit zentralen Informationen
ist dann die erste Tabelle vor dem Testablauf (ohne Deckblatt), die alle diese
Bezeichner als Zelltext enthält. Fehlt sie, gibt es die Warnung
`info_table_not_found`. Ohne Angabe gilt die letzte Tabelle vor der Schritttabelle.

Unbekannte Schlüssel werden beim Laden als Konfigurationsfehler abgewiesen –
auch in `field_aliases` (Tippfehler fallen so sofort auf). Erlaubte Schlüssel in
`field_aliases`:

- `testfallname` – Bezeichner der Testfall-Zeile auf dem Deckblatt (`Testfall: <Name>`)
- Schrittspalten: `schritt_nummer`, `system`, `action`, `data`, `expected_result`, `actual_result`
  (`data` wird nach `steps[].data` übernommen; Spalten ohne Schlüssel werden ignoriert)

`required_markers` werden als wörtlicher Text im Dokumentkörper gesucht (nicht in
Kopf-/Fußzeilen). Eine Pflichtspalte gilt auch dann als vorhanden, wenn die
Tabelle einen Alias derselben Gruppe verwendet (z. B. „System / Transaktion“ für
„System / Komponente“).

Mitgelieferte Profile:

| Profil | Vorlage | Unterscheidungsmerkmal | Info-Tabelle |
|---|---|---|---|
| `lunar_standard_v1` | Standard-Testfall (LUNAR und EDDI, früher auch `eddi_standard_v1`) | Marker „Kurzbeschreibung“, Systemspalte „System / Komponente“ oder „System / Transaktion“ | enthält „Kurzbeschreibung“ |
| `lunar_legacy_v1` | ältere Vorlage | Spalte „Eingabedaten“, keine Systemspalte; die Spalte „Feld“ wird bewusst nicht übernommen | enthält „Beschreibung“ |
| `gh_standard_v1` | GH-Vorlage | Marker „Geschäftsvorfall“, Spalten „Geschäftsprozess-Schritte“, „Eingabedaten / besondere Angaben“, „Ausgabedaten / erwartetes Ergebnis“; keine Systemspalte; „Variante“ und „Feld“ werden nicht übernommen; die Tabellen „Varianten“, „Tatsächliche Ergebnisse“ und „Fehlerbeschreibung“ ebenfalls nicht | enthält „Geschäftsvorfall“ |

## 10a. Dokumentaufbau und Abbildung ins Ziel-JSON

Jedes Testdokument besteht grundsätzlich aus Deckblatt, Tabelle mit zentralen
Informationen und Schritttabelle:

| Dokumentteil | Erkennung | Ziel-JSON |
|---|---|---|
| Deckblatt: Label | Erster nichtleerer Absatz der ersten Zelle der Deckblatt-Tabelle | `labels`: einzelnes Label, z. B. `RWWS` oder `RWWS-GH` |
| Deckblatt: Prozesspfad | Nummerierte Prozesszeilen und alle nichtleeren Deckblattabsätze danach bis zur Testfall-Zeile, in Dokumentreihenfolge | `custom_fields.customfield_15909` = `"Prozess1/Prozess2/Prozess3/…"` |
| Deckblatt: optionale Testfallbeschreibung | Zeilen zwischen letzter Prozesszeile und Testfall-Zeile, ohne Zeilenumbrüche zusammengeführt | `description`: `h1. <Testfallbeschreibung>` |
| Deckblatt: `Testfall: <Name>` | Alias `testfallname` | kein Zielfeld; `summary` kommt aus dem Quelldateinamen |
| Tabelle mit zentralen Informationen | Tabelle mit den Bezeichnern aus `info_table.required_labels`, sonst letzte Tabelle vor der Schritttabelle; nie die Deckblatt-Tabelle | an `description` angehängte Jira-Wiki-Tabelle |
| Schritttabelle | Pflichtspalten des Profils | `steps[]` (wie bisher) |

Wiki-Tabelle: jede Zeile `|Zelle|Zelle|`, leere Zellen als `| |`, Zeilenumbrüche
in Zellen als `\\`, `|` im Text als `\|`. Vollständig leere Zeilen entfallen.
Bilder in der Tabelle werden an ihrer Position als `!0001.png!` verankert und als
globale Screenshots exportiert.
Bekannte Jira-Emoticon-Kürzel in Quelldaten werden escaped, damit Jira sie als
Text darstellt. Word-Checkboxen in der Info-Tabelle werden gezielt als `(/)`
(angekreuzt) bzw. `(x)` (leer) ausgegeben.

Beispiel:

```text
h1. Berücksichtigung von rechnungswirksamen Konditionen in der Bestellaktualisierung

|Kurzbeschreibung|Berücksichtigung von neuangelegten …|
|Voraussetzungen| |
|Stammdaten|Betriebe, Lieferanten, Artikel, Einkaufskonditionen|
```

**Testrepository-Pfad:** Der aus den Prozesszeilen gebildete Pfad wird unter
`custom_fields.customfield_15909` ausgegeben. Das ist die konfigurierte Xray-
Feld-ID für den Testrepository-Pfad.

Nicht mehr übernommen werden weitere Deckblattzeilen (z. B. „Verantwortlicher: …“)
und Inhalte weiterer Tabellen vor der Schritttabelle (z. B. Status/Version).

## 11. Neues Profil ergänzen

1. Bestehendes Profil nach `config/profiles/<neue_id>.json` kopieren.
2. `id` (eindeutig, `A-Z a-z 0-9 _ . -`) und `name` anpassen.
3. `required_markers` so wählen, dass sie das neue Layout **eindeutig** von
   anderen Profilen unterscheiden – sonst entsteht `ambiguous_profile`.
4. `step_table.required_columns` und `field_aliases` an die Bezeichnungen des
   neuen Layouts anpassen.
5. Mit `--dry-run --log-level DEBUG` gegen Beispieldokumente prüfen; die
   Gründe je Profil stehen im Log und unter `checked_profiles` im Report.
6. Einen Test in `tests/integration/test_profile_detection.py` ergänzen.

## 12. Bilder: Extraktion, Zuordnung, Jira-Wiki-Anker

- Ausgewertet wird nur der Dokumentkörper (`word/document.xml`). Bilder in
  Kopf- und Fußzeilen (z. B. Logos, auch bei abweichender erster Seite oder
  gerade/ungerade Seiten) werden **nie** extrahiert, erhalten keine Nummer und
  keinen Anker. Sie werden nur gezählt, im Log (`INFO`) gemeldet und im Report
  unter `ignored_header_footer_image_count` ausgewiesen.
- Erkannt werden DrawingML (`a:blip`) und VML (`v:imagedata`, ältere bzw.
  aus `.doc` konvertierte Dokumente). Bei `mc:AlternateContent` zählt nur die
  erste Variante, damit Bilder nicht doppelt erscheinen.
- Unterstützt: PNG, JPG, JPEG. PNG wird unverändert übernommen, JPG/JPEG nach
  PNG konvertiert (Pillow). Andere Formate (EMF, WMF, GIF, …), nur verknüpfte
  oder unlesbare Bilder: kein Export, kein Anker, Warnung im Log und Report.
- Dateinamen `0001.png`, `0002.png`, … je Dokument neu ab `0001`, in natürlicher
  Dokumentreihenfolge (oben → unten, in Tabellen zeilenweise links → rechts).
- Jede sichtbare Platzierung derselben Mediendatei erzeugt eine eigene PNG-Datei.

| Fall | Ergebnis |
|---|---|
| Bild in der Tabelle mit zentralen Informationen | globales `screenshots` + Anker in der Wiki-Tabelle der `description` |
| Bild in Zelle „Beschreibung des Testschritts“ | `steps[].attachments` + Anker in `steps[].action` |
| Bild in Zelle „Erwartete Ergebnisse“ | `steps[].attachments` + Anker in `steps[].expected_result` |
| Bild in Zelle „Tatsächliche Ergebnisse“ | `steps[].attachments` + Anker im Block „Tatsächliches Ergebnis“ |
| Bild außerhalb von Feldern, in Spalte System/Schritt-Nr., im Testfallnamen, frei positioniert (schwebend) | globales `screenshots`, **kein** Anker, Warnung `image_assignment_unclear` |

Anker haben exakt das Format `!0001.png!` und werden an der Bildposition im
Text eingefügt. Grundsatz: bei Unsicherheit Testfall-Anhang statt Schritt-Anhang.
Ein Schritt-Anhang erscheint nie zusätzlich in den globalen `screenshots`.
`steps[].screenshots` wird derzeit immer leer ausgegeben.

## 13. `conversion.log` und `conversion-report.json`

Logformat (UTF-8): `Zeitstempel Level [Datei: …] [Profil: …] Meldung`.

- `INFO`: Start, Ende, erkannte Profile, erfolgreiche Exporte,
- `WARNING`: nicht unterstützte Bilder, unklare Bildzuordnung, ignorierte Layout-Zeilen,
- `ERROR`: Datei-, Profil-, Pflichtfeld-, Validierungs- und Exportfehler,
- `DEBUG`: Profilprüfungen im Detail, technische Details, Stacktraces.

Fehlermeldung (Beispiel):

```text
Datei: input\TFB_001.docx
Profil: lunar_standard_v1
Fehler: required_field_missing
Feld: steps[2].expected_result
Ursache: Erwartetes Ergebnis im 3. Testschritt ist leer.
```

Report (Auszug):

```json
{
  "started_at": "2026-09-30T13:23:24+02:00",
  "finished_at": "2026-09-30T13:23:25+02:00",
  "input_dir": "input",
  "output_dir": "output\\lauf1",
  "dry_run": false,
  "profile_override": null,
  "summary": {"total": 2, "success": 1, "failed": 1, "skipped": 0},
  "files": [
    {
      "input_file": "input\\a.docx",
      "status": "success",
      "profile_detection_status": "matched",
      "detected_profile": "lunar_standard_v1",
      "checked_profiles": [{"profile": "lunar_standard_v1", "matched": true, "reasons": ["Alle Marker gefunden."]}],
      "output_directory": "output\\lauf1\\a",
      "planned_output_directory": null,
      "step_count": 9,
      "exported_image_count": 3,
      "ignored_header_footer_image_count": 5,
      "warnings": [],
      "errors": []
    }
  ]
}
```

`warnings`/`errors` enthalten Objekte mit `code`, `message` sowie optional
`field` und `details`. `profile_detection_status` ist `null`, wenn die Datei vor
der Profilerkennung scheiterte (z. B. `.doc`-Umwandlung oder beschädigte Datei). `planned_output_directory`
ist nur bei `--dry-run` gesetzt. `step_count` und `exported_image_count` werden
nur bei Erfolg befüllt.

## 14. Fehlerverhalten und Exit-Codes

| Exit-Code | Bedeutung |
|---|---|
| `0` | alle verarbeiteten Dateien erfolgreich (auch bei 0 Dateien) |
| `1` | mindestens eine Datei fehlgeschlagen |
| `2` | globaler Fehler vor der Verarbeitung: Output-Ordner nicht leer, Eingabeordner fehlt, Profil/Schema ungültig, unbekanntes `--profile`; ebenso ungültige CLI-Parameter |

Fehlercodes je Datei: `input_file_error`, `doc_conversion_failed`,
`no_matching_profile`, `ambiguous_profile`,
`required_field_missing`, `schema_validation_failed`, `missing_image_reference`,
`output_name_conflict`, `export_failed`, `unexpected_error`.

Fehler einer Datei brechen den Batch nicht ab. Mit `--fail-fast` werden die
restlichen Dateien als `skipped` gemeldet.

Pflichtregeln: `summary` (= Quelldateiname ohne `.doc`/`.docx`) nicht leer, mindestens ein Schritt,
je Schritt `system` und `action` nicht leer. Alle Pflichtfeldfehler eines
Dokuments werden gemeinsam gemeldet.

Übergreifende Platzhalterregeln (profilunabhängig, nur im Ziel-Renderer):

| Situation | Wert im Ziel-JSON |
|---|---|
| Erwartetes Ergebnis leer | `expected_result = "-"` (keine Warnung) |
| Schritttabelle hat keine Systemspalte (z. B. `lunar_legacy_v1`, `gh_standard_v1`) | `system = "nicht definiert"` |
| Systemspalte vorhanden, Zelle aber leer (z. B. Folgezeilen ohne Schritt-Nr.) | `system = "nicht definiert"` |
| Eingabedaten vorhanden | Werden mit `\n` an `action` angehängt; `data = ""` |
| Keine Eingabedaten-Spalte | `data = ""` |

## 15. Übergangslösung für tatsächliche Ergebnisse

Das Ziel-JSON hat kein eigenes Feld für tatsächliche Ergebnisse. Der
Ziel-Renderer (nur dort!) hängt für alle Profile einen Block an:

```text
<Erwartetes Ergebnis>

h3. Tatsächliches Ergebnis
<Tatsächliches Ergebnis>
```

Der Block wird angehängt, wenn ein tatsächliches Ergebnis vorhanden ist. Ist das
erwartete Ergebnis leer, steht davor der Platzhalter `-`.

## 16. Ziel-JSON und Schema anpassen

- Das Ziel-JSON entsteht ausschließlich in `target_renderer.py`
  (`XrayImportRenderer`). Für Custom-Field-Mappings `render_custom_fields`
  (bzw. `render_labels`, `render_components`) überschreiben oder erweitern –
  die Extraktion bleibt unverändert.
- Die `description` besteht aus `h1. <Testfallbeschreibung>` (falls vorhanden)
  und der Tabelle mit zentralen Informationen als Jira-Wiki-Tabelle (Abschnitt 10a).
- Die Geschäftsprozessstruktur steht unter der Xray-Feld-ID für den
  Testrepository-Pfad `custom_fields.customfield_15909` (`PROCESS_PATH_FIELD`).
- `schema/testcase.schema.json` (JSON Schema 2020-12) muss bei Änderungen am
  Renderer mitgepflegt werden. Bildreferenzen müssen dem Muster `*.png` ohne
  Pfadanteile entsprechen. Neue Felder sind wegen `additionalProperties: false`
  im Schema zu ergänzen.

## 17. Tests

```powershell
pip install -r requirements.txt
python -m pytest
```

Struktur:

```text
tests/
  unit/          Normalisierung, Profile, Validator, Dateisystem, Renderer, Extraktion
  integration/   CLI-Läufe: Batch, DOC (Word-Aufruf gemockt), Profile, Pflichtfelder, Bilder
  fixtures/      docx_factory.py – programmatisch erzeugte DOCX-Dokumente
```

Es werden keine echten Kundendokumente eingecheckt. `docx_factory.py` bildet
die Struktur echter LUNAR-Dokumente sowie das POC-Beispiel nach. Liegt
`input/sample.docx` lokal vor, prüft `tests/integration/test_real_sample.py`
zusätzlich dessen Extraktion (sonst übersprungen). Der Word-Aufruf wird in
den Tests gemockt. Ein echter Word-Test läuft nur, wenn die Umgebungsvariable
`LUNAR_TEST_WORD_DOC` auf eine `.doc`-Datei zeigt:

```powershell
$env:LUNAR_TEST_WORD_DOC = "C:\temp\input\beispiel.doc"; python -m pytest -k echte_word
```

## 18. Bekannte Grenzen

- Bildpositionen werden aus der Struktur (Absatz/Zelle) abgeleitet, nicht aus
  dem gerenderten Seitenlayout. Frei positionierte (schwebende) Bilder und Bilder
  außerhalb der Zellen (z. B. unter der Tabelle) werden daher bewusst als
  globale Screenshots ohne Anker behandelt.
- Textfelder/Textboxen werden in den Text übernommen, Formen, SmartArt,
  Diagramme und eingebettete OLE-Objekte nicht (deren Vorschaubilder sind meist
  EMF/WMF und damit nicht unterstützt).
- Vertikal verbundene Zellen (`vMerge`) übernehmen den Text der Ursprungszelle,
  Bilder werden nicht dupliziert. Zeilen mit einer über mehrere Inhaltsspalten
  verbundenen Zelle gelten als Layout-Zeile und werden mit Warnung ignoriert.
- Prozesszeilen werden nur an der Nummerierung (`NN.NN …`, `NN.NN.NNN …`)
  erkannt; ohne Prozesszeilen gibt es auch keine Testfallbeschreibung (`h1.`).
- Ohne `info_table.required_labels` gilt die letzte Tabelle vor der
  Schritttabelle als Info-Tabelle; weitere Tabellen davor werden nicht übernommen. Jira-Wiki-
  Tabellen kennen keine verbundenen Zellen; verbundene Zellen erscheinen als eine Zelle.
- Nummerierungen/Aufzählungszeichen aus Word-Listen werden nicht als Text
  übernommen.
- Kopf- und Fußzeilen werden nicht ausgewertet.
- Spalten der Schritttabelle ohne Zuordnung in `field_aliases` (z. B. „Feld“ in
  `lunar_legacy_v1`) werden ohne Warnung ignoriert.
- `scripts/create_sample_doc.py` schreibt nach `input/sample.docx` und
  überschreibt dort ein vorhandenes Dokument.
- `.doc` wird nur mit installiertem Microsoft Word unterstützt (Windows,
  interaktiver Benutzer). Beschädigte Dokumente können Word trotz
  unterdrükter Dialoge blockieren; nach dem Zeitlimit wird die Datei abgelehnt,
  ein hängender Word-Prozess muss dann ggf. im Task-Manager beendet werden.
