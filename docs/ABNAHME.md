# Abnahmestand und VM-Protokoll

Stand der lokalen Prüfung: 11. September 2026 auf `codex/entra-schueler-sync-impl`, einschließlich der Korrekturen aus der abschließenden Codeprüfung.

Die Implementierung liegt zur abschließenden Codeprüfung und zur Abnahme in der Windows-Parallels-VM vor. Eine produktive Freigabe ist damit noch nicht erteilt. Lokal wurden **keine Pester-Tests erfolgreich ausgeführt und keine Live-Mandantentests durchgeführt**.

## Lokal ausgeführte Prüfungen

| Prüfung | Ergebnis |
|---|---|
| `git status --short` | Arbeitsbaum vor diesem Protokoll sauber |
| `git diff --check` und `git diff 5303d37..HEAD --check` | Keine Whitespace-Fehler |
| `git check-ignore -v Schueler.xlsx` und synthetische Zielpfade | XLSX-Regeln greifen im Root und in Unterordnern, auch für Sicherungen, temporäre Kopien und Großschreibung der Endung |
| `git ls-files '*.xlsx'` | Ausschließlich `tests/fixtures/Schueler-Testdaten.xlsx` verfolgt |
| Quellcode-Suche nach Passwortzuweisungen, privaten Schlüsseln und Tokens | Kein literales Geheimnis gefunden. Die Formatvorlage im Generator ist kein Passwort |
| Quellcode-Suche und Sichtprüfung der Passwort-Ausgabepfade | Kein direkter Passwortwert in den Ausgabeaufrufen. Das ersetzt keinen Laufzeittest aller PowerShell-Ausgabekanäle |
| Modulmanifest und Konfiguration, statisch gelesen | Export `Invoke-SchuelerSync`, Version `0.1.0`, PowerShell-Mindestversion `7.0`, Domain, Schüler-Rollen-ID, Passwortlänge `12`, EXO-Wiederholungen `5` und Wartezeit `60` konsistent |
| Lokale Markdown-Dateiziele | 20 Verweise auf vorhandene Dateien geprüft. Sieben bestehende relative GitHub-Issue-/Security-Verweise sind Hosting-Links und wurden nicht als lokale Dateien bewertet |
| `file docs/assets/entra-schueler-sync-hero.png` | PNG, 2048 × 768 Pixel |
| `unzip -t tests/fixtures/Schueler-Testdaten.xlsx` | Alle ZIP-Einträge fehlerfrei |
| Fixture-XML gelesen | Fünf Pflichtspalten, zwei erfundene Schülerzeilen, keine Passwort-, UPN- oder Objekt-ID-Werte |
| Testinventar | 14 Testdateien, 177 statische `It`-Deklarationen und ein separater Discovery-Prüfer. Dies ist keine Anzahl ausgeführter oder bestandener Tests |
| Passwort-Auswahlraum, unabhängig aus dem Wortkatalog berechnet | 847 eindeutige Wörter, 220.505 eindeutige Zehnbuchstaben-Präfixe, 19.845.450 mögliche Ausgaben, rund 24,24 Bit |
| Discovery-Struktur, statisch geprüft | Alle zehn Testdateien mit `InModuleScope` importieren in `BeforeDiscovery`. Fixturepfade entstehen innerhalb der Run-Phase, kein Test ersetzt das Modul mit `-Force` |

Zusätzlich wurden alle 27 `.ps1`-, `.psm1`- und `.psd1`-Dateien mit `tree-sitter-powershell 0.26.4` untersucht. Der Ersatzparser meldet keine Fehler- oder Missing-Knoten. Leere Mock-Scriptblöcke verwenden nun explizit `{ return }`, damit sie auch mit diesem Parser lesbar sind. Die native PowerShell-Parserprüfung sowie die neue Fresh-Session-Discovery-Prüfung bleiben Aufgabe der VM.

Diese drei lokalen Befehle wurden separat versucht:

```sh
pwsh -NoLogo -NoProfile -Command 'Invoke-Pester ./tests -Output Detailed -PassThru'
pwsh -NoLogo -NoProfile -Command 'Invoke-ScriptAnalyzer -Path ./Sync-SchuelerEntra.ps1 -Settings ./PSScriptAnalyzerSettings.psd1; Invoke-ScriptAnalyzer -Path ./src -Recurse -Settings ./PSScriptAnalyzerSettings.psd1'
pwsh -NoLogo -NoProfile -Command 'Test-ModuleManifest ./src/SchuelerSync/SchuelerSync.psd1'
```

Alle drei endeten vor dem Start von PowerShell mit Exitcode `127` und exakt:

```text
zsh:1: operation not permitted: pwsh
```

Damit sind Pester, PSScriptAnalyzer und die native Manifestprüfung lokal **nicht geprüft**. Die echte Schülerdatei wurde weder gelesen noch kopiert oder verändert. Es fand keine Anmeldung und keine Schreiboperation bei Graph oder Exchange Online statt.

## Reproduzierbare VM-Prüfung

Verwende PowerShell 7 im Repository-Root. Installiere zuerst die in [ENTRA-SCHUELER-SYNC.md](ENTRA-SCHUELER-SYNC.md) aufgeführten Module. Das Repository enthält keine getestete Versionssperre. Dokumentiere die tatsächlich geladenen Versionen und den geprüften Commit, statt ungeprüfte Versionsnummern zu übernehmen.

```powershell
$ErrorActionPreference = 'Stop'
$PSVersionTable
git rev-parse HEAD

$required = @(
  'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users',
  'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Groups',
  'ExchangeOnlineManagement', 'ImportExcel', 'Pester', 'PSScriptAnalyzer'
)
Import-Module Pester -MinimumVersion 5.0 -Force
Import-Module PSScriptAnalyzer -Force
Get-Module -ListAvailable $required |
  Sort-Object Name,Version -Descending |
  Select-Object Name,Version,Path

$parseErrors = foreach ($relativePath in git ls-files '*.ps1' '*.psm1' '*.psd1') {
  $tokens = $null
  $errors = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PWD $relativePath), [ref]$tokens, [ref]$errors
  )
  $errors
}
$parseErrors | Format-List
if ($parseErrors) { throw 'Native PowerShell-Parserprüfung fehlgeschlagen.' }

Test-ModuleManifest .\src\SchuelerSync\SchuelerSync.psd1
$pesterResult = Invoke-Pester .\tests -Output Detailed -PassThru
if ($pesterResult.Result -ne 'Passed' -or $pesterResult.TotalCount -eq 0) {
  throw 'Pester nicht erfolgreich oder keine Tests ausgeführt.'
}

$analyzerFindings = @(
  Invoke-ScriptAnalyzer -Path '.\Sync-SchuelerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\src' -Recurse -Settings '.\PSScriptAnalyzerSettings.psd1'
)
$analyzerFindings | Format-Table RuleName,Severity,ScriptName,Line,Message -Wrap
if ($analyzerFindings.Count -gt 0) { throw 'Analyzer-Befunde sind vor der Freigabe zu klären.' }
Get-Module $required | Select-Object Name,Version
```

Die Live-Befehle und Sollwerte stehen vollständig in [TESTING.md](TESTING.md). Erstelle dafür eine Arbeitskopie der erfundenen Fixture. Nutze einen vollständig isolierten Mandanten oder eine konfigurierte Schüler-Rollengruppe mit ausschließlich synthetischen Mitgliedern. Eine Testklasse in der produktiven Schüler-Rollengruppe reicht nicht aus, weil Abgänge für die gesamte konfigurierte Rollenpopulation berechnet werden.

## Abnahme-Checkliste

Jede Position benötigt Datum, tatsächliches Ergebnis und gegebenenfalls eine bereinigte Fehlermeldung. Schreibe keine Passwörter oder Tokens in das Protokoll.

- [ ] Native Parserprüfung aller 25 PowerShell-Dateien, insbesondere `tests/Entra.Tests.ps1`
- [ ] `Test-ModuleManifest` erfolgreich
- [ ] Gesamte Pester-Suite erfolgreich, tatsächliche Anzahl Passed/Failed/Skipped dokumentiert
- [ ] PSScriptAnalyzer ohne ungeklärte Befunde
- [ ] Standardvergleich zeigt alle Tabellen, unveränderter Excel-Hash, keine Graph-/EXO-Schreiboperationen
- [ ] Vollständiges und selektives `-WhatIf` ohne Passworterzeugung, Backup, Dateiänderung, Mandantenänderung oder EXO-Wartezeit
- [ ] Isolierter Neuzugang bleibt bis zur verifizierten Excel-Rückschreibung deaktiviert
- [ ] Initialpasswort exakt 12 Zeichen, kontrollierte Erstanmeldung ohne erzwungenen Passwortwechsel
- [ ] Alle Entra-Attribute, Manager und drei Pflichtgruppen nachgelesen
- [ ] UPN-Kollision über Graph-/EXO-Adressen erkannt, alternatives Präfix und numerischer Fallback geprüft
- [ ] Klassen- und Lehrerwechsel geprüft, Zielgruppen vor Entfernung alter Gruppen nachgelesen, bestehendes Passwort unverändert
- [ ] Synthetischer Abgang deaktiviert und Sitzungswiderruf separat bestätigt
- [ ] EXO-Sollwerte vollständig nachgelesen, einschließlich Richtlinien, AuditDelegate, AuditOwner, beider AuditAdmin-Varianten und sämtlicher CAS-Werte
- [ ] Sofortiges Postfach und verzögerte Bereitstellung geprüft, höchstens sechs Prüfungen und fünf Batch-Wartezeiten von jeweils 60 Sekunden
- [ ] EXO-Schreibfehler wird sofort gemeldet und nicht als Bereitstellungsproblem wiederholt
- [ ] `-ConfigureExchangeOnlineOnly -Mail <synthetischer-UPN>` einschließlich Aliase und Idempotenz geprüft
- [ ] Zweiter lesender Vergleich ohne Änderungen für konforme aktive Schüler, deaktivierter Abgang darf sichtbar bleiben
- [ ] Wiederherstellung nach gesperrter oder konkurrierend veränderter Excel-Datei geprüft, neue Konten bleiben deaktiviert, erhaltene Datei- und Backup-Varianten zugeordnet
- [ ] Windows-Dateisperr-, Umbenennungs- und Wiederherstellungsverhalten aus `WorkbookIdentitySafety.Tests.ps1` bestätigt
- [ ] Ausgabeprüfung aller Kanäle mit ausschließlich erfundenen Passwortwerten, einschließlich Passwort-Retry-Fehlern
- [ ] Git-Hygiene nach VM-Test bestätigt und ausschließlich dokumentierte synthetische Testobjekte kontrolliert bereinigt

## Ergebnis eintragen

| Feld | Eintrag |
|---|---|
| Datum / Prüfer | offen |
| Geprüfter Commit | offen |
| Windows / PowerShell | offen |
| Tatsächlich geladene Modulversionen | offen |
| Isolierter Testmandant / Rollenpopulation | offen |
| Pester Passed / Failed / Skipped | offen |
| Analyzer / nativer Parser / Manifest | offen |
| Lesender Vergleich / WhatIf | offen |
| Synthetischer Lebenszyklus / EXO / Wiederherstellung | offen |
| Verbleibende Abweichungen | offen |
| Produktive Freigabe | ausstehend |
