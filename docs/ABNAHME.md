# Abnahmestand und VM-Protokoll

Stand der Release-Prüfung: 12. September 2026 auf `codex/v0.1.1-recovery-fix` vor dem Merge für `v0.1.1`.

Die automatisierten PowerShell-Prüfungen wurden auf macOS erfolgreich ausgeführt. Ein lesender Lauf gegen den vorgesehenen Mandanten wurde vom Betreiber durchgeführt und lieferte die erwarteten Vergleichstabellen. Produktive Schreibaktionen wurden im Rahmen dieser lokalen Prüfung nicht unabhängig ausgeführt oder bestätigt. Die kontrollierte Windows- und Mandantenabnahme bleibt deshalb getrennt dokumentiert.

## Lokal ausgeführte Prüfungen

| Prüfung | Ergebnis |
|---|---|
| PowerShell | PowerShell 7.6.6 auf macOS |
| Pester | 228 Tests erkannt, 227 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| PSScriptAnalyzer | Keine Befunde für Einstiegsskript und Modul mit `PSScriptAnalyzerSettings.psd1` |
| Nativer PowerShell-Parser | 29 versionierte `.ps1`-, `.psm1`- und `.psd1`-Dateien ohne Parserfehler |
| `Test-ModuleManifest` | Erfolgreich, Modulversion `0.1.1`, PowerShell-Mindestversion `7.0` |
| Manueller Wiederanlauf | Unit- und Integrationstests bestätigen `-Add -EntraObjectId` für ein passendes, noch nicht gruppiertes Teilkonto sowie die erneute Graph-Prüfung vor jeder Mutationsphase. Normaler Abgleich, Fremdkonten und ein zwischenzeitlich verändertes Konto bleiben blockiert |
| `git diff --check` | Keine Whitespace-Fehler |
| `git check-ignore -v Schueler.xlsx` und synthetische Zielpfade | XLSX-Regeln greifen im Root und in Unterordnern, auch für Sicherungen, temporäre Kopien und Großschreibung der Endung |
| `git ls-files '*.xlsx'` | Ausschließlich `tests/fixtures/Schueler-Testdaten.xlsx` verfolgt |
| Modulinventar | Graph SDK 2.39.0, ExchangeOnlineManagement 3.10.1, ImportExcel 7.8.10, Pester 5.7.1 und PSScriptAnalyzer 1.25.0 |
| Lesender Mandantenvergleich | Vom Betreiber erfolgreich ausgeführt, keine lokale Bestätigung produktiver Schreibaktionen |
| Passwort-Auswahlraum, unabhängig aus dem Wortkatalog berechnet | 847 eindeutige Wörter, 220.505 eindeutige Zehnbuchstaben-Präfixe, 19.845.450 mögliche Ausgaben, rund 24,24 Bit |

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

- [x] Native Parserprüfung aller 29 versionierten PowerShell-Dateien
- [x] `Test-ModuleManifest` auf macOS erfolgreich
- [x] Gesamte Pester-Suite auf macOS erfolgreich, tatsächliche Anzahl Passed/Failed/Skipped dokumentiert
- [x] PSScriptAnalyzer ohne ungeklärte Befunde
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
