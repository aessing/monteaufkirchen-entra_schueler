# Abnahmestand und VM-Protokoll

## Personalabgleich 0.2.0

Die Erweiterung benötigt eine eigene kontrollierte Windows- und Mandantenabnahme mit synthetischen Konten. Zu prüfen sind beide direkten Personalrollen, die A3-/A1-Gruppen-SKUs, Lizenzbestandsschutz bei Jobwechsel, ABP und Adresslistenfilter, Exchange-Richtlinien, die tatsächliche Purview-Auditaufbewahrung und der Erstwechsel bei genau zwölf Passwortzeichen. Ein voller Abgangstest darf nur mit einer vollständig isolierten Personalpopulation erfolgen. Die folgenden 0.1.1-Zahlen dokumentieren den früheren Schülerstand und sind keine Ergebnisse für 0.2.0.

Die lokalen Ergebnisse für 0.2.0 stehen unten in einem eigenen Abschnitt. Windows-Tests wurden für 0.2.0 nicht ausgeführt. Ein vom Betreiber ausgeführter lesender Mandantenvergleich meldete zunächst 28 `Preflight.Identity`-Fehler. Nach der lokalen Korrektur meldete der Betreiber keine Fehler mehr und zeigte eine Abgangsliste mit 14 Konten. Deren Aktivierungszustände wurden nicht übermittelt. Der Vergleich wurde nach der Änderung der Abgangslogik noch nicht wiederholt.

### Lokal ausgeführte Prüfungen für 0.2.0 am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| PowerShell | 7.6.6 auf macOS |
| Pester-Gesamtsuite | 285 erkannt, 284 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Frische Pester-Discovery | 285 Tests in einem neuen PowerShell-Prozess erkannt, keine Testausführung |
| Nativer Parser | 35 `.ps1`-, `.psm1`- und `.psd1`-Dateien geprüft, 0 Parserfehler |
| `Test-ModuleManifest` | Version `0.2.0`, Exporte `Invoke-SchuelerSync` und `Invoke-LehrerSync` |
| PSScriptAnalyzer | 0 Befunde für beide Einstiegsskripte und `src`, mit `PSScriptAnalyzerSettings.psd1` |
| Git-Prüfung | `git diff --check` ohne Befund, nur synthetische XLSX-Fixtures ausgenommen |
| Lehrerarbeitsmappe | Nach der gemeldeten Sperrfehlermeldung ließ sich `Lehrer2026-27.xlsx` lokal exklusiv zum Lesen und Schreiben öffnen. Die Datei wurde dabei nicht verändert |
| Mandant und Windows | Betreiber meldete nach der Lizenzkorrektur einen lesenden Vergleich ohne Fehler und mit 14 Abgängen. Aktivierungszustände offen. Kein Mandantenlauf und kein Windows-Test durch Codex |

Die lokalen Tests nutzten Pester 5.7.1, ImportExcel 7.8.10 und PSScriptAnalyzer 1.25.0 aus einem temporären Modulverzeichnis. Die Graph- und Exchange-Integrationen wurden in den Tests gemockt. Die produktiven Gruppen, Lizenz-SKUs und Richtlinien bleiben bis zur kontrollierten Mandantenabnahme unbestätigt.

Die zusätzlichen Tests prüfen die Lehrer-Fortschrittsphasen, den Vergleich vor Kontoaktionen, den vollständigen `-OutputFile`-Bericht sowie die Schreibprüfung bei gesperrter Arbeitsmappe mit und ohne geplante Rückschreibung. Die neue Fortschrittsanzeige wurde lokal mit Mocks geprüft, nicht in einer interaktiven Windows-Konsole oder im Mandanten.

Die Schülerregression wurde ebenfalls lokal geprüft: Bereits deaktivierte Konten ohne Excel-Zeile fehlen in der Abgangsliste. Ein eindeutig zugeordnetes deaktiviertes Konto mit Excel-Zeile erscheint als Änderung und wird nach erfolgreicher Update-Prüfung aktiviert. Der Integrationstest bestätigt zudem, dass ein bereits abgegangenes und gesperrtes Konto im folgenden lesenden Vergleich nicht erneut erscheint. Der vom Betreiber gezeigte Mandantenfall wurde nicht erneut live ausgeführt.

Ein zusätzlicher Berichtstest reproduzierte den mehrzeiligen Umbruch der Lehrer-Kategorie „Bestehende“ mit langen technischen Spalten. Nach der kompakten Spaltenauswahl enthält der Bericht für diesen synthetischen Eintrag eine Datenzeile. Die Live-Ausgabe mit 30 Personen wurde nicht erneut ausgeführt.

Die reale `Lehrer2026-27.xlsx` wurde lokal nur lesend importiert: 30 Zeilen, einschließlich `L, PA` und `PA, JAS`, ohne Importfehler. Eine synthetische Kopie bestätigte die Rückschreibung von Passwort, Objekt-ID, UPN und Mail bei erhaltenem Mehrfachjob. In die reale Datei wurde nicht zurückgeschrieben. Der vom Betreiber gemeldete erste lesende Mandantenvergleich führte wegen 28 Lizenz-Vorprüfungsfehlern zu keiner Personeneinteilung. Ein fokussierter Test reproduzierte den Fehler mit `State = Active` und `Error = None`. Nach dessen Korrektur meldete der Betreiber keine Fehler und 14 Abgänge, darunter ein Funktionskonto. Die lokale Abgangslogik wurde daraufhin so angepasst, dass gesperrte Konten ohne Excel-Zeile nicht als Abgänge erscheinen und passende gesperrte Bestandskonten als Reaktivierung geplant werden. Diese Änderung wurde noch nicht im Mandanten geprüft.

### Zusätzliche lokale TUI-Prüfung am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester-Gesamtsuite nach TUI-Erweiterung | 318 erkannt, 317 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| PSScriptAnalyzer und Parser | 0 Befunde für TUI-Modul und Einstiegsskript, 0 Parserfehler |
| Modulmanifest | Version `0.2.0`, vier Exporte einschließlich `Invoke-DisabledAccountTui` |
| Synthetische TUI in einem Pseudoterminal | Leertaste, Enter und Ja führten mit `-WhatIf` zu `WhatIf` ohne Löschaufruf. Ohne `-WhatIf` wurde ausschließlich ein gemocktes `Remove-MgUser` für die synthetische Objekt-ID aufgerufen und `Deleted` angezeigt |
| Lesender CLI-Modus | `-List` und `-PassThru` lieferten mit synthetischen Graph- und Exchange-Antworten die bisherige Prüfliste |
| Git-Prüfung | `git diff --check` ohne Befund |

Die Tests verwendeten keine produktive Verbindung. Die echte Kontolöschung, die Microsoft-Graph-Berechtigung im Mandanten, eine Windows-Konsole und die manuelle Prüfung von Personenzuordnung und Sperrfrist sind nicht live verifiziert.

### Alle gesperrten Entra-Konten in der Prüfliste

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester nach Erweiterung des Inventars | 334 erkannt, 333 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Gezielte Inventar- und TUI-Tests | 45 bestanden, 0 fehlgeschlagen. Zwei weitere Schutztests für fehlende Exchange-Empfänger bestanden danach im Gesamtlauf. |
| Synthetischer `-List`-Aufruf | 3 gesperrte Konten angezeigt: `UserMailbox`, `SharedMailbox` und ein Konto ohne Exchange-Empfänger. Kopfzeile: 2 ohne persönliches Postfach. |
| Synthetische TUI im Pseudoterminal mit `-WhatIf` | Vor der späteren Ausblendung von `SharedMailbox` waren alle 3 Konten sichtbar. Leertaste beim Konto ohne Exchange-Empfänger ließ `Markiert: 0` stehen und zeigte den Hinweis „nur zur Prüfung sichtbar“. Esc beendete die TUI ohne Aktion. |
| PSScriptAnalyzer | 0 Fehler oder Warnungen außer der bestehenden Stilregel für `Write-Host` bei farbiger Konsolenausgabe |
| Git-Prüfung | `git diff --check` ohne Befund |

Die neuen Inventar- und TUI-Fälle wurden nur mit synthetischen Graph- und Exchange-Antworten geprüft. Ein erneuter Lauf im produktiven Mandanten fand nicht statt.

### Entsperren von Konten mit Besitzobjekten

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regressionstest vor der Korrektur | Ein gesperrtes Konto mit sieben Besitzobjekten wurde als `Skipped` gemeldet. 28 von 29 gezielten Tests bestanden. |
| Gezielte TUI-Tests nach der Korrektur | 29 von 29 bestanden. Der Entsperrpfad fragte Besitzobjekte nicht erneut ab und setzte bei gemocktem Graph `AccountEnabled = true`. Der Löschpfad übersprang Eigentümerkonten weiterhin. |
| Pester-Gesamtsuite | 335 erkannt, 334 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI mit `-WhatIf` | Ein gesperrtes `UserMailbox`-Konto mit einem Besitzobjekt war markierbar. Nach `U` und Ja erschien `WhatIf`, ohne Änderung am Konto. |

Im produktiven Mandanten wurden keine Konten entsperrt. Die Wirkung der Korrektur für die vom Betreiber genannten Konten wurde nicht live geprüft.

### Freigegebene Postfächer aus der TUI ausgeblendet

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Gezielte TUI-Tests | 30 von 30 bestanden. `SharedMailbox` fehlte in der übergebenen TUI-Auswahl und konnte nicht in eine Entsperraktion gelangen. |
| Pester-Gesamtsuite | 336 erkannt, 335 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI mit `-WhatIf` | Vor der späteren Freigabe empfängerloser Konten zeigte die TUI 2 von 3 gesperrten Entra-Konten. `SharedMailbox` fehlte, das Konto ohne Exchange-Empfänger blieb als „Nur prüfen“ sichtbar. Esc beendete die Auswahl ohne Aktion. |
| Synthetisches `-List` | Alle 3 Konten blieben in der lesenden Inventarliste, einschließlich `SharedMailbox`. |

Im produktiven Mandanten wurde die TUI nach dieser Änderung nicht ausgeführt.

### Konten ohne Exchange-Empfänger auswählen

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regressionstest vor der Änderung | 52 gezielte Tests, davon 7 fehlgeschlagen. `RecipientStatus` fehlte, und die TUI sowie der Löschpfad übersprangen ein Konto ohne Empfänger. |
| Gezielte Check- und TUI-Tests nach der Änderung | 54 bestanden, 0 fehlgeschlagen. Der Check unterscheidet `None`, `Unique`, `Ambiguous` und `Unreadable`. Nur `None` und persönliche Benutzerpostfächer sind auswählbar. |
| Schutz vor `SharedMailbox` | Eine nach der Auswahl im frischen Exchange-Bestand gefundene `SharedMailbox` führte zu `Skipped`, auch mit großgeschriebener und von Leerzeichen umgebener Objekt-ID. Uneindeutige Zuordnung und Fehler beim erneuten Bestandsabruf führten ebenfalls zu `Skipped`. `Remove-MgUser` wurde in diesen Fällen nicht aufgerufen. |
| Pester-Gesamtsuite | 341 erkannt, 340 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI mit `-WhatIf` | Das Konto ohne Exchange-Empfänger war markierbar. Nach Enter und Ja meldete der erneute Bestandsabgleich `WhatIf`. `SharedMailbox` blieb ausgeblendet. Es gab keinen Löschaufruf. |

Ein echter Lösch- oder Entsperrlauf im produktiven Mandanten fand nicht statt. Die TUI kann nicht selbst belegen, ob ein empfängerloses Konto einer realen Person gehört oder seit mindestens 90 Tagen gesperrt ist.

### Fortschrittsanzeige bei mehreren ausgewählten Konten

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Gezielte Check- und TUI-Tests | 56 bestanden, 0 fehlgeschlagen. Für Löschen und Entsperren wurden die Schritte „1 von 2“ und „2 von 2“ sowie der Abschluss des Fortschrittsbalkens geprüft. Der Abschluss wurde auch bei übersprungenen Konten und Aktionsfehlern geprüft. |
| Pester-Gesamtsuite | 343 erkannt, 342 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI im Pseudoterminal | Je zwei Konten wurden für Löschen und Entsperren markiert. Beide Bestätigungen listeten die zwei Namen. Beide `-WhatIf`-Läufe meldeten zweimal `WhatIf` und zeigten während der Verarbeitung den Fortschritt. `SharedMailbox` blieb ausgeblendet. |
| PSScriptAnalyzer | 0 Fehler oder Warnungen in den drei betroffenen PowerShell-Dateien, abgesehen von der bestehenden Stilregel für `Write-Host` bei farbiger Konsolenausgabe |
| Git-Prüfung | `git diff --check` ohne Befund |

Die synthetischen Terminalaufrufe verwendeten keine produktive Verbindung und führten weder Löschung noch Entsperrung aus. Die Darstellung in einer Windows-Konsole und die echte Graph-Aktion bleiben ungeprüft.

### Löschung trotz Besitzobjekten nach Warnung

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regressionstest vor der Änderung | 40 gezielte TUI-Tests, davon 38 bestanden und 2 fehlgeschlagen. Der Bestätigungsdialog zeigte kein Besitzobjekt, und der Löschpfad meldete `Skipped`. |
| Gezielte TUI-Tests nach der Änderung | 40 bestanden, 0 fehlgeschlagen. Die Bestätigung nennt Typ, Name und ID eines Besitzobjekts. Ein erneut bestätigter Eigentümer wird nach Warnung gelöscht. Ein Fehler bei der erneuten Besitzabfrage führt weiter zu `Skipped`. |
| Pester-Gesamtsuite | 346 erkannt, 345 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI im Pseudoterminal | Ein Eigentümerkonto war markierbar. Der Dialog zeigte `group: Schulteam (group-1)` und warnte vor fehlender Eigentümerschaft. Nach Ja erschien die Warnung vor dem `-WhatIf`-Löschaufruf. Es gab keinen Schreibzugriff. |
| PSScriptAnalyzer | 0 Fehler oder Warnungen im geänderten TUI-Modul, abgesehen von der ausgeschlossenen bestehenden Stilregel für `Write-Host` |
| Git-Prüfung | `git diff --check` ohne Befund |

Die aktuelle Regel wurde mit gemocktem Graph und synthetischen Konten geprüft. Eine echte Löschung im Mandanten wurde nicht durchgeführt.

### Lokale Prüfung der Entsperroption am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester-Gesamtsuite | 329 erkannt, 328 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI im Pseudoterminal | Leertaste, `U` und Ja führten mit `-WhatIf` zu `WhatIf` ohne Änderung |
| Entsperrpfad mit gemocktem Graph | `Update-MgUser` erhielt ausschließlich die ausgewählte Objekt-ID und `AccountEnabled = true`. Aktive Konten, geänderte Identität und Shared Mailboxes wurden übersprungen. Graph-Fehler wurden als `Failed` gemeldet. |
| Berechtigung und Syntax | `User.EnableDisableAccount.All` wurde für die bestätigte Entsperraktion angefordert, kein Schreibscope unter `-WhatIf`. PowerShell-Parser ohne Befund. |
| PSScriptAnalyzer | 0 Befunde im TUI-Modul. Im Einstiegsskript bestehen 8 Stilwarnungen für `Write-Host`, das die farbigen Konsolenüberschriften ausgibt. Keine Fehler. |
| Git-Prüfung | `git diff --check` ohne Befund |

Es gab keine produktive Verbindung oder Entsperrung. Die echte Graph-Berechtigung und die Wirkung im Mandanten sind nicht live verifiziert.

### Lokale Prüfung der Exchange-Bestätigungen am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester-Gesamtsuite | 331 erkannt, 330 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Lehrer- und Schülerpostfächer | Fokussierte Tests bestätigen `Compliant` ohne Konfigurationsaufruf bei leeren Exchange-Abweichungen. Abweichende Postfächer erreichen weiter den bestätigungspflichtigen Pfad. |
| Schüler-Integration | Ein konformes Bestands-Postfach wurde im vollständigen Update nicht erneut konfiguriert. Zwei abweichende Postfächer wurden weiterhin verarbeitet. |
| Parser und PSScriptAnalyzer | Geänderte PowerShell-Dateien ohne Parserfehler. PSScriptAnalyzer meldet 2 bestehende `Write-Host`-Stilwarnungen in der Lehrerorchestrierung und 1 Information an anderer Stelle im Exchange-Modul, keine Fehler. |
| Git-Prüfung | `git diff --check` ohne Befund |

Der vom Betreiber gezeigte Lauf und die Exchange-Einstellungen im Live-Mandanten wurden nach der Korrektur nicht erneut geprüft.

## Historischer Schülerabgleich 0.1.1

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
  Invoke-ScriptAnalyzer -Path '.\Sync-LehrerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
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
- [ ] Zweiter lesender Vergleich ohne Änderungen für konforme aktive Schüler, deaktivierte Konten ohne Excel-Zeile erscheinen nicht erneut als Abgänge
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

## Auditkorrekturen vom 27.09.2026, Version 0.2.0

Die folgenden Ergebnisse wurden nach den Korrekturen lokal unter macOS mit PowerShell 7.6.6, Pester 5.7.1 und PSScriptAnalyzer 1.23.0 ermittelt. Die bereits vorhandenen Arbeitsänderungen blieben erhalten. Es wurden keine produktiven Konten verändert und keine Änderungen gepusht oder veröffentlicht.

| Prüfung | Tatsächliches Ergebnis |
| --- | --- |
| `tests/Invoke-LocalVerification.ps1` mit lokalen Pester-/Analyzer-Manifesten | Erfolgreich, 374 Tests, 373 bestanden, 0 fehlgeschlagen, 1 übersprungen, Laufzeit 29,19 Sekunden |
| Discovery in frischer PowerShell-Sitzung | 374 Tests gefunden, keine Discovery-Fehler |
| PowerShell-Syntax der Quellen, Einstiegsskripte und Konfiguration | Keine Parserfehler |
| PSScriptAnalyzer für diese Dateien | Keine Fehler und keine ungeklärten Warnungen. 12 bekannte `PSAvoidUsingWriteHost`-Hinweise für farbige Überschriften |
| Windows-Dateisperrprüfung | Unter macOS übersprungen, weiterhin auf Windows auszuführen |
| Echte CLI-Aufrufe beider Sync-Skripte mit Excel als Berichtsziel | Abgewiesen, Exitcode 1, Quellhash unverändert, keine Cloud-Anmeldung |
| Echter Schüler-CLI-Aufruf mit ausschließlich Excel-Überschriften und `-Update` | Abgewiesen, Exitcode 1, Quellhash unverändert, keine Cloud-Anmeldung |
| Audit-Regressionen | Leerschutz, Mandantenabgleich, Rollen-/Lizenzänderungen, Wiederanlauf, Bestätigungslogik, Berichtsschutz und manuelles Add bestanden |
| Neues Schülerpasswortformat | Genau zwölf Zeichen, alle drei Zeichenklassen, mehr als 64 Bit Auswahlraum, Einmaligkeit und begrenzte Wiederholungen geprüft |
| Abschließender Secret-Musterscan | 74 aktuelle Textdateien geprüft, keine Treffer für private Schlüssel, GitHub-/AWS-Token oder feste Client-Secrets |
| Git-Schutz und Diff | Produktive Arbeitsmappen weiterhin ignoriert, `git diff --check` ohne Befund |
| Dokumentation nach dem abschließenden Abgleich | 11 Dokumentationstests bestanden, 0 fehlgeschlagen |
| Workflow-YAML | Lokal erfolgreich geparst |
| GitHub-Workflow | Vorbereitet, nicht ausgeführt, da kein Push erfolgt ist |

Die frühen Regressionstests reproduzierten zunächst fünf Auditfehler. Zwischenläufe fanden unvollständige Test-Mocks, eine unerwünschte Bestätigungsweitergabe durch `ForEach-Object -MemberName` sowie Kodierungswarnungen in zwei Konfigurationsdateien. Diese Punkte wurden korrigiert. Ein früher Exchange-Reparaturtest versuchte wegen eines fehlenden Mocks eine Graph-Anmeldung. Der Browseraufruf scheiterte, der Prozess wurde abgebrochen. Es wurde keine Anmeldung abgeschlossen. Der reine Schüler-Exchange-Reparaturmodus verwendet weiterhin ausschließlich Exchange. Die abschließenden Prüfungen liefen ohne Anmeldung und ausschließlich mit synthetischen Testdaten.

Die frühere Messung von rund 24,24 Bit in diesem Dokument beschreibt das historische Zweiwortformat. Ab diesem Korrekturstand gilt für neue Schülerkonten das zufällige Zeichenformat. Bestehende Passwörter sowie die Erstwechselregel bleiben unverändert.

## Abschließender lokaler Prüflauf vor Git-Freigabe 0.2.0

Am 27.09.2026 wurde der vollständige lokale Prüflauf mit PowerShell 7.6.6, Pester 5.7.1 und PSScriptAnalyzer 1.23.0 erneut ausgeführt:

- 374 Tests, davon 373 bestanden, 0 fehlgeschlagen und 1 Windows-spezifischer Test unter macOS übersprungen.
- Laufzeit der Pester-Tests: 26,65 Sekunden.
- Keine PowerShell-Parserfehler und keine ungeklärten Analyzer-Befunde.
- Secret-Musterscan über 70 vorhandene, für Git vorgesehene Textdateien ohne Treffer. Produktive Arbeitsmappen bleiben ignoriert. Die neue Lehrer-Fixture enthält ausschließlich erfundene Personen.
- `git diff --check` ohne Befund.

Dieser lokale Prüflauf führte keine produktiven Microsoft-365-Änderungen aus. Die Windows- und macOS-Jobs des neuen GitHub-Workflows laufen nach dem Push. Ihre Ergebnisse sind separat in GitHub Actions zu prüfen und ersetzen keine kontrollierte Mandantenabnahme.
