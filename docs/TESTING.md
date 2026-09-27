# Test- und Abnahmeplan für die Parallels-VM

Für den Personalabgleich 0.2.0 die [Personalregeln](ENTRA-LEHRER-SYNC.md) zusätzlich prüfen. Benutze eine Arbeitskopie der synthetischen `tests/fixtures/Lehrer-Testdaten.xlsx` in einem isolierten Mandanten oder mit vollständig isolierten Personalrollen. Teste L, CO-L, PA und OGTS, den Wechsel zwischen Schul- und Ganztagsrolle, bestehende Direkt- und Gruppenlizenzen, die zwei Exchange-Profile und die Passwortregel mit genau zwölf Zeichen und erzwungenem Erstwechsel. Prüfe `-Remove` auf Deaktivierung und Sitzungswiderruf bei erhaltenem Konto, Postfach, Gruppen und Lizenzen. `-WhatIf` darf keine Mandanten- oder Excel-Mutation auslösen. Halte den Personaltest getrennt vom historischen Schülerprotokoll fest.

Der [Check gesperrter Konten](GESPERRTE-KONTEN-VERWALTUNG.md) lädt zunächst nur lesend. Die Pester-Tests prüfen aktive Konten, die Aufnahme von freigegebenen und Ressourcenpostfächern sowie Konten ohne Exchange-Empfänger, die eindeutige Zuordnung über Objekt-IDs, den Abbruch bei einem Exchange-Bestandsfehler, `UserMailbox` und `RemoteUserMailbox`, Gäste, synchronisierte Konten, Verzeichnisrollen, fehlende Anmeldedaten, besessene Entra-Objekte mit Typ, Name und ID aus Graph-`AdditionalProperties` und den abgeschlossenen Fortschrittsbalken. Die TUI-Tests prüfen das Ausblenden von `SharedMailbox`, Pfeile, Markierung mit Leertaste für persönliche Benutzerpostfächer und eindeutig empfängerlose Konten, Nein und Esc ohne Aktion, Enter zum Löschen, `U` zum Entsperren, Ja nur für markierte Konten sowie `-WhatIf`. Bei mehreren ausgewählten Konten müssen Löschen und Entsperren jeden Schritt im Fortschrittsbalken anzeigen und den Balken auch nach `Skipped` oder `Failed` abschließen. Beide Schreibpfade müssen vor jeder Aktion Identität, Sperrstatus, Exchange-Zuordnung und Verzeichnisrollen erneut lesen. Vor der Aktion bei einem empfängerlosen Konto muss ein neuer vollständiger Exchange-Bestand ohne Treffer für seine Objekt-ID geladen werden. Eine inzwischen zugeordnete `SharedMailbox`, eine unklare Zuordnung oder ein Exchange-Fehler muss die Aktion überspringen. Der Löschpfad nennt bekannte Besitzobjekte in der Bestätigung, fragt sie unmittelbar vor dem Löschen erneut ab und warnt bei Treffern. Eigentümerkonten dürfen nach Bestätigung gelöscht werden. Fehler der erneuten Besitzabfrage führen weiterhin zu `Skipped`. Der Entsperrpfad aktiviert Eigentümerkonten, ohne eine Besitzabfrage zu benötigen. Aktive Konten und freigegebene Postfächer werden bei beiden Aktionen übersprungen. Prüfe bei einem synthetischen Skriptaufruf, dass die TUI nur in einem interaktiven Terminal startet und `-List` und `-PassThru` lesend bleiben. Für einen kontrollierten Mandantentest zuerst ausschließlich `-List` oder `-WhatIf` verwenden und die Tenant-ID prüfen. Das Skript kann weder eine 90-Tage-Frist noch eine echte Person automatisch belegen. Vor einer echten Löschung muss der Betreiber beides außerhalb der TUI anhand verlässlicher Daten prüfen.

Für die Personalabnahme zuerst die Mitglieder beider Rollen und alle Abgänge im lesenden Vergleich kontrollieren. Prüfe die tatsächlichen SKUs der Lizenzgruppen und die Exchange-Richtlinien im Mandanten. Verwende für einen Test mit einer unvollständigen Personalliste ausschließlich `-Update -CreateNewUsers -UpdateUsers`. Vergleiche den Dateihash vor und nach `-WhatIf`. Führe den vollständigen Lebenszyklus erst mit vollständig isolierten Personalrollen aus und dokumentiere die Ergebnisse getrennt nach Graph, Exchange und Excel. Der lokale Testlauf ersetzt weder die Windows-Dateisperrprüfung noch den Mandantennachweis.

Prüfe bei Schülern und Personal einen vollständigen `-Update`-Lauf mit einem bereits konformen und einem abweichenden bestehenden Postfach. Für das konforme Postfach darf keine Konfigurationsbestätigung erscheinen und das Ergebnis muss `Compliant` sein. Für die Abweichung muss die Bestätigung vor der Exchange-Änderung erhalten bleiben. Auch eine erst nach dem Vergleich entstandene Abweichung muss bei der erneuten Prüfung berücksichtigt werden.

Prüfe zusätzlich `L, PA` und `PA, JAS` aus einer einzigen Excel-Zeile. Beide müssen mit dem Schulprofil verarbeitet werden, ohne ein zweites Konto oder eine zweite Standardlizenz anzulegen. `JAS` allein und eine Kombination aus Schul- und Ganztagsjob müssen vor jeder Mutation scheitern. Kontrolliere nach einer synthetischen Neuanlage die Rückschreibung von Passwort, Objekt-ID, UPN und Mail, während die Jobzelle erhalten bleibt.

Die Tests verwenden ausschließlich erfundene Konten. Führe keine Live-Mutation mit echten Schülerdaten aus.

Der aktuelle lokale Prüfstand und die auszufüllende Abnahme-Checkliste stehen in [ABNAHME.md](ABNAHME.md).

> [!CAUTION]
> Abgänge sind alle aktiven, nicht zugeordneten Mitglieder der konfigurierten Gruppe `SEC-A-ROL-Schule_Schüler`. Die Erkennung ist nicht auf eine Testklasse oder einen Arbeitsmappenausschnitt begrenzt. Für jede Abgangsaktion muss die Excel-Datei die vollständige konfigurierte Schülerpopulation enthalten. Eine Testklasse in einer gemeinsam genutzten produktiven Rollengruppe ist nicht isoliert und kann reale Schüler als Abgänge markieren. Live-Tests mit `-Update`, `-DisableUsers` oder `-RevokeSessions` sind deshalb nur in einem vollständig isolierten Mandanten oder mit einer konfigurierten Schüler-Rollengruppe erlaubt, deren Mitglieder ausschließlich synthetische Konten sind.

## Freigabekriterien

Die Implementierung ist erst für den produktiven Einsatz freigegeben, wenn:

- alle Pester-Tests in PowerShell 7 unter Windows bestehen
- PSScriptAnalyzer keine ungeklärten Fehler meldet
- alle PowerShell-Dateien ohne Parserfehler sind
- Standardvergleich und `-WhatIf` nachweislich nichts schreiben
- Neuzugang, Änderung, Abgang und Exchange-Reparatur mit synthetischen Konten geprüft wurden
- keine produktive XLSX oder Passwortdatei von Git verfolgt wird

## 1. VM und Module dokumentieren

```powershell
$PSVersionTable

$required = @(
  'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users',
  'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Groups',
  'ExchangeOnlineManagement', 'ImportExcel', 'Pester', 'PSScriptAnalyzer'
)
Get-Module -ListAvailable $required |
  Sort-Object Name,Version -Descending |
  Select-Object Name,Version,Path |
  Format-Table -AutoSize
```

Dokumentiere Windows-Version, PowerShell-Version und die höchste gefundene Version jedes Moduls im Abnahmeprotokoll.

## 2. Repository-Tests

Im Repository-Root:

```powershell
pwsh -NoLogo -NoProfile -NonInteractive -File .\tests\Test-PesterDiscovery.ps1
Invoke-Pester .\tests -Output Detailed

$analyzerFindings = @(
  Invoke-ScriptAnalyzer -Path '.\Sync-SchuelerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\Sync-LehrerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\Verwalte-GesperrteEntraKonten.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\src' -Recurse -Settings '.\PSScriptAnalyzerSettings.psd1'
)
$analyzerFindings | Format-Table RuleName,Severity,ScriptName,Line,Message -Wrap
$unexpectedFindings = @($analyzerFindings | Where-Object RuleName -ne PSAvoidUsingWriteHost)
if ($unexpectedFindings) { throw 'Ungeklärte PSScriptAnalyzer-Befunde gefunden.' }

$parseErrors = foreach ($file in Get-ChildItem . -Recurse -File | Where-Object Extension -in '.ps1','.psm1','.psd1') {
  $tokens = $null
  $errors = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile(
    $file.FullName,
    [ref]$tokens,
    [ref]$errors
  )
  $errors
}
$parseErrors | Format-List
if ($parseErrors) { throw 'PowerShell-Parserfehler gefunden.' }
```

Erwartung: Pester erfolgreich, Analyzer ohne ungeklärte Fehler, `$parseErrors` leer.

Die separate Discovery-Prüfung startet ohne vorab importiertes Anwendungsmodul. Tests mit `InModuleScope` importieren das Modul ausdrücklich in `BeforeDiscovery`. Fixturepfade, Zustände und Mocks entstehen erst in der Run-Phase in `BeforeAll` beziehungsweise `BeforeEach`. Discovery-Blöcke verwenden keine dort erst entstehenden Werte. Testdateien ersetzen das Modul nicht mit `Import-Module -Force`, da das bereits gebundene Scriptblöcke ungültig machen könnte. Der Discovery-Test prüft diese Struktur und startet zusätzlich selbst eine frische PowerShell-Sitzung.

## 3. Git- und Datenschutzprüfung

```powershell
git status --short
git ls-files '*.xlsx'
git check-ignore .\Schueler.xlsx
git grep -n -I -E 'Passwort|PasswordProfile' -- ':!docs/**' ':!tests/**'
```

Erwartung:

- `Schueler.xlsx` wird ignoriert
- eigene Root- und Unterordner-Arbeitsmappen sowie konkrete Sicherungs- und temporäre Pfade werden ignoriert
- unter Git liegen ausschließlich die synthetischen Fixtures `tests/fixtures/Schueler-Testdaten.xlsx` und `tests/fixtures/Lehrer-Testdaten.xlsx`
- keine produktiven Namen, UPNs oder Passwörter sind versioniert
- Quellcode-Treffer für Passwortlogik enthalten keine echten Geheimnisse

Die Git-Regressionen verwenden temporäre Test-Repositories. Sie prüfen auch den Fehlerfall, dass nur die Quelle ignoriert ist. Dann darf weder eine vertrauliche Kopie noch eine Sicherung entstehen. Der Passworttest prüft das Format, die Einmaligkeit im Lauf und mehr als 64 Bit tatsächlich erreichbaren Auswahlraum.

## 4. Synthetische Testdaten

Verwende erfundene Personen, zum Beispiel:

| Fall | Vorname | Nachname | Klasse | Klassenlehrer |
|---|---|---|---|---|
| Neuzugang | Lina | Testwald | `JK1-3g2_1` | UPN eines Test-Managers |
| UPN-Kollision | Luis | Testwald | `JK1-3g2_1` | UPN eines Test-Managers |
| Klassenwechsel | Mira | Beispielstern | `JK4-6m2_4` | UPN eines Test-Managers |
| Abgang | Theo | Demoklang | nur im Mandanten | UPN eines Test-Managers |

Alle Testkonten und Gruppen müssen eindeutig als synthetisch gekennzeichnet und nach der Abnahme kontrolliert bereinigt werden. Nutze keine Namen existierender Kinder oder Beschäftigter. Erstelle diese Live-Testpopulation nur im vollständig isolierten Mandanten oder in der ausschließlich synthetischen konfigurierten Schüler-Rollengruppe. Eine bloße Testklasse innerhalb der produktiven Rollengruppe reicht nicht aus.

Erzeuge eine Arbeitskopie außerhalb des Repository-Roots:

```powershell
$testRoot = 'C:\Temp\EntraSchuelerSync-Abnahme'
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
Copy-Item .\tests\fixtures\Schueler-Testdaten.xlsx (Join-Path $testRoot 'Schueler-Abnahme.xlsx')
$testFile = Join-Path $testRoot 'Schueler-Abnahme.xlsx'
```

Passe nur die erfundenen Zeilen und den vorgesehenen Test-Manager an.

## 5. Lesenden Standardmodus beweisen

```powershell
$before = Get-FileHash $testFile -Algorithm SHA256
$result = .\Sync-SchuelerEntra.ps1 -File $testFile
$after = Get-FileHash $testFile -Algorithm SHA256

$before.Hash -eq $after.Hash
$result.Mode
$result.HasErrors
```

Erwartung:

- Hash bleibt identisch
- `Mode` ist `Compare`
- keine Graph- oder Exchange-Schreiboperation
- kein Warten auf fehlende Postfächer
- Tabellen für Neuzugänge, Abgänge, Änderungen und Bestehende

## 6. WhatIf beweisen

```powershell
$before = Get-FileHash $testFile -Algorithm SHA256
$result = .\Sync-SchuelerEntra.ps1 -File $testFile -Update -WhatIf
$after = Get-FileHash $testFile -Algorithm SHA256

$before.Hash -eq $after.Hash
$result.Actions | Format-Table Phase,Status,UserPrincipalName
```

Erwartung:

- Excel-Hash bleibt identisch
- keine neuen Konten, Attribute, Manager oder Gruppen
- keine Passworterzeugung, kein Backup und keine Berichtsdatei, auch mit `-OutputFile`
- keine Exchange-Schreiboperation
- keine 60-Sekunden-Wartezeit
- geplante Phasen tragen Status `WhatIf`

Wiederhole mit selektiven Schaltern:

```powershell
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -CreateNewUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -UpdateUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -DisableUsers -RevokeSessions -WhatIf
```

## 7. Neuzugang live testen

Dieser Schritt verändert den Testmandanten. Prüfe zuerst Tenant-ID, UPN und Zielgruppen.

```powershell
$result = .\Sync-SchuelerEntra.ps1 -File $testFile -Update -CreateNewUsers
$result.Actions | Format-Table Phase,Status,UserId,UserPrincipalName,RecoveryCommand -Wrap
```

Prüfe für den erfundenen Neuzugang:

```powershell
$row = Import-Excel $testFile | Where-Object Vorname -eq 'Lina'
$row.Passwort.Length
$user = Get-MgUser -UserId $row.EntraObjectId -Property id,displayName,givenName,surname,userPrincipalName,mail,mailNickname,department,officeLocation,companyName,employeeType,usageLocation,ageGroup,consentProvidedForMinor,legalAgeGroupClassification,accountEnabled
$user | Format-List
Get-MgUserManager -UserId $user.Id
Get-MgUserMemberOfAsGroup -UserId $user.Id -All | Select-Object Id,DisplayName
```

Erwartung:

- Passwortlänge genau 12
- zwölf zufällige, gut unterscheidbare ASCII-Buchstaben und Ziffern mit mindestens einem Zeichen jeder Klasse
- `ForceChangePasswordNextSignIn = false`, über das erstellte PasswordProfile oder einen kontrollierten Anmeldetest bestätigt
- Name bleibt mit Umlauten korrekt, UPN ist normalisiert
- Konto erst nach erfolgreicher Excel-Rückschreibung aktiviert
- Gruppen `SEC-A-LIC-O365A1Student`, `SEC-A-ROL-Schule_Schüler` und `SEC-A-CLS-JK1-3g2_1`
- Office Location `G2`
- `ageGroup = Minor`, `consentProvidedForMinor = Granted`, `usageLocation = DE`
- `legalAgeGroupClassification = MinorWithParentalConsent` wird nur gelesen
- genau eine Backup-Datei enthält den vorherigen Arbeitsmappenstand

Melde das Testkonto einmal mit dem Initialpasswort an. Es darf kein erzwungener Passwortwechsel erscheinen.

### Manuellen Add-Modus testen

Verwende ausschließlich ein erfundenes Testkonto und starte ohne Transcript:

```powershell
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin' -WhatIf
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin'
```

Erwartung bei der Neuanlage:

- keine Excel-Datei wird gelesen oder verändert
- das Passwort hat genau 12 Zeichen und erscheint genau einmal im Terminal, auch bei einem simulierten Teilfehler nach der Kontoerstellung
- Ergebnisobjekt und `-OutputFile` enthalten das Passwort nicht
- Konto, Pflichtattribute, Manager und Gruppen werden vor der Aktivierung verifiziert
- Exchange wird sofort und danach höchstens fünfmal mit jeweils 60 Sekunden Wartezeit geprüft

Führe denselben Befehl danach erneut mit einer geänderten Klasse aus. Erwartung: Der vorhandene Schüler wird aktualisiert. Es entsteht kein weiteres Konto und kein neues Passwort. Der vorhandene UPN bleibt unverändert. Deaktiviere das synthetische Konto und wiederhole `-Add`. Das Konto darf erst nach geprüftem Pflichtzustand aktiviert werden. Simuliere außerdem einen Fehler nach der Kontoerstellung und prüfe den ausgegebenen Wiederanlauf mit `-EntraObjectId`.

## 8. UPN-Kollision testen

Reserviere den ersten UPN-Kandidaten mit einem synthetischen Benutzer oder Exchange-Empfänger. Führe den Vergleich aus und prüfe:

- Warnung enthält den kollidierenden Kandidaten und den gewählten Ersatz
- Präfix wächst vom ersten zum zweiten Buchstaben und danach weiter
- nach vollständigem Vornamen beginnt der numerische Fallback bei `2`
- Kollisionsquellen aus UPN, Mail, SMTP-Proxyadresse und Exchange-Empfänger werden erkannt
- eine Adresse desselben Objekt-IDs wird wiederverwendet

Nutze den tatsächlich erzeugten UPN für alle weiteren Prüfungen.

## 9. Änderung und Gruppennormalisierung testen

Ändere die Klasse eines synthetischen bestehenden Schülers von `JK1-3g2_1` auf `JK4-6m2_4`. Füge ihm außerdem direkt eine fremde statische `SEC-A-ROL-...`-Gruppe und eine alte `SEC-A-CLS-...`-Gruppe hinzu.

```powershell
.\Sync-SchuelerEntra.ps1 -File $testFile
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -UpdateUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -UpdateUsers
```

Erwartung:

- `department = JK4-6m2_4`
- `officeLocation = M2`
- neue Klassengruppe wurde vor Entfernung der alten Gruppe bestätigt
- nur `SEC-A-ROL-Schule_Schüler` bleibt als direkte Rollengruppe
- nur `SEC-A-CLS-JK4-6m2_4` bleibt als direkte Klassengruppe
- Lizenzgruppe bleibt vorhanden
- Passwortspalte bleibt bytegenau unverändert
- dynamische oder geerbte Testgruppe wird nicht entfernt und als Warnung ausgegeben

## 10. Abgang testen

Entferne nur die erfundene Abgangszeile aus der Testarbeitsmappe. Mindestens eine andere gültige Zeile muss verbleiben, eine leere Schülerliste wird abgewiesen. Prüfe zuerst Compare und WhatIf, danach Sperren und Sitzungswiderruf im selben Lauf:

```powershell
.\Sync-SchuelerEntra.ps1 -File $testFile
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -DisableUsers -RevokeSessions -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -DisableUsers -RevokeSessions
```

Erwartung:

- Konto ist deaktiviert
- Sitzungswiderruf wurde von Graph bestätigt
- Konto, Lizenz und Gruppen wurden nicht gelöscht
- ein erneuter Vergleich zeigt das deaktivierte Konto ohne Excel-Zeile nicht mehr als Abgang, auch wenn es weiterhin Mitglied der Schüler-Rollengruppe ist
- ein erneuter lesender Lauf führt keine Mutation aus

### Manuellen Remove-Modus testen

Verwende einen erfundenen UPN, der direkt Mitglied der Schüler-Rollengruppe ist:

```powershell
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com'
```

Erwartung: Genau dieses Konto wird deaktiviert und seine Sitzungen werden widerrufen. Der Benutzer wird nicht gelöscht. Gruppen und Lizenzen bleiben erhalten. Wiederhole den Test mit einem Benutzer außerhalb der direkten Schüler-Rollengruppe. Dieser Lauf muss vor jeder Mutation fehlschlagen.

## 11. Exchange Online testen

### Sofort vorhandenes Postfach

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com'
```

Lies danach `Get-Mailbox` und `Get-CASMailbox` aus und gleiche jeden Wert mit [ENTRA-SCHUELER-SYNC.md](ENTRA-SCHUELER-SYNC.md) ab. Prüfe beide AuditAdmin-Varianten, einmal ohne und einmal mit `CommunicationsCompliance`.

### Verzögerte Bereitstellung

Verwende einen synthetischen Neuzugang, dessen Postfach bei der ersten Prüfung noch fehlt. Miss Zeit und Aufrufe. Erwartung:

- eine sofortige Prüfung
- höchstens fünf weitere Prüfungen
- 60 Sekunden zwischen den Wiederholungen
- nur eine Wartezeit pro Batch-Runde
- gefundene Postfächer werden aus dem Batch entfernt
- nach sechs erfolglosen Prüfungen Status `Pending` mit fertigem Reparaturbefehl

### Schreibfehler

Entziehe in einer kontrollierten Testkonstellation einen benötigten Exchange-Parameter oder simuliere ihn ausschließlich im Pester-Test. Erwartung: `Set-Mailbox` oder `Set-CASMailbox` meldet sofort `Failed`. Der Schreibfehler wird nicht als vermeintlich fehlendes Postfach wiederholt.

## 12. Idempotenz

Nach einem vollständig erfolgreichen Testlauf:

```powershell
$before = Get-FileHash $testFile -Algorithm SHA256
$result = .\Sync-SchuelerEntra.ps1 -File $testFile
$after = Get-FileHash $testFile -Algorithm SHA256

$result.Comparison.ChangedStudents.Count
$before.Hash -eq $after.Hash
```

Erwartung: keine Änderungen für aktive, vollständig konforme Schüler und identischer Excel-Hash. Ein deaktiviertes Konto ohne Excel-Zeile darf nicht erneut als Abgang erscheinen. Ein deaktiviertes Konto mit eindeutig passender Excel-Zeile muss dagegen als Reaktivierung unter Änderungen erscheinen.

## 13. Aufräumen

Vor dem Löschen erfundener Konten sichere das Abnahmeprotokoll ohne Passwörter. Entferne danach ausschließlich die dokumentierten synthetischen Objekte. Lösche keine produktiven Benutzer oder Gruppen.

```powershell
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false

Get-ChildItem $testRoot -Force
```

Entferne die lokale Testarbeitsmappe, Backups und temporären Dateien nach der vereinbarten sicheren Löschregel. Prüfe abschließend:

```powershell
git status --short
git ls-files '*.xlsx'
```

## Abnahmeprotokoll

Halte fest:

- Datum, VM, Windows- und PowerShell-Version
- getestete Modulversionen
- Tenant-ID und Testbereich, ohne Tokens oder Passwörter
- Commit-ID des getesteten Stands
- Ergebnis jeder Sektion 1 bis 13
- IDs ausschließlich der erfundenen Testobjekte
- offene Abweichungen und verantwortliche Freigabe

## Lokale Gesamtprüfung und CI

```powershell
./tests/Test-PesterDiscovery.ps1
./tests/Invoke-LocalVerification.ps1
```

Der gemeinsame Prüflauf parst die PowerShell-Quelldateien, prüft sie mit PSScriptAnalyzer und führt Pester aus. Nur die bewusst eingesetzten farbigen Überschriften (`PSAvoidUsingWriteHost`) sind als Analyzer-Ausnahme zugelassen. Jeder andere Analyzer-Befund und jeder fehlgeschlagene Test lässt den Lauf scheitern.

`.github/workflows/powershell-tests.yml` führt Discovery und Gesamtprüfung auf Windows und macOS aus. Die Modulversionen sind festgelegt. Der Workflow hat ausschließlich Leserechte auf Repository-Inhalte und benötigt keine Mandantenzugangsdaten. Er wird lokal vorbereitet, ein erfolgreicher lokaler Lauf belegt keinen ausgeführten GitHub-Workflow.

Die Audit-Regressionen prüfen leere Imports, einen fremden oder mehrdeutigen Exchange-Mandanten, erneute Mandantenprüfung vor der TUI-Aktion, nicht mehr gültige Rollenmitgliedschaften, nachträglich lizenzierende Rollen, gesperrte Wiederanlaufkonten, fehlgeschlagene Excel-Sicherung, abgelehnte Lehrerupdates, manuelles Add ohne falsche Abgänge und vertrauliche Berichtspfade. Der Berichtsschutz wird auch gegen bereits vorhandene Dateien und nicht existierende Unterverzeichnisse in temporären Git-Repositories geprüft.
