![3D-Comic-Illustration eines sicheren Schülerabgleichs zwischen Tabelle und Cloud-Verzeichnis](docs/assets/entra-schueler-sync-hero.png)

# Entra-Synchronisation für Schüler und Personal

Aktuelle Version: **0.2.0**

Dieses PowerShell-Tool vergleicht eine Excel-Schülerliste mit Microsoft Entra ID und Exchange Online. Der Standardlauf ist rein lesend. Änderungen benötigen ausdrücklich `-Update`, `-Add`, `-Remove` oder den getrennten Exchange-Reparaturmodus.

Für Lehrer, Co-Lehrer, Pädagogen und OGTS gibt es zusätzlich `Sync-LehrerEntra.ps1`. Es liest `Lehrer.xlsx` mit den Spalten `Name mit Rufname`, `Vorname`, `Nachname` und `Job`. Der Standardlauf zeigt ebenfalls nur den Vergleich. Der Lehrerabgleich zeigt den Fortschritt wie der Schülerabgleich und gibt bei `-Update` den Vergleich vor den Aktionen aus.

Ein vollständiges `-Update` prüft bei Schülern und Personal auch bestehende Exchange-Postfächer. Sind ihre Einstellungen bereits korrekt, meldet der Lauf `Compliant` ohne Bestätigungsfrage. Nur eine aktuell festgestellte Abweichung führt zur Bestätigung einer Exchange-Änderung.

`Verwalte-GesperrteEntraKonten.ps1` erfasst alle aktuell gesperrten Entra-Konten, auch ohne Exchange-Empfänger sowie freigegebene und Ressourcenpostfächer. Die TUI blendet `SharedMailbox` aus. Persönliche Benutzerpostfächer und eindeutig als postfachlos erkannte Konten sind auswählbar. Andere Empfängertypen und unklare Zuordnungen bleiben nur zur Prüfung sichtbar. Pfeiltasten bewegen die Auswahl, Leertaste markiert zulässige Konten, Enter öffnet die Löschbestätigung und `U` die Entsperrbestätigung mit allen ausgewählten Namen. Esc beendet die Auswahl. Erst Ja startet die erneute Einzelprüfung und die gewählte Aktion mit Fortschrittsanzeige pro Konto. `-List` und `-PassThru` bleiben lesend und enthalten auch `SharedMailbox`. Ein Sperrdatum wird bisher nicht gespeichert, deshalb kann das Skript eine 90-Tage-Frist nicht bestätigen. [Details zur Kontenverwaltung](docs/GESPERRTE-KONTEN-VERWALTUNG.md).

Die drei Konsolenausgaben trennen Bereiche mit Leerzeilen und färben Überschriften. `-OutputFile` speichert die Vergleichsberichte mit denselben Abständen als reinen Text ohne Farbcodes.

> [!IMPORTANT]
> Die Excel-Datei enthält personenbezogene Daten und nach Neuanlagen auch Initialpasswörter. Arbeitsmappen im Root und in Unterordnern sowie Sicherungen und temporäre Kopien sind per `.gitignore` ausgeschlossen. Nur die beiden mitgelieferten synthetischen Fixtures sind ausgenommen. Vor der Rückschreibung prüft das Skript jeden konkreten Dateipfad. Committe keine produktiven Schüler- oder Personallisten, Sicherungen oder Konsolenausgaben mit personenbezogenen Daten.

## Schutzprüfungen in 0.2.0

- Leere Schülerlisten lösen keine Abgänge oder Kontoänderungen aus. Leere Personallisten sind für Updates ebenfalls gesperrt.
- Bei kombinierten Graph-/Exchange-Läufen müssen beide Verbindungen denselben Mandanten verwenden. Mehrere aktive Exchange-Verbindungen oder fehlende Tenant-IDs führen zum Abbruch. Die TUI prüft den Exchange-Mandanten vor jeder Einzelaktion erneut.
- Rollenmitgliedschaften und die Lizenzfreiheit von Personalrollen werden unmittelbar vor den jeweiligen Änderungen erneut gelesen. Die Schülerrolle kommt auch im Lehrerabgleich aus `config/SchuelerSync.psd1`.
- `-OutputFile` akzeptiert ausschließlich neue `.txt`-Dateien. Innerhalb eines Git-Repositories müssen sie ignoriert sein, zum Beispiel unter `Berichte/`. Bestehende Dateien und symbolische Links werden abgewiesen. Verwende pro Lauf einen neuen Dateinamen.
- Neue Lehrerkonten werden gesperrt angelegt. Passwort und Objekt-ID werden vor der weiteren Konfiguration in Excel gesichert. Details zum Wiederanlauf stehen in der [Personalanleitung](docs/ENTRA-LEHRER-SYNC.md).
- Schülerpasswörter neuer Konten verwenden zwölf zufällige, gut unterscheidbare Buchstaben und Ziffern. Bestehende Passwörter und der ausgeschaltete Erstwechsel bleiben erhalten.

Syntax, Analyzer und Tests laufen lokal mit `./tests/Invoke-LocalVerification.ps1`. Der Workflow `powershell-tests.yml` führt dieselben Prüfungen auf Windows und macOS aus, sobald die Änderungen in GitHub vorliegen. Lokale Ergebnisse stehen in [ABNAHME.md](docs/ABNAHME.md).

## Das erledigt das Tool

- Tabellen für Neuzugänge, Abgänge, Änderungen und bestehende Schüler
- Bereits deaktivierte Schülerkonten ohne Excel-Zeile werden nicht erneut als Abgänge gezeigt. Eindeutig zugeordnete gesperrte Konten in Excel erscheinen als Reaktivierung unter Änderungen.
- Zuordnung ohne gelieferte UPN über eindeutigen Rufnamen und Nachnamen innerhalb der Schüler-Rollengruppe
- Feldgenaue Abweichungen für Entra, Manager, Gruppen und Exchange Online
- Eindeutige UPNs unter `@monteaufkirchen.com`, inklusive Umlautumschrift und Kollisionsprüfung
- Sichere Neuanlage mit deaktiviertem Konto, Pflichtzustand, geschützter Excel-Rückschreibung und erst anschließender Aktivierung
- Exakt 12 Zeichen lange, kryptografisch zufällige und gut lesbare Initialpasswörter nur für neue Schüler, ohne erzwungenen Wechsel bei der ersten Anmeldung
- Genau eine Schüler-Rollengruppe und eine aktuelle Klassengruppe pro aktivem Schüler
- Automatischer Exchange-Abgleich mit sofortiger Prüfung und bis zu fünf Wiederholungen im Abstand von 60 Sekunden
- Separater Exchange-Reparaturlauf für einzelne oder mehrere UPNs
- Manuelles Anlegen oder Aktualisieren eines einzelnen Schülers mit `-Add`
- Sicheres Deaktivieren eines einzelnen Schülers und Widerrufen seiner Sitzungen mit `-Remove`

## Voraussetzungen

- Windows oder macOS mit PowerShell 7.0 oder neuer
- Git für die Prüfung vertraulicher Excel- und Berichtspfade
- Interaktive Anmeldung bei Microsoft Graph und Exchange Online
- Berechtigtes Administratorkonto im richtigen Mandanten
- Die Module `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions`, `Microsoft.Graph.Groups`, `ExchangeOnlineManagement` und `ImportExcel`

```powershell
Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Users.Actions,Microsoft.Graph.Groups -Scope CurrentUser
Install-Module ExchangeOnlineManagement,ImportExcel -Scope CurrentUser
```

## Schnellstart

### Gesperrte Konten verwalten

```powershell
.\Verwalte-GesperrteEntraKonten.ps1
.\Verwalte-GesperrteEntraKonten.ps1 -List
.\Verwalte-GesperrteEntraKonten.ps1 -WhatIf
```

`-List` zeigt `DisplayName`, `UPN`, `UserId`, Exchange-Empfängertyp, `SEC-A Gruppen`, `Last Login`, Besitzobjekte mit Typ, Name und ID sowie Schutz- und Prüfgründe. Die Kopfzeile zählt Konten ohne persönliches Postfach, darunter freigegebene Postfächer und Konten ohne Exchange-Empfänger. Ein fehlender Empfänger erzeugt keine Warnung. Exchange wird einmal als Bestand gelesen und über die Entra-Objekt-ID zugeordnet. Mit `-PassThru` erhältst du zusätzlich das Ergebnisobjekt mit `RecipientStatus` und `OwnedObjects` je Konto. Ohne Schalter benötigt die TUI ein interaktives Terminal. Die Löschung benötigt `User.ReadWrite.All`, das Entsperren `User.EnableDisableAccount.All` und beide eine passende Entra-Administratorrolle. Unmittelbar vor jeder Aktion prüft das Skript das Konto erneut. Bei einem Konto ohne Exchange-Empfänger muss eine neue vollständige Exchange-Abfrage diesen Zustand bestätigen. Ein inzwischen zugeordnetes Postfach, einschließlich `SharedMailbox`, blockiert die Aktion. Aktive, Gast- oder lokal synchronisierte Konten sowie Konten mit Verzeichnisrollen werden übersprungen. Besitzobjekte erscheinen vor der Löschbestätigung mit Typ, Name und ID. Ihr Besitz blockiert die Löschung nicht mehr. Eine erneute Abfrage warnt unmittelbar vor dem Löschaufruf vor möglicherweise besitzerlosen Objekten. Scheitert diese Abfrage, wird das Konto übersprungen. Entsperren setzt nur `AccountEnabled = true`. Das bestehende Passwort bleibt gültig. Weder ein Empfängertyp noch ein fehlendes Postfach belegt eine echte Person. Du musst die Zuordnung und vor einer Löschung die gewünschte Aufbewahrungsfrist selbst prüfen.

### Personal

```powershell
.\Sync-LehrerEntra.ps1 -File '.\Lehrer-2026.xlsx'
.\Sync-LehrerEntra.ps1 -File '.\Lehrer-2026.xlsx' -Update -CreateNewUsers -UpdateUsers -WhatIf
.\Sync-LehrerEntra.ps1 -Add -Vorname 'Lea' -Nachname 'Testwald' -Job 'CO-L' -WhatIf
.\Sync-LehrerEntra.ps1 -Remove -UPN 'ltestwald@monteaufkirchen.com' -WhatIf
.\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@monteaufkirchen.com' -WhatIf
```

`L`, `CO-L` und `PA` erhalten `EmployeeType = Pädagogisches Team` und die entsprechende Schulrolle. `OGTS` erhält `EmployeeType = Ganztag` und die Ganztagsrolle. Mehrere Jobs desselben Profils können kommasepariert angegeben werden. `JAS` ist nur ein zusätzlicher Job und erhält kein eigenes Entra-Profil. `JobTitle` enthält die Kürzel, etwa `L, PA` oder `PA, JAS`. Neue Personalkonten erhalten ein zufälliges Initialpasswort mit genau 12 Zeichen und müssen es bei der ersten Anmeldung ändern. Bestehende Lizenzen bleiben erhalten. `-Remove` sperrt das Konto und widerruft Sitzungen, ohne es zu löschen.

Ein vollständiger `-Update`-Lauf verarbeitet auch aktive Rollenmitglieder ohne Excel-Zeile als Abgänge. Bereits gesperrte Konten werden dabei übergangen. Die Exceldatei muss beide Personalrollen vollständig abdecken. Details und Grenzen stehen in der [Personaldokumentation](docs/ENTRA-LEHRER-SYNC.md).

Die Tabelle „Bestehende“ im Personalvergleich zeigt `NameMitRufname`, `UserId`, `DisplayName`, `UserPrincipalName` und `Job` in dieser Reihenfolge. Die Excel-Zeilennummer erscheint dort nicht.

### Schüler

Lege `Schueler.xlsx` im Repository-Root ab. Ohne Parameter zeigt das Skript nur den Vergleich und nimmt keine Änderungen vor:

```powershell
.\Sync-SchuelerEntra.ps1
```

Eine andere `.xlsx`-Datei kannst du mit `-File` angeben. Relative Pfade beziehen sich auf das aktuelle PowerShell-Arbeitsverzeichnis:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Schuelerimport\Schueler-2026.xlsx'
```

Mit `-OutputFile` speicherst du den vollständigen, passwortfreien Laufbericht zusätzlich als UTF-8-Textdatei. Das Ziel muss eine neue `.txt`-Datei sein. Git muss zur Pfadprüfung verfügbar sein. Vorhandene Dateien sowie symbolische Links im Ziel oder seinen Elternordnern werden abgewiesen. Innerhalb eines Git-Repositories muss das Ziel ignoriert sein. Der empfohlene Ordner `Berichte` und der Rootname `output.txt` sind dafür bereits ausgeschlossen. Unter `-WhatIf` wird auch kein Bericht geschrieben:

```powershell
.\Sync-SchuelerEntra.ps1 -OutputFile "./Berichte/Schueler-$(Get-Date -Format yyyyMMdd-HHmmss).txt"
```

Prüfe einen vollständigen Lauf zuerst mit `-WhatIf`:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Schuelerimport\Schueler-2026.xlsx' -Update -WhatIf
```

Führe danach alle geplanten Aktionen aus:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Schuelerimport\Schueler-2026.xlsx' -Update
```

Bei `-Update` beziehungsweise `-UpdateUsers` schreibt das Skript für alle eindeutig zugeordneten Bestandskonten die aktuelle `EntraObjectId`, den tatsächlich vorhandenen `UPN` und `Mail` nach Excel. Manuell vergebene UPNs bleiben dabei unverändert. Neue Schüler erhalten zusätzlich ihr Initialpasswort in der Spalte `Passwort`.

Sobald du einen Aktionsschalter angibst, werden nur die ausgewählten Graph-Aktionen ausgeführt. Erstellen und Aktualisieren ziehen die Exchange-Konfiguration für erfolgreich bearbeitete Benutzer automatisch nach:

```powershell
.\Sync-SchuelerEntra.ps1 -Update -CreateNewUsers -UpdateUsers
.\Sync-SchuelerEntra.ps1 -Update -DisableUsers -RevokeSessions
```

Eine ausstehende Exchange-Konfiguration kannst du später gezielt nachholen:

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test.schueler@monteaufkirchen.com'
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'test1@monteaufkirchen.com','test2@monteaufkirchen.com' -WhatIf
```

`-Mail` akzeptiert auch die Aliase `-UPN` und `-UserPrincipalName`.

Einen einzelnen Schüler kannst du ohne Excel-Datei anlegen oder aktualisieren:

```powershell
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin' -WhatIf
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin'
```

Ist der Schüler bereits eindeutig in der Schüler-Rollengruppe vorhanden, aktualisiert das Skript seine abweichenden Daten. Ein deaktiviertes Bestandskonto wird erst nach erfolgreicher Prüfung des Pflichtzustands aktiviert. Nur bei einer echten Neuanlage entsteht ein Passwort. Es wird genau einmal im Terminal angezeigt, nicht in Excel und nicht in `-OutputFile`. Bei einem Teilfehler enthält das Ergebnis die Objekt-ID und einen sicheren Wiederanlaufbefehl. Dieser `-Add -EntraObjectId`-Wiederanlauf funktioniert auch dann, wenn die Schüler-Rollengruppe vor dem Fehler noch nicht gesetzt werden konnte. Unmittelbar vor jeder Mutationsphase liest das Skript das Konto erneut aus Graph. Name, Firma und Mitarbeitertyp müssen weiterhin exakt zum angegebenen Schüler passen.

Einen einzelnen Schüler deaktivierst du per UPN. Dabei werden zusätzlich alle Sitzungen widerrufen. Der Benutzer wird nicht gelöscht und seine Gruppen oder Lizenzen bleiben erhalten:

```powershell
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com'
```

## Sicherheitsmodell

- Vergleich ist immer der Standard.
- `-WhatIf` führt keine Graph-, Excel- oder Exchange-Schreiboperation aus, schreibt keine Berichtsdatei und wartet nicht auf Postfächer.
- Ein Vorprüfungsfehler blockiert den gesamten schreibenden Lauf.
- Neue Konten bleiben deaktiviert, bis Attribute, Manager, Gruppen und Excel-Rückschreibung verifiziert sind.
- Bestehende Passwörter werden weder neu erzeugt noch geändert.
- Abgänge werden deaktiviert und optional von Sitzungen getrennt. Sie werden nicht gelöscht und behalten ihre Gruppen und Lizenzen.
- `legalAgeGroupClassification` wird nur gegen `MinorWithParentalConsent` geprüft. Das schreibgeschützte Graph-Feld wird nicht gesetzt.
- Berichte, Ergebnisobjekte und `-OutputFile` enthalten keine Passwörter. Nur `-Add` zeigt das Initialpasswort einer echten Neuanlage einmalig im Terminal an.

## Dokumentation

- [Vollständige Fach- und Schnittstellendokumentation](docs/ENTRA-SCHUELER-SYNC.md)
- [Lehrer- und OGTS-Abgleich](docs/ENTRA-LEHRER-SYNC.md)
- [Verwaltung gesperrter Konten](docs/GESPERRTE-KONTEN-VERWALTUNG.md)
- [Betrieb, Fehlerbehebung und Wiederanlauf](docs/BETRIEB.md)
- [Testplan für die Parallels-VM](docs/TESTING.md)
- [Abnahmestand](docs/ABNAHME.md)
- [Changelog](CHANGELOG.md)
- [Sicherheitsrichtlinie](docs/SECURITY.md)

Die automatisierte Suite läuft unter PowerShell 7 auf macOS und Windows. Vor produktiven Schreibläufen bleiben die kontrollierten Windows- und Mandantentests aus dem Testplan erforderlich.
