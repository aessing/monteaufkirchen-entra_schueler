![3D-Comic-Illustration eines sicheren Schülerabgleichs zwischen Tabelle und Cloud-Verzeichnis](docs/assets/entra-schueler-sync-hero.png)

# Entra-Schülersynchronisation

Aktuelle Version: **0.1.1**

Dieses PowerShell-Tool vergleicht eine Excel-Schülerliste mit Microsoft Entra ID und Exchange Online. Der Standardlauf ist rein lesend. Änderungen benötigen ausdrücklich `-Update`, `-Add`, `-Remove` oder den getrennten Exchange-Reparaturmodus.

> [!IMPORTANT]
> Die Excel-Datei enthält personenbezogene Daten und nach Neuanlagen auch Initialpasswörter. Arbeitsmappen im Root und in Unterordnern sowie Sicherungen und temporäre Kopien sind per `.gitignore` ausgeschlossen. Nur die mitgelieferte synthetische Fixture ist ausgenommen. Vor der Rückschreibung prüft das Skript jeden konkreten Dateipfad. Committe keine produktiven Schülerlisten, Sicherungen oder Konsolenausgaben mit personenbezogenen Daten.

## Das erledigt das Tool

- Tabellen für Neuzugänge, Abgänge, Änderungen und bestehende Schüler
- Zuordnung ohne gelieferte UPN über eindeutigen Rufnamen und Nachnamen innerhalb der Schüler-Rollengruppe
- Feldgenaue Abweichungen für Entra, Manager, Gruppen und Exchange Online
- Eindeutige UPNs unter `@monteaufkirchen.com`, inklusive Umlautumschrift und Kollisionsprüfung
- Sichere Neuanlage mit deaktiviertem Konto, Pflichtzustand, geschützter Excel-Rückschreibung und erst anschließender Aktivierung
- Exakt 12 Zeichen lange, kindgerechte Initialpasswörter nur für neue Schüler, ohne erzwungenen Wechsel bei der ersten Anmeldung
- Genau eine Schüler-Rollengruppe und eine aktuelle Klassengruppe pro aktivem Schüler
- Automatischer Exchange-Abgleich mit sofortiger Prüfung und bis zu fünf Wiederholungen im Abstand von 60 Sekunden
- Separater Exchange-Reparaturlauf für einzelne oder mehrere UPNs
- Manuelles Anlegen oder Aktualisieren eines einzelnen Schülers mit `-Add`
- Sicheres Deaktivieren eines einzelnen Schülers und Widerrufen seiner Sitzungen mit `-Remove`

## Voraussetzungen

- Windows oder macOS mit PowerShell 7.0 oder neuer
- Interaktive Anmeldung bei Microsoft Graph und Exchange Online
- Berechtigtes Administratorkonto im richtigen Mandanten
- Die Module `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions`, `Microsoft.Graph.Groups`, `ExchangeOnlineManagement` und `ImportExcel`

```powershell
Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Users.Actions,Microsoft.Graph.Groups -Scope CurrentUser
Install-Module ExchangeOnlineManagement,ImportExcel -Scope CurrentUser
```

## Schnellstart

Lege `Schueler.xlsx` im Repository-Root ab. Ohne Parameter zeigt das Skript nur den Vergleich und nimmt keine Änderungen vor:

```powershell
.\Sync-SchuelerEntra.ps1
```

Eine andere `.xlsx`-Datei kannst du mit `-File` angeben. Relative Pfade beziehen sich auf das aktuelle PowerShell-Arbeitsverzeichnis:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Schuelerimport\Schueler-2026.xlsx'
```

Mit `-OutputFile` speicherst du den vollständigen, passwortfreien Laufbericht zusätzlich als UTF-8-Textdatei. Eine vorhandene Datei wird ersetzt. Der empfohlene Ordner `Berichte` und der häufig verwendete Rootname `output.txt` sind wegen der enthaltenen personenbezogenen Daten von Git ausgeschlossen:

```powershell
.\Sync-SchuelerEntra.ps1 -OutputFile '.\Berichte\Schueler-Abgleich.txt'
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
- `-WhatIf` führt keine Graph-, Excel- oder Exchange-Schreiboperation aus und wartet nicht auf Postfächer.
- Ein Vorprüfungsfehler blockiert den gesamten schreibenden Lauf.
- Neue Konten bleiben deaktiviert, bis Attribute, Manager, Gruppen und Excel-Rückschreibung verifiziert sind.
- Bestehende Passwörter werden weder neu erzeugt noch geändert.
- Abgänge werden deaktiviert und optional von Sitzungen getrennt. Sie werden nicht gelöscht und behalten ihre Gruppen und Lizenzen.
- `legalAgeGroupClassification` wird nur gegen `MinorWithParentalConsent` geprüft. Das schreibgeschützte Graph-Feld wird nicht gesetzt.
- Berichte, Ergebnisobjekte und `-OutputFile` enthalten keine Passwörter. Nur `-Add` zeigt das Initialpasswort einer echten Neuanlage einmalig im Terminal an.

## Dokumentation

- [Vollständige Fach- und Schnittstellendokumentation](docs/ENTRA-SCHUELER-SYNC.md)
- [Betrieb, Fehlerbehebung und Wiederanlauf](docs/BETRIEB.md)
- [Testplan für die Parallels-VM](docs/TESTING.md)
- [Abnahmestand](docs/ABNAHME.md)
- [Changelog](CHANGELOG.md)
- [Sicherheitsrichtlinie](docs/SECURITY.md)

Die automatisierte Suite läuft unter PowerShell 7 auf macOS und Windows. Vor produktiven Schreibläufen bleiben die kontrollierten Windows- und Mandantentests aus dem Testplan erforderlich.
