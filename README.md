![3D-Comic-Illustration eines sicheren Schülerabgleichs zwischen Tabelle und Cloud-Verzeichnis](docs/assets/entra-schueler-sync-hero.png)

# Entra-Schülersynchronisation

Dieses PowerShell-Tool vergleicht eine Excel-Schülerliste mit Microsoft Entra ID und Exchange Online. Der Standardlauf ist rein lesend. Änderungen benötigen ausdrücklich `-Update` oder den getrennten Exchange-Reparaturmodus.

> [!IMPORTANT]
> Die Excel-Datei enthält personenbezogene Daten und nach Neuanlagen auch Initialpasswörter. Arbeitsmappen im Root und in Unterordnern sowie Sicherungen und temporäre Kopien sind per `.gitignore` ausgeschlossen. Nur die mitgelieferte synthetische Fixture ist ausgenommen. Vor der Rückschreibung prüft das Skript jeden konkreten Dateipfad. Committe keine produktiven Schülerlisten, Sicherungen oder Konsolenausgaben mit personenbezogenen Daten.

## Das erledigt das Tool

- Tabellen für Neuzugänge, Abgänge, Änderungen und bestehende Schüler
- Feldgenaue Abweichungen für Entra, Manager, Gruppen und Exchange Online
- Eindeutige UPNs unter `@monteaufkirchen.com`, inklusive Umlautumschrift und Kollisionsprüfung
- Sichere Neuanlage mit deaktiviertem Konto, Pflichtzustand, geschützter Excel-Rückschreibung und erst anschließender Aktivierung
- Exakt 12 Zeichen lange, kindgerechte Initialpasswörter nur für neue Schüler, ohne erzwungenen Wechsel bei der ersten Anmeldung
- Genau eine Schüler-Rollengruppe und eine aktuelle Klassengruppe pro aktivem Schüler
- Automatischer Exchange-Abgleich mit sofortiger Prüfung und bis zu fünf Wiederholungen im Abstand von 60 Sekunden
- Separater Exchange-Reparaturlauf für einzelne oder mehrere UPNs

## Voraussetzungen

- Windows mit PowerShell 7.0 oder neuer
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

Prüfe einen vollständigen Lauf zuerst mit `-WhatIf`:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Schuelerimport\Schueler-2026.xlsx' -Update -WhatIf
```

Führe danach alle geplanten Aktionen aus:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Schuelerimport\Schueler-2026.xlsx' -Update
```

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

## Sicherheitsmodell

- Vergleich ist immer der Standard.
- `-WhatIf` führt keine Graph-, Excel- oder Exchange-Schreiboperation aus und wartet nicht auf Postfächer.
- Ein Vorprüfungsfehler blockiert den gesamten schreibenden Lauf.
- Neue Konten bleiben deaktiviert, bis Attribute, Manager, Gruppen und Excel-Rückschreibung verifiziert sind.
- Bestehende Passwörter werden weder neu erzeugt noch geändert.
- Abgänge werden deaktiviert und optional von Sitzungen getrennt. Sie werden nicht gelöscht und behalten ihre Gruppen und Lizenzen.
- `legalAgeGroupClassification` wird nur gegen `MinorWithParentalConsent` geprüft. Das schreibgeschützte Graph-Feld wird nicht gesetzt.
- Die Ausgabe enthält keine Passwörter.

## Dokumentation

- [Vollständige Fach- und Schnittstellendokumentation](docs/ENTRA-SCHUELER-SYNC.md)
- [Betrieb, Fehlerbehebung und Wiederanlauf](docs/BETRIEB.md)
- [Testplan für die Parallels-VM](docs/TESTING.md)
- [Sicherheitsrichtlinie](docs/SECURITY.md)

Die PowerShell- und Live-Mandantentests müssen vor dem produktiven Einsatz in der vorgesehenen Windows-Parallels-VM abgeschlossen werden.
