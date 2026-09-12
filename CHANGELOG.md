# Changelog

Alle wesentlichen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.
Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung folgt [Semantic Versioning](https://semver.org/lang/de/).

## [Unreleased]

## [0.1.1] - 2026-09-12

### Fixed

- Der ausdrückliche Wiederanlauf mit `-Add -EntraObjectId` erreicht ein teilweise angelegtes, noch nicht der Schüler-Rollengruppe zugeordnetes Konto wieder.
- Der Wiederanlauf liest das Ziel unmittelbar vor jeder Mutationsphase erneut aus Graph und prüft die exakte Objekt-ID, den normalisierten Vor- und Nachnamen, `CompanyName` und `EmployeeType`.
- Der normale Excel- und Namensabgleich blockiert unverändert alle Konten außerhalb der direkten Schüler-Rollengruppe.

## [0.1.0] - 2026-09-12

### Added

- Lesender Standardvergleich zwischen einer `.xlsx`-Schülerliste, Microsoft Entra ID und Exchange Online
- Tabellen für Neuzugänge, Abgänge, Änderungen, bestehende Schüler sowie Warnungen und Fehler
- Selektive und vollständige Update-Läufe mit `-Update`, `-CreateNewUsers`, `-UpdateUsers`, `-DisableUsers` und `-RevokeSessions`
- Manuelle Einzeloperationen mit `-Add` und `-Remove`
- Separater Exchange-Reparaturmodus mit `-ConfigureExchangeOnlineOnly`
- Frei wählbare Eingabedatei mit `-File` und passwortfreier Textbericht mit `-OutputFile`
- Fortschrittsanzeige für länger laufende Entra- und Exchange-Abfragen
- Kindgerechte Initialpasswörter mit exakt 12 Zeichen für neue Schüler
- Automatische Exchange-Online-Konfiguration mit sofortiger Prüfung und bis zu fünf Wiederholungen im Abstand von 60 Sekunden
- Vollständige Betriebs-, Test-, Sicherheits- und Schnittstellendokumentation

### Changed

- Bestehende Benutzer werden ohne gelieferten UPN über eindeutigen Rufnamen und Nachnamen innerhalb der Schüler-Rollengruppe zugeordnet
- Manuell vergebene UPNs bestehender Benutzer bleiben erhalten
- Gespeicherte Objekt-IDs und UPNs werden nur akzeptiert, wenn das Ziel direktes Mitglied der Schüler-Rollengruppe ist
- `EntraObjectId`, tatsächlicher `UPN` und `Mail` werden bei autorisierten Updates sicher nach Excel zurückgeschrieben
- Neue Benutzer werden zuerst deaktiviert angelegt und erst nach verifizierten Attributen, Gruppen, Manager und Excel-Rückschreibung aktiviert

### Security

- Der Standardlauf ist vollständig lesend, Schreibzugriffe benötigen einen ausdrücklichen Aktionsparameter
- `-WhatIf` verhindert Graph-, Excel- und Exchange-Schreiboperationen sowie Passworterzeugung und Wartezeiten
- Produktive XLSX-Dateien, Sicherungen, temporäre Dateien, der empfohlene Berichtsordner und `output.txt` sind von Git ausgeschlossen
- Passwörter werden nicht in Berichte, Ergebnisobjekte oder `-OutputFile` geschrieben
- Abgänge werden deaktiviert und Sitzungen widerrufen, Benutzerkonten, Gruppen und Lizenzen werden nicht gelöscht

[Unreleased]: https://github.com/aessing/monteaufkirchen-entra_schueler/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/aessing/monteaufkirchen-entra_schueler/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/aessing/monteaufkirchen-entra_schueler/releases/tag/v0.1.0
