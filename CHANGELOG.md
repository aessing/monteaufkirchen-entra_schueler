# Changelog

Alle wesentlichen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.
Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung folgt [Semantic Versioning](https://semver.org/lang/de/).

## [0.2.0] - 2026-09-27

### Added

- Gemeinsamer lokaler Prüflauf sowie GitHub-Workflow für Syntax, PSScriptAnalyzer und Pester auf Windows und macOS.
- Getrennter, standardmäßig lesender Personalabgleich für L, CO-L, PA und OGTS mit `.xlsx`-Import, manuellem Add/Remove und Exchange-Reparaturmodus.
- Rollenabhängige Entra- und Exchange-Profile mit Jobkürzel in `JobTitle` und bestehenden Gruppenwerten in `EmployeeType`.
- Prüfung direkter und gruppenbasierter Lizenzzuweisungen. Bestehende Lizenzen und Lizenzgruppen werden auch bei Jobwechsel erhalten.
- Initialpasswörter mit genau zwölf zufälligen Zeichen und erzwungenem Erstwechsel ausschließlich für neue Personalkonten.
- Synthetische Personal-Fixture und fokussierte Tests für Personalregeln und Regressionsschutz.
- Kommaseparierte Mehrfachjobs innerhalb eines Profils. `JAS` ist als Zusatzjob ohne eigene Rolle oder Lizenzierung möglich.
- Lesender Check für gesperrte Entra-Konten mit Fortschrittsanzeige, SEC-A-Gruppen, letzter erfolgreicher Anmeldung und Schutzmarkierungen für Sonderkonten. Ohne belegtes Sperrdatum wird keine 90-Tage-Löschfreigabe ausgegeben.
- Tastaturauswahl für gesperrte Entra-Konten mit Pfeiltasten, Leertaste für persönliche Benutzerpostfächer und Konten ohne Exchange-Empfänger, Enter-Bestätigung aller markierten Namen und Esc zum Beenden. Nach Ja prüft das Skript jedes Konto erneut und löscht es mit Microsoft Graph. `-List` und `-PassThru` erhalten die lesende Ausgabe, `-WhatIf` simuliert den Löschpfad.
- Die TUI kann markierte Konten mit `U` nach eigener Namensbestätigung entsperren. Vor `AccountEnabled = true` werden Mandant, Identität und Schutzkriterien erneut geprüft. Das Passwort, Gruppen und Lizenzen bleiben unverändert, `-WhatIf` simuliert die Aktion.

### Changed

- Sichere Excel-Lese- und Rückschreibemechanik unterstützt zusätzlich einen getrennten Personalspaltenvertrag.
- Das Modul exportiert auch `Invoke-LehrerSync`. Der Schüleraufruf bleibt erhalten. Neue Schülerpasswörter verwenden zwölf zufällige, gut lesbare Buchstaben und Ziffern mit mehr als 64 Bit Auswahlraum. Bestehende Passwörter und `ForceChangePasswordNextSignIn = false` bleiben erhalten.
- Die lesende Prüfliste enthält alle gesperrten Entra-Benutzer. `SharedMailbox` wird in der TUI ausgeblendet. Eindeutig empfängerlose Konten sind dort markierbar, Ressourcenpostfächer und unklare Zuordnungen bleiben nur zur Prüfung sichtbar. Die Kopfzeile der lesenden Liste zählt Konten ohne persönliches Benutzerpostfach.
- Für Löschen und Entsperren braucht ein Konto ein persönliches Benutzerpostfach oder den Inventarstatus `None`. Vor einer Löschung müssen die Zuordnung zu einer realen Person und die gewünschte Sperrfrist manuell belegt werden.
- Die Prüfliste zeigt besessene Entra-Objekte mit Typ, Name und Objekt-ID. Das Ergebnisobjekt stellt sie je Konto unter `OwnedObjects` bereit.
- Besitzobjekte blockieren eine bestätigte Kontolöschung nicht mehr. Die Löschbestätigung nennt die Objekte und das Risiko fehlender Eigentümer. Unmittelbar vor dem Löschaufruf werden sie erneut gelesen und als Warnung gemeldet. Schlägt die Abfrage fehl, wird das Konto übersprungen.
- Der Einstieg für die interaktive Kontenverwaltung heißt `Verwalte-GesperrteEntraKonten.ps1`. Die Anleitung und alle Aufrufbeispiele verwenden den neuen Namen.
- Schüler- und Personalvergleich sowie die Kontenverwaltung trennen Ausgabebereiche durch Leerzeilen und zeigen farbige Konsolenüberschriften. Gespeicherte Vergleichsberichte bleiben reiner Text.

### Fixed

- Leere Schülerimporte werden abgewiesen und erzeugen keine Abgänge.
- Kombinierte Graph-/Exchange-Läufe prüfen denselben Mandanten. Die Kontenverwaltung prüft dies vor jeder Einzelaktion erneut, damit ein fremder Exchange-Bestand den SharedMailbox-Schutz nicht umgehen kann.
- Vergleichsberichte schreiben ausschließlich neue `.txt`-Dateien und prüfen den Git-Schutz. Arbeitsmappen, bestehende Dateien und symbolische Links, auch in Elternordnern, werden nicht überschrieben. `-WhatIf` schreibt keine Berichte.
- Lehrer-Neuanlagen sichern Passwort und Objekt-ID in Excel vor der weiteren Konfiguration. Der Wiederanlauf unterstützt diese gesperrten Konten. Bei fehlgeschlagener Sicherung werden manueller Wiederanlauf und erforderlicher Passwortreset ausdrücklich genannt.
- Schüler- und Lehrerupdates prüfen aktuelle Rollenmitgliedschaften. Personalrollen werden unmittelbar vor Änderungen auf Lizenzen und Änderbarkeit geprüft. Die konfigurierte Schülerrolle wird auch im Lehrerabgleich verwendet.
- Lehrer-Identitäten werden nur nach eigener Excel-Bestätigung nachgetragen. Abgelehnte Kontoänderungen sind davon ausgeschlossen.
- Manuelles Lehrer-Add zeigt keine unbeteiligten Personalmitglieder als Abgänge.

- Ein gesperrtes Entra-Konto ohne Exchange-Empfänger kann nach Bestätigung gelöscht oder entsperrt werden. Vor der Aktion liest das Skript den vollständigen Exchange-Bestand erneut. Ein neuer Empfänger, auch `SharedMailbox`, sowie Abfragefehler blockieren die Aktion.
- Löschen und Entsperren mehrerer ausgewählter Konten zeigen den Fortschritt pro Konto. Der Balken wird nach dem Lauf abgeschlossen, auch wenn einzelne Konten übersprungen werden oder fehlschlagen.
- Beim Entsperren blockiert Objektbesitz nicht. Vor einer Löschung wird Besitz weiterhin erneut geprüft und gemeldet.
- Ein aktiver Graph-Lizenzzustand mit `Error = None` blockiert die Personal-Vorprüfung nicht mehr. Tatsächliche Lizenzfehler und Fehlerstatus bleiben blockiert.
- Bereits gesperrte Personalrollenmitglieder ohne Excel-Zeile erscheinen nicht mehr als Abgänge. Eindeutig zugeordnete gesperrte Konten in Excel werden als Reaktivierung unter Änderungen geplant. Abgänge zeigen Quellrolle und Kontostatus.
- Die Schreibprüfung nennt eine gesperrte Lehrerarbeitsmappe nun „Lehrerdatei“. Gemeinsame Excel-Fehlertexte verwenden neutrale Bezeichnungen.
- Der Personalabgleich zeigt wie der Schülerabgleich Fortschritt für Verbindung, Einzelvergleich und Postfachbereitstellung. Bei `-Update` erscheint der Vergleich vor den Aktionen, während `-OutputFile` den vollständigen Bericht erhält.
- Die Lehrerarbeitsmappe muss für einen Update-Lauf nur dann exklusiv beschreibbar sein, wenn Neuanlagen oder Identitätsdaten zurückgeschrieben werden sollen.
- Bereits deaktivierte Schülerrollenmitglieder ohne Excel-Zeile erscheinen nicht mehr als Abgänge. Ein eindeutig zugeordnetes gesperrtes Schülerkonto in Excel wird als Reaktivierung unter Änderungen geplant und bei autorisiertem `-UpdateUsers` nach den Pflichtprüfungen aktiviert.
- Die Tabelle „Bestehende“ im Lehrerbericht zeigt `NameMitRufname`, `UserId`, `DisplayName`, `UserPrincipalName` und `Job` in dieser Reihenfolge wie die entsprechende Schülertabelle. `RowNumber` und lange technische Profilfelder entfallen aus der Tabellenansicht.
- Der Check liest Graph-`AdditionalProperties` auch aus generischen Dictionaries. Prüfungsfehler brechen den Skripteinstieg ab, ohne eine nachfolgende irreführende `Accounts`-Meldung auszugeben.
- Der Check lädt Exchange-Empfänger einmal als Bestand und ordnet sie per Entra-Objekt-ID zu. Gesperrte Konten ohne Exchange-Empfänger bleiben ohne 404-Warnung sichtbar. Bei einem Fehler der Bestandsabfrage bricht der Lauf ab.
- Schüler- und Personalabgleich fragen bei einem vollständigen `-Update` für konforme bestehende Postfächer nicht mehr nach einer Exchange-Änderung. Nur eine aktuell festgestellte Abweichung erreicht den bestätigungspflichtigen Konfigurationspfad.

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

[0.2.0]: https://github.com/aessing/monteaufkirchen-entra_schueler/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/aessing/monteaufkirchen-entra_schueler/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/aessing/monteaufkirchen-entra_schueler/releases/tag/v0.1.0
