# Entra-Schülersynchronisation – Designspezifikation

## Ziel

Das Repository erhält einen sicheren PowerShell-Workflow, der Schülerdaten aus einer frei wählbaren `.xlsx`-Datei mit Microsoft Entra ID und Exchange Online vergleicht. Ohne Update-Schalter arbeitet der Workflow ausschließlich lesend. Mit ausdrücklich gewählten Update-Parametern erstellt, aktualisiert oder deaktiviert er Schülerkonten, korrigiert direkte Gruppenmitgliedschaften, widerruft Sitzungen und gleicht die Exchange-Online-Konfiguration ab.

Der Workflow ist wiederholbar. Ein zweiter Lauf mit unveränderten Quelldaten soll keine weiteren Änderungen planen oder ausführen.

## Nicht-Ziele

- Keine Benutzer werden gelöscht.
- Abgänge verlieren nicht automatisch ihre Lizenz-, Rollen- oder Klassengruppen. Sie werden ausschließlich deaktiviert und bei gewählter Aktion von ihren Sitzungen getrennt.
- Das Skript verwaltet keine Mitarbeiter, Pädagogen, Sekretariatskonten oder sonstigen Benutzerattribute außerhalb der Schülerpopulation.
- Das Skript verarbeitet keine alten `.xls`-Dateien und keine CSV-Dateien.
- Das Skript speichert keine Zugangsdaten, Tokens oder Mandantengeheimnisse im Repository.
- Der reine Exchange-Reparaturmodus verändert keine Entra-Benutzerattribute oder Gruppenmitgliedschaften.

## Laufzeit und Abhängigkeiten

Zielumgebung ist PowerShell 7 unter Windows. Die fachliche Logik bleibt plattformunabhängig und wird mit Pester getestet. Die Live-Integration wird in einer Windows-Parallels-VM geprüft.

Benötigte PowerShell-Module:

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Users`
- `Microsoft.Graph.Users.Actions`
- `Microsoft.Graph.Groups`
- `ExchangeOnlineManagement`
- `ImportExcel`
- `Pester` für Tests
- `PSScriptAnalyzer` für statische Prüfung

Der Workflow verwendet ausschließlich Microsoft Graph v1.0 und Exchange-Online-Cmdlets. Beta-Endpunkte sind nicht vorgesehen.

Für einen vollständigen delegierten Lauf werden mindestens die Graph-Berechtigungen benötigt, die das Lesen und Ändern von Benutzern, Mailadressen und Gruppenmitgliedschaften sowie das Widerrufen von Sitzungen erlauben. Die Dokumentation nennt die konkret getesteten Scopes und die dazu erforderlichen Entra- und Exchange-Administratorrollen. Der angemeldete Administrator muss außerdem die von Microsoft geforderten Verzeichnisrollen besitzen.

## Repository-Struktur

```text
Sync-SchuelerEntra.ps1
config/
  SchuelerSync.psd1
src/
  SchuelerSync/
    SchuelerSync.psd1
    SchuelerSync.psm1
    Private/
      Comparison.ps1
      Entra.ps1
      Excel.ps1
      ExchangeOnline.ps1
      Normalization.ps1
      Reporting.ps1
      Update.ps1
tests/
  Comparison.Tests.ps1
  Normalization.Tests.ps1
  Password.Tests.ps1
  UpdateSelection.Tests.ps1
  fixtures/
    Schueler-Testdaten.xlsx
docs/
  ENTRA-SCHUELER-SYNC.md
  assets/
    entra-schueler-sync-hero.png
README.md
.gitignore
```

`Sync-SchuelerEntra.ps1` ist der einzige Einstiegspunkt. Das interne Modul kapselt Excel-, Graph-, Exchange-, Vergleichs- und Reporting-Funktionen. `config/SchuelerSync.psd1` enthält ausschließlich nicht geheime Konstanten wie Domain, Gruppen, Gruppen-IDs, Präfixe und Richtliniennamen.

## Kommandozeilenschnittstelle

### Vergleichs- und Update-Modus

```powershell
.\Sync-SchuelerEntra.ps1 [-File <Pfad>] [-Update]
    [-CreateNewUsers] [-DisableUsers] [-UpdateUsers] [-RevokeSessions]
    [-WhatIf] [-Confirm] [-Verbose]
```

Ohne `-Update` wird ausschließlich verglichen. Die Aktionsschalter ohne `-Update` führen zu einem Parameterfehler.

Der Standardvergleich liest neben Entra auch den vorhandenen Exchange-Online-Zustand der zugeordneten Schüler. Er prüft die Postfachverfügbarkeit einmal und wartet nicht auf noch nicht bereitgestellte Postfächer. Ein fehlendes Postfach wird als `EXO-Konfiguration ausstehend` gemeldet. Es findet keine Exchange-Schreiboperation statt.

`-Update` ohne Aktionsschalter aktiviert den vollständigen Ablauf:

1. Neuzugänge erstellen
2. Änderungen an bestehenden Schülern anwenden
3. Abgänge deaktivieren
4. Sitzungen der Abgänge widerrufen
5. Gruppen der aktiven Schüler normalisieren
6. Exchange Online für alle aktiven Excel-Schüler abgleichen

Sobald mindestens ein Aktionsschalter angegeben wird, werden nur die gewählten Graph-Aktionen ausgeführt:

- `-CreateNewUsers` erstellt Neuzugänge, setzt ihre Pflichtgruppen und konfiguriert anschließend automatisch ihr Exchange-Postfach.
- `-UpdateUsers` aktualisiert geänderte aktive Schüler, normalisiert ihre Pflichtgruppen und konfiguriert anschließend automatisch ihr Exchange-Postfach.
- `-DisableUsers` deaktiviert Abgänge.
- `-RevokeSessions` widerruft die Sitzungen der Abgänge.

`-DisableUsers` und `-RevokeSessions` allein starten keine EXO-Konfiguration für aktive Schüler.

Das Skript ist ein Advanced Script mit `SupportsShouldProcess`. `-WhatIf` führt keine Graph-, Exchange- oder Excel-Schreiboperation aus. `-Confirm` und `-Verbose` verwenden das übliche PowerShell-Verhalten.

### Frei wählbare Excel-Datei

`-File` akzeptiert einen relativen oder absoluten Pfad zu einer `.xlsx`-Datei:

```powershell
.\Sync-SchuelerEntra.ps1 -File "C:\Import\Schueler-2026.xlsx"
```

Ohne `-File` wird `<Repository>\Schueler.xlsx` verwendet. Relative Pfade werden relativ zum aktuellen PowerShell-Arbeitsverzeichnis aufgelöst. Das Skript prüft Dateiexistenz, Erweiterung, Lesbarkeit und vor einer Rückschreibung auch die Schreibbarkeit beziehungsweise eine bestehende Dateisperre.

### Reiner Exchange-Reparaturmodus

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail <UPN[,UPN...]>
    [-WhatIf] [-Confirm] [-Verbose]
```

`-Mail` ist ein String-Array und besitzt die Aliase `-UPN` und `-UserPrincipalName`. Dieser Parameter-Satz ist mit `-File`, `-Update` und allen Graph-Aktionsschaltern unvereinbar. Er liest keine Excel-Datei, verändert keine Entra-Attribute und verändert keine Gruppenmitgliedschaften. Er gleicht ausschließlich die vollständige Schüler-Postfachkonfiguration für die angegebenen Identitäten ab.

## Excel-Vertrag

Das Skript verwendet das erste Arbeitsblatt, das alle Pflichtspalten besitzt. Gibt es kein oder mehr als ein passendes Arbeitsblatt, bricht die Vorprüfung mit einer eindeutigen Meldung ab.

Pflichtspalten:

- `Name mit Rufname`
- `Vorname`
- `Nachname`
- `Klassen`
- `Klassenlehrer`

Vom Skript verwaltete Zusatzspalten:

- `Passwort`
- `EntraObjectId`
- `UPN`
- `Mail`

Fehlende Zusatzspalten werden erst bei einem autorisierten Update am rechten Tabellenende ergänzt. Im Vergleichsmodus wird die Arbeitsmappe nie verändert.

`Name mit Rufname` dient als benutzerfreundliche Bezeichnung in Ausgaben. Entra-Vorname, Entra-Nachname und DisplayName stammen ausschließlich aus `Vorname` und `Nachname`.

Leere Zeilen werden ignoriert. Fehlende Pflichtwerte, doppelte Objekt-IDs, doppelte gespeicherte UPNs und mehrdeutige Schüleridentitäten sind Fehler der Vorprüfung. Bei solchen Fehlern findet keine Mutation statt.

Vor jeder Excel-Rückschreibung wird neben der Quelldatei eine Sicherung nach dem Muster `<Basisname>.backup-YYYYMMDD-HHmmss.xlsx` erstellt. Änderungen werden zuerst in eine temporäre Datei geschrieben, erneut geöffnet und geprüft und danach atomar an die Stelle der Quelldatei verschoben. Scheitert die Rückschreibung nach einer Entra-Neuanlage, bleibt das neue Konto deaktiviert und der Fehlerbericht enthält einen konkreten Wiederherstellungsschritt.

## Git-Schutz für personenbezogene Daten

Die Root-`.gitignore` enthält:

```gitignore
/*.xlsx
```

Damit sind `Schueler.xlsx`, frei benannte Root-Arbeitsmappen und Sicherungen ignoriert. Synthetische Testdateien in `tests/fixtures/` bleiben versionierbar.

Vor einer Excel-Schreiboperation prüft das Skript, sofern Git verfügbar ist, ob eine innerhalb eines Git-Repositories liegende Quelldatei ignoriert wird. Eine verfolgte oder nicht ignorierte Datei innerhalb des Repositories blockiert die Rückschreibung, damit Passwörter nicht versehentlich committed werden. Dateien außerhalb eines Git-Repositories sind davon nicht betroffen.

## Entra-Zielattribute

| Ziel | Sollwert |
|---|---|
| `givenName` | Excel `Vorname`, unverändert |
| `surname` | Excel `Nachname`, unverändert |
| `displayName` | `<Vorname> <Nachname>` |
| `userPrincipalName` | Generierter UPN |
| `mail` | Identisch zum UPN |
| `mailNickname` | Lokaler Teil des UPN |
| `department` | Excel `Klassen`, unverändert |
| `officeLocation` | Aus Excel `Klassen` abgeleitet |
| `companyName` | `Montessori Schule Aufkirchen` |
| `employeeType` | `Schüler` |
| `ageGroup` | `Minor` |
| `consentProvidedForMinor` | `Granted` |
| `usageLocation` | `DE` |
| Manager | Eindeutig aufgelöster Excel-`Klassenlehrer` |

`legalAgeGroupClassification` ist in Microsoft Graph schreibgeschützt. Das Skript erwartet nach dem Setzen von `ageGroup = Minor` und `consentProvidedForMinor = Granted` den berechneten Wert `MinorWithParentalConsent` und meldet eine Abweichung. Es versucht nicht, die Eigenschaft direkt zu schreiben.

## Identitätszuordnung

Eine Excel-Zeile wird in dieser Reihenfolge zugeordnet:

1. Gültige `EntraObjectId`, wenn vorhanden und das Ziel direktes Mitglied der Schüler-Rollengruppe ist
2. Gespeicherter `UPN`, wenn vorhanden und der eindeutige Treffer direktes Mitglied der Schüler-Rollengruppe ist
3. Eindeutige, normalisierte Kombination aus `givenName` und `surname` innerhalb der Mitglieder der Schüler-Rollengruppe

Vor einer Neuanlage prüft das Skript zusätzlich das gesamte Entra-Benutzerverzeichnis. Ein möglicher Treffer außerhalb der Schüler-Rollengruppe wird als Konflikt gemeldet und nicht automatisch zum Schüler umgewidmet. Mehrere mögliche Treffer blockieren alle Mutationen und insbesondere die Deaktivierung vermeintlicher Abgänge.

Die einzige Ausnahme ist ein ausdrücklicher manueller Wiederanlauf mit `-Add -EntraObjectId`. Er darf das exakt bezeichnete, teilweise angelegte Konto auch dann auflösen, wenn die Schüler-Rollengruppe noch fehlt. Unmittelbar vor jeder Mutationsphase wird das Konto erneut aus Graph gelesen. Normalisierter Vor- und Nachname, `CompanyName` und `EmployeeType` müssen weiterhin zum angegebenen Schüler passen. Excel-Abgleich, gespeicherter UPN und Namenssuche erhalten diese Ausnahme nicht.

Eine gespeicherte Objekt-ID ist die stabile Identität. Dadurch bleibt ein Schüler bei einer späteren Änderung des Vor- oder Nachnamens zuordenbar.

## UPN- und Mail-Algorithmus

Domain ist `monteaufkirchen.com`.

Für den lokalen Teil werden Vor- und Nachname wie folgt normalisiert:

1. Kleinschreibung mit invarianter Kultur
2. `ä → ae`, `ö → oe`, `ü → ue`, `ß → ss`, einschließlich Großbuchstaben
3. Zerlegung und Entfernung weiterer diakritischer Zeichen
4. Entfernung von Leerzeichen, Apostrophen und allen Zeichen außerhalb `a-z` und `0-9`

DisplayName, Vorname und Nachname behalten ihre originale Schreibweise.

UPN-Kandidaten werden in dieser Reihenfolge gebildet:

1. Erster Buchstabe des normalisierten Vornamens plus normalisierter Nachname
2. Erste zwei Buchstaben des Vornamens plus Nachname
3. Fortsetzung bis zum vollständigen Vornamen
4. Wenn alle Varianten belegt sind, vollständiger Vorname plus Nachname plus aufsteigende Zahl ab `2`

Die Belegung wird gegen alle Entra-UPNs, Entra-Mailadressen, vorhandene Proxyadressen und Exchange-Empfänger sowie gegen bereits für denselben Lauf reservierte Adressen geprüft. Eine Adresse, die demselben Entra-Objekt gehört, gilt nicht als Kollision.

Der UPN eines bestehenden Schülers bleibt immer unverändert, auch bei einer Namensänderung oder wenn eine früher verursachende Kollision nicht mehr existiert. Der Algorithmus wählt nur für Neuzugänge einen neuen UPN. Jede Abweichung vom ersten Kandidaten wird mit Ursache und gewähltem UPN in der Warnungstabelle ausgegeben.

## Klassen- und Office-Location-Regel

`department` entspricht exakt dem Wert aus `Klassen`.

Die Office Location wird case-insensitiv aus dem Token direkt vor dem Unterstrich oder dem Klassenende ermittelt. Erlaubte Ergebnisse:

- `G1`, `G2`, `G3`, `G4`
- `M1`, `M2`, `M3`
- `O1`, `O2`
- `A1`, `A2`

Beispiele:

```text
JK1-3g2_1  -> G2
JK4-6m2_4  -> M2
```

Eine leere, mehrdeutige oder nicht erlaubte Office Location ist ein Vorprüfungsfehler.

## Manager-Regel

Entra unterstützt genau einen Manager. Der vollständige Inhalt von `Klassenlehrer` wird als eine Identität behandelt.

Auflösung:

1. Wenn der Wert wie eine Mailadresse aussieht, exakter Vergleich mit UPN oder Mail
2. Andernfalls case-insensitiver exakter Vergleich mit `displayName`
3. Genau ein Treffer ist erforderlich

Kein Treffer und mehrere Treffer sind Vorprüfungsfehler. Das Skript meldet die betroffene Excel-Zeile und den Suchwert. Ein Manager wird mit `Set-MgUserManagerByRef` gesetzt und anschließend erneut ausgelesen.

## Pflichtgruppen und Exklusivität

Jeder aktive Excel-Schüler muss direktes Mitglied folgender Gruppen sein:

- `SEC-A-LIC-O365A1Student`
- `SEC-A-ROL-Schule_Schüler`
- `SEC-A-CLS-<Klassenwert>`

Die Schüler-Rollengruppe besitzt die bekannte ID `cebc1326-1174-4126-ba84-7a8960850e0a`. Das Skript prüft, ob ID und Anzeigename zusammenpassen. Lizenz- und Klassengruppen werden über einen exakten Anzeigenamen aufgelöst und müssen eindeutig sein.

Der bekannte Rollengruppenkatalog aus dem bisherigen manuellen Skript wird als nicht geheime Konfiguration übernommen:

| Gruppe | Objekt-ID |
|---|---|
| `SEC-A-ROL-Schule_PädagogischesTeam` | `103a1c4c-036b-40bc-bbe8-e3f7a866cb01` |
| `SEC-A-ROL-Schule_Schüler` | `cebc1326-1174-4126-ba84-7a8960850e0a` |
| `SEC-A-ROL-Schule_Sekretariat` | `534f94ed-420d-4794-a374-20c572def8cb` |
| `SEC-A-ROL-Schule_Vertretungskräfte` | `4ed4f87b-374f-4679-ac11-5e51212e9a91` |
| `SEC-A-ROL-Kinderhaus_PädagogischesTeam` | `cfc6d358-875d-441b-b45a-0d9734b7d788` |
| `SEC-A-ROL-Kinderhaus_Sekretariat` | `43165f2c-3ce5-471e-a7be-e16b66c03406` |
| `SEC-A-ROL-Kinderhaus_Vertretungskräfte` | `d2e7e424-83d2-481d-9226-bf529bc9640a` |
| `SEC-A-ROL-Ganztag` | `3997db35-2914-4b0f-92b4-08038ef8f84a` |
| `SEC-A-ROL-Unterstützung` | `a4a00442-bc04-48ee-b659-cb3bde194630` |
| `SEC-A-ROL-ExterneBenutzer` | `9c620843-0071-4d44-ba23-318e19cd09a0` |
| `SEC-A-ROL-ExterneAdmins` | `a246dfde-eccf-489b-ac1c-8d92f67145ca` |

Der Katalog dient der Validierung und der verständlichen Ausgabe. Zusätzlich werden direkte Mitgliedschaften in später hinzugekommenen Gruppen mit dem Präfix `SEC-A-ROL-` erkannt. Für aktive Schüler werden alle direkten Rollengruppen außer `SEC-A-ROL-Schule_Schüler` als zu entfernende Abweichung behandelt.

Ein aktiver Schüler darf direkt nur in der Schüler-Rollengruppe mit Präfix `SEC-A-ROL-` und nur in seiner aktuellen Klassengruppe mit Präfix `SEC-A-CLS-` sein. Der mitgelieferte Rollengruppenkatalog dient als zusätzliche Plausibilitätsprüfung. Attribute der Benutzer in den anderen Rollengruppen werden nicht verändert.

Reihenfolge der Gruppenänderungen:

1. Ziel-Rollengruppe und Ziel-Klassengruppe eindeutig auflösen
2. Fehlende Zielgruppen hinzufügen
3. Erfolg erneut lesen
4. Erst danach andere direkte Rollen- beziehungsweise Klassengruppen entfernen
5. Endzustand erneut lesen

Dynamische und indirekte Mitgliedschaften werden nicht entfernt. Sie erscheinen als nicht automatisch korrigierbare Abweichung. Eine fehlende Zielgruppe führt niemals zur Entfernung einer bisherigen Gruppe.

## Vergleichskategorien

Der Vergleich erzeugt getrennte Tabellen:

### Neuzugänge

Excel-Zeilen ohne eindeutig zugeordnetes Entra-Konto.

### Abgänge

Direkte Mitglieder von `SEC-A-ROL-Schule_Schüler`, die keiner Excel-Zeile eindeutig zugeordnet sind.

### Änderungen

Zugeordnete Schüler mit mindestens einer Abweichung in:

- Entra-Zielattributen
- UPN oder Mail
- Manager
- Pflichtgruppen
- unerlaubten direkten Rollen- oder Klassengruppen
- Exchange-Online-Sollwerten, sofern ein Postfach existiert

Jede Abweichung zeigt Feld, Istwert und Sollwert. Gruppenänderungen zeigen hinzuzufügende und zu entfernende Gruppen getrennt.

### Bestehende

Zugeordnete Schüler ohne festgestellte Abweichung.

### Warnungen und Fehler

Mindestens folgende Sachverhalte werden gemeldet:

- UPN- oder Mailkollision und verwendeter Alternativwert
- fehlender oder mehrdeutiger Klassenlehrer
- fehlende oder mehrdeutige Zielgruppe
- ungültige Klasse oder Office Location
- doppelte Excel-Identität
- Entra-Treffer außerhalb der Schüler-Rollengruppe
- dynamische oder indirekte Gruppenabweichung
- nicht vorhandenes oder noch nicht bereitgestelltes Postfach
- fehlende Berechtigung oder nicht unterstützter Exchange-Parameter

Passwörter werden in keiner Tabelle ausgegeben.

## Update-Ablauf

Vor jeder Mutation wird der vollständige Datenbestand gelesen und validiert. Ein Fehler der Vorprüfung verhindert den gesamten schreibenden Lauf. UPN-Kollisionen mit eindeutig gewähltem Alternativwert sind Warnungen und blockieren den Lauf nicht.

### Neuzugänge

Ein neuer Schüler wird fail-safe angelegt:

1. Exakt zwölf Zeichen langes Passwort generieren
2. Benutzer zunächst mit `accountEnabled = false` erstellen
3. `passwordProfile.forceChangePasswordNextSignIn = false` setzen
4. Entra-Sollattribute setzen
5. Manager setzen und prüfen
6. Lizenz-, Rollen- und Klassengruppe hinzufügen und prüfen
7. `Passwort`, `EntraObjectId`, tatsächlichen `UPN` und `Mail` sicher in Excel zurückschreiben
8. Konto aktivieren und Aktivierung prüfen
9. Postfachverfügbarkeit gesammelt abwarten und EXO konfigurieren

Scheitert ein Pflichtschritt vor der Aktivierung, bleibt das Konto deaktiviert. Das Skript löscht den Benutzer nicht automatisch.

### Abgänge

Bei `-DisableUsers` wird `accountEnabled = false` gesetzt und erneut gelesen. Bei `-RevokeSessions` wird `Revoke-MgUserSignInSession` aufgerufen und das API-Ergebnis geprüft. Eine fehlgeschlagene Sitzungswiderrufung wird nicht als erfolgreiche Sperrung ausgegeben.

### Änderungen

Bei `-UpdateUsers` werden ausschließlich abweichende Eigenschaften geschrieben. UPN, Mail, Manager und Gruppen werden jeweils nach dem Schreiben erneut gelesen. Danach schreibt das Skript für alle eindeutig zugeordneten Bestandskonten die aktuelle `EntraObjectId`, den tatsächlich vorhandenen `UPN` und `Mail` nach Excel. Bestehende Passwörter und die Excel-Passwortspalte bleiben unverändert.

### Laufzeitfehler

Nach erfolgreicher Vorprüfung werden Benutzeraktionen einzeln mit Fehlerbehandlung ausgeführt. Ein Fehler bei einem Benutzer stoppt nicht automatisch alle bereits unabhängigen Benutzer, wird aber als fehlgeschlagene Aktion mit Objekt-ID und Wiederanlaufhinweis ausgegeben. Es gibt keinen automatischen tenantweiten Rollback, weil Graph- und Exchange-Operationen nicht zuverlässig transaktional rückgängig gemacht werden können.

## Passwortgenerator

Jedes neue Passwort ist exakt zwölf Zeichen lang.

Regeln:

- Zwei oder drei kindgerechte deutsche Wörter in CamelCase
- Wörter ergeben zusammen exakt zehn ASCII-Buchstaben
- Zwei kryptografisch zufällige Ziffern von `10` bis `99`
- Mindestens ein Großbuchstabe, Kleinbuchstaben und Ziffern
- Keine Umlaute oder Sonderzeichen
- Keine doppelte Ausgabe innerhalb eines Laufs
- Zufall ausschließlich über `System.Security.Cryptography.RandomNumberGenerator`
- Keine Konsolenausgabe, keine Transcript-Ausgabe und keine Speicherung außerhalb der gewählten Excel-Datei und ihrer Sicherung

Beispiele für das Format sind `TigerWiese56` und `LegoGarten24`. Die tatsächliche Wortauswahl erfolgt zufällig aus kuratierten Listen. Lehnt Entra ein Passwort aufgrund der Mandantenrichtlinie ab, wird bis zu fünfmal ein neues Passwort erzeugt. Andere API-Fehler werden nicht als Passwortfehler behandelt.

## Exchange-Online-Sollzustand

Der EXO-Abgleich gilt für Schüler und verwendet folgende Werte.

### Mailbox

- `CustomAttribute1 = Montessori Schule Aufkirchen - Schüler`
- `AddressBookPolicy = MON-EXO-ABP-Schule_Schüler`
- `AuditEnabled = true`
- `AuditLogAgeLimit = 365 Tage`
- `RetainDeletedItemsFor = 30 Tage`
- `RoleAssignmentPolicy = MON-EXO-UserRoles-Default`
- `SharingPolicy = MON-EXO-Sharing-Default`
- `RetentionPolicy = MON-EXO-Retention-Default`

### AuditDelegate

Folgende Aktionen müssen mindestens vorhanden sein. Weitere vorhandene Aktionen werden nicht entfernt:

- `Create`
- `FolderBind`
- `HardDelete`
- `Move`
- `MoveToDeletedItems`
- `SendAs`
- `SendOnBehalf`
- `SoftDelete`
- `Update`
- `UpdateFolderPermissions`
- `UpdateInboxRules`

### AuditOwner

Folgende Aktionen müssen mindestens vorhanden sein. Weitere vorhandene Aktionen werden nicht entfernt:

- `Create`
- `HardDelete`
- `Move`
- `MailboxLogin`
- `MoveToDeletedItems`
- `SoftDelete`
- `Update`
- `UpdateFolderPermissions`
- `UpdateInboxRules`
- `UpdateCalendarDelegation`

### AuditAdmin

Ohne `CommunicationsCompliance` müssen mindestens vorhanden sein:

- `Copy`
- `Create`
- `FolderBind`
- `HardDelete`
- `Move`
- `MoveToDeletedItems`
- `SendAs`
- `SendOnBehalf`
- `SoftDelete`
- `Update`
- `UpdateFolderPermissions`
- `UpdateInboxRules`
- `UpdateCalendarDelegation`

Wenn `PersistedCapabilities` den Wert `CommunicationsCompliance` enthält, muss zusätzlich `MailItemsAccessed` vorhanden sein. Weitere vorhandene Aktionen werden nicht entfernt.

### CAS-Mailbox

- `ActiveSyncEnabled = false`
- `ImapEnabled = false`
- `MAPIEnabled = true`
- `OWAEnabled = true`
- `OWAforDevicesEnabled = false`
- `OwaMailboxPolicy = MON-EXO-OWA-Default`
- `PopEnabled = false`
- `SmtpClientAuthenticationDisabled = true`

Alle Eigenschaften werden zuerst gelesen. Nur abweichende Werte oder fehlende Audit-Aktionen werden geschrieben. Anschließend wird der Zustand erneut gelesen. Ein vom aktuellen Exchange-Online-Modul nicht unterstützter Parameter wird mit dem exakten Cmdlet-Fehler gemeldet und nicht stillschweigend übersprungen oder ersetzt.

## EXO-Bereitstellungswartezeit

Bei einem autorisierten vollständigen Update und bei `-ConfigureExchangeOnlineOnly` wird zuerst sofort nach den betroffenen Postfächern gesucht. Fehlende Postfächer werden anschließend bis zu fünf weitere Male mit jeweils 60 Sekunden Abstand geprüft. Damit gibt es höchstens sechs Prüfungen und fünf Minuten zusätzliche Wartezeit.

Bei mehreren Benutzern arbeitet die Wiederholung gesammelt:

1. Alle betroffenen UPNs sammeln
2. In jeder Runde nur noch fehlende Postfächer prüfen
3. Gefundene Postfächer sofort konfigurieren und aus der Restmenge entfernen
4. Nur einmal pro Runde 60 Sekunden warten

`-WhatIf` führt nur die erste Verfügbarkeitsprüfung aus und wartet nicht, da keine Konfiguration geschrieben werden darf.

Ist ein Postfach nach dem letzten Versuch nicht vorhanden, wird die Graph-Aktion nicht rückgängig gemacht. Der Benutzer erhält den Status `EXO-Konfiguration ausstehend`. Die Ausgabe enthält einen fertigen Befehl für `-ConfigureExchangeOnlineOnly -Mail <UPN>`.

Die Wiederholung gilt ausschließlich für noch nicht bereitgestellte Postfächer. Fehler von `Set-Mailbox` oder `Set-CASMailbox` werden sofort gemeldet und nicht blind wiederholt.

## Sicherheit und Datenschutz

- Vergleich ist der Standardmodus.
- Alle Schreibaktionen erfordern `-Update` oder den expliziten Parameter-Satz `-ConfigureExchangeOnlineOnly`.
- `ShouldProcess`, `-WhatIf` und `-Confirm` sichern jede Mutation zusätzlich ab.
- Passwörter erscheinen ausschließlich in der ausgewählten Excel-Datei und ihrer lokalen Sicherung.
- Fehlerobjekte und Tabellen enthalten keine Passwörter.
- Das Skript aktiviert neue Benutzer erst nach erfolgreicher Pflichtkonfiguration und Excel-Rückschreibung.
- Vor dem Entfernen alter Gruppen muss die neue Zielgruppe nachweislich hinzugefügt sein.
- Mehrdeutige Identitäten blockieren alle Mutationen und damit auch potenziell falsche Abgänge.
- Das Skript verwendet `Set-StrictMode` und `ErrorAction Stop` innerhalb schreibender Operationen.
- Das Skript prüft die Mandanten-ID nach der Anmeldung gegen eine optional konfigurierte erwartete Tenant-ID, sobald diese eingetragen wurde. Ohne konfigurierte Tenant-ID zeigt es die verbundene Organisation deutlich vor dem Update an.

## Tests

Alle Repository-Tests verwenden synthetische Personen und gemockte Microsoft-Dienste.

### Pester

- Umlaut- und Unicode-Normalisierung
- Entfernung ungeeigneter UPN-Zeichen
- UPN-Kollisionen über alle Vornamenspräfixe
- Numerischer Fallback nach vollständigem Vornamen
- Kollisionen innerhalb eines Excel-Laufs
- Stabilität eines bereits vergebenen UPN
- Klassenerkennung für alle erlaubten Office Locations
- Ablehnung ungültiger und mehrdeutiger Klassen
- Manager-Auflösung über Anzeigename, Mail und UPN
- Fehlender und mehrdeutiger Manager
- Identitätszuordnung über Objekt-ID, UPN und Namen
- Globaler Mutationsstopp bei mehrdeutiger Identität
- Feldgenaue Änderungslisten
- Pflichtgruppen und sichere Reihenfolge der Gruppenänderungen
- Dynamische und indirekte Gruppenabweichungen
- Exakt zwölf Zeichen lange Passwörter
- Keine Passworterzeugung oder Änderung für bestehende Schüler
- Selektive Aktionsschalter und unzulässige Parameterkombinationen
- Idempotenter zweiter Lauf
- Gesammelte EXO-Wiederholungen ohne benutzerweise Wartezeit
- EXO-AuditAdmin mit und ohne `CommunicationsCompliance`
- `-WhatIf` ohne Schreiboperation und ohne EXO-Wartezeit

### Statische Prüfung

- `PSScriptAnalyzer` für Skript und Modul
- PowerShell-Parserprüfung für alle `.ps1`, `.psm1` und `.psd1`
- Suche nach versehentlich eingecheckten Passwörtern, realen Schülerdaten und Excel-Quelldateien

### Windows-Parallels-VM

1. Module installieren und Versionen dokumentieren
2. Pester und PSScriptAnalyzer ausführen
3. Standardvergleich gegen den Mandanten ausführen
4. Vollständigen Lauf mit `-Update -WhatIf` prüfen
5. Einen synthetischen Testschüler zunächst deaktiviert erstellen
6. Attribute, Manager und Gruppen prüfen
7. Excel-Rückschreibung und exakte Passwortlänge prüfen
8. Anmeldung ohne erzwungenen Passwortwechsel prüfen
9. Klassenwechsel und Entfernung der vorherigen Klassengruppe prüfen
10. Zusätzliche Rollengruppe erkennen und entfernen
11. Abgang deaktivieren und Sitzungen widerrufen
12. Sofortige und verzögerte EXO-Bereitstellung prüfen
13. `-ConfigureExchangeOnlineOnly -Mail <UPN>` separat prüfen
14. Alle Mailbox-, Audit-, Richtlinien- und CAS-Werte nachlesen

Die VM-Prüfung verwendet keine produktiven Schülerdaten in Testfixtures. Live-Änderungen erfolgen nur durch den Benutzer mit einem dafür vorgesehenen Testkonto.

## Dokumentation und README

`README.md` wird vollständig auf das Projekt zugeschnitten. Sie enthält:

- ein breites 3D-Comic-Hero-Bild
- Zweck und Sicherheitsmodell
- Voraussetzungen
- Schnellstart für Vergleich, vollständiges Update, Teilaktionen und EXO-Reparatur
- Link zur vollständigen Dokumentation
- deutlichen Hinweis auf die vertrauliche Excel-Datei

Das Hero-Bild wird generisch gestaltet. Es zeigt eine freundliche, stilisierte Schulverwaltungs-Szene mit Laptop, Tabellenkarten, Identitäts-Cloud und Schulmaterialien. Es enthält keine realen Kinder, keine Namen, keine lesbaren Schülerdaten und keine fremden Markenlogos. Ablageort ist `docs/assets/entra-schueler-sync-hero.png`.

`docs/ENTRA-SCHUELER-SYNC.md` beschreibt:

- Installation und Modulabhängigkeiten
- benötigte Berechtigungen und Administratorrollen
- Excel-Schema und `-File`
- alle Parameter-Sätze mit Beispielen
- Matching-, UPN-, Klassen-, Manager- und Gruppenregeln
- Vergleichstabellen
- Update- und Wiederanlaufverhalten
- vollständige EXO-Sollwerte
- Datenschutz und Passwortbehandlung
- Parallels-VM-Testanleitung
- Fehlerdiagnose und manuelle Wiederherstellung

## Verifikation und Abnahme

Die Implementierung gilt erst als abgeschlossen, wenn:

- alle lokal ausführbaren statischen Prüfungen erfolgreich sind
- alle Pester-Tests in der Windows-Parallels-VM erfolgreich sind
- der Standardmodus nachweislich keine Datei und keinen Mandantenwert verändert
- `-WhatIf` keine Schreiboperation ausführt
- ein synthetischer Neuzugang korrekt erstellt und in Excel dokumentiert wurde
- ein synthetischer Klassenwechsel korrekt normalisiert wurde
- ein synthetischer Abgang deaktiviert und seine Sitzung widerrufen wurde
- der EXO-Abgleich mit sofortigem und verzögertem Postfach getestet wurde
- der reine EXO-Reparaturmodus verifiziert wurde
- keine produktiven Schülerdaten oder Passwörter in Git enthalten sind
- README und Benutzerdokumentation mit den implementierten Parametern übereinstimmen

## Primärquellen

- Microsoft Graph user resource: <https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0>
- Microsoft Graph create user: <https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0>
- Microsoft Graph update user: <https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0>
- Microsoft Graph password profile: <https://learn.microsoft.com/en-us/graph/api/resources/passwordprofile?view=graph-rest-1.0>
- Microsoft Graph assign manager: <https://learn.microsoft.com/en-us/graph/api/user-post-manager?view=graph-rest-1.0>
- Microsoft Graph add group member: <https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0>
- Microsoft Graph remove group member: <https://learn.microsoft.com/en-us/graph/api/group-delete-members?view=graph-rest-1.0>
- Microsoft Graph revoke sessions: <https://learn.microsoft.com/en-us/graph/api/user-revokesigninsessions?view=graph-rest-1.0>
- Exchange Online `Set-Mailbox`: <https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailbox>
- Exchange Online address book policies: <https://learn.microsoft.com/en-us/exchange/address-books/address-book-policies/assign-an-address-book-policy-to-mail-users>
- ImportExcel existing workbook workflow: <https://github.com/dfinke/ImportExcel/blob/master/FAQ/How%20to%20Write%20to%20an%20Existing%20Excel%20File.md>
