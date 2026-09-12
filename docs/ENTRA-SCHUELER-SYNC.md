# Entra-Schülersynchronisation

Diese Dokumentation beschreibt den fachlichen Vertrag und die tatsächlich implementierte Kommandozeile. [BETRIEB.md](BETRIEB.md) führt durch einen sicheren Lauf. [TESTING.md](TESTING.md) enthält die Abnahme in der Windows-Parallels-VM.

## Zweck und Grenzen

`Sync-SchuelerEntra.ps1` gleicht Schülerdaten aus genau einer `.xlsx`-Datei mit Microsoft Entra ID und Exchange Online ab. Der Lauf ohne Aktionsparameter ist der Standard und bleibt vollständig lesend. Schreibzugriffe benötigen ausdrücklich `-Update`, `-Add`, `-Remove` oder den Exchange-Reparaturmodus. Das Skript löscht keine Benutzer. Abgänge werden nur bei entsprechender Auswahl deaktiviert und ihre Sitzungen optional widerrufen. Gruppen und Lizenzen eines Abgangs bleiben erhalten.

Der reine Exchange-Modus liest keine Excel-Datei und ändert weder Entra-Benutzer noch Gruppen.

## Laufzeit und Module

Unterstützt wird PowerShell 7.0 oder neuer auf Windows und macOS. Das Modulmanifest fordert mindestens PowerShell 7.0. Die automatisierte Suite kann auf beiden Plattformen laufen. Windows-spezifisches Dateisperrverhalten und produktive Mandantenschreibläufe werden zusätzlich in einer Windows-Parallels-VM abgenommen.

Die Abhängigkeiten sind nicht auf feste Versionen gepinnt. Installiere aktuelle, vom jeweiligen Hersteller unterstützte Versionen und dokumentiere die tatsächlich getesteten Versionen vor dem produktiven Einsatz.

| Zweck | Modul |
|---|---|
| Graph-Anmeldung | `Microsoft.Graph.Authentication` |
| Benutzer lesen und ändern | `Microsoft.Graph.Users` |
| Sitzungen widerrufen | `Microsoft.Graph.Users.Actions` |
| Gruppen lesen und ändern | `Microsoft.Graph.Groups` |
| Postfächer verwalten | `ExchangeOnlineManagement` |
| XLSX lesen und schreiben | `ImportExcel` |
| Tests | `Pester` |
| Statische Analyse | `PSScriptAnalyzer` |

Installation für den angemeldeten Benutzer:

```powershell
Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Users.Actions,Microsoft.Graph.Groups -Scope CurrentUser
Install-Module ExchangeOnlineManagement,ImportExcel,Pester,PSScriptAnalyzer -Scope CurrentUser

$required = @(
  'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users',
  'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Groups',
  'ExchangeOnlineManagement', 'ImportExcel', 'Pester', 'PSScriptAnalyzer'
)
Get-Module -ListAvailable $required |
  Sort-Object Name, Version -Descending |
  Select-Object Name, Version, Path
```

## Anmeldung, Graph-Scopes und Rollen

Das Skript verwendet eine interaktive delegierte Anmeldung. `Connect-SchuelerGraph` fordert genau diese Scopes an und prüft sie nach der Anmeldung:

- `User.ReadWrite.All`
- `User-Mail.ReadWrite.All`
- `Group.Read.All`
- `GroupMember.ReadWrite.All`
- `User.RevokeSessions.All`

Die Zustimmung zu Scopes ersetzt nicht die Rollen des angemeldeten Administrators. Für normale Schülerkonten ist als praktische Aufgabentrennung mindestens eine Kombination aus **User Administrator**, **Groups Administrator** und **Exchange Administrator** vorgesehen. Alternativ kann Exchange RBAC über `Organization Management` oder `Recipient Management` bereitgestellt werden. Ein Global Administrator ist technisch breiter berechtigt, sollte aber nicht als tägliches Betriebskonto verwendet werden. Bei administrativen Zielkonten oder restriktiven Administrative Units können zusätzliche oder höher privilegierte Rollen erforderlich sein.

Exchange Online entscheidet über verfügbare Cmdlets und Parameter anhand von RBAC. Ein nicht unterstützter oder nicht erlaubter Parameter wird als Fehler gemeldet und nicht stillschweigend ersetzt.

Vor einem Update zeigt das Skript Tenant-ID und Konto an. Setze optional `ExpectedTenantId` in `config/SchuelerSync.psd1`. Bei Abweichung bricht die Vorprüfung vor jeder Mutation ab.

## Excel-Vertrag

### Datei und Arbeitsblatt

Ohne `-File` liest das Skript `Schueler.xlsx` im Repository-Root. `-File` akzeptiert einen relativen oder absoluten Pfad. Relative Pfade werden gegen das aktuelle PowerShell-Arbeitsverzeichnis aufgelöst. Zulässig ist ausschließlich `.xlsx`.

Die Kopfzeile steht in Zeile 1. Die Reihenfolge der Spalten ist frei. Das Skript verwendet genau ein Arbeitsblatt, das alle fünf Pflichtspalten enthält. Der Name des Arbeitsblatts ist frei. Gibt es kein passendes oder mehr als ein passendes Arbeitsblatt, ist die Vorprüfung fehlgeschlagen.

Pflichtspalten:

| Spalte | Verwendung |
|---|---|
| `Name mit Rufname` | Lesbare Zuordnung in Berichten |
| `Vorname` | Entra `givenName` und DisplayName |
| `Nachname` | Entra `surname` und DisplayName |
| `Klassen` | Department, Office Location und Klassengruppe |
| `Klassenlehrer` | Eindeutig aufzulösender Manager |

Vom Skript verwaltete Zusatzspalten:

| Spalte | Verhalten |
|---|---|
| `Passwort` | Nur bei neuen Schülern befüllt, bei bestehenden nie geändert |
| `EntraObjectId` | Stabile und bevorzugte Identität |
| `UPN` | Tatsächlich verwendeter UPN |
| `Mail` | Aktuelle Entra-Mailadresse |

Fehlende Zusatzspalten werden nur bei einer autorisierten Rückschreibung rechts ergänzt. Vollständig leere Zeilen werden ignoriert. Eine teilweise gefüllte Datenzeile führt zum Abbruch.

So darf die Tabelle beispielsweise aussehen:

| Name mit Rufname | Vorname | Nachname | Klassen | Klassenlehrer | Passwort | EntraObjectId | UPN | Mail |
|---|---|---|---|---|---|---|---|---|
| Muster, Mia | Mia | Muster | JK1-3g2_1 | Lea Lehrerin | | | | |
| Beispiel, Ömer | Ömer | Beispiel | JK4-6m2_4 | lehrer@monteaufkirchen.com | | | | |

Für jede nicht vollständig leere Datenzeile müssen alle fünf Pflichtfelder gefüllt sein. `Klassenlehrer` enthält entweder den exakten Entra-DisplayName oder die exakte Mail-Adresse beziehungsweise den UPN der Lehrkraft. `Passwort`, `EntraObjectId`, `UPN` und `Mail` sind optional. Lasse sie für neue Schüler leer. Die Schule muss keinen UPN liefern. Ohne gespeicherte Identität erfolgt die Zuordnung über den eindeutigen Rufnamen aus `Vorname` und `Nachname` innerhalb der Schüler-Rollengruppe. Ändere vorhandene Werte in den verwalteten Spalten nicht manuell während eines Laufs.

### Git- und Dateischutz

Die Root-`.gitignore` enthält `/*.xlsx` und zusätzlich `*.[xX][lL][sS][xX]` für Root und Unterordner, unabhängig von der Großschreibung der Dateiendung. Nur `tests/fixtures/Schueler-Testdaten.xlsx` ist als synthetische Fixture ausgenommen. Sicherungen und temporäre XLSX-Dateien bleiben durch zusätzliche Muster ausgeschlossen. Der empfohlene Ordner `Berichte` und der häufig verwendete Rootbericht `output.txt` sind ebenfalls ausgeschlossen.

Vor einer Rückschreibung prüft das Skript die Schreibbarkeit und eine exklusive Dateisperre. Innerhalb eines Git-Worktrees müssen die Quelldatei, der konkret gewählte Sicherungspfad und der temporäre Zielpfad jeweils ignoriert und unverfolgt sein. Alle drei Prüfungen erfolgen vor der ersten vertraulichen Kopie. Eine Ignore-Regel nur für die Quelle reicht nicht aus, auch in einem fremden Repository mit eigener `.gitignore`. Das gilt ebenso für eine mit `-File` gewählte Datei in einem Unterordner. Dateien außerhalb eines Git-Repositories benötigen diese Git-Prüfung nicht, müssen aber sicher gespeichert werden.

Vor dem Schreiben wird der SHA-256-Stand erneut geprüft. Danach entsteht zunächst eine temporäre Kopie im gleichen Ordner. Sie wird geändert, gespeichert und wieder geöffnet. Erst nach erfolgreicher Inhaltsprüfung wird die vorherige Quelldatei ohne Überschreiben nach `<Name>.backup-YYYYMMDD-HHmmss.xlsx` verschoben und die vorbereitete Datei eingesetzt. Bei einer Race Condition oder einem Restore-Fehler bleiben wiederherstellbare Varianten erhalten. Die Fehlermeldung nennt ihre vollständigen Pfade.

## Identität und Matching

Eine Excel-Zeile wird in dieser Reihenfolge zugeordnet:

1. `EntraObjectId`, wenn vorhanden und gültig
2. gespeicherter `UPN`, wenn vorhanden
3. eindeutige normalisierte Kombination aus Vorname und Nachname innerhalb der Schüler-Rollengruppe

Die gespeicherte Objekt-ID ist autoritativ. Sie bleibt stabil, wenn sich Name oder UPN ändern. Doppelte Objekt-IDs, doppelte gespeicherte UPNs, mehrdeutige Namen oder Treffer außerhalb der Schüler-Rollengruppe blockieren die Mutationen. So wird kein vermeintlicher Abgang aufgrund einer unsicheren Zuordnung deaktiviert.

## DisplayName, UPN und Mail

`displayName` ist `<Vorname> <Nachname>`. Vorname, Nachname und DisplayName behalten Umlaute und originale Schreibweise.

Nur für UPN, Mail und `mailNickname` wird normalisiert:

1. Kleinschreibung
2. `ä` zu `ae`, `ö` zu `oe`, `ü` zu `ue`, `ß` zu `ss`
3. weitere diakritische Zeichen entfernen
4. alle Zeichen außerhalb `a-z` und `0-9` entfernen

Die Domain ist `monteaufkirchen.com`. Die Kandidaten beginnen mit dem ersten Buchstaben des normalisierten Vornamens plus Nachname. Danach wird das Vornamenspräfix schrittweise um einen Buchstaben erweitert. Sind alle Präfixe belegt, folgt auf den vollständigen Vornamen und Nachnamen eine Zahl ab `2`.

Beispiel für `Ömer Groß`:

```text
ogross@monteaufkirchen.com
oegross@monteaufkirchen.com
oemgross@monteaufkirchen.com
...
oemergross@monteaufkirchen.com
oemergross2@monteaufkirchen.com
```

Die Kollisionsprüfung umfasst:

- alle Entra-UPNs
- das Entra-Feld `mail`
- alle SMTP-Proxyadressen
- alle Exchange-Empfängeradressen
- alle bereits im aktuellen Lauf reservierten Adressen

Eine Adresse desselben Entra-Objekts ist keine Kollision. Ein bereits regelkonformer UPN bleibt bei unverändertem Vor- und Nachnamen stabil. Jede Abweichung vom ersten Kandidaten erscheint als Warnung mit den kollidierenden Kandidaten.

## Entra-Sollzustand

| Eigenschaft | Sollwert |
|---|---|
| `givenName` | Excel `Vorname` |
| `surname` | Excel `Nachname` |
| `displayName` | `<Vorname> <Nachname>` |
| `userPrincipalName` | gewählter UPN |
| `mail` | gewählter UPN |
| `mailNickname` | lokaler Teil des UPN |
| `department` | Excel `Klassen` |
| `officeLocation` | aus `Klassen` abgeleitet |
| `companyName` | `Montessori Schule Aufkirchen` |
| `employeeType` | `Schüler` |
| `ageGroup` | `Minor` |
| `consentProvidedForMinor` | `Granted` |
| `usageLocation` | `DE` |
| Manager | aufgelöster `Klassenlehrer` |

`legalAgeGroupClassification` wird gegen `MinorWithParentalConsent` verglichen. Diese Eigenschaft ist in Microsoft Graph schreibgeschützt. Das Skript schreibt sie nicht direkt und meldet eine Abweichung als Warnung.

### Office Location

Erlaubt sind `G1`, `G2`, `G3`, `G4`, `M1`, `M2`, `M3`, `O1`, `O2`, `A1` und `A2`. Ermittelt wird das eindeutige Token direkt vor `_` oder am Klassenende, ohne Beachtung der Großschreibung.

```text
JK1-3g2_1  -> G2
JK4-6m2_4  -> M2
```

Kein oder mehr als ein erlaubtes Token führt zu einem Vorprüfungsfehler.

### Manager

Der komplette Inhalt von `Klassenlehrer` bezeichnet genau eine Person. Enthält der Wert `@`, wird exakt und ohne Beachtung der Großschreibung gegen UPN oder Mail gesucht. Andernfalls wird exakt gegen DisplayName gesucht. Null oder mehrere Treffer blockieren den Update-Lauf.

### Gruppen

Jeder aktive Excel-Schüler benötigt direkte Mitgliedschaften in:

- `SEC-A-LIC-O365A1Student`
- `SEC-A-ROL-Schule_Schüler`
- `SEC-A-CLS-<Klassenwert>`, zum Beispiel `SEC-A-CLS-JK4-6m3_5`

Die Schüler-Rollengruppe wird zusätzlich über die konfigurierte ID `cebc1326-1174-4126-ba84-7a8960850e0a` geprüft. Lizenz- und Klassengruppe müssen per DisplayName eindeutig auflösbar sein.

Fehlende Zielgruppen werden zuerst hinzugefügt und nachgelesen. Erst nach dieser Bestätigung entfernt das Skript andere direkte, statische Gruppen mit Präfix `SEC-A-ROL-` oder `SEC-A-CLS-`. Der Endzustand wird nochmals gelesen. Dynamische und geerbte Mitgliedschaften werden nicht entfernt und erscheinen als Warnung. Eine fehlende Zielgruppe löst niemals die Entfernung der alten Gruppe aus.

## Initialpasswörter

Ein Initialpasswort entsteht ausschließlich für einen Neuzugang. Es besteht aus zwei kindgerechten CamelCase-Wörtern mit zusammen genau zehn ASCII-Buchstaben und zwei kryptografisch zufälligen Ziffern von `10` bis `99`. Damit ist es exakt 12 Zeichen lang und enthält Großbuchstaben, Kleinbuchstaben und Ziffern. Umlaute und Sonderzeichen kommen nicht vor. Passwörter wiederholen sich innerhalb eines Laufs nicht.

Der kuratierte Wortschatz enthält 847 vertraute Wörter und einfache Wortformen. Daraus entstehen 220.505 verschiedene Zweiwort-Präfixe mit zehn Buchstaben. Zusammen mit 90 Zahlenwerten ergibt das vor der ersten Reservierung genau 19.845.450 mögliche Passwörter, entsprechend rund 24,24 Bit Auswahlentropie. Beide Zufallsentscheidungen verwenden `RandomNumberGenerator.GetInt32`, jedes Präfix ist gleich wahrscheinlich. Der Regressionstest berechnet den erreichbaren Raum aus den tatsächlich verwendeten Präfixen und verlangt mindestens 24 Bit. Das merkbare Format ist keine Folge aus zwölf unabhängig zufälligen Zeichen und bietet nicht deren Entropie.

Das Passwortprofil setzt `ForceChangePasswordNextSignIn = false`. Lehnt Entra das Passwort eindeutig wegen der Passwort-Richtlinie ab, werden höchstens fünf neue Passwörter versucht. Einschließlich Erstversuch sind das maximal sechs Versuche. Andere Graph-Fehler werden nicht als Passwortfehler wiederholt.

Im Excel-basierten Modus erscheinen Passwörter nicht in der Ausgabe, also weder in Tabellen noch in Informationsmeldungen oder Aktionsfehlern. Sie werden nur in die ausgewählte Excel-Datei und deren lokale Sicherung geschrieben. Bestehende Schülerpasswörter werden nie erzeugt, ersetzt oder zurückgesetzt.

Im manuellen `-Add`-Modus gibt es keine Excel-Rückschreibung. Bei einer echten Neuanlage zeigt das Skript das Startpasswort genau einmal im Terminal an. Im Erfolgsfall geschieht dies nach verifizierter Aktivierung. Scheitert ein Schritt nach der Kontoerstellung, wird das Passwort für den sicheren Wiederanlauf ebenfalls einmal angezeigt und der Kontostatus ausdrücklich als unbekannt gemeldet. Das Passwort erscheint nicht im Ergebnisobjekt und nicht in `-OutputFile`. Bei einem bereits vorhandenen Schüler wird kein Passwort erzeugt oder angezeigt. Verwende für eine manuelle Neuanlage kein PowerShell-Transcript und kopiere das Passwort unmittelbar in einen geeigneten geschützten Kanal.

## Vergleichsausgabe

Jeder Lauf erzeugt diese Bereiche:

- `Neuzugänge`, nur in Excel
- `Abgänge`, nur in der Schüler-Rollengruppe
- `Änderungen`, in beiden Quellen und mit feldgenauer Abweichung
- `Bestehende`, in beiden Quellen und ohne Abweichung
- `Warnungen und Fehler`, etwa Kollisionen, Manager-, Gruppen- oder Postfachprobleme
- `Aktionsergebnisse`, nur relevant für Update und Exchange-Reparatur

Ein fehlendes Postfach wird im Vergleich nur einmal geprüft. Es erscheint als `EXO-Konfiguration ausstehend`. Der reine Vergleich wartet nicht und schreibt nichts.

## Parameter und Modi

### Lesender Standardvergleich

```powershell
.\Sync-SchuelerEntra.ps1
.\Sync-SchuelerEntra.ps1 -File 'C:\Import\Schueler-2026.xlsx'
```

### Vollständiger Update-Lauf

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\Import\Schueler-2026.xlsx' -Update -WhatIf
.\Sync-SchuelerEntra.ps1 -File 'C:\Import\Schueler-2026.xlsx' -Update
```

`-Update` ohne Selektor führt in dieser Reihenfolge aus:

1. Neuzugänge erstellen
2. Änderungen aktiver Schüler anwenden
3. Abgänge deaktivieren
4. Sitzungen der Abgänge widerrufen
5. Gruppen aktiver Schüler normalisieren
6. Exchange Online für alle aktiven Excel-Schüler abgleichen

### Selektiver Update-Lauf

```powershell
.\Sync-SchuelerEntra.ps1 -Update -CreateNewUsers
.\Sync-SchuelerEntra.ps1 -Update -UpdateUsers
.\Sync-SchuelerEntra.ps1 -Update -DisableUsers -RevokeSessions
```

Sobald mindestens einer der folgenden Schalter angegeben wird, führt das Skript nur die gewählten Graph-Aktionen aus:

| Schalter | Aktion |
|---|---|
| `-CreateNewUsers` | Neuzugänge erstellen, Pflichtgruppen setzen, erfolgreiche Benutzer in EXO konfigurieren |
| `-UpdateUsers` | Abweichungen und Gruppen korrigieren, erfolgreiche Benutzer in EXO konfigurieren |
| `-DisableUsers` | Abgänge deaktivieren |
| `-RevokeSessions` | Sitzungen der Abgänge widerrufen |

`-DisableUsers` und `-RevokeSessions` allein lösen keinen Exchange-Abgleich für aktive Schüler aus. Aktionsschalter ohne `-Update` sind ungültig.

### Reiner Exchange-Reparaturmodus

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com'
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'test1@monteaufkirchen.com','test2@monteaufkirchen.com'
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -UserPrincipalName 'test1@monteaufkirchen.com' -WhatIf
```

`-Mail` ist ein Array. Die Aliase sind `-UPN` und `-UserPrincipalName`. Dieser Parametersatz ist nicht mit `-File`, `-Update` oder den vier Graph-Aktionsschaltern kombinierbar. Er ändert nur die unten beschriebene Exchange-Konfiguration.

### Einzelnen Schüler anlegen oder aktualisieren

```powershell
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin' -WhatIf
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'lehrerin@monteaufkirchen.com'
```

`-Add` liest und schreibt keine Excel-Datei. Das Skript sucht den Schüler über den eindeutig normalisierten Vor- und Nachnamen innerhalb der direkten Schüler-Rollengruppe. Gibt es noch keinen Treffer, legt es den Schüler mit allen Pflichtattributen und Gruppen an. Gibt es genau einen Treffer, bleiben dessen UPN und Passwort erhalten und das Skript aktualisiert abweichende Attribute, Manager und verwaltete Gruppen. Ein deaktiviertes Bestandskonto wird erst nach erfolgreicher Prüfung dieses Pflichtzustands aktiviert. Mehrdeutige Namen oder ein gleichnamiger Benutzer außerhalb der Schüler-Rollengruppe blockieren den Lauf. Nach einer erfolgreichen Neuanlage oder Aktualisierung wird Exchange Online automatisch geprüft und konfiguriert, einschließlich der konfigurierten Wiederholungen.

`-EntraObjectId <ID>` ist ausschließlich für den vom Skript ausgegebenen Wiederanlauf nach einer teilweise erfolgreichen manuellen Neuanlage vorgesehen. Damit wird exakt das bereits erzeugte Entra-Objekt weiterverarbeitet, auch wenn die Schüler-Rollengruppe beim ersten Versuch noch nicht gesetzt werden konnte. Diese Ausnahme gilt nicht für Excel-Zeilen, gespeicherte UPNs oder die normale Namenssuche. Unmittelbar vor jeder Mutationsphase liest der Wiederanlauf das Objekt erneut aus Graph und verlangt weiterhin denselben normalisierten Vor- und Nachnamen sowie exakt `CompanyName = Montessori Schule Aufkirchen` und `EmployeeType = Schüler`. Übernimm die Objekt-ID nur aus dem vorherigen Aktionsergebnis und prüfe sie vor der Bestätigung.

### Einzelnen Schüler deaktivieren

```powershell
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com'
```

`-Remove` akzeptiert genau einen UPN und liest keine Excel-Datei. Der gefundene Benutzer muss direktes Mitglied von `SEC-A-ROL-Schule_Schüler` sein. Das Skript deaktiviert ausschließlich dieses Konto und widerruft anschließend dessen Sitzungen. Es löscht den Benutzer nicht und entfernt weder Gruppen noch Lizenzen. Auch wenn die Deaktivierung fehlschlägt, versucht es den Sitzungswiderruf und meldet beide Ergebnisse getrennt.

`-OutputFile <Pfad>` schreibt den vollständigen, passwortfreien Laufbericht zusätzlich als UTF-8-Textdatei. Relative Pfade beziehen sich auf das aktuelle PowerShell-Arbeitsverzeichnis. Fehlende Unterordner werden angelegt und eine vorhandene Datei wird für jeden Lauf ersetzt. Der Parameter funktioniert im Vergleichs-, Update-, Add-, Remove- und Exchange-Only-Modus. Die laufend aktualisierte Fortschrittsanzeige bleibt ausschließlich im Terminal. Verwende bevorzugt den per `.gitignore` ausgeschlossenen Ordner `Berichte`, da der Bericht personenbezogene Schülerdaten enthält.

Bestehende Entra-Benutzer behalten ihren aktuellen UPN auch dann, wenn sich Vorname oder Nachname in Excel ändern. Mail und `mailNickname` werden für den Sollvergleich aus diesem bestehenden UPN abgeleitet. Nur bei Neuzugängen erzeugt das Skript einen neuen kollisionsfreien UPN.

### Gemeinsame PowerShell-Schalter

`-WhatIf` führt keine Graph-, Excel- oder Exchange-Schreiboperation aus. Für Exchange erfolgt nur die sofortige Verfügbarkeitsprüfung, ohne Wartezeit. `-Confirm` aktiviert die üblichen PowerShell-Rückfragen. `-Verbose` aktiviert die üblichen ausführlichen Meldungen.

## Fail-safe-Lebenszyklus eines Neuzugangs

1. Initialpasswort erzeugen
2. Entra-Benutzer mit `accountEnabled = false` erstellen
3. Attribute schreiben und lesen
4. Manager setzen und lesen
5. Lizenz-, Rollen- und Klassengruppe hinzufügen und lesen
6. `Passwort`, `EntraObjectId`, tatsächlichen `UPN` und `Mail` in die geschützte Arbeitsmappe schreiben und lesen
7. Benutzer aktivieren und Aktivierung lesen
8. Postfach gesammelt abwarten und konfigurieren

Scheitert ein Pflichtschritt bis einschließlich Excel-Rückschreibung, bleibt das Konto deaktiviert. Das Skript löscht es nicht. Nach einem Laufzeitfehler werden unabhängige Benutzer weiterverarbeitet. Es gibt keinen tenantweiten Rollback.

Bei Namensänderungen speichert das Skript die stabile Objekt-ID vor einem UPN-Wechsel und danach die tatsächlich aus Entra gelesene Identität. So bleibt ein teilweise erfolgreicher Graph-Schreibvorgang wieder zuordenbar.

## Exchange-Online-Sollzustand

Das Skript liest zuerst `Get-Mailbox` und `Get-CASMailbox`, schreibt nur Abweichungen und liest den Zustand danach erneut. Vorhandene zusätzliche Audit-Aktionen bleiben bestehen.

### Mailbox und Richtlinien

| Eigenschaft | Sollwert |
|---|---|
| `CustomAttribute1` | `Montessori Schule Aufkirchen - Schüler` |
| `AddressBookPolicy` | `MON-EXO-ABP-Schule_Schüler` |
| `AuditEnabled` | `true` |
| `AuditLogAgeLimit` | 365 Tage |
| `RetainDeletedItemsFor` | 30 Tage |
| `RoleAssignmentPolicy` | `MON-EXO-UserRoles-Default` |
| `SharingPolicy` | `MON-EXO-Sharing-Default` |
| `RetentionPolicy` | `MON-EXO-Retention-Default` |

### AuditDelegate

Mindestens vorhanden sein müssen `Create`, `FolderBind`, `HardDelete`, `Move`, `MoveToDeletedItems`, `SendAs`, `SendOnBehalf`, `SoftDelete`, `Update`, `UpdateFolderPermissions` und `UpdateInboxRules`.

### AuditOwner

Mindestens vorhanden sein müssen `Create`, `HardDelete`, `Move`, `MailboxLogin`, `MoveToDeletedItems`, `SoftDelete`, `Update`, `UpdateFolderPermissions`, `UpdateInboxRules` und `UpdateCalendarDelegation`.

### AuditAdmin

Mindestens vorhanden sein müssen `Copy`, `Create`, `FolderBind`, `HardDelete`, `Move`, `MoveToDeletedItems`, `SendAs`, `SendOnBehalf`, `SoftDelete`, `Update`, `UpdateFolderPermissions`, `UpdateInboxRules` und `UpdateCalendarDelegation`. Enthält `PersistedCapabilities` den Wert `CommunicationsCompliance`, kommt `MailItemsAccessed` hinzu.

### CAS-Mailbox

| Eigenschaft | Sollwert |
|---|---|
| `ActiveSyncEnabled` | `false` |
| `ImapEnabled` | `false` |
| `MAPIEnabled` | `true` |
| `OWAEnabled` | `true` |
| `OWAforDevicesEnabled` | `false` |
| `OwaMailboxPolicy` | `MON-EXO-OWA-Default` |
| `PopEnabled` | `false` |
| `SmtpClientAuthenticationDisabled` | `true` |

### Bereitstellungswiederholung

Bei einem autorisierten Update und im Exchange-Reparaturmodus erfolgt eine sofortige Prüfung und danach höchstens fünf weitere Prüfungen. Zwischen den Wiederholungen liegen jeweils 60 Sekunden. Damit gibt es maximal sechs Prüfungen und fünf Minuten zusätzliche Wartezeit.

Mehrere UPNs werden als Batch verarbeitet. Pro Runde werden nur fehlende Postfächer geprüft. Gefundene Postfächer werden sofort konfiguriert und aus der Restmenge entfernt. Das Skript wartet nur einmal pro Runde, nicht einmal pro Benutzer. Ein Fehler bei `Set-Mailbox` oder `Set-CASMailbox` wird sofort gemeldet und nicht blind wiederholt.

Bleibt ein Postfach ausstehend, enthält das Ergebnis einen Wiederanlaufbefehl:

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com'
```

## Datenschutz und Protokollierung

- Verwende nur erfundene Personen in Repository-Tests.
- Speichere produktive Arbeitsmappen nur in einem geschützten lokalen Ordner.
- Versioniere keine produktive XLSX-Datei, auch nicht aus Unterordnern. Nur die benannte synthetische Fixture ist von den XLSX-Ignore-Regeln ausgenommen.
- Sichere Arbeitsmappen und Backups nach denselben Regeln wie Passwörter.
- Aktiviere für diesen Lauf kein Transcript, wenn dessen Zugriffsschutz nicht geprüft ist.
- Das Skript redigiert bekannte Passwortwerte aus öffentlichen Ergebnissen. Behandle Fehlermeldungen trotzdem als personenbezogene Betriebsdaten.
- Entferne temporäre und Backup-Dateien erst nach geprüfter Wiederherstellung.

## Fehler und Wiederanlauf

Ein Fehler der Vorprüfung verhindert jede Mutation. Dazu gehören unsichere Identitäten, ungültige Klassen, nicht eindeutige Manager, fehlende Zielgruppen, falscher Tenant und eine unsichere oder veränderte Excel-Datei.

Nach einem Laufzeitfehler:

1. Prüfe `Aktionsergebnisse`, `RecoveryCommand` und alle genannten Dateipfade.
2. Prüfe Entra über die gespeicherte oder ausgegebene Objekt-ID. Verlasse dich nicht nur auf den Namen.
3. Bei Excel-Fehlern vergleiche Quelle, `.backup-...xlsx` und erhaltene `.tmp.xlsx`, bevor du etwas löschst oder umbenennst.
4. Starte zuerst den lesenden Vergleich mit derselben `-File`-Angabe.
5. Führe nur noch die notwendige Teilaktion aus.
6. Nutze für ein ausstehendes Postfach den Exchange-Reparaturmodus.

Ausführliche Ablaufbeispiele stehen in [BETRIEB.md](BETRIEB.md).

## Bekannte Einschränkungen

- Nur `.xlsx`, kein `.xls` und kein CSV
- Nur interaktive delegierte Anmeldung
- Kein Löschen von Benutzern
- Kein automatisches Entfernen von Gruppen oder Lizenzen bei Abgängen
- Keine automatische Korrektur dynamischer oder geerbter Gruppen
- Kein tenantweiter Rollback nach bereits erfolgreichen Einzelaktionen
- `legalAgeGroupClassification` ist schreibgeschützt und wird nur validiert
- Produktive Schreibläufe erfordern die kontrollierten Windows- und Mandantentests aus [TESTING.md](TESTING.md)

## Primärquellen

- [Microsoft Graph Benutzer aktualisieren](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0)
- [Microsoft Graph Sitzungen widerrufen](https://learn.microsoft.com/en-us/graph/api/user-revokesigninsessions?view=graph-rest-1.0)
- [Microsoft Graph Berechtigungsreferenz](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Exchange Online PowerShell verbinden](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps)
- [Set-Mailbox](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailbox?view=exchange-ps)
- [Set-CASMailbox](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-casmailbox?view=exchange-ps)
