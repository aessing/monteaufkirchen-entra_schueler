# Personal 0.3.0: Spezifikation

Zusätzlich beauftragt ist `-ListToExcel` in den bestehenden Schüler- und Lehrereinstiegen. Die [Schüleranleitung](ENTRA-SCHUELER-SYNC.md) beschreibt zehn Spalten mit Klasse, Klassenlehrer und Login-Zeitstempel, die [Lehreranleitung](ENTRA-LEHRER-SYNC.md) neun Spalten mit dem vorhandenen Job und Login-Zeitstempel. Der gemeinsame verifizierte Schreiber wird wiederverwendet. Der nachfolgende Personalvertrag und die bisherigen Schüler-/Lehrer-Sync-Regeln bleiben erhalten.

Stand: 10. Oktober 2026. Vom Betreiber freigegeben und auf `codex/0.3.0-personal` lokal implementiert. Ergänzt um Rollenanzeige, Gastaktionen, gemeinsame Konten-TUIs, Adressregeln für Neuanlagen und lesbare Jobnamen mit eindeutiger Rückzuordnung. Diese Spezifikation beschreibt den umgesetzten Vertrag. Lokale Ergebnisse und offene Windows-/Mandantenprüfungen stehen in [ABNAHME.md](ABNAHME.md).

## 1. Ziel und Eingrenzung

Zusätzlich beauftragt am 30. September 2026: `Sync-LehrerEntra.ps1 -ManageAccounts` verwendet dieselbe Inventar-, TUI- und Aktionsumsetzung mit einer ausdrücklich auf direkte Schulteam-/Ganztagsmitglieder begrenzten Population. Die direkte Rollenmitgliedschaft wird auch vor jeder Mutation geprüft, die Population bleibt nach dem Neuladen erhalten. Der Personalvertrag mit ausgeschlossenen Lehrern bleibt bestehen. Der eigene Lehrer-Parametersatz benötigt keine Arbeitsmappe und verändert die bisherigen Lehrer-Sync-Regeln nicht. [Bedienung und Schutzregeln](ENTRA-LEHRER-SYNC.md).

`Sync-PersonalEntra.ps1` verwaltet die übrigen Benutzer außerhalb von Lehrern, CO-L, Schulpädagogen, OGTS und Schülern. Einzelanlage ist der Hauptfall. Ein Excel-Sync bleibt möglich. Bestandsliste, Excel-Export und TUI erfassen auch Konten ohne Rollengruppe.

Das bestehende Modul bleibt `SchuelerSync`. Die öffentlichen Aufrufe für Schüler, Lehrer und gesperrte Konten bleiben kompatibel. Bestehende private Funktionen werden gezielt wiederverwendet oder um einen notwendigen Konfigurationskontext ergänzt.

## 2. Rollen und Profilwerte

Die folgenden IDs und Werte stammen aus der vom Betreiber gelieferten Vorlage. Sie wurden nicht gegen den Mandanten geprüft. Der interne `Job` ist jeweils der Gruppenname ohne `SEC-A-ROL-`, mit der hier angegebenen Schreibweise. Der Personalexport ersetzt Unterstriche durch Leerzeichen und trennt CamelCase, etwa `Kinderhaus Pädagogisches Team`. Import und Einzelpflege erkennen die lesbaren Namen sowie die internen Werte anhand der konfigurierten Jobs. Groß-/Kleinschreibung, äußere und mehrfache Leerzeichen werden normalisiert. Unbekannte oder mehrdeutige Zuordnungen werden abgewiesen, geschützte Jobs bleiben geschützt. Lehrerjobkürzel sind im neuen Skript ungültig. Das tatsächliche Personal-`JobTitle` verwendet bei Neuanlage und normaler Profilpflege den lesbaren Namen. Bereits vorhandene technische Jobtitel werden im Vergleich als Änderung angezeigt, lesende Aufrufe ändern sie nicht. Bestandsliste und Vergleich zeigen ebenfalls lesbare Jobnamen. Interne Profil- und Rollenschlüssel bleiben kanonisch.

| Job | GroupId | CompanyName | EmployeeType | AddressBookPolicy |
|---|---|---|---|---|
| `Schule_Sekretariat` | `534f94ed-420d-4794-a374-20c572def8cb` | Montessori Schule Aufkirchen | Sekretariat | `MON-EXO-ABP-Schule_Sekretariat` |
| `Schule_Vertretungskräfte` | `4ed4f87b-374f-4679-ac11-5e51212e9a91` | Montessori Schule Aufkirchen | Vertretungskräfte | `MON-EXO-ABP-Schule_Vertretungskräfte` |
| `Kinderhaus_PädagogischesTeam` | `cfc6d358-875d-441b-b45a-0d9734b7d788` | Montessori Kinderhaus Erding | Pädagogisches Team | `MON-EXO-ABP-Kinderhaus_PädagogischesTeam` |
| `Kinderhaus_Sekretariat` | `43165f2c-3ce5-471e-a7be-e16b66c03406` | Montessori Kinderhaus Erding | Sekretariat | `MON-EXO-ABP-Kinderhaus_Sekretariat` |
| `Kinderhaus_Vertretungskräfte` | `d2e7e424-83d2-481d-9226-bf529bc9640a` | Montessori Kinderhaus Erding | Vertretungskräfte | `MON-EXO-ABP-Kinderhaus_Vertretungskräfte` |
| `Unterstützung` | `a4a00442-bc04-48ee-b659-cb3bde194630` | Montessori Verein Landkreis Erding e.V. | Unterstützung | `MON-EXO-ABP-Unterstützung` |
| `ExterneBenutzer` | `9c620843-0071-4d44-ba23-318e19cd09a0` | Montessori Verein Landkreis Erding e.V. | Externe Benutzer | `MON-EXO-ABP-ExterneBenutzer` |
| `ExterneAdmins` | `a246dfde-eccf-489b-ac1c-8d92f67145ca` | Montessori Verein Landkreis Erding e.V. | Externe Admins | `$null`, bestehende ABP nicht verwalten |

`ExterneAdmins` ist gemäß Freigabe nur ein bekanntes, geschütztes Profil. Die Zielwerte werden dokumentiert, aber nicht automatisch angewandt.

Für freigegebene Profile gelten `AgeGroup = Adult`, `ConsentProvidedForMinor = NotRequired` und `UsageLocation = DE`. `CustomAttribute1` lautet `<CompanyName> - <EmployeeType>`. `legalAgeGroupClassification` wird wie bisher nur verglichen.

`SEC-A-ADM-GlobalAdministrators` wird als geschützte Gruppe erfasst. Es gibt keine gelieferte ID und keine Profilwerte. Zur Schutzprüfung ist eine eindeutige Namensauflösung zulässig. Fehlt diese Gruppe oder ist der Name mehrdeutig, ist der Schutz unvollständig und der Personal-Schreiblauf bricht vor Mutationen ab. Inventaransichten dürfen mit klarer Warnung weiterlaufen.

### Ausschlüsse

| Population | Gruppen-ID | Konfigurationsquelle |
|---|---|---|
| Schule-Pädagogik mit L, CO-L und PA | `103a1c4c-036b-40bc-bbe8-e3f7a866cb01` | `config/LehrerSync.psd1`, `Roles.School` |
| Ganztag/OGTS | `3997db35-2914-4b0f-92b4-08038ef8f84a` | `config/LehrerSync.psd1`, `Roles.Ganztag` |
| Schüler | `cebc1326-1174-4126-ba84-7a8960850e0a` | `config/SchuelerSync.psd1`, `StudentRoleGroup` |

Diese IDs werden im Programm aus ihren vorhandenen Konfigurationen gelesen. Die Tabelle dient der Prüfung und ist keine zweite Konfigurationsquelle.

## 3. Gemeinsame Bestandsauswahl

1. Alle aktiven und gesperrten Entra-Benutzer lesen, einschließlich benötigter Identitäts- und Statusfelder.
2. Die drei Ausschlusspopulationen vollständig auflösen. Direkte und verschachtelte Mitgliedschaften zählen für den Ausschluss. Ein Ausschluss hat Vorrang vor einer zusätzlichen Personalrolle.
3. Exchange-Empfänger einmal vollständig lesen und anhand `ExternalDirectoryObjectId` zuordnen. Doppelte oder unlesbare Zuordnungen erhalten einen Prüfstatus.
4. Die drei ausgeschlossenen Populationen und `SharedMailbox` aus Liste, Export und TUI entfernen.
5. Die übrigen Konten unabhängig von Rollenzugehörigkeit zeigen und nach `DisplayName`, dann UPN sortieren.

Die Auswahl darf nicht allein auf `JobTitle`, `EmployeeType`, Namensmustern oder der Mitgliedschaft in Personalrollen beruhen. Rollenlose Konten bleiben enthalten. Fehler beim vollständigen Laden der Ausschlusspopulationen oder des Exchange-Bestands brechen ab, damit eine unvollständige Auswahl nicht als vollständig erscheint.

Gäste, synchronisierte Konten, Ressourcenpostfächer, geschützte Admins sowie nicht eindeutig prüfbare Konten bleiben mit Grund sichtbar. Vor jedem Schreibzugriff werden Identität, Kontostatus, Ausschlüsse, Schutzgruppen, relevante Mitgliedschaften, Verzeichnisrollen und Exchange-Zuordnung erneut geprüft. Fehlende Prüfdaten führen zu `Skipped` oder zum Vorprüfungsabbruch.

## 4. Aufrufe und Parameter

| Modus | Parameter | Vertrag |
|---|---|---|
| Vergleich, Standard | `-File`, `-OutputFile` | `Personal.xlsx` im Repository-Root, nur lesen und vergleichen |
| Sync | `-Update`, optional `-CreateNewUsers`, `-UpdateUsers`, `-DisableUsers`, `-RevokeSessions` | Bekannte Aktionsschalter, Unterschiede aus Abschnitt 5 |
| Einzelanlage/-pflege | `-Add -Vorname -Nachname -Job`, optional `-EntraObjectId`, `-ResumeCreation` | Neues Konto oder explizit zugeordnetes Bestandskonto |
| Einzelkonto sperren | `-Remove -Mail` | Genau ein Konto sperren und Sitzungen widerrufen |
| Exchange reparieren | `-ConfigureExchangeOnlineOnly -Mail` | Ein oder mehrere explizite Ziele mit eindeutigem freigegebenem Profil |
| Inventar | `-List`, optional `-PassThru`, `-OutputFile` | Keine Exceldatei erforderlich, vollständige lesende Bestandsliste |
| Ergebnisobjekt | `-PassThru` ohne `-List` | Lesende Bestandsabfrage, keine TUI |
| Excel-Export | `-ListToExcel`, optional `-File`, `-PassThru` | Neue `.xlsx`-Datei, standardmäßig `Personal.xlsx` im Repository-Root |
| Kontenverwaltung | `-ManageAccounts` | Interaktive TUI ohne Exceldatei |

`-Mail` behält die Aliase `-UPN` und `-UserPrincipalName`. `-File` nimmt `.xlsx` an, `-OutputFile` bleibt ausschließlich ein neuer passwortfreier `.txt`-Bericht. Gemeinsame PowerShell-Parameter einschließlich `-WhatIf` und `-Confirm` werden bis zum tatsächlichen Schreibzugriff weitergegeben. Schreibmodi verwenden `SupportsShouldProcess`.

Unvereinbare Modi werden über Parametersätze abgewiesen. `-PassThru` ist nur in Inventar und Export vorgesehen. `-ListToExcel` braucht kein zusätzliches `-Update`, weil der Schalter die lokale Dateierzeugung bereits ausdrücklich auswählt. Unter `-WhatIf` wird auch diese Datei nicht erzeugt.

```powershell
# Implementierte Aufrufe mit erfundenen Beispieldaten
.\Sync-PersonalEntra.ps1 -Add -Vorname 'Lea' -Nachname 'Testwald' -Job 'Schule_Sekretariat' -WhatIf
.\Sync-PersonalEntra.ps1 -File '.\Personal.xlsx'
.\Sync-PersonalEntra.ps1 -File '.\Personal.xlsx' -Update -WhatIf
.\Sync-PersonalEntra.ps1 -List
.\Sync-PersonalEntra.ps1 -ListToExcel
.\Sync-PersonalEntra.ps1 -ListToExcel -File '.\Personal-Export.xlsx' -WhatIf
.\Sync-PersonalEntra.ps1 -ManageAccounts -WhatIf
.\Sync-PersonalEntra.ps1 -Remove -UPN 'ltestwald@montessori-erding.de' -WhatIf
.\Sync-PersonalEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@montessori-erding.de' -WhatIf
```

## 5. Sync, Identität und Kontozustand

Der Vergleich zeigt wie der Lehrerabgleich Neuzugänge, Abgänge, Änderungen, Bestehende sowie Warnungen und Fehler. Bei `-Update` erscheint der Vergleich vor den Aktionen. Abgänge werden nur innerhalb der verwalteten freigegebenen Personalrollen ermittelt.

Freigegebene Personalregel: `-Update` ohne Aktionsschalter wählt Erstellen und Aktualisieren. Sobald ein Aktionsschalter vorliegt, werden wie bisher nur ausdrücklich ausgewählte Aktionen ausgeführt. `-DisableUsers` und `-RevokeSessions` benötigen `-Update` und eine vollständige Liste aller aktiven Konten der freigegebenen Personalrollen. Beide Schalter arbeiten nur auf den so ermittelten Abgängen. Eine leere Importliste blockiert Updates. Rollenlose und geschützte Konten sind keine automatischen Abgänge.

Identität wird über Objekt-ID, UPN oder eindeutig normalisierten Vor- und Nachnamen innerhalb der verwalteten Profile zugeordnet. Fremdkonten werden beim Namensvergleich als Konflikt gemeldet. Bestehende UPNs, Mailadressen, Mail-Nicknames und Aliase werden erhalten. Neue Personal-UPNs und Mailadressen verwenden `montessori-erding.de`. Angestellte beginnen mit Vornamensinitiale plus Nachname und erweitern das Präfix bei Belegung. Freigegebene Rollen `Externe*` verwenden `vorname.nachname` mit Zahl ab `2` bei Belegung. `ExterneAdmins` bleibt geschützt. Der gemeinsame [Adressvertrag](ADRESSREGELN.md) beschreibt Domains, Kollisionsprüfung und Bestandsschutz.

Ein rollenloses Konto wird nur über eine ausdrücklich angegebene `EntraObjectId` plus vollständigen Namen und `Job` in die Profilpflege übernommen. Das gilt bei `-Add` und einer entsprechend bearbeiteten Excel-Zeile. Vor einer Übernahme müssen alle Ausschlüsse und Schutzprüfungen bestehen. Ein gleicher Name allein reicht nicht. Leere oder mehrdeutige Jobs blockieren den schreibenden Import vor der ersten Mutation.

Der rein lesende Personalvergleich nimmt einen unveränderten Gesamtexport entgegen. Unvollständige Pflichtwerte, leere/unbekannte/geschützte Jobs und explizit identifizierte, nicht schreibbare Bestandskonten erscheinen als `Workbook.Review` mit Zeilennummer und Identität. Zulässige Zeilen werden weiter verglichen. Eine Prüfzeile erzeugt keinen Neuzugang, eine bekannte explizite Identität dieser Zeile wird nicht als Abgang gewertet. Diese Ausnahme gilt ausschließlich für `Compare`, nicht für Update einschließlich WhatIf, Einzelpflege oder Lehrer-/Schülerimporte. Die strikten Vorprüfungen und sämtliche Mutationprüfungen bleiben erhalten.

Neue Konten werden deaktiviert angelegt. Bei Sync werden Passwort und Identität früh in Excel gesichert. Danach folgen Attribut- und Gruppenprüfung und die Aktivierung nach erfolgreichem Pflichtzustand. Manuelles `-Add` zeigt das Initialpasswort genau einmal im Terminal. Passwörter haben wie bei Lehrern zwölf zufällige Zeichen und erfordern den Erstwechsel. Bestehende Passwörter werden nicht geändert.

Bereits gesperrte Bestandskonten behalten beim neuen Personal-Sync ihren Status. Eine gewollte Entsperrung erfolgt über die TUI. Der Wiederanlauf einer fehlgeschlagenen Neuanlage mit expliziter ID und passendem Profil darf die ursprüngliche Aktivierung nachholen. Dafür ist `-ResumeCreation` zusammen mit `-Add -EntraObjectId` erforderlich. Ohne diesen Schalter bleibt ein gesperrtes Bestandskonto gesperrt. Der Wiederanlauf prüft den Namen und das bereits passende Firmen-/Mitarbeitertyp-Profil.

Rollenwechsel betrifft ausschließlich explizit konfigurierte verwaltete Personalrollen. Erst Zielmitgliedschaft hinzufügen und bestätigen, dann die bekannte alte Rolle entfernen. Fremde `SEC-A-ROL-*` und andere Gruppen bleiben erhalten. Mehrere konkurrierende Personalprofile, dynamische Rollen und selbst lizenzierende Rollen blockieren Änderungen. Vor dem Schreiben werden Mitgliedschaften, Rollenlizenzen und Änderbarkeit erneut gelesen. Admin-Schutzgruppen dürfen weder hinzugefügt noch entfernt werden.

## 6. Lizenzierung

Freigegebenes Personal erhält `SEC-A-LIC-O365A1Faculty`, wenn nachweislich keine direkte Lizenz, keine Gruppenlizenz und keine ausstehende gruppenbasierte Zuweisung vorliegt. Die Gruppe wird vorab eindeutig aufgelöst und auf tatsächliche Lizenzzuweisung sowie Änderbarkeit geprüft.

Die Entscheidung verwendet die vorhandene Lehrerlogik und wird unmittelbar vor einem nötigen Gruppenbeitritt wiederholt. `Error = None` ist kein Fehler. Tatsächliche Fehler oder unvollständige Lizenzdaten blockieren die automatische Zuweisung. Vorhandene Lizenzen und Lizenzgruppen bleiben bei Profilwechsel, Sperrung und Entsperrung bestehen.

Graph beschreibt Herkunft und Zustand einer Zuweisung über `assignedByGroup`, `state` und `error`. Das stützt die bestehende kombinierte Prüfung, keine reine Kontrolle des Gruppennamens. [Microsoft Graph: licenseAssignmentState](https://learn.microsoft.com/en-us/graph/api/resources/licenseassignmentstate?view=graph-rest-1.0).

## 7. Exchange-Verhalten

Die gemeinsamen Funktionen für Istzustand, feldgenauen Vergleich, Schreibparameter, Auditaktionen und Verifikation bleiben die Grundlage. Geschrieben werden nur Abweichungen. Ein aktuell konformes Postfach liefert `Compliant` ohne Änderungsbestätigung.

| Einstellung | Ziel |
|---|---|
| `AddressBookPolicy` | Profilwert, bei `$null` oder Leerwert überspringen |
| `CustomAttribute1` | `<CompanyName> - <EmployeeType>` |
| `AuditEnabled` | `$true` |
| `AuditLogAgeLimit` | 365 Tage als `TimeSpan` |
| `RetainDeletedItemsFor` | 30 Tage als `TimeSpan` |
| `RoleAssignmentPolicy` | `MON-EXO-UserRoles-Default` |
| `SharingPolicy` | `MON-EXO-Sharing-Default` |
| `RetentionPolicy` | `MON-EXO-Retention-Default` |
| `OwaMailboxPolicy` | `MON-EXO-OWA-Default` |
| `ActiveSyncEnabled`, `ImapEnabled`, `OWAforDevicesEnabled`, `PopEnabled` | `$false` |
| `MAPIEnabled`, `OWAEnabled`, `SmtpClientAuthenticationDisabled` | `$true` |

`AuditDelegate`, `AuditOwner` und `AuditAdmin` übernehmen exakt die bestehenden Listen aus `Get-RequiredExchangeAuditAction`. Fehlende Aktionen werden mit `@{ Add = ... }` ergänzt, vorhandene zusätzliche Aktionen bleiben erhalten. `MailItemsAccessed` wird wie bisher für `AuditAdmin` bei vorhandener `CommunicationsCompliance`-Capability ergänzt. Eine Umstellung der Auditstrategie ist kein Teil dieser Erweiterung.

Eine nicht konfigurierte ABP wird in Vorprüfung, Vergleich, Parametersammlung und Nachprüfung ausgelassen. Das ist wesentlich: Ein explizites `Set-Mailbox -AddressBookPolicy $null` würde die Zuweisung entfernen. [Microsoft: ABP zuweisen oder entfernen](https://learn.microsoft.com/en-us/exchange/address-books/address-book-policies/assign-an-address-book-policy-to-mail-users).

Die bestehenden `TimeSpan`-Werte werden wiederverwendet. Microsoft dokumentiert diesen Parametertyp und maximal 30 Tage für gelöschte Elemente in Exchange Online. [Microsoft: Set-Mailbox](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailbox?view=exchange-ps).

Für erfolgreich erstellte oder aktualisierte Konten mit erwartetem Benutzerpostfach gilt weiterhin die sofortige Prüfung plus bis zu fünf Wiederholungen im Abstand von 60 Sekunden. Fehlende Postfächer werden mit `Pending` und passendem Wiederanlaufbefehl gemeldet. Eine vorhandene Lizenz garantiert kein Exchange-Postfach. Liste, TUI und Export warten deshalb nie auf eine Bereitstellung. Unter `-WhatIf` entstehen weder Wartezeiten noch Schreibzugriffe.

## 8. Liste und TUI

Ergänzung vom 30. September 2026: Der Schülereinstieg verwendet denselben Populationsfilter, Renderer, Bestätigungs- und Lebenszykluscode mit `-ManageAccounts`. Die Sonderaktionen K (vorhandene CLS-Gruppe plus Klassenattribute) und L (Manager aus dem aktiven internen Schul-Pädagogikteam) stehen ausschließlich in der Schülerpopulation bereit. Schutzregeln und frische Prüfung gelten weiter. Quelldateien bleiben außerhalb der TUI und müssen vor einem späteren Sync abgeglichen werden. Der genaue Vertrag steht in der [Schüleranleitung](ENTRA-SCHUELER-SYNC.md#schüler-tui-manageaccounts).

`-List` bietet mindestens die vorhandenen Felder `DisplayName`, `UPN`, `UserId`, Exchange-Empfängertyp/-status, `SEC-A Gruppen`, `Last Login`, Besitzobjekte mit Typ/Name/ID und Schutz-/Prüfgründe. Zusätzlich erscheinen `AccountEnabled`, erkanntes Personalprofil und die Lizenzentscheidung. Nicht abrufbare Werte werden als unbekannt ausgewiesen. `-PassThru` liefert dieselben Informationen strukturiert und passwortfrei.

In der TUI wird die letzte erfolgreiche Anmeldung im Detailbereich des aktuell ausgewählten Kontos als `Last Login Timestamp (UTC)` angezeigt. Fehlende Daten erscheinen als `Kein Wert`, nicht abrufbare Anmeldedaten als `Nicht abrufbar`. Die Anzeige stammt aus dem geladenen Inventar und wird bei dessen Neuladen nach einer Aktion aktualisiert.

Zusätzlich zeigt der Detailbereich `Rolle` mit sämtlichen erkannten direkten oder transitiven `SEC-A-ROL-`-Gruppen, einschließlich unbekannter und mehrfacher Rollen, oder `Keine Rollengruppe`. Lizenz- und Klassengruppen zählen nicht als Rolle. Die Personal-TUI zeigt `Kontotyp` aus `UserType`.

Die TUI übernimmt Pfeiltasten, Leertaste, Seitenaufteilung, Namensbestätigung, Esc, Aktionsresultate und Fortschritt pro Konto. Tasten:

- `B`: aktive markierte Konten sperren und Sitzungen widerrufen.
- `U`: gesperrte markierte Konten entsperren.
- `Backspace` oder `Entf`: gesperrte markierte Konten löschen.
- `Esc`: abbrechen oder zur Auswahl zurückkehren.

Jede Aktion zeigt vor der Bestätigung alle konkreten Zielkonten. Bei gemischten Kontoständen werden passende Ziele und übersprungene Konten mit Gründen getrennt aufgeführt. Aktive Konten werden nicht in derselben Aktion gesperrt und gelöscht. Nach einer Aktion wird der betroffene Zustand frisch gelesen und die Markierung zurückgesetzt.

Rollenlose interne Cloudkonten mit persönlichem Postfach oder sicher bestätigtem Empfängerstatus `None` sind für passende Lebenszyklusaktionen zulässig. Zusätzlich beauftragt am 30. September 2026: Gäste dürfen in der Personal-TUI gesperrt und nach bereits erfolgter Sperrung gelöscht werden. Für Gäste sind außerdem eindeutige Empfänger vom Typ `MailUser` und `GuestMailUser` zulässig. Der gewählte und frisch gelesene Kontotyp sowie eine vorhandene Empfängerzuordnung müssen übereinstimmen. Gastentsperrung, Gastpflege im normalen Sync und die bisherigen gesperrte-Konten-Aktionen bleiben gesperrt. Synchronisierte Konten, Ressourcenempfänger, unklare Zuordnungen, Verzeichnisrollenträger und geschützte Gruppenmitglieder sind weiterhin nicht schreibbar. Vor jeder Aktion werden auch neu hinzugekommene Lehrer-, Ganztags-, Schüler- und SharedMailbox-Zuordnungen erneut ausgeschlossen. `MutationAllowed` gilt bei Gästen ausschließlich für TUI-Sperren und -Löschen, der Prüfhinweis `Gastkonto` bleibt erhalten.

Beim Sperren gilt `AccountEnabled = false` plus Sitzungswiderruf. Ein fehlgeschlagener Widerruf wird als Teilfehler mit bereits erfolgter Sperrung gemeldet. Entsperren setzt nur `AccountEnabled = true`. Löschen nutzt den bestehenden Graph-Löschpfad nach erneuter Einzelprüfung. Besitzobjekte werden wie heute gezeigt und unmittelbar vor der Löschung erneut geprüft. Fehler dieser Prüfung blockieren die Löschung. Ein unbekanntes Sperrdatum wird nicht aus `Last Login` abgeleitet.

Lesende Modi benötigen keine Schreibberechtigungen. Für eine echte Kontostatusänderung dokumentiert Graph `User.EnableDisableAccount.All` zusammen mit `User.Read.All` als minimale Berechtigungskombination. Die vorhandenen aktionsbezogenen Verbindungen werden entsprechend verwendet und für das Sperren um den Sitzungswiderruf ergänzt. [Microsoft Graph: Benutzer aktualisieren](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0).

## 9. Excel-Vertrag

Referenz ist die lokal vorhandene `Lehrer2026-27.xlsx`, am 29. September 2026 ausschließlich strukturell geprüft. `Lehrer.xlsx` ist der Standarddateiname im bisherigen Skript, im aktuellen Arbeitsverzeichnis liegt die Jahresdatei. Die synthetische Lehrer-Fixture enthält nur die vier Pflichtspalten und ist keine vollständige Formatvorlage.

Der Export erzeugt ein Blatt und eine Tabelle namens `Tabelle1`, Kopfzeile 1, Stil `TableStyleMedium14` mit alternierenden Zeilen, ohne zusätzliche Berichtsblätter oder Statusspalten. Spaltenfolge:

| Spalte | Überschrift | Exportwert |
|---|---|---|
| A | `Name mit Rufname` | `Nachname, Vorname` wie bei manueller Lehreranlage, bei fehlenden Namen `DisplayName` |
| B | `Vorname` | `GivenName`, fehlende Werte leer |
| C | `Nachname` | `Surname`, fehlende Werte leer |
| D | `Job` | Lesbarer Profilname bei eindeutiger Rolle (Unterstriche und CamelCase als Leerzeichen), ansonsten leer und im Exportbericht erklärt |
| E | `Passwort` | Immer leer |
| F | `EntraObjectId` | Tatsächliche Graph-Objekt-ID |
| G | `UPN` | Tatsächlicher `UserPrincipalName` |
| H | `Mail` | Tatsächlicher `Mail`-Wert, fehlende Werte leer |
| I | `Last Login Timestamp` | `signInActivity.lastSuccessfulSignInDateTime` als UTC-Text, ohne Wert leer |

Die zusätzlich beauftragte Timestamp-Spalte gilt für alle drei Einstiege. Schüler exportieren zehn Spalten, Lehrer und Personal neun. Das Format ist UTC mit Sekunden, beispielsweise `2026-01-02T10:30:00Z`. Fehlt ein Wert, bleibt die Zelle leer. Nicht abrufbare Anmeldedaten werden gewarnt und der Benutzerbestand wird ohne `signInActivity` erneut geladen. Die neue Spalte beeinflusst keine Sync-Aktion und bleibt bei einer Rückschreibung erhalten. Schüler- und Lehrerexporte fordern für diese Abfrage zusätzlich `AuditLog.Read.All` an.

Die Exportmenge entspricht exakt `-List`, einschließlich gesperrter und rollenloser Konten sowie sichtbar geschützter Konten. Weitere Listendetails erscheinen im Terminal beziehungsweise Ergebnisobjekt, damit die neun Exportspalten erhalten bleiben. Ungeklärte, geschützte oder unvollständige Zeilen werden vor einem schreibenden Reimport als solche gemeldet. Sie werden weder still übersprungen noch mit erfundenen Rollen ergänzt. Ein unbearbeiteter Gesamtexport mit solchen Zeilen ist eine Bestandsdatei und blockiert einen schreibenden Sync, bis die Auswahl bereinigt oder korrekt zugeordnet ist.

Die Referenz besitzt keine fixierten Fenster und keine verbundenen Zellen. Diese Struktur wird übernommen. Die sichtbaren Spaltenbreiten A–H sind ungefähr `21.664`, `11.164`, `15`, `11`, `11`, `21.5`, `12.5`, `35.5`. Breiten dürfen für lange Personaljobs und IDs vergrößert werden, bei erhaltenen acht Basisspalten und rechts angehängtem Login-Zeitstempel. ImportExcel/EPPlus wird als vorhandene Abhängigkeit verwendet, ohne neue Laufzeitbibliothek und ohne das auf macOS problematische `System.Drawing`-AutoSize.

Ein Export liest keine Passwörter aus einer vorhandenen Datei und generiert keine neuen. Graph-Zellwerte werden als Text geschrieben, auch bei führenden Formelzeichen. Vorhandene Exportziele und symbolische Links werden abgewiesen. Temporäre Dateien und Ziel müssen Git-ignoriert sein. Nach Schreiben einer temporären Datei werden Spalten, Tabellenstruktur, Zeilenanzahl, IDs und leere Passwortzellen geprüft. Erst danach wird das neue Ziel ohne Überschreiben bereitgestellt.

Beim späteren Sync gelten weiterhin Quellhash, Sicherung, temporäre Verifikation und exklusiver Dateizugriff nur bei tatsächlich geplanter Rückschreibung. Fremde Spalten und andere Arbeitsblätter eines bestehenden Imports bleiben erhalten.

## 10. Abnahmekriterien

1. Alle drei Ansichten liefern dieselbe Menge, auch bei rollenlosen Konten. Lehrer, OGTS, Schüler und Shared Mailboxes fehlen immer.
2. Kein Schreibpfad kann einen Ausschluss oder eine geschützte Identität über UPN, Objekt-ID, Wiederanlauf, Rollenwechsel oder eine alte TUI-Auswahl umgehen.
3. Bestehende direkte, gruppenbasierte und ausstehende Lizenzen verhindern A1-Zusatzzuweisungen. Echte Lizenzfehler blockieren.
4. Die acht gelieferten Profile sind vollständig abgebildet. Das Adminprofil bleibt gemäß Freigabe geschützt. Unbekannte Profile werden nicht erfunden.
5. ABP `$null` erzeugt keinen ABP-Vergleich und keinen ABP-Schreibparameter. Konforme Postfächer lösen keine Änderungsbestätigung aus.
6. Ein Teillisten-Update sperrt keine anderen Konten. Rollenlose Konten werden nie automatische Abgänge. Gesperrte Bestandskonten bleiben beim Reimport gesperrt.
7. Neuanlage, früher Excel-Checkpoint, Passwortausgabe und Wiederanlauf entsprechen den bestehenden Schutzmechanismen.
8. TUI-Aktionen betreffen ausschließlich bestätigte, frisch geprüfte Konten. Fortschritt wird auch nach Fehlern abgeschlossen.
9. Export entspricht dem dokumentierten Lehreraufbau. Alle Passwörter sind leer. Vorhandene Dateien bleiben erhalten. Formelzeichen bleiben Text.
10. Standardvergleich, `-List`, `-PassThru` und sämtliche `-WhatIf`-Pfade ändern keine Konten. `-WhatIf` schreibt auch keine Dateien und wartet nicht auf Postfächer.
11. Bestehende Schüler-, Lehrer- und Kontenverwaltungstests bleiben erfolgreich. Die echte TUI wurde synthetisch unter macOS im Pseudoterminal geprüft. Die Windows-Prüfung bleibt ein eigener offener Abnahmeschritt.

## 11. Bestätigte Entscheidungen und offene Integrationsprüfung

- ExterneAdmins und GlobalAdministrators werden nur angezeigt und geschützt.
- Die sprechenden Jobwerte, keine impliziten Abgangssperrungen und der Erhalt gesperrter Bestandskonten sind bestätigt und umgesetzt.
- Ein explizites `-ResumeCreation` unterscheidet den Wiederanlauf einer Teilanlage von normaler Bestandskontenpflege.
- Die tatsächliche ID der GlobalAdministrators-Gruppe, die produktiven Gruppen-/Lizenzwerte, Richtlinien und Adresslistenfilter wurden nicht live geprüft. Die eindeutige Namensauflösung bleibt deshalb die implementierte Schutzregel. Eine spätere kontrollierte Mandantenabnahme benötigt eigene Autorisierung.

Siehe [Variantenentscheidung](PERSONAL-0.3.0-BRAINSTORMING.md), [Umsetzungsplan](PERSONAL-0.3.0-PLAN.md) und [Bedienungsanleitung](ENTRA-PERSONAL-SYNC.md).
