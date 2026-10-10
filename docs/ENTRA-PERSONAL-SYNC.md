# Übrige Personalkonten verwalten (0.3.0)

Die Konten-TUI zeigt die verfügbaren Tastenbefehle als letzte Ausgabezeile unter Liste, Details und Hinweisen. Bestätigungsdialoge enden mit den Ja-/Nein-Befehlen. Backspace und Entf öffnen die Löschbestätigung, Enter hat in der Kontenliste keine Aktion.

Neue UPNs und Mailadressen verwenden `montessori-erding.de`. Angestellte erhalten Vornamensinitiale plus Nachname, bei Belegung ein längeres Vornamenspräfix. Freigegebene Rollen `Externe*` erhalten `vorname.nachname`, bei Belegung mit Zahl ab `2`. `ExterneAdmins` bleibt geschützt. Bestehende UPNs, Mailadressen, Mail-Nicknames und Aliase bleiben unverändert, einschließlich fehlender Werte. Details stehen in den [Adressregeln](ADRESSREGELN.md).

Auch [Schüler](ENTRA-SCHUELER-SYNC.md) und [Lehrer](ENTRA-LEHRER-SYNC.md) unterstützen jetzt `-ListToExcel`. Alle drei Einstiege verwenden denselben verifizierten Excel-Schreiber. Der Personalexport ergänzt seine acht Basisspalten um `Last Login Timestamp` und behält seinen Bestand und seine Schutzregeln. Schüler-/Lehrerbestände und deren unveränderte Reaktivierungsregeln stehen in den jeweiligen Anleitungen.

`Sync-PersonalEntra.ps1` verwaltet Sekretariate, Vertretungskräfte, das pädagogische Team im Kinderhaus, Unterstützung und externe Benutzer. Lehrer, CO-L, Schulpädagogen, OGTS/Ganztag, Schüler und Shared Mailboxes sind ausgeschlossen. Diese Ausschlüsse gelten auch bei verschachtelten Mitgliedschaften und bei Einzelaufrufen mit Objekt-ID oder UPN.

Die [Spezifikation](PERSONAL-0.3.0-SPEC.md) beschreibt alle gelieferten Rollen-IDs und Profilwerte. Das Skript verwendet `config/PersonalSync.psd1` und die vorhandenen Lehrer- und Schülerkonfigurationen. `ExpectedTenantId` sollte vor dem Einsatz auf den gewünschten Mandanten gesetzt werden. Graph und Exchange müssen denselben Mandanten verwenden.

## Rollen und Jobwerte

| `Job` in Excel | Interner Jobwert | `JobTitle` in Entra | Firma | Mitarbeitertyp |
|---|---|---|---|---|
| `Schule Sekretariat` | `Schule_Sekretariat` | `Schule Sekretariat` | Montessori Schule Aufkirchen | Sekretariat |
| `Schule Vertretungskräfte` | `Schule_Vertretungskräfte` | `Schule Vertretungskräfte` | Montessori Schule Aufkirchen | Vertretungskräfte |
| `Kinderhaus Pädagogisches Team` | `Kinderhaus_PädagogischesTeam` | `Pädagogisches Team` | Montessori Kinderhaus Erding | Pädagogisches Team |
| `Kinderhaus Sekretariat` | `Kinderhaus_Sekretariat` | `Sekretariat` | Montessori Kinderhaus Erding | Sekretariat |
| `Kinderhaus Vertretungskräfte` | `Kinderhaus_Vertretungskräfte` | `Vertretungskräfte` | Montessori Kinderhaus Erding | Vertretungskräfte |
| `Unterstützung` | `Unterstützung` | `Unterstützung` | Montessori Verein Landkreis Erding e.V. | Unterstützung |
| `Externe Benutzer` | `ExterneBenutzer` | `Externe Benutzer` | Montessori Verein Landkreis Erding e.V. | Externe Benutzer |

Der Personalexport ersetzt Unterstriche durch Leerzeichen und trennt CamelCase an Wortgrenzen, einschließlich Umlauten. Beim Einlesen werden die lesbaren Namen eindeutig auf die konfigurierten internen Jobwerte zurückgeführt. Bisherige Werte mit Unterstrichen oder CamelCase bleiben gültig, auch bei `-Add -Job`. Äußere und mehrfache Leerzeichen sowie Groß-/Kleinschreibung werden normalisiert. Unbekannte Namen oder mehrdeutige Zuordnungen blockieren den schreibenden Import. `JobTitle` erhält bei Neuanlage und normaler Profilpflege den lesbaren Namen. Bei Kinderhausprofilen entfällt der Präfix `Kinderhaus`, etwa `Pädagogisches Team`, da `CompanyName` bereits `Montessori Kinderhaus Erding` enthält. Die Excel-Spalte `Job` behält den Präfix für die eindeutige Zuordnung. Vorhandene technische oder mit `Kinderhaus` beginnende Jobtitel werden im Vergleich als Änderung angezeigt und erst beim ausdrücklich gestarteten Update oder bei Einzelpflege korrigiert. Auch Bestandsliste und Vergleich zeigen lesbare Jobnamen, die interne Profilzuordnung bleibt kanonisch. Die Kurzwerte `L`, `CO-L`, `PA` und `OGTS` gehören weiterhin zu `Sync-LehrerEntra.ps1`. Mehrere konkurrierende Personalprofile blockieren Änderungen. Zusätzliche Gruppen außerhalb der verwalteten Rollen bleiben erhalten.

`ExterneAdmins` (Export: `Externe Admins`) und `SEC-A-ADM-GlobalAdministrators` sind sichtbar, aber vor Änderungen geschützt. Das gelieferte Adminprofil wird nicht angewandt. Eindeutig anhand von Objekt-ID oder UPN zugeordnete Konten mit bestätigter geschützter Gruppen- oder Verzeichnisrollenmitgliedschaft werden in Compare und Update mit einem `Workbook.Review`-Hinweis übersprungen. Eine Zeile `Externe Admins` im Gesamtexport blockiert dann die übrigen gültigen Konten nicht. Die GlobalAdministrators-Gruppe wird über den eindeutigen exakten Namen aufgelöst. Fehlt eine Schutzgruppe oder ist ihre Zuordnung mehrdeutig, laufen Liste und Export mit Warnungen weiter. Schreibaktionen sind dann gesperrt.

## Voraussetzungen und Berechtigungen

Es gelten dieselben PowerShell-7- und Modulvoraussetzungen wie im [Betriebshandbuch](BETRIEB.md). Es wurden keine Laufzeitabhängigkeiten ergänzt. ImportExcel/EPPlus erzeugt die Arbeitsmappe ohne AutoSize.

Liste, Export und TUI-Bestandsaufnahme verwenden `User.Read.All`, `Directory.Read.All` und `AuditLog.Read.All`. Ist `signInActivity` nicht abrufbar, wird der Bestand ohne Anmeldedaten erneut geladen und die letzte Anmeldung als nicht abrufbar ausgewiesen. Fehler beim vollständigen Benutzer-, Gruppen-, Ausschluss- oder Exchange-Bestand brechen die Abfrage ab.

Der Sync verwendet die vorhandenen Graph-Schreibscopes `User.ReadWrite.All`, `User-Mail.ReadWrite.All`, `Group.Read.All`, `GroupMember.ReadWrite.All` und `User.RevokeSessions.All`. Im Personalmodus kommt `Directory.Read.All` für die vollständige Schutzprüfung hinzu. Die TUI fordert Schreibscopes erst bei einer bestätigten echten Aktion an. Sperren und Entsperren verwenden `User.EnableDisableAccount.All` sowie die Lesescopes. Sperren benötigt zusätzlich `User.RevokeSessions.All`, Löschen `User.ReadWrite.All`.

Gemeinsame Aufrufkonventionen sind `-WhatIf`, `-Confirm`, `-OutputFile` und `-Mail` mit den Aliasen `-UPN` und `-UserPrincipalName`. Für einzelne Sync-Phasen stehen `-CreateNewUsers`, `-UpdateUsers`, `-DisableUsers` und `-RevokeSessions` zusammen mit `-Update` bereit.

## Einzelanlage und Einzelpflege

Alle Namen und Adressen in den Beispielen sind erfunden. Führe zuerst die Vorschau aus:

```powershell
.\Sync-PersonalEntra.ps1 -Add -Vorname Lea -Nachname Testwald -Job 'Schule Sekretariat' -WhatIf
.\Sync-PersonalEntra.ps1 -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat
```

Neue Konten werden zuerst gesperrt angelegt. Das Startpasswort hat zwölf zufällige Zeichen, der erste Login verlangt einen Passwortwechsel. Nach verifizierten Attributen, Rolle und Lizenzzustand wird das neue Konto aktiviert. Ein manuelles `-Add` zeigt das Startpasswort einmal im Terminal. Es erscheint nicht im Vergleichsbericht oder Ergebnisobjekt.

Ein eindeutig zugeordnetes bestehendes Konto wird gepflegt, sein UPN und Passwort bleiben erhalten. Ein rollenloses Konto erfordert eine ausdrückliche Objekt-ID und den passenden vollständigen Namen:

```powershell
.\Sync-PersonalEntra.ps1 -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat `
  -EntraObjectId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -WhatIf
```

Ein normales `-Add` und der Sync erhalten den Sperrstatus eines Bestandskontos. Für eine gewollte Entsperrung benutze `U` in der TUI.

## Liste und Ergebnisobjekt

```powershell
.\Sync-PersonalEntra.ps1 -List
.\Sync-PersonalEntra.ps1 -List -OutputFile '.\Berichte\Personal-Bestand.txt'
$bestand = .\Sync-PersonalEntra.ps1 -PassThru
$bestand.Accounts | Select-Object DisplayName,UPN,Job,AccountEnabled,MutationAllowed
```

Der Bestand umfasst aktive und gesperrte Konten, auch ohne Rollengruppe. Er zeigt Objekt-ID, Status, Personalprofil, Lizenzentscheidung, Exchange-Empfängertyp und Prüfstatus, SEC-A-Gruppen, letzte Anmeldung, Besitzobjekte und Schutzgründe. Gäste, lokal synchronisierte Konten, Ressourcenpostfächer, Admins und unklare Exchange-Zuordnungen bleiben zur Prüfung sichtbar. Das Ergebnisobjekt enthält zusätzlich `UserType`. `MutationAllowed` beschreibt die grundsätzliche Zulässigkeit für Lebenszyklusaktionen, die aktuelle Einzelprüfung vor einer Aktion bleibt erforderlich. Bei Gästen betrifft die Zulässigkeit ausschließlich Sperren und Löschen, nicht Entsperren oder Profilpflege. `Gastkonto` ist dabei ein Prüfhinweis.

Ein fehlender Empfänger erhält `None`, ein eindeutiger `Unique`, eine doppelte Zuordnung `Ambiguous`, ein fehlender Empfängertyp `Unreadable`. Ein Sperrdatum und eine Löschfrist werden nicht aus der letzten Anmeldung abgeleitet.

## Excel-Export und optionaler Sync

```powershell
.\Sync-PersonalEntra.ps1 -ListToExcel
.\Sync-PersonalEntra.ps1 -ListToExcel -File '.\Personal-Export.xlsx' -WhatIf
.\Sync-PersonalEntra.ps1 -File '.\Personal.xlsx'
.\Sync-PersonalEntra.ps1 -File '.\Personal.xlsx' -Update -WhatIf
.\Sync-PersonalEntra.ps1 -File '.\Personal.xlsx' -Update
```

Der Export enthält dieselben Konten wie die Liste. Er erzeugt ein Blatt und eine Tabelle `Tabelle1`, `TableStyleMedium14`, Kopfzeile 1 und alternierende Zeilen. Die Spalten entsprechen der Lehrerdatei:

| Spalte | Inhalt |
|---|---|
| `Name mit Rufname` | `Nachname, Vorname`, sonst `DisplayName` |
| `Vorname`, `Nachname` | Tatsächliche Namensfelder |
| `Job` | Lesbarer Name des eindeutigen Profils, sonst leer |
| `Passwort` | Beim Export immer leer |
| `EntraObjectId`, `UPN`, `Mail` | Tatsächliche Verzeichniswerte, fehlende Mail bleibt leer |
| `Last Login Timestamp` | Rechts angehängter UTC-Zeitstempel der letzten erfolgreichen Anmeldung, sonst leer |

`Last Login Timestamp` verwendet dieselbe Quelle wie die Bestandsliste: `signInActivity.lastSuccessfulSignInDateTime`, die letzte erfolgreiche interaktive oder nicht interaktive Anmeldung. Der Wert wird als UTC-Text mit Sekunden (`2026-01-02T10:30:00Z`) exportiert. Ohne Wert bleibt die Zelle leer, das beweist keine fehlende frühere Anmeldung. Ist `signInActivity` nicht abrufbar, wird der Benutzerbestand ohne diese Eigenschaft erneut geladen und eine Warnung ausgegeben. Der Export läuft mit leeren Timestamp-Zellen weiter. Für die Abfrage sind `AuditLog.Read.All`, passende Administratorrechte und laut [Microsoft-Dokumentation](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0#example-11-get-users-including-their-last-sign-in-time) eine Entra-ID-P1-/P2-Lizenz erforderlich. Die zusätzliche Spalte wird vom Import ignoriert und bei Rückschreibungen erhalten.

Lange Jobs und IDs erhalten breitere Spalten. Zellwerte werden als Text geschrieben, Formelzeichen bleiben Text. Ein leerer Bestand ergibt eine Kopfzeile. Vorhandene Ziele und symbolische Links werden abgewiesen. Die temporäre Datei wird vor Bereitstellung wieder eingelesen und verifiziert. Ziel und temporäre Datei müssen in Git-Repositories ignoriert sein. `-WhatIf` erzeugt keine Datei und kein Zielverzeichnis.

Ein Gesamtexport ist zunächst eine Bestandsdatei. Du kannst ihn mit `-File '.\Personal.xlsx'` ohne `-Update` direkt lesend vergleichen. Leere, unbekannte oder geschützte Jobs und unvollständige Namen führen dabei zu `Workbook.Review` mit Zeilennummer und Identität. Geschützte Admins, lokal synchronisierte Konten oder unzulässige Empfänger werden als Prüfzeilen übersprungen. Eindeutig zugeordnete Entra-Gäste werden ohne Datenwarnung oder Prüfzeile ausgelassen. Die übrigen Zeilen werden normal verglichen. Prüfzeilen werden weder als Neuanlage eingeplant noch als Abgang gewertet, wenn ihre Identität im Bestand bekannt ist. Warnungen machen die Datei nicht zu einer freigegebenen Sync-Quelle.

Eindeutig über Objekt-ID oder UPN erkannte Mitgliedskonten ohne Job und ohne verwaltete Personalrolle werden in Compare und Update einschließlich WhatIf mit `Workbook.Review` übersprungen. Fehlende Namen dieser Konten blockieren die Aktualisierung gültiger Personalzeilen nicht. Die Ausnahme setzt einen vollständigen Schutzkontext und vollständig lesbare aktuelle transitive Entra-Mitgliedschaften voraus. Beide Identitätswerte müssen bei gemeinsamer Angabe dasselbe Konto bezeichnen. Eine Datei ausschließlich mit übersprungenen Konten erzeugt keine Abgänge. Sobald ein Job eingetragen wird, gilt die reguläre Importprüfung und für eine Übernahme die ausdrückliche Objekt-ID.

Für die übrigen internen und nicht eindeutig erkannten Konten gelten bei `-Update`, auch mit `-WhatIf`, weiterhin die strikten Import- und Schutzprüfungen. Leere, unbekannte oder geschützte Jobs sowie unvollständige Pflichtdaten blockieren vor Mutationen. Weise einem zulässigen rollenlosen Konto ein freigegebenes Profil ausdrücklich zu und verwende die exportierte Objekt-ID. Bestätigte geschützte Adminkonten können in der Datei bleiben, ihre Zeilen werden mit Hinweis übersprungen. Entferne andere nicht schreibbare Konten aus der Arbeitskopie. Eine andere Jobzelle hebt den Admin- oder Gastschutz nicht auf. Exceldatei und neunspaltiger Exportaufbau werden durch den lesenden Vergleich nicht verändert.

Gastzeilen werden in Compare und Update einschließlich WhatIf vor der Pflichtdaten- und Jobprüfung übersprungen. Entscheidend ist der Entra-Kontotyp `Guest` anhand von `EntraObjectId` oder `UPN`. Sind beide Werte vorhanden, müssen sie auf dasselbe Konto zeigen. Namen oder E-Mail-Muster allein geben keine Gastausnahme. Gäste bleiben in Liste und Export enthalten, erhalten dort keine Import-Prüfwarnung und werden weder als Neuanlage, Profiländerung noch als Abgang behandelt. Eine Datei ausschließlich mit Gastzeilen oder bestätigten geschützten Adminkonten führt keine Sync-Aktionen aus und liefert keine Abgänge interner Konten. Einzelpflege von Gästen bleibt gesperrt, ihre TUI-Aktionen bleiben verfügbar.

**`-Update` allein erstellt und aktualisiert. Es sperrt keine Abgänge und widerruft keine Abgangssitzungen.** Sobald ein Aktionsschalter angegeben wird, werden nur ausdrücklich gewählte Aktionen ausgeführt. Abgänge werden nur unter aktiven direkten Mitgliedern der sieben verwalteten Personalrollen ermittelt. Rollenlose Konten sind nie automatische Abgänge. Eine leere Importliste blockiert Updates.

`-Update -DisableUsers` und `-Update -RevokeSessions` sind nur mit einer vollständigen Liste der verwalteten aktiven Personalpopulation zulässig. Sie können alle aktiven Rollenmitglieder außerhalb dieser Liste betreffen. Eine Teilliste reicht dafür nicht. Diese bewusste Abgangsverarbeitung folgt den vorhandenen Fortschritts- und Berichtskonventionen.

Bei geplanten Rückschreibungen gelten Quellhash, exklusive Dateiprüfung, Sicherung und temporäre Verifikation. Eine Neuanlage schreibt Passwort, Objekt-ID, UPN und Mail früh in die Quellzeile. Fremde Spalten und Arbeitsblätter bleiben erhalten. Die Arbeitsmappe muss vor einem schreibenden Lauf in Excel geschlossen sein.

## Lizenz- und Exchange-Verhalten

Die Standardgruppe ist `SEC-A-LIC-O365A1Faculty`. Vorhandene direkte oder gruppenbasierte Lizenzen sowie ausstehende Gruppenzuweisungen werden erhalten. Die Standardgruppe wird nur bei nachweislich fehlender Lizenz ergänzt. Unvollständige Daten und echte Zuweisungsfehler blockieren die automatische Lizenzpflege. `Error = None` gilt nicht als Fehler. Eine zwischen Vergleich und Schreiben hinzugekommene Lizenz wird nochmals berücksichtigt.

Das Exchange-Profil setzt die jeweilige `MON-EXO-ABP-<Job>` und `CustomAttribute1 = <CompanyName> - <EmployeeType>`. Auditaktionen, 365 Tage AuditLogAgeLimit, 30 Tage RetainDeletedItemsFor, die vier bestehenden Standardrichtlinien und CAS-Einstellungen werden mit den vorhandenen Funktionen verglichen und nur bei Abweichungen geschrieben. Eine nicht konfigurierte ABP wird in Vorprüfung, Vergleich und Schreibparametern ausgelassen. Eine bestehende Zuweisung wird nicht entfernt.

Konforme Postfächer liefern `Compliant` ohne Änderungsbestätigung. Nach einer Kontoänderung wird das Postfach sofort und gegebenenfalls fünfmal im Abstand von 60 Sekunden geprüft. Fehlende Postfächer ergeben `Pending` mit einem Personal-Wiederanlaufbefehl. Liste, Export, TUI und `-WhatIf` warten nicht auf eine Bereitstellung.

```powershell
.\Sync-PersonalEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@example.invalid' -WhatIf
.\Sync-PersonalEntra.ps1 -Remove -UserPrincipalName 'ltestwald@example.invalid' -WhatIf
```

Der Reparaturmodus benötigt genau ein freigegebenes direktes Personalprofil je Ziel. `-Remove` akzeptiert genau ein Konto, sperrt die Anmeldung und widerruft Sitzungen. Konto, Gruppen, Lizenz und Postfach bleiben erhalten. Ein rollenloses, zulässiges Cloudkonto ist für diese gezielte Sperrung erlaubt.

## TUI

Der Detailbereich zeigt `Rolle` mit allen direkt oder transitiv erkannten `SEC-A-ROL-`-Gruppen. Mehrere und unbekannte Rollengruppen werden mit ihrem tatsächlichen Namen angezeigt. Ohne Rollengruppe erscheint `Keine Rollengruppe`. Lizenz- und Klassengruppen werden hier nicht als Rolle ausgegeben. `Kontotyp` zeigt `Member` oder `Guest`.

Der Detailbereich des aktuell mit den Pfeiltasten ausgewählten Kontos zeigt `Last Login Timestamp (UTC)`. Der Zeitstempel bezeichnet dieselbe letzte erfolgreiche Anmeldung wie im Excel-Export, beispielsweise `2026-01-02T10:30:00Z`. Ohne gespeicherten Wert erscheint `Kein Wert`, bei nicht abrufbaren Anmeldedaten `Nicht abrufbar`. Der Wert wird aus dem bereits geladenen Bestand angezeigt und bei dessen Neuladen nach einer Aktion ebenfalls aktualisiert.

```powershell
.\Sync-PersonalEntra.ps1 -ManageAccounts -WhatIf
.\Sync-PersonalEntra.ps1 -ManageAccounts
```

| Taste | Aktion |
|---|---|
| Pfeile | Auswahl bewegen |
| Leertaste | Konto markieren oder Markierung entfernen |
| `B` | Markierte aktive Konten sperren und Sitzungen widerrufen |
| `U` | Markierte gesperrte Mitgliedskonten entsperren, Gäste überspringen |
| `Backspace` oder `Entf` | Markierte gesperrte Konten löschen |
| `Esc` | Bestätigung abbrechen oder TUI verlassen |

Auch Gäste können mit `B` gesperrt und mit `Backspace` oder `Entf` gelöscht werden. Zum Löschen müssen sie bereits gesperrt sein. Zulässig sind Gäste ohne Exchange-Empfänger, mit eindeutigem `UserMailbox`/`RemoteUserMailbox` oder `MailUser`/`GuestMailUser`. Rollenlose Gäste sind ebenfalls enthalten. Ressourcenpostfächer, unklare Empfänger, lokale AD-Synchronisierung, ausgeschlossene Rollen, geschützte Adminrollen und unvollständige Schutzprüfungen blockieren weiterhin. Ein seit der Auswahl geänderter Kontotyp oder Empfängertyp blockiert die Aktion. Ein zuerst gesperrter Gast erscheint nach dem Neuladen mit dem neuen Status und kann erneut zum Löschen markiert werden.

Diese Gastfreigabe gilt nur für Sperren und Löschen über `-ManageAccounts`. `-Add`, Excel-Sync, `-Remove -UPN`, Exchange-Pflege und der bisherige gesperrte-Konten-Einstieg behalten ihre bestehenden Gastschutzregeln.

Die Bestätigung zeigt konkrete Namen und UPNs. Unpassende Status werden mit Grund übersprungen. `Y` oder `J` bestätigt, `N` oder `Esc` kehrt zurück. Vor jeder Mutation werden Identität, Mandant, Status, direkte und transitive Mitgliedschaften, Adminschutz und Exchange-Zuordnung frisch geprüft. Besitzobjekte erscheinen in der Löschbestätigung und werden unmittelbar vor der Löschung erneut abgefragt. Fehler dieser Abfrage blockieren die Löschung.

Backspace oder Entf auf einem markierten aktiven Konto zeigt den konkreten Hinweis, zuerst mit `B` zu sperren und danach erneut zum Löschen zu markieren. Dabei bleiben Markierung und Cursor erhalten und der Bestand wird nicht erneut geladen. Dasselbe gilt bei einer vollständig übersprungenen Aktion oder einer abgebrochenen Bestätigung. `B` kann nach dem Hinweis unmittelbar für die erhaltene Auswahl verwendet werden.

Nach ausgeführten Aktionen, Vorschauen oder möglichen Teilfehlern wird der Bestand neu geladen, die Markierung zurückgesetzt und jedes Ergebnis angezeigt. Bei fehlgeschlagenem Sitzungswiderruf bleibt eine bereits erfolgte Sperrung bestehen und wird als Teilfehler gemeldet. Die Fortschrittsanzeige wird auch bei Fehlern geschlossen. Umleitung der Tastatureingabe ist nicht unterstützt, verwende dann `-List`.

## Wiederanlauf und Prüfstand

Eine fehlgeschlagene Neuanlage bleibt gesperrt. Nutze den ausgegebenen Wiederanlaufbefehl nur für diese konkrete Teilanlage:

```powershell
.\Sync-PersonalEntra.ps1 -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat `
  -EntraObjectId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ResumeCreation -WhatIf
```

`-ResumeCreation` erfordert eine explizite ID, ein gesperrtes Konto, den passenden Namen und das bereits passende Neuanlageprofil. Es erzeugt kein neues Passwort und kann nach erfolgreicher Pflichtkonfiguration die Aktivierung nachholen. Ohne diesen Schalter bleibt ein gesperrtes Bestandskonto gesperrt. Wenn der Excel-Checkpoint fehlgeschlagen ist, muss das Startpasswort administrativ zurückgesetzt und sicher übergeben werden. Trage anschließend die Identität in die Arbeitsmappe nach.

Berichte brauchen einen neuen ignorierten `.txt`-Pfad. Ergebnisobjekte und Berichte enthalten keine Startpasswörter. Die tatsächlich ausgeführten lokalen Prüfungen sowie offene Windows- und Mandantenprüfungen stehen in [ABNAHME.md](ABNAHME.md). Die bereitgestellten Gruppen-IDs und Richtlinien wurden in dieser Umsetzung nicht im produktiven Mandanten geprüft.
