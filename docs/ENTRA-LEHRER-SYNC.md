# Lehrer- und Ganztagsabgleich (seit 0.2.0)

Die Konten-TUI zeigt die verfügbaren Tastenbefehle als letzte Ausgabezeile unter Liste, Details und Hinweisen. Bestätigungsdialoge enden mit den Ja-/Nein-Befehlen. Backspace und Entf öffnen die Löschbestätigung, Enter hat in der Kontenliste keine Aktion.

Für die übrigen Personalkonten ab 0.3.0 siehe [ENTRA-PERSONAL-SYNC.md](ENTRA-PERSONAL-SYNC.md). Die Jobkürzel und bisherigen Aufrufe dieses Lehrer-Einstiegs bleiben kompatibel. Der Lehrerabgleich behält seine bisherige volle Update-Auswahl und die Reaktivierung passend zugeordneter gesperrter Bestandskonten. Die abweichenden Personalregeln gelten ausschließlich für den neuen Einstieg.

`Sync-LehrerEntra.ps1` gleicht Lehrer, Co-Lehrer, Pädagogen und OGTS mit Microsoft Entra ID und Exchange Online ab. Der Aufruf ohne Aktionsparameter zeigt nur Neuzugänge, Abgänge, Änderungen, bestehende Konten und Fehler. Ein ausdrücklich gewähltes `-OutputFile` schreibt einen passwortfreien Bericht. Abgänge stammen aus der direkten Mitgliedschaft der beiden Personalrollen. Nur aktive Konten ohne Excel-Zeile erscheinen dort. Die Tabelle zeigt die Quellrolle und `AccountEnabled`.

Wie beim Schülerabgleich zeigt das Terminal den Fortschritt beim Verbinden, Vergleichen der einzelnen Personen, Prüfen bestehender Postfächer und Warten auf neue Postfächer. Ein `-Update`-Lauf zeigt den vollständigen Vergleich vor den Aktionen und danach die Aktionsergebnisse. `-OutputFile` enthält weiterhin den vollständigen Vergleich und die Aktionsergebnisse. Die Fortschrittsanzeige wird nicht in die Datei geschrieben.

Ein vollständiges `-Update` prüft Exchange auch für bestehende Personalkonten erneut. Ein konformes Postfach wird ohne Bestätigungsfrage als `Compliant` gemeldet. Nur wenn die aktuelle Prüfung eine Exchange-Abweichung findet, fragt der Lauf vor der Konfiguration. Daher kann ein Postfach auch dann eine Bestätigung benötigen, wenn sich sein Zustand seit dem ersten Vergleich geändert hat.

Leerzeilen trennen die Vergleichsbereiche im Terminal und in `-OutputFile`. Überschriften sind nur im Terminal farbig, der gespeicherte Bericht bleibt reiner Text.

Die Tabelle „Bestehende“ zeigt wie beim Schülerabgleich `NameMitRufname`, `UserId`, `DisplayName`, `UserPrincipalName` und `Job` in genau dieser Reihenfolge. `RowNumber` und lange technische Profilfelder entfallen in dieser Tabellenansicht. Das Ergebnisobjekt enthält weiterhin die übrigen Vergleichsdaten.

## Datei und Identität

Standarddatei ist `Lehrer.xlsx` im Repository-Root. `-File` akzeptiert eine andere `.xlsx`-Datei. Alte `.xls`-Dateien werden abgewiesen. Genau ein Arbeitsblatt muss die Pflichtspalten `Name mit Rufname`, `Vorname`, `Nachname` und `Job` besitzen. Erlaubte Hauptjobs sind `L`, `CO-L`, `PA` und `OGTS`, unabhängig von der Großschreibung. Mehrere Kürzel stehen kommasepariert in einer Zelle, etwa `L, PA` oder `PA, JAS`. `JAS` bezeichnet Jugendsozialarbeit und ist nur als Zusatzjob zulässig. Die Kürzel werden einheitlich geschrieben, doppelte Kürzel entfernt und in stabiler Reihenfolge in `JobTitle` übernommen. Leere Zeilen werden übersprungen. Eine leere Datei kann keinen Update-Lauf auslösen.

Optional sind `EntraObjectId`, `UPN`, `Mail` und `Passwort`. Der autorisierte Update-Lauf schreibt die tatsächliche Objekt-ID und Identität zurück. Nur bei Neuanlagen kommt das Initialpasswort in die Datei. Die Rückschreibung prüft den Quellhash, schützt die Datei vor Git-Verfolgung, sichert sie und verifiziert die temporäre Kopie vor dem Austausch. Andere Spalten und Arbeitsblätter bleiben erhalten. Produktive Arbeitsmappen und Sicherungen sind von Git ausgeschlossen.

Vor einem Update prüft das Skript, ob die Arbeitsmappe für eine geplante Neuanlage oder Identitätsrückschreibung exklusiv beschreibbar ist. Wenn der Lauf nichts in Excel zurückschreiben muss, verlangt er keine exklusive Schreibfreigabe. Die Prüfung des Quellhashs bleibt auch dann bestehen.

Die eindeutige Zuordnung erfolgt über Objekt-ID, gespeicherten UPN oder normalisierten Vor- und Nachnamen in der Vereinigung beider Personalrollen. Ein gleichnamiges Fremdkonto, Schülerkonto oder Gast wird nicht automatisch übernommen. Bestehende UPNs, Mailadressen, Mail-Nicknames und Aliase bleiben erhalten. Neue UPNs und Mailadressen verwenden `montessori-erding.de`. Das Schulteam (`L`, `CO-L`, `PA`, auch mit `JAS`) erhält zusätzlich denselben lokalen Namen als Alias unter `monteaufkirchen.com`. OGTS erhält keinen Schulalias. Hauptadresse und Alias werden gemeinsam gegen Graph- und Exchange-Adressen geprüft. Details und Umgang mit Teilfehlern stehen in den [Adressregeln](ADRESSREGELN.md).

## Profile

| JobTitle | EmployeeType | Rolle | Standardlizenzgruppe bei neuer Lizenzierung |
|---|---|---|---|
| `L` | `Pädagogisches Team` | `SEC-A-ROL-Schule_PädagogischesTeam` | `SEC-A-LIC-M365A3Faculty` |
| `CO-L` | `Pädagogisches Team` | `SEC-A-ROL-Schule_PädagogischesTeam` | `SEC-A-LIC-M365A3Faculty` |
| `PA` | `Pädagogisches Team` | `SEC-A-ROL-Schule_PädagogischesTeam` | `SEC-A-LIC-M365A3Faculty` |
| `OGTS` | `Ganztag` | `SEC-A-ROL-Ganztag` | `SEC-A-LIC-O365A1Faculty` |

Alle Hauptjobs einer Person müssen zum selben Profil gehören. `L, PA` nutzt daher eine Schulrolle und eine Lizenzentscheidung. `PA, JAS` nutzt ebenfalls das Schulprofil. `JAS` fügt keine eigene Entra-Rolle, Lizenzgruppe oder Exchange-Richtlinie hinzu. Kombinationen aus Schul- und Ganztagsprofil sowie `JAS` ohne Hauptjob werden vor Änderungen abgewiesen.

Das Schulprofil verwendet `CompanyName = Montessori Schule Aufkirchen`, `MON-EXO-ABP-Schule_PädagogischesTeam` und `CustomAttribute1 = Montessori Schule Aufkirchen - Pädagogisches Team`. Das Ganztagsprofil verwendet `CompanyName = Montessori Verein Landkreis Erding e.V.`, `MON-EXO-ABP-Ganztag` und `CustomAttribute1 = Montessori Verein Landkreis Erding e.V. - Ganztag`. Beide setzen `AgeGroup = Adult`, `ConsentProvidedForMinor = NotRequired` und `UsageLocation = DE`. Die schreibgeschützte `legalAgeGroupClassification` wird nur verglichen.

Der Abgleich korrigiert direkte verwaltete Rollen erst nach Bestätigung der Zielrolle. Klassengruppen und andere Gruppen bleiben erhalten. Eine lizenzierende Rolle oder nicht bearbeitbare dynamische Rolle blockiert einen unsicheren Rollenwechsel. Die Rolle einschließlich ihrer Lizenzzuweisungen wird unmittelbar vor dem Hinzufügen oder Entfernen frisch gelesen. Ein alter Vergleichsstand reicht dafür nicht aus.

## Lizenzschutz

Neue Konten treten der passenden Standardlizenzgruppe bei. Bestandskonten erhalten diese Gruppe nur, wenn Graph nachweislich keine Lizenz und keine ausstehende gruppenbasierte Zuweisung meldet. Jede vorhandene Direkt- oder Gruppenlizenz bleibt erhalten. Beim Wechsel von OGTS zu L oder zurück ändern sich Job und Rolle, die bisherige Lizenz bleibt bestehen. Der Bericht zeigt diese bewahrte Entscheidung. Vor einem nötigen Beitritt zur Lizenzgruppe liest das Skript die Zuweisungen erneut. Fehlerhafte oder unvollständige Lizenzdaten blockieren die automatische Entscheidung. Das Werkzeug entzieht keine Lizenz und vergibt keine direkte Benutzerlizenz.

## Bedienung

### Konten in der TUI verwalten (0.3.0)

```powershell
.\Sync-LehrerEntra.ps1 -ManageAccounts -WhatIf
.\Sync-LehrerEntra.ps1 -ManageAccounts
```

`-ManageAccounts` startet die gemeinsame Kontenverwaltung ohne Exceldatei. Die Überschrift lautet `Lehrer- und Ganztagskonten`. Die Population enthält aktive und gesperrte direkte Benutzer-Mitglieder der konfigurierten Rollen `SEC-A-ROL-Schule_PädagogischesTeam` und `SEC-A-ROL-Ganztag`, unabhängig vom gespeicherten Jobkürzel. Rein indirekte Mitglieder und rollenlose Konten gehören nicht zu diesem Einstieg. Schülerzuordnungen werden transitiv ausgeschlossen, Shared Mailboxes ebenfalls. Derselbe Rollenfilter gilt nach jeder Bestandsaktualisierung.

| Taste | Aktion |
|---|---|
| Pfeile | Konto auswählen |
| Leertaste | Konto markieren oder Markierung entfernen |
| `B` | Aktive Konten sperren und Sitzungen widerrufen |
| `U` | Gesperrte Mitgliedskonten entsperren |
| `Backspace` oder `Entf` | Gesperrte Konten löschen |
| `Esc` | Bestätigung abbrechen oder zur CLI zurückkehren |

Im Detailbereich erscheinen die tatsächlichen `SEC-A-ROL-`-Gruppen, der Kontotyp, `Last Login Timestamp (UTC)`, Exchange-Empfänger, Prüfgründe und Besitzobjekte. Fehlende Anmeldedaten werden als `Kein Wert` oder `Nicht abrufbar` ausgewiesen. Die Rollenanzeige nutzt direkte und transitive Mitgliedschaften, während die Aufnahme in die Lehrer-TUI eine direkte Lehrer-/Ganztagsmitgliedschaft benötigt.

Die Aktionen nennen die ausgewählten Namen und UPNs. `Y` oder `J` bestätigt, `N` oder Esc kehrt zurück. Vor jeder Mutation werden Objekt-ID, Name, UPN, Kontotyp, Status, Mandant, Schülerausschluss, direkte Lehrer-/Ganztagsmitgliedschaft, Adminschutz und Exchange-Zuordnung frisch geprüft. Verlorene Rollenmitgliedschaft führt zu `Skipped`. ExterneAdmins, GlobalAdministrators und Träger einer Verzeichnisrolle bleiben sichtbar und geschützt. Fehlende oder mehrdeutige Schutzgruppen blockieren Schreibaktionen. Die Schutzgruppennamen und ExterneAdmins-ID stammen aus der gemeinsam genutzten `config/PersonalSync.psd1`, die Rollen und `ExpectedTenantId` aus `config/LehrerSync.psd1`.

Persönliche Postfächer und eindeutig fehlende Exchange-Empfänger sind zulässig. Gäste in einer direkten Lehrer-/Ganztagsrolle können wie in der Personal-TUI gesperrt und nach erfolgter Sperrung gelöscht werden, auch mit eindeutigem `MailUser`/`GuestMailUser`. Gäste werden beim Entsperren übersprungen. Lokal synchronisierte Konten, Ressourcenpostfächer, unklare Empfänger sowie konkurrierende direkte Schul-/Ganztagsprofile bleiben geschützt. Ein Postfach oder eine Rolle belegt keine echte Person und kein Sperrdatum. Prüfe dies vor einer Löschung selbst.

Die Fortschrittsanzeige und Einzelresultate verwenden dieselbe Umsetzung wie die Personal-TUI. Nach einer Aktion wird der Bestand neu geladen und die Markierung entfernt. Unter `-WhatIf` erfolgen keine Mutationen oder Anfragen nach zusätzlichen Schreibscopes. Der Modus benötigt ein interaktives Terminal. Er ist nicht mit `-File`, Update-Selektoren, Einzelpflege, Export, Exchange-Reparatur oder `-OutputFile` kombinierbar. Das Ergebnisobjekt enthält `Mode = Manage`, `Actions` und `HasErrors`.

Die Bestandsaufnahme fordert `User.Read.All`, `Directory.Read.All` und `AuditLog.Read.All` an und liest Exchange im selben Mandanten. Erst bestätigte echte Aktionen fordern zusätzliche Schreibscopes an. Sperren und Entsperren verwenden `User.EnableDisableAccount.All`, Sperren zusätzlich `User.RevokeSessions.All`, Löschen `User.ReadWrite.All`.

### Bestand nach Excel exportieren (0.3.0)

```powershell
.\Sync-LehrerEntra.ps1 -ListToExcel
.\Sync-LehrerEntra.ps1 -ListToExcel -File '.\Lehrer-Bestand.xlsx'
.\Sync-LehrerEntra.ps1 -ListToExcel -File '.\Lehrer-Vorschau.xlsx' -WhatIf
```

Standardziel ist eine neue `Lehrer.xlsx` im Repository-Root. `-File` bezeichnet im Exportmodus das Ziel, keine Eingabedatei. Exportiert werden aktive und gesperrte direkte Benutzer-Mitglieder von Schulteam und Ganztag, ohne Shared Mailboxes. Ein Konto in beiden Rollen erscheint einmal. Konten ohne Postfach bleiben enthalten, verschachtelte Gruppen werden nicht als Personen exportiert. Fehler beim vollständigen Rollen-, Benutzer- oder Exchange-Bestand führen zum Abbruch.

Die neun Spalten erweitern das bestehende Lehrerformat: `Name mit Rufname`, `Vorname`, `Nachname`, `Job`, `Passwort`, `EntraObjectId`, `UPN`, `Mail`, `Last Login Timestamp`. Blatt und Tabelle heißen `Tabelle1`, Stil ist `TableStyleMedium14`. Passwörter bleiben leer, Werte werden als Text verifiziert, IDs und bestehende Identitäten bleiben erhalten. Der gemeinsame Exportcode wird auch für Schüler und übriges Personal verwendet.

`Last Login Timestamp` verwendet dieselbe Quelle wie die Bestandsliste: `signInActivity.lastSuccessfulSignInDateTime`, die letzte erfolgreiche interaktive oder nicht interaktive Anmeldung. Der Wert wird als UTC-Text mit Sekunden (`2026-01-02T10:30:00Z`) exportiert. Ohne Wert bleibt die Zelle leer, das beweist keine fehlende frühere Anmeldung. Ist `signInActivity` nicht abrufbar, wird der Benutzerbestand ohne diese Eigenschaft erneut geladen und eine Warnung ausgegeben. Der Export läuft mit leeren Timestamp-Zellen weiter. Für die Abfrage sind `AuditLog.Read.All`, passende Administratorrechte und laut [Microsoft-Dokumentation](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0#example-11-get-users-including-their-last-sign-in-time) eine Entra-ID-P1-/P2-Lizenz erforderlich. Die zusätzliche Spalte wird vom Import ignoriert und bei Rückschreibungen erhalten.

`Job` stammt aus dem vorhandenen `JobTitle` und wird mit der bestehenden Profilauflösung normalisiert. Der Job muss zu genau einer direkt zugewiesenen Schul-/Ganztagsrolle passen. Ein fehlender, ungültiger oder widersprüchlicher Job bleibt leer und erzeugt eine Warnung. Es wird kein `L` oder `OGTS` aus der Rollengruppe geraten. Prüfe solche Zeilen vor einem schreibenden Import. **Der bisherige Lehrerupdate kann enthaltene gesperrte Konten reaktivieren.** Der Export enthält keine zusätzliche Sperrstatusspalte.

Der Export fordert `User.Read.All`, `Group.Read.All` und für die Anmeldedaten `AuditLog.Read.All` an und liest Exchange im selben Mandanten. Er benötigt keine Graph-Schreibscopes und verändert keine Mandantendaten. `-ListToExcel` ist nicht mit Update-Selektoren, Einzelpflege, Exchange-Reparatur oder `-OutputFile` kombinierbar.

Bestehende Dateien und symbolische Links werden abgewiesen. Innerhalb eines Git-Repositories müssen Ziel und temporäre Kopie ignoriert und unverfolgt sein. Git und ImportExcel werden benötigt. Vor dem Verschieben wird die neue Arbeitsmappe vollständig verifiziert. Unter `-WhatIf` werden nur Bestand und geplante Dateiaktion geprüft, ohne Datei oder Zielordner anzulegen. Verwende einen neuen Dateinamen, wenn `Lehrer.xlsx` bereits vorhanden ist. Eine Fortschrittsanzeige begleitet den Export und wird auch nach einem Fehler abgeschlossen.

### Vergleich und Pflege

```powershell
# Lesender Vergleich
.\Sync-LehrerEntra.ps1 -File '.\Lehrer-2026.xlsx'

# Erstellen und Aktualisieren, ohne Abgänge
.\Sync-LehrerEntra.ps1 -File '.\Lehrer-2026.xlsx' -Update -CreateNewUsers -UpdateUsers -WhatIf

# Vollständiger Lauf mit Abgängen, erst prüfen
.\Sync-LehrerEntra.ps1 -File '.\Lehrer-2026.xlsx' -Update -WhatIf

# Einzelperson ohne Excel
.\Sync-LehrerEntra.ps1 -Add -Vorname 'Lea' -Nachname 'Testwald' -Job 'L' -WhatIf
.\Sync-LehrerEntra.ps1 -Add -Vorname 'Lea' -Nachname 'Testwald' -Job 'PA, JAS' -WhatIf
.\Sync-LehrerEntra.ps1 -Remove -UPN 'ltestwald@montessori-erding.de' -WhatIf

# Exchange-Profil aus dem vorhandenen Job und der Rolle bestimmen
.\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@montessori-erding.de' -WhatIf
```

Ein vollständiger `-Update`-Lauf erfordert eine Datei mit der gesamten aktiven Population beider Personalrollen. Andernfalls würden nicht enthaltene aktive Rollenmitglieder als Abgänge gelten. Auch Funktionskonten in diesen Rollen erscheinen als Abgänge, wenn sie aktiv sind und in Excel fehlen. Prüfe die Liste vor einem Sperrlauf. Bereits gesperrte Rollenmitglieder ohne Excel-Zeile werden ignoriert. Passt ein gesperrtes Rollenmitglied eindeutig zu einer Excel-Zeile, erscheint seine Aktivierung unter „Änderungen“. Beim Update wird dieses Konto nach Attribut- und Gruppenprüfung wieder aktiviert, ohne ein neues Konto oder Passwort anzulegen. Für Teillisten nur `-CreateNewUsers` und/oder `-UpdateUsers` auswählen. `-DisableUsers` und `-RevokeSessions` erfordern `-Update`. `-Remove` nimmt genau einen UPN und deaktiviert dieses Konto nach erneuter Prüfung der direkten Personalrolle. Außerdem widerruft es die Sitzungen. Das Konto, sein Postfach, seine Gruppen und Lizenzen bleiben bestehen.

Eine Neuanlage erstellt das fehlende Konto zunächst deaktiviert. Beim Dateiimport sichert das Skript unmittelbar danach Passwort und Identität in Excel. Anschließend prüft es Attribute und Gruppen und aktiviert das Konto erst nach erfolgreicher Pflichtkonfiguration. Manuelles `-Add` benötigt keine Arbeitsmappe. Nur echte Neuanlagen erhalten ein Passwort mit genau zwölf kryptografisch zufälligen Zeichen und `ForceChangePasswordNextSignIn = true`. Bei manuellem Add erscheint dieses Passwort einmal im Terminal, bei Dateiimport in der geschützten Arbeitsmappe. Berichte und Ergebnisobjekte enthalten es nicht. Ein Bestandskonto erhält weder ein neues Passwort noch ein neues Passwortprofil. Die Schülerregel bleibt bei zwölf Zeichen und ohne erzwungenen Erstwechsel.

## Exchange und Betrieb

Je nach Rolle verwendet der Abgleich die passende Adressbuchrichtlinie und `CustomAttribute1`. Auditaktionen werden ergänzend gesetzt. Die vorhandenen Werte für `RetainDeletedItemsFor`, Rollen-, Freigabe-, Aufbewahrungs- und OWA-Richtlinien sowie CAS-Protokolle werden übernommen. Ein fehlendes oder noch nicht repliziertes Postfach wird sofort und danach höchstens dreimal im Abstand von 150 Sekunden (2,5 Minuten) geprüft. Das sind vier Prüfungen und höchstens 7,5 Minuten Wartezeit. Der Batch wartet nur einmal pro Runde. Die Werte stehen unter `Exchange` in `config/LehrerSync.psd1` und gelten für Anlage, Bestandsabgleich und `-ConfigureExchangeOnlineOnly`.

Eindeutige Exchange-Meldungen über ein nicht gefundenes Empfängerobjekt werden auch bei generischen Ausnahmeinformationen als Bereitstellungsverzögerung erkannt. Das gilt beim Lesen von Postfach und CAS-Zustand sowie beim Schreiben und Rücklesen der Konfiguration. Nach ausgeschöpften Wiederholungen folgt `Pending` mit Reparaturbefehl. Berechtigungs-, Verbindungs-, Parameterfehler und bestätigte verbleibende Konfigurationsabweichungen bleiben sofortige Fehler. `-WhatIf` prüft einmal und wartet oder schreibt nicht.

`AuditLogAgeLimit = 365` bestätigt keine tatsächliche 365-Tage-Auditaufbewahrung. Diese wird in Microsoft Purview verwaltet. Die Postfachprüfung ersetzt weder die Kontrolle der Adresslistenfilter noch der tatsächlichen A1-/A3-Lizenzverarbeitung.

Vor produktiven Schreibläufen sind die eindeutigen Gruppen, ihre tatsächlichen SKUs, die Exchange-Richtlinien, Adresslistenfilter und mögliche dynamische Lizenzregeln im Mandanten zu prüfen. Die vorhandenen Gruppenskripte verwenden dieselben Werte für `EmployeeType`, `CompanyName` und `CustomAttribute1`. Bei einem Rollenwechsel darf ein paralleler Alt-Skriptlauf keinen Zwischenzustand auswerten.

## Sicherheitsprüfungen und Wiederanlauf

Wenn Kontoanlage, Attribute, Rollen und Lizenzzuweisung erfolgreich waren und nur Exchange fehlgeschlagen ist, verwende den vorhandenen UPN für die gezielte Reparatur. Der Befehl legt kein weiteres Konto an und ändert kein Passwort:

```powershell
.\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@example.invalid' -WhatIf
.\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@example.invalid'
```

Ersetze den Beispiel-UPN durch das betroffene Konto. Die Vorschau zeigt die aktuellen Abweichungen, der zweite Aufruf konfiguriert nach Bestätigung das vorhandene Postfach. Wiederholte Aufrufe ändern nur noch abweichende Exchange-Einstellungen. Ist das Postfach weiterhin nicht verfügbar, prüfe die tatsächliche Lizenzverarbeitung und wiederhole anschließend die Reparatur. Der Reparaturmodus benötigt einen eindeutigen vorhandenen Lehrerjob und die passende direkte Lehrer-/Ganztagsrolle.

Sind auch frühere Entra-Schritte fehlgeschlagen, behalte die zurückgeschriebenen Objekt-IDs, UPNs und Passwortzellen in der Arbeitsmappe und nutze den selektiven Wiederanlauf. Diese Auswahl führt keine Abgangssperrungen oder Sitzungswiderrufe aus:

```powershell
.\Sync-LehrerEntra.ps1 -File '.\Lehrer.xlsx' -Update -CreateNewUsers -UpdateUsers -WhatIf
.\Sync-LehrerEntra.ps1 -File '.\Lehrer.xlsx' -Update -CreateNewUsers -UpdateUsers
```

Der Graph- und Exchange-Mandant müssen übereinstimmen. Eine fehlende Tenant-ID oder mehrere aktive Exchange-Verbindungen führen zum Abbruch. Die Schülerrolle stammt aus `config/SchuelerSync.psd1`. Fehlt diese Gruppe im Bestand oder stimmt ihr Name nicht, wird kein Personalabgleich ausgeführt. Vor einem Lehrerupdate werden die aktuellen direkten Mitgliedschaften erneut geprüft, insbesondere ein möglicher Wechsel in die Schülerrolle.

Bei Dateiimport werden Passwort, Objekt-ID, UPN und Mail unmittelbar nach der gesperrten Kontoanlage sicher in Excel geschrieben. Erst danach folgen Attributprüfung, Rollen, Lizenzgruppe und Aktivierung. Scheitert die weitere Konfiguration, kann der ausgegebene selektive Wiederanlaufbefehl das gespeicherte, gesperrte Konto fortsetzen. Ein Konto außerhalb der Personalrollen wird dabei nur über eine explizite Objekt-ID, eine vorhandene Passwortzelle und ein passendes Personalprofil akzeptiert. Namensübereinstimmung allein reicht nicht aus.

Scheitert bereits die Excel-Sicherung, bleibt das neue Konto gesperrt. Die Fehlermeldung enthält seine Objekt-ID und einen manuellen Wiederanlaufbefehl. Vor dessen Ausführung muss das Startpasswort administrativ zurückgesetzt und sicher übergeben werden. Danach die Objekt-ID in Excel nachtragen. Passwörter werden bei Dateiimport auch im Fehlerfall nicht ins Terminal oder in Berichte geschrieben.

Das Nachtragen bestehender Identitäten in Excel hat eine eigene Bestätigung. Abgelehnte oder fehlgeschlagene Kontoänderungen werden nicht nachgetragen. Manuelles `-Add` zeigt keine unbeteiligten Personen als Abgänge.

`-OutputFile` akzeptiert neue `.txt`-Dateien, die innerhalb eines Git-Repositories ignoriert sein müssen. Bestehende Dateien und symbolische Links, auch in Elternordnern, sind ausgeschlossen. Git muss zur Pfadprüfung verfügbar sein. Verwende beispielsweise `Berichte/Lehrer-2026-09-27-1430.txt`. Unter `-WhatIf` entsteht keine Datei.
