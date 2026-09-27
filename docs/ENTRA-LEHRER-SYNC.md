# Personalabgleich 0.2.0

`Sync-LehrerEntra.ps1` gleicht Lehrer, Co-Lehrer, Pädagogen und OGTS mit Microsoft Entra ID und Exchange Online ab. Der Aufruf ohne Aktionsparameter zeigt nur Neuzugänge, Abgänge, Änderungen, bestehende Konten und Fehler. Ein ausdrücklich gewähltes `-OutputFile` schreibt einen passwortfreien Bericht. Abgänge stammen aus der direkten Mitgliedschaft der beiden Personalrollen. Nur aktive Konten ohne Excel-Zeile erscheinen dort. Die Tabelle zeigt die Quellrolle und `AccountEnabled`.

Wie beim Schülerabgleich zeigt das Terminal den Fortschritt beim Verbinden, Vergleichen der einzelnen Personen, Prüfen bestehender Postfächer und Warten auf neue Postfächer. Ein `-Update`-Lauf zeigt den vollständigen Vergleich vor den Aktionen und danach die Aktionsergebnisse. `-OutputFile` enthält weiterhin den vollständigen Vergleich und die Aktionsergebnisse. Die Fortschrittsanzeige wird nicht in die Datei geschrieben.

Ein vollständiges `-Update` prüft Exchange auch für bestehende Personalkonten erneut. Ein konformes Postfach wird ohne Bestätigungsfrage als `Compliant` gemeldet. Nur wenn die aktuelle Prüfung eine Exchange-Abweichung findet, fragt der Lauf vor der Konfiguration. Daher kann ein Postfach auch dann eine Bestätigung benötigen, wenn sich sein Zustand seit dem ersten Vergleich geändert hat.

Leerzeilen trennen die Vergleichsbereiche im Terminal und in `-OutputFile`. Überschriften sind nur im Terminal farbig, der gespeicherte Bericht bleibt reiner Text.

Die Tabelle „Bestehende“ zeigt wie beim Schülerabgleich `NameMitRufname`, `UserId`, `DisplayName`, `UserPrincipalName` und `Job` in genau dieser Reihenfolge. `RowNumber` und lange technische Profilfelder entfallen in dieser Tabellenansicht. Das Ergebnisobjekt enthält weiterhin die übrigen Vergleichsdaten.

## Datei und Identität

Standarddatei ist `Lehrer.xlsx` im Repository-Root. `-File` akzeptiert eine andere `.xlsx`-Datei. Alte `.xls`-Dateien werden abgewiesen. Genau ein Arbeitsblatt muss die Pflichtspalten `Name mit Rufname`, `Vorname`, `Nachname` und `Job` besitzen. Erlaubte Hauptjobs sind `L`, `CO-L`, `PA` und `OGTS`, unabhängig von der Großschreibung. Mehrere Kürzel stehen kommasepariert in einer Zelle, etwa `L, PA` oder `PA, JAS`. `JAS` bezeichnet Jugendsozialarbeit und ist nur als Zusatzjob zulässig. Die Kürzel werden einheitlich geschrieben, doppelte Kürzel entfernt und in stabiler Reihenfolge in `JobTitle` übernommen. Leere Zeilen werden übersprungen. Eine leere Datei kann keinen Update-Lauf auslösen.

Optional sind `EntraObjectId`, `UPN`, `Mail` und `Passwort`. Der autorisierte Update-Lauf schreibt die tatsächliche Objekt-ID und Identität zurück. Nur bei Neuanlagen kommt das Initialpasswort in die Datei. Die Rückschreibung prüft den Quellhash, schützt die Datei vor Git-Verfolgung, sichert sie und verifiziert die temporäre Kopie vor dem Austausch. Andere Spalten und Arbeitsblätter bleiben erhalten. Produktive Arbeitsmappen und Sicherungen sind von Git ausgeschlossen.

Vor einem Update prüft das Skript, ob die Arbeitsmappe für eine geplante Neuanlage oder Identitätsrückschreibung exklusiv beschreibbar ist. Wenn der Lauf nichts in Excel zurückschreiben muss, verlangt er keine exklusive Schreibfreigabe. Die Prüfung des Quellhashs bleibt auch dann bestehen.

Die eindeutige Zuordnung erfolgt über Objekt-ID, gespeicherten UPN oder normalisierten Vor- und Nachnamen in der Vereinigung beider Personalrollen. Ein gleichnamiges Fremdkonto, Schülerkonto oder Gast wird nicht automatisch übernommen. Bestehende UPNs bleiben erhalten. Neue UPNs werden erst nach Abgleich mit Graph- und Exchange-Adressen reserviert.

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
.\Sync-LehrerEntra.ps1 -Remove -UPN 'ltestwald@monteaufkirchen.com' -WhatIf

# Exchange-Profil aus dem vorhandenen Job und der Rolle bestimmen
.\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN 'ltestwald@monteaufkirchen.com' -WhatIf
```

Ein vollständiger `-Update`-Lauf erfordert eine Datei mit der gesamten aktiven Population beider Personalrollen. Andernfalls würden nicht enthaltene aktive Rollenmitglieder als Abgänge gelten. Auch Funktionskonten in diesen Rollen erscheinen als Abgänge, wenn sie aktiv sind und in Excel fehlen. Prüfe die Liste vor einem Sperrlauf. Bereits gesperrte Rollenmitglieder ohne Excel-Zeile werden ignoriert. Passt ein gesperrtes Rollenmitglied eindeutig zu einer Excel-Zeile, erscheint seine Aktivierung unter „Änderungen“. Beim Update wird dieses Konto nach Attribut- und Gruppenprüfung wieder aktiviert, ohne ein neues Konto oder Passwort anzulegen. Für Teillisten nur `-CreateNewUsers` und/oder `-UpdateUsers` auswählen. `-DisableUsers` und `-RevokeSessions` erfordern `-Update`. `-Remove` nimmt genau einen UPN und deaktiviert dieses Konto nach erneuter Prüfung der direkten Personalrolle. Außerdem widerruft es die Sitzungen. Das Konto, sein Postfach, seine Gruppen und Lizenzen bleiben bestehen.

Eine Neuanlage erstellt das fehlende Konto zunächst deaktiviert. Beim Dateiimport sichert das Skript unmittelbar danach Passwort und Identität in Excel. Anschließend prüft es Attribute und Gruppen und aktiviert das Konto erst nach erfolgreicher Pflichtkonfiguration. Manuelles `-Add` benötigt keine Arbeitsmappe. Nur echte Neuanlagen erhalten ein Passwort mit genau zwölf kryptografisch zufälligen Zeichen und `ForceChangePasswordNextSignIn = true`. Bei manuellem Add erscheint dieses Passwort einmal im Terminal, bei Dateiimport in der geschützten Arbeitsmappe. Berichte und Ergebnisobjekte enthalten es nicht. Ein Bestandskonto erhält weder ein neues Passwort noch ein neues Passwortprofil. Die Schülerregel bleibt bei zwölf Zeichen und ohne erzwungenen Erstwechsel.

## Exchange und Betrieb

Je nach Rolle verwendet der Abgleich die passende Adressbuchrichtlinie und `CustomAttribute1`. Auditaktionen werden ergänzend gesetzt. Die vorhandenen Werte für `RetainDeletedItemsFor`, Rollen-, Freigabe-, Aufbewahrungs- und OWA-Richtlinien sowie CAS-Protokolle werden übernommen. Ein fehlendes Postfach wird einmal sofort und nach autorisierten Änderungen höchstens fünfmal im Abstand von 60 Sekunden geprüft. Der Batch wartet nur einmal pro Runde. Ein weiterhin fehlendes Postfach erhält einen Reparaturbefehl, ein Schreibfehler wird als Fehler gemeldet.

`AuditLogAgeLimit = 365` bestätigt keine tatsächliche 365-Tage-Auditaufbewahrung. Diese wird in Microsoft Purview verwaltet. Die Postfachprüfung ersetzt weder die Kontrolle der Adresslistenfilter noch der tatsächlichen A1-/A3-Lizenzverarbeitung.

Vor produktiven Schreibläufen sind die eindeutigen Gruppen, ihre tatsächlichen SKUs, die Exchange-Richtlinien, Adresslistenfilter und mögliche dynamische Lizenzregeln im Mandanten zu prüfen. Die vorhandenen Gruppenskripte verwenden dieselben Werte für `EmployeeType`, `CompanyName` und `CustomAttribute1`. Bei einem Rollenwechsel darf ein paralleler Alt-Skriptlauf keinen Zwischenzustand auswerten.

## Sicherheitsprüfungen und Wiederanlauf

Der Graph- und Exchange-Mandant müssen übereinstimmen. Eine fehlende Tenant-ID oder mehrere aktive Exchange-Verbindungen führen zum Abbruch. Die Schülerrolle stammt aus `config/SchuelerSync.psd1`. Fehlt diese Gruppe im Bestand oder stimmt ihr Name nicht, wird kein Personalabgleich ausgeführt. Vor einem Lehrerupdate werden die aktuellen direkten Mitgliedschaften erneut geprüft, insbesondere ein möglicher Wechsel in die Schülerrolle.

Bei Dateiimport werden Passwort, Objekt-ID, UPN und Mail unmittelbar nach der gesperrten Kontoanlage sicher in Excel geschrieben. Erst danach folgen Attributprüfung, Rollen, Lizenzgruppe und Aktivierung. Scheitert die weitere Konfiguration, kann der ausgegebene selektive Wiederanlaufbefehl das gespeicherte, gesperrte Konto fortsetzen. Ein Konto außerhalb der Personalrollen wird dabei nur über eine explizite Objekt-ID, eine vorhandene Passwortzelle und ein passendes Personalprofil akzeptiert. Namensübereinstimmung allein reicht nicht aus.

Scheitert bereits die Excel-Sicherung, bleibt das neue Konto gesperrt. Die Fehlermeldung enthält seine Objekt-ID und einen manuellen Wiederanlaufbefehl. Vor dessen Ausführung muss das Startpasswort administrativ zurückgesetzt und sicher übergeben werden. Danach die Objekt-ID in Excel nachtragen. Passwörter werden bei Dateiimport auch im Fehlerfall nicht ins Terminal oder in Berichte geschrieben.

Das Nachtragen bestehender Identitäten in Excel hat eine eigene Bestätigung. Abgelehnte oder fehlgeschlagene Kontoänderungen werden nicht nachgetragen. Manuelles `-Add` zeigt keine unbeteiligten Personen als Abgänge.

`-OutputFile` akzeptiert neue `.txt`-Dateien, die innerhalb eines Git-Repositories ignoriert sein müssen. Bestehende Dateien und symbolische Links, auch in Elternordnern, sind ausgeschlossen. Git muss zur Pfadprüfung verfügbar sein. Verwende beispielsweise `Berichte/Lehrer-2026-09-27-1430.txt`. Unter `-WhatIf` entsteht keine Datei.
