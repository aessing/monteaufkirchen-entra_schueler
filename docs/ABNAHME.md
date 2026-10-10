# Abnahmestand und VM-Protokoll

## Ergänzung 0.3.0: Kinderhauspräfix aus Entra-Jobtiteln entfernt, 10. Oktober 2026

Die drei Kinderhausprofile erhalten bei Neuanlage und normaler Profilpflege die Entra-Jobtitel `Pädagogisches Team`, `Sekretariat` und `Vertretungskräfte`. `CompanyName` bleibt `Montessori Kinderhaus Erding`. Die Excel-Spalte `Job`, Bestandsliste und Profilzuordnung behalten die eindeutigen Jobnamen einschließlich `Kinderhaus`. Vorhandene technische oder lesbare Titel mit Kinderhauspräfix erscheinen im Vergleich als Änderung und werden erst beim gestarteten Update beziehungsweise bei Einzelpflege korrigiert. Lesende Aufrufe ändern keine Konten.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| TDD | Vor Umsetzung die drei Kinderhausfälle fehlgeschlagen, vier andere Profile bestanden. Nach Umsetzung alle sieben fokussierten Fälle bestanden |
| Personalregressionen | Neuanlage sowie Update technischer und bisher lesbarer Titel geprüft. Synthetische Graph-Schreibparameter, unveränderte Firma, Rollen und Adressen sowie zweiter Vergleich ohne Änderungen geprüft |
| Gesamtprüfung | 663 Tests erkannt, 662 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. Discovery erfolgreich, Parser und statische Analyse ohne ungeklärte Befunde. Tatsächliche isolierte CLI-Aufrufe der Suite bestanden. Git-Diff ohne Whitespace-Befunde |
| Dokumentation | README, Personal-Anleitung, Betriebs-/Supportanleitung, Spezifikation, Umsetzungsplan, Testbeschreibung, Changelog und Abnahme aktualisiert |

Es wurden ausschließlich synthetische Konten und Fernbefehle geprüft. Keine produktiven Konten oder Arbeitsmappen geändert. Frühere Nachweise unten dokumentieren den Zustand vor dieser Korrektur.

## Ergänzung 0.3.0: Lesbare Personal-Jobtitel in Entra und Berichten, 10. Oktober 2026

Die vorherige Formatierung betraf die Jobzelle im Personalexport, das tatsächliche Entra-`JobTitle` erhielt weiterhin den technischen Profilwert. Neuanlage und normale Personal-Profilpflege schreiben jetzt den lesbaren Namen mit Leerzeichen statt Unterstrichen und CamelCase, etwa `Kinderhaus Pädagogisches Team` und `Externe Benutzer`. Technische Bestandstitel erscheinen im lesenden Vergleich als Änderung und werden erst beim gestarteten Update beziehungsweise bei Einzelpflege korrigiert. Bestandsliste und Vergleich zeigen ebenfalls lesbare Jobnamen. Interne Profilwerte, Rollen-IDs, Exchange-Profile und Lehrerjobkürzel bleiben erhalten. Bereits lesbare Titel erzeugen keine erneute Jobtiteländerung.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| TDD | Vor Umsetzung acht neue Fälle fehlgeschlagen, der bereits identische Titel `Unterstützung` bestand. Danach alle neun fokussierten Regressionen bestanden |
| Personalregressionen | Alle sieben freigegebenen Profile auf lesbares JobTitle bei Neuanlage und Update geprüft. Synthetische Graph-Schreibparameter, anschließender Vergleich ohne Änderungen und fehlende Adressmutationen geprüft. Bestandsliste und Vergleich bleiben lesbar, interne Jobschlüssel unverändert |
| Gesamtprüfung | 663 Tests erkannt, 662 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. Discovery erfolgreich, Parser und statische Analyse ohne ungeklärte Befunde. Tatsächliche isolierte Personal-, Lehrer- und Schüler-CLI-Aufrufe der Suite bestanden. Git-Diff ohne Whitespace-Befunde |
| Dokumentation | README, Personal-Anleitung, Betriebs-/Supportanleitung, Spezifikation, Umsetzungsplan, Testbeschreibung, Changelog und Abnahme aktualisiert |

Es wurden ausschließlich synthetische Konten und Fernbefehle geprüft. Keine produktiven Konten oder Arbeitsmappen geändert. Frühere Nachweise unten dokumentieren den Zustand vor dieser Korrektur.

## Ergänzung 0.3.0: Lesbare Jobnamen im Personalexport, 10. Oktober 2026

Der Personalexport ersetzt Unterstriche durch Leerzeichen und trennt CamelCase an Wortgrenzen, einschließlich Umlauten. Beispiele: `Schule_Sekretariat` wird `Schule Sekretariat`, `Kinderhaus_PädagogischesTeam` wird `Kinderhaus Pädagogisches Team`, `ExterneBenutzer` wird `Externe Benutzer`. Beim Einlesen und bei Einzelpflege werden lesbare Namen sowie bisherige interne Werte anhand der konfigurierten Jobs eindeutig auf das kanonische Profil zurückgeführt. Groß-/Kleinschreibung und mehrfache Leerzeichen werden normalisiert. Mehrdeutige Anzeigenamen werden auch bei anders geschriebenen internen Werten abgewiesen. Der kanonische `JobTitle`, Gruppen-IDs und Exchange-Profilwerte bleiben erhalten. `Externe Admins` bleibt geschützt und wird im lesenden Vergleich als kanonisch identifizierte Prüfzeile erfasst.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| TDD und fokussierte Tests | Regressionen für lesbare Namen vor der Umsetzung fehlgeschlagen. Nach Korrektur der wiederverwendeten Testdateinamen 272 Personal-/Export-/Lehrertests bestanden. Vier gezielt ausgeführte Mehrdeutigkeitsfälle bestanden, zwei davon zuvor mit anderer Groß-/Kleinschreibung reproduziert |
| Native XLSX-Prüfung | Alle acht konfigurierten Jobs exportiert und rückgelesen. Lesbare Zellen, Umlaute, leere Jobs, kanonische Werte, zusätzliche Leerzeichen, Groß-/Kleinschreibung, wiederholte identische Jobs, Schutzprofile und unveränderte Inventarwerte geprüft |
| Tatsächlicher Skriptaufruf | Isolierter Personal-CLI-Export mit `Schule Sekretariat` in der Jobzelle. Anschließender Compare und Update-WhatIf mit Exitcode 0, HasErrors false und unverändertem Quelldateihash. Gesamtexport mit geschützten, rollenlosen und Gastzeilen bleibt lesend vergleichbar, schreibender Import wird weiterhin abgewiesen |
| Gesamtprüfung | 654 Tests erkannt, 653 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. Discovery erfolgreich, Parser und statische Analyse ohne ungeklärte Befunde. Git-Diff ohne Whitespace-Befunde |
| Dokumentation | Job-Mappingtabelle, Personal-Anleitung, README, Betriebs-/Supportanleitung, Spezifikation, Umsetzungsplan, Testbeschreibung, Changelog und Abnahme aktualisiert |

Alle Konten und Fernbefehle in den Tests waren synthetisch. Bestehende produktive Arbeitsmappen und Konten wurden nicht verändert, kein Push. Ein tatsächlicher Windows-Terminal- oder Mandantentest war nicht Bestandteil dieser lokalen Prüfung.

## Ergänzung 0.3.0: Backspace und Entf zum Löschen, 10. Oktober 2026

Backspace und Entf öffnen in allen Konten-TUIs die bestehende Löschbestätigung für markierte gesperrte Konten. Enter startet in keiner Kontenliste mehr eine Aktion. Die Schüler-Zielauswahl übernimmt weiterhin mit Enter. Y/J bestätigt, N/Esc bricht ab. Befehlszeilen, Statushinweise und aktuelle Anleitungen verwenden die neue Belegung. Frühere Prüfstände unten behalten ihre damalige Tastaturbelegung.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| TDD | Vor Umsetzung 26 fehlgeschlagene fokussierte Fälle, danach alle 238 TUI-/Personaltests bestanden. Beide Löschpfade, beide Tasten, Enter ohne Aktion in allen drei Populationen, Nein mit erhaltener Auswahl und aktiver Gast mit Sperrhinweis geprüft |
| Tatsächlicher Terminalaufruf | Isolierter Personal-Einstieg mit gesperrtem synthetischem Konto im macOS-Pseudoterminal. Markieren, Enter ohne Aktion, Backspace und Entf mit Löschbestätigung, jeweils N zum Abbrechen, anschließend Backspace und Y mit WhatIf sowie Esc geprüft. Exitcode 0, genau ein WhatIf-Ergebnis und HasErrors false |
| Gesamtprüfung | 630 Tests erkannt, 629 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. Discovery erfolgreich, Parser und statische Analyse ohne ungeklärte Befunde. Git-Diff ohne Whitespace-Befunde |
| Dokumentation | README, Changelog, Betriebs-/Supportanleitung, vier TUI-Anleitungen, Spezifikation, Umsetzungsplan, Testbeschreibung und Abnahme aktualisiert |

Keine produktiven Konten oder Arbeitsmappen verändert, kein Push. Tastaturverhalten unter Windows wurde nicht im tatsächlichen Terminal geprüft.

## Ergänzung 0.3.0: Tastenbefehle in der letzten TUI-Zeile, 10. Oktober 2026

Alle Kontenansichten zeigen ihre verfügbaren Tastenbefehle als letzte Ausgabezeile unter Liste, Details und Hinweisen. Die gemeinsame Ansicht deckt gesperrte Konten, Personal, Lehrer/Ganztag und Schüler ab. Schüler ergänzen K und L in derselben Befehlszeile. Klassen-/Klassenlehrer-Zielauswahl zeigt die Navigation ebenfalls unten. Hinweise zu übersprungenen Konten erscheinen in Sperr-/Entsperr-/Löschbestätigungen vor den Ja-/Nein-Befehlen.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| TDD und fokussierte Tests | Neue Reihenfolgetests vor den jeweiligen Änderungen fehlgeschlagen. Abschließend 226 TUI-/Personaltests bestanden, darunter leere Listen, Zielauswahl und Bestätigungen mit geschützten Konten |
| Tatsächlicher Terminalaufruf | Isolierter Schüler-Einstieg mit synthetischen Daten und WhatIf im macOS-Pseudoterminal. Kontenliste, Markieren, K-Zielauswahl, Esc zur erhaltenen Auswahl und Esc zum Beenden geprüft. Befehle jeweils zuletzt sichtbar, Exitcode 0 und HasErrors false |
| Gesamtprüfung | 618 Tests erkannt, 617 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. Discovery erfolgreich, Parser und statische Analyse ohne ungeklärte Befunde. Git-Diff ohne Whitespace-Befunde |
| Dokumentation | README, Changelog, vier TUI-Anleitungen, Testbeschreibung und Abnahme aktualisiert |

Keine produktiven Konten oder Arbeitsmappen verändert, kein Push. Die neue Ausgabe wurde unter macOS geprüft, nicht auf einem echten Windows-Terminal.

## Ergänzung 0.3.0: Adressregeln für Neuanlagen, 10. Oktober 2026

Schüler behalten `monteaufkirchen.com` als Domain für neue UPNs und Mailadressen. Schulteam, OGTS und übriges Personal verwenden bei Neuanlage `montessori-erding.de`. Das Schulteam (`L`, `CO-L`, `PA`, auch mit `JAS`) erhält zusätzlich denselben lokalen Namen als Alias unter `monteaufkirchen.com`. OGTS erhält keinen Schulalias. Angestellte verwenden eine Vornamensinitiale mit wachsendem Präfix bei Kollisionen, freigegebene Rollen `Externe*` verwenden `vorname.nachname` mit Zahlenfallback. ExterneAdmins bleibt geschützt.

Hauptadresse und Lehrer-Alias werden gemeinsam gegen Graph- und Exchange-Belegungen geprüft und im Batch reserviert. Neue Exchange-Adressen werden unter Erhalt weiterer SMTP-/X500-Adressen geschrieben und nachgelesen. Die Schüler-Postfachschleife erhält die Adressregel nur für im aktuellen Lauf erfolgreich neu angelegte Konten. Bestehende UPNs, Mailadressen, Mail-Nicknames und Aliase bleiben unverändert. Der Schüler-Schreibpfad ignoriert diese Felder auch bei ausdrücklich angelieferten Änderungsplänen. Normale Exchange-Reparatur verändert keine Adressen. Der gezielte Abschluss eines fehlenden Alias nach einer Teilanlage steht in den [Adressregeln](ADRESSREGELN.md).

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| TDD vor Umsetzung | 17 fehlgeschlagene Adressregressionen. Danach jeweils ein weiterer fehlgeschlagener Test für den Schüler-Schreibschutz und die SMTP-Regel im gemischten Schülerbatch vor der jeweiligen Umsetzung |
| Fokussierte Prüfung | 27 Entra-Mutationstests bestanden. Anschließend 74 Exchange- und Orchestrierungstests bestanden, einschließlich gemischtem Schülerbatch |
| Vollständige lokale Verifikation | 607 Tests erkannt, 606 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. PowerShell-Parser erfolgreich, keine ungeklärten Analyzer-Befunde |
| CLI und Dateien | Die Gesamtsuite umfasst die vorhandenen tatsächlichen isolierten Skriptaufrufe sowie synthetische Graph-/Exchange-Mutationen und XLSX-Roundtrips. Keine produktiven Arbeitsmappen geändert |
| Dokumentation | README, Changelog, Schüler-/Lehrer-/Personalanleitung, Betrieb, Spezifikation, Plan, Variantenentscheidung, Testing, Contribution, Security und Support aktualisiert. Neue gemeinsame Adressregeln. Lokale Dateilinks in 17 Markdown-Dateien ohne fehlende Ziele, vorhandene relative GitHub-Navigationslinks separat eingeordnet. Git-Diff ohne Whitespace-Befunde |

Der vorhandene Arbeitsstand auf `codex/0.3.0-personal` wurde erhalten und ergänzt. Keine produktiven Konten geändert, kein Push und keine Veröffentlichung. Domainfreigabe, echte Exchange-Bereitstellung und Replikation sowie Windows-Dateisperren wurden nicht live geprüft. Frühere Prüfergebnisse folgen unten.

## Ergänzung 0.3.0: Enter auf aktivem Gast in der Personal-TUI, 30. September 2026

Der Betreiber bestätigte, dass der ausgewählte Gast aktiv war. Der bestehende Löschpfad verlangt ein gesperrtes Konto und übersprang die Auswahl deshalb vor der Bestätigung. Die TUI lud auch nach diesem vollständig übersprungenen Vorgang den gesamten Bestand erneut und entfernte die Markierung. Der allgemeine Statushinweis erklärte den nötigen nächsten Schritt nicht.

Aktive Konten erhalten jetzt den konkreten Hinweis, zuerst mit B zu sperren und danach erneut zum Löschen zu markieren. Vollständig übersprungene Aktionen und abgebrochene Bestätigungen erhalten Markierung und Cursor, ohne den Bestand erneut zu laden. Leere Aktionsergebnisse werden explizit als Array behandelt, damit der Abbruch unter StrictMode keine Count-Ausnahme erzeugt. Schutzprüfungen, Sperrpflicht und bestätigte Löschung gesperrter Gäste bleiben erhalten. Ausgeführte Aktionen, Vorschauen und mögliche Teilfehler laden den Bestand weiterhin neu.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regression vor Änderung | Zwei neue Tests fehlgeschlagen. Unnötiges Neuladen bei Enter auf aktivem Gast und verlorene Auswahl vor anschließendem B reproduziert |
| Fokussierte Prüfung | Sieben Tests bestanden. Aktiver Gast mit Hinweis und erhaltener Auswahl, anschließendes Sperren und Löschen, Nein ohne Neuladen sowie Gastaktionen für fehlenden Empfänger, MailUser, GuestMailUser und UserMailbox geprüft |
| Tatsächliche TUI | Zwei isolierte macOS-Pseudoterminalprozesse über den echten Personal-Einstieg mit synthetischen Gastkonten. Aktiv: Leertaste, Enter, erhaltener Hinweis und Markierung, B, N, erhaltener Auswahlzustand, Esc. Gesperrt: Leertaste, Enter, Löschbestätigung, Y, Löschvorschau unter WhatIf und Esc. Beide Prozesse Exitcode 0 und HasErrors false |
| Vollständige lokale Verifikation | 584 Tests erkannt, 583 bestanden, 0 fehlgeschlagen, ein vorhandener Windows-Dateisperrtest übersprungen. Parser für 52 PowerShell-Dateien ohne Fehler, keine ungeklärten Analyzer-Befunde |
| Dokumentation und Git | Personal-Anleitung, Support, Changelog, Testbeschreibung und dieses Protokoll aktualisiert. Dateilinks in 16 Markdown-Dateien ohne fehlende Ziele. Git-Diff ohne Whitespace-Befunde |

Keine produktiven Konten geändert oder gelöscht. Keine produktiven Daten in Tests, kein Push und keine Veröffentlichung. Die simulierten Schreibtests und tatsächlichen Terminalvorschauen ersetzen keine gesondert autorisierte Mandantenabnahme. Windows-Dateisperren und Windows-Tastaturdarstellung bleiben offen. Frühere Prüfergebnisse folgen unten.

## Ergänzung 0.3.0: Lehrer-Exchange-Wiederholungen und Wiederanlauf, 30. September 2026

Der Betreiber meldete nach einer Lehreranlage einen Exchange-Fehler, weil das Empfängerobjekt auf einem Exchange-Server nicht gefunden wurde. Die bisherige Schleife wartete fünfmal mit 60 Sekunden, behandelte jedoch eine generische Ausnahme mit dieser konkreten Servermeldung sowie manche CAS- und Schreibfehler als endgültiges `Failed`. Die tatsächlichen Ausnahmeinformationen des produktiven Fehlers wurden nicht live untersucht. Die lokale Regression bildet die gemeldete Nachricht mit synthetischem Empfänger und generischer Ausnahme nach.

`config/LehrerSync.psd1` und die Standardwerte des gemeinsamen Lehrer-Batchhelfers verwenden jetzt drei Wiederholungen mit jeweils 150 Sekunden nach einer sofortigen Prüfung. Eindeutige Exchange-Empfänger-nicht-gefunden-Fehler bleiben beim Lesen, Schreiben und Rücklesen wiederholbar. Ein interner `Retryable`-Wert erhält diese Information aus den Konfigurationsergebnissen. Nach ausgeschöpften Wiederholungen folgt `Pending` mit Reparaturbefehl. Parameter-, Berechtigungs- und Verbindungsfehler sowie bestätigte Konfigurationsabweichungen bleiben sofortige Fehler. WhatIf wartet und schreibt nicht. Schüler und übriges Personal behalten ihre explizit konfigurierten fünf Wiederholungen mit 60 Sekunden.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regression vor Änderung | 7 fokussierte Tests, 1 bestanden und 6 fehlgeschlagen. Zu kurze Standardintervalle, nicht erkannte generische Servermeldung, fehlende CAS-Behandlung und endgültiger Abbruch eines wiederholbaren Konfigurationsfehlers reproduziert |
| Fokussierte Prüfung nach Änderung | 14 bestanden, 0 fehlgeschlagen. Erfolgreiche Fortsetzung nach Lese-/Konfigurationsfehler, vier Prüfungen mit drei simulierten 150-Sekunden-Wartezeiten, Pending-Reparaturbefehl, Mailbox-/CAS-Schreibfehler, Rückleseverlust, unverändert endgültiger Parameter-/Konfigurationsfehler und WhatIf geprüft |
| Tatsächliche CLI | Separater Prozess mit `Invoke-SyntheticPersonalCli.ps1 -Population Teacher -Mode Exchange` ruft den echten Lehrer-Reparatureinstieg für ein bestehendes synthetisches Konto unter WhatIf auf. Exitcode 0, `Mode = ExchangeOnly`, Exchange-Aktionsvorschau und `HasErrors = false`. Fernverbindungen und Mutationen ersetzt oder gesperrt |
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich. 582 Tests erkannt, 581 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Parser und Analyzer | 52 PowerShell-Dateien einschließlich Tests ohne Parserfehler. 0 ungeklärte Analyzer-Befunde, 25 bestehende erlaubte Write-Host-Warnungen |
| Dokumentation und Git | README, Lehreranleitung, Betrieb, Support, Testbeschreibung, Changelog und dieses Protokoll aktualisiert. Lokale Dateilinks in 16 Markdown-Dateien ohne fehlende Ziele, sieben bestehende GitHub-Issue-/Security-Routen vom lokalen Dateicheck ausgenommen. `git diff --check` ohne Befund |

Alle Prüfungen verwendeten synthetische Konten und ersetzte Wartezeiten. Reale 150-Sekunden-Wartezeiten, tatsächliche Exchange-Replikation, Windows-Dateisperren und eine separat autorisierte Mandantenabnahme bleiben offen. Vorhandene Änderungen auf `codex/0.3.0-personal` wurden erhalten. Keine produktiven Mandantenänderungen, keine produktiven Testdaten, kein Push und keine Veröffentlichung. Die folgenden Abschnitte dokumentieren frühere Prüfstände.

## Ergänzung 0.3.0: Administrator-Schutzgruppe umbenannt, 30. September 2026

Die vom Betreiber gemeldete Gruppe heißt jetzt `SEC-A-ADM-GlobalAdministrators`. Die gemeinsame Schutzkonfiguration in `config/PersonalSync.psd1`, die synthetischen Gruppenfixtures, README, Personal-Anleitung, Spezifikation, Testbeschreibung und Changelog verwenden diesen Namen. Die vorhandene eindeutige Namensauflösung bleibt erhalten. Mitglieder der Gruppe sind in Personal-, Lehrer- und Schülerverwaltung geschützt. Eine feste Gruppen-ID wurde nicht ergänzt.

162 Tests aus `PersonalSync.Tests.ps1` bestanden, 0 fehlgeschlagen. Darin enthalten sind fünf zusätzliche Fälle für direkte Administrator-Mitgliedschaft in allen drei Populationen, einen ausschließlich vorhandenen veralteten Namen und einen mehrdeutigen aktuellen Namen. Die tatsächlichen isolierten CLI-Aufrufe der Suite verwenden ebenfalls die umbenannte synthetische Gruppe. Ein fehlender oder mehrdeutiger aktueller Gruppenname sperrt Schreibaktionen weiterhin. Diese gezielte Prüfung ersetzt keine erneute vollständige Gesamtsuite, deren vorheriger Stand folgt unten.

Die Prüfungen verwendeten synthetische Konten und ersetzte Fernbefehle. Keine produktiven Mandantenänderungen, keine Bearbeitung produktiver Arbeitsmappen, kein Push und keine Veröffentlichung. Die tatsächliche Gruppe und ihre ID wurden nicht live geprüft.

## Ergänzung 0.3.0: Personalexport lesend vergleichen, 30. September 2026

Der Betreiber meldete einen Vorprüfungsabbruch nach `-ListToExcel`, weil die exportierte Jobzelle `ExterneAdmins` kein schreibbares Profil ist. Der Export enthält bewusst geschützte und rollenlose Konten, der bisherige Adapter verlangte jedoch auch für den rein lesenden Vergleich vollständige Pflichtwerte und ein freigegebenes Profil. Die Regression wurde mit einem vollständigen synthetischen Export reproduziert.

Nur der Personalmodus Compare aktiviert jetzt die Prüfzeilenverarbeitung im vorhandenen Exceladapter und Vergleich. Unvollständige Namen sowie leere, unbekannte oder geschützte Jobs werden als `Workbook.Review` mit Zeilennummer und Identität gemeldet. Auch explizit zugeordnete nicht schreibbare Konten werden so übersprungen. Zulässige Zeilen werden weiter verglichen. Bekannte explizite Prüfidentitäten sind keine Abgangskandidaten. Update einschließlich WhatIf, Einzelpflege und Lehrer-/Schülerimporte behalten ihre strikten Vorprüfungen. Admin-/Gastschutz bleibt erforderlich, eine geänderte Jobzelle gibt kein geschütztes Konto frei.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Reproduktion vor dem Fix | Beide neuen Regressionstests fehlgeschlagen. Der tatsächliche Export-Compare-Aufruf reproduzierte den Vorprüfungsabbruch beim geschützten Job, die neue Vergleichsoption war noch nicht vorhanden |
| Fokussierte Personal-/Exporttests | 181 bestanden, 0 fehlgeschlagen |
| Tatsächliche Export-/Compare-/Update-CLI | Isolierter Harness mit ReviewFixture exportiert ein zulässiges Konto, ExterneAdmins, ein rollenloses Konto ohne Namensbestandteile und einen Gast mit freigegebenem Job. Compare liefert 1 bestehendes Konto, 3 Workbook.Review-Warnungen, HasErrors=false und Exitcode 0. Update-WhatIf mit derselben Datei bleibt vor Mutationen gesperrt, Exitcode 1. Kein unerwarteter Fernbefehl |
| Datei und Adapter | Dateihash nach Export, Compare und abgewiesenem Update unverändert. Prüfzeilen bleiben mit Job und Identität im lesenden Adapter erhalten. Der Standardimport wird vor und nach dem lesenden Import weiterhin abgewiesen. Lehrer akzeptieren die Prüfzeilenoption nicht |
| Schutz und Abgänge | Gast, Adminschutz, lokale AD-Synchronisierung und Shared Mailbox erzeugen nur im ausdrücklich lesenden Vergleich Prüfhinweise. Der strikte Vergleich derselben Ziele meldet weiterhin Fehler. Eine explizit identifizierte Prüfzeile erzeugt weder Neuzugang noch falschen Abgang |
| Vollständige lokale Verifikation | Invoke-LocalVerification erfolgreich, 566 Tests, 565 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen. Schüler-, Lehrer-, Personal- und TUI-Regressionen enthalten |
| Frische Discovery | 566 Tests erkannt, keine Discovery-Fehler |
| Parser und Analyzer | 52 PowerShell-Dateien ohne Parserfehler. 0 ungeklärte Analyzer-Befunde, 25 weiterhin erlaubte Warnungen für Terminalausgabe |
| Dokumentation und Git | README, Personal-Anleitung, Spezifikation, Plan, Betrieb, Support, Sicherheit, Testing und Changelog angepasst. Lokale Dateilinks in 16 Markdown-Dateien ohne fehlende Ziele, git diff --check ohne Befund |

Die Tests verwendeten ausschließlich synthetische Arbeitsmappen und ersetzte Fernbefehle. Die vorhandene produktive Personal.xlsx wurde nicht geöffnet oder verändert. Vorhandene Änderungen auf codex/0.3.0-personal wurden erhalten. Keine produktiven Mandantenänderungen, kein Push und keine Veröffentlichung. Windows-Dateisperrtest und eine separat autorisierte Mandantenprüfung bleiben offen. Die folgenden Abschnitte dokumentieren frühere Prüfstände.

## Ergänzung 0.3.0: Schüler-TUI, 30. September 2026

`Sync-SchuelerEntra.ps1` und `Invoke-SchuelerSync` unterstützen `-ManageAccounts` mit dem gemeinsamen Schülerbestand, Renderer und Lebenszykluscode. Direkte aktive und gesperrte Schüler-Mitgliedskonten sind verwaltbar. Lehrer-/Ganztagsmitgliedschaften und Shared Mailboxes sind ausgeschlossen. Gäste, lokale AD-Synchronisierung, geschützte Administratoren und Verzeichnisrollen bleiben gegen Änderungen geschützt. Der Detailbereich zeigt Rolle, Kontotyp, tatsächliche CLS-Gruppen, Department/Jahrgang, Klassenlehrer und UTC-Login.

K wählt eine bestehende eindeutige, nicht dynamische oder lizenzierende CLS-Gruppe. Die Zielmitgliedschaft wird hinzugefügt und bestätigt, dann werden mit dem bestehenden Attributhelfer ausschließlich `Department` und `OfficeLocation` aktualisiert. Alte CLS-Mitgliedschaften werden erst nach Bestätigung dieser Schritte entfernt. L wählt ein aktives internes direktes Mitglied des Schul-Pädagogikteams und aktualisiert ausschließlich die Manager-Zuordnung mit dem bestehenden Helfer und Rückleseprüfung. Gesperrte Konten bleiben bei beiden Aktionen gesperrt. B, U und Enter verwenden die bestehenden Sperr-, Entsperr- und Löschaktionen.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Fokussierte Personal-/TUI-Tests | 200 bestanden, 0 fehlgeschlagen |
| Schülerbestand und Anzeige | Direkte Schülerpopulation mit aktuellen CLS-Gruppen, Klassenattributen und Manager geprüft. Renderer zeigt K/L, Klasse, Jahrgang und Klassenlehrer neben Rolle, Status und Login. Fehlende CLS-Zuordnung wird ausdrücklich angezeigt |
| Klassenwechsel | Neue Mitgliedschaft, beide Klassenattribute und Entfernung alter CLS-Mitgliedschaft geprüft. Schülerrolle, zusätzliche Lizenzgruppe und gesperrter Status erhalten. SDK-serialisierter Benutzer-Payload enthält genau die beiden Klassenattribute. Erste Klasse ohne bisherigen CLS-Bestand, bereits passende Klasse und zwei ausgewählte Schüler geprüft |
| Klassenlehrerwechsel | Ausschließliche Manager-Änderung auf frisch bestätigtes Schulteam-Mitglied und bereits passende Manager-Zuordnung geprüft. Keine Gruppen- oder Attributmutation |
| Frische Schutzprüfungen | Verlorene Schülerrolle, neuer Adminschutz, neue Shared Mailbox, geänderte Ausgangsklasse/Manager-Zuordnung sowie veränderte Klassen-/Lehrerziele blockieren beide Zuordnungsaktionen. Dynamische und lizenzierende Ausgangsklassen blockieren vor dem Hinzufügen |
| Teilfehler | Nicht bestätigte neue Mitgliedschaft, Attributfehler und neuer Adminschutz nach dem Hinzufügen melden mögliche Teiländerung. Alte CLS-Mitgliedschaft wird in diesen Fällen nicht entfernt |
| Lebenszyklus und Berechtigungen | Schüler-Sperren, Entsperren und Löschen mit synthetischen Mutationen geprüft. Nach Verlust der direkten Schülerrolle werden alle drei ohne Mutation übersprungen. Klassenwechsel fordert Benutzer- und Gruppenmitgliedschafts-Schreibscope, Klassenlehrerwechsel nur Benutzer-Schreibscope an |
| Vorschau und Abbruch | Beide Zuordnungen unter WhatIf ohne Mutation. Esc bei Zielwahl und Nein bei Bestätigung ohne Schreibaufruf. Zielauswahl, Namensbestätigung und Aktualisierung bleiben auf die Schülerpopulation begrenzt |
| Tatsächliche CLI | Echter Schülereinstieg im isolierten Harness erkennt ManageAccounts und weist umgeleitete Eingabe vor einer Verbindung ab. Kombinationen mit Update, ListToExcel und File werden als ungültiger Parametersatz abgewiesen |
| Tatsächliche Terminalbedienung | Zwei macOS-Pseudoterminalaufrufe mit `Invoke-SyntheticPersonalCli.ps1 -Population Student -Mode Tui`. Leertaste, K/L, Pfeil nach unten, Enter, Y, Aktualisierung und Esc geprüft. Beide liefern die zwei Vorschauen mit `Status = WhatIf`, `Changed = false`, `HasErrors = false` und Exitcode 0. Der zweite Aufruf bestätigt die endgültige Schülerüberschrift und Bedienhinweise |
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich, 559 Tests, 558 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Discovery | 559 Tests erkannt, keine Discovery-Fehler |
| Parser und Analyzer | 52 PowerShell-Dateien in src, tests, config und Root-Einstiegen ohne Parserfehler. 0 ungeklärte Analyzer-Befunde, 25 erlaubte Warnungen für Terminalausgabe |
| Dokumentation und Git | README, Schüleranleitung, Betrieb, Sicherheit, Support, Testplan, Changelog, Personal-Spezifikation und Umsetzungsplan ergänzt. Lokale Dateilinks in 16 Markdown-Dateien ohne fehlende Ziele. Git-Diff ohne Whitespacefehler |

Die TUI aktualisiert keine Quelldatei. Vor einem späteren Excel-Sync müssen Klasse und Klassenlehrer in der Quelle abgeglichen werden. Eine veraltete Quelle kann TUI-Zuordnungen zurücksetzen und der bisherige Schüler-Sync kann gesperrte Schüler reaktivieren. Es gibt keinen automatischen Rollback bei Teiländerungen.

Alle Prüfungen verwendeten ausschließlich synthetische Konten und ersetzte Fernbefehle. Vorhandene Änderungen auf `codex/0.3.0-personal` wurden erhalten. Keine produktiven Mandantenänderungen, kein Push und keine Veröffentlichung. Windows-Tastatur/Darstellung, der übersprungene Windows-Dateisperrtest und eine gesondert autorisierte Mandantenabnahme bleiben offen. Die folgenden Abschnitte dokumentieren frühere Prüfstände.

## Ergänzung 0.3.0: Lehrer-ManageAccounts, 30. September 2026

`Sync-LehrerEntra.ps1` und `Invoke-LehrerSync` unterstützen `-ManageAccounts` als eigenen Parametersatz. Der Modus verwendet dieselbe Inventar-, TUI- und Aktionsumsetzung wie die Personalverwaltung, begrenzt auf aktive und gesperrte direkte Mitglieder von Schulteam und Ganztag. Schüler und Shared Mailboxes sind ausgeschlossen. Die direkte Rollenmitgliedschaft wird vor jeder Aktion erneut geprüft, der Rollenfilter bleibt nach einer Aktualisierung erhalten. Lehrer-Sync, Einzelpflege und Export behalten ihre bestehenden Aufrufe.

Die Lehrerüberschrift lautet `Lehrer- und Ganztagskonten`. Rollen, Kontotyp, Status, letzte erfolgreiche Anmeldung, Bestätigungen und Fortschritt werden gemeinsam angezeigt. B sperrt und widerruft Sitzungen, U entsperrt Mitgliedskonten, Enter löscht gesperrte Konten. Administrator-, Identitäts-, Status-, Mandanten- und Empfängerschutz gelten weiterhin.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Fokussierte Personal-/TUI-Tests | 159 bestanden, 0 fehlgeschlagen |
| Lehrerpopulation | Aktive und gesperrte direkte Schulteam-/Ganztagsmitglieder geprüft. Übrige Personalkonten, Schülerzuordnungen, Shared Mailboxes und rein indirekte Lehrer-Mitglieder ausgeschlossen |
| Aktionen und frische Prüfung | Sperren, Entsperren und Löschen erfolgreich mit synthetischen Mutationen geprüft. Verlorene direkte Lehrerrolle, neue Admin-/Schülerzuordnung und neue Shared Mailbox blockieren jede Aktion vor der Mutation |
| TUI-Aktualisierung | Ganztagskonto nach Sperrung weiterhin in der Lehrerpopulation, mit gesperrtem Status und korrekter Lehrerüberschrift |
| Tatsächliche CLI | Echter Lehrereinstieg über den isolierten Harness erkennt `-ManageAccounts` und weist umgeleitete Eingabe mit der Interaktivitätsmeldung ab. Kombinationen mit `-Update`, `-ListToExcel` und `-File` werden vor einer Verbindung als ungültiger Parametersatz abgewiesen |
| Tatsächliche Tastaturprüfung | Zwei macOS-Pseudoterminalaufrufe mit `Invoke-SyntheticPersonalCli.ps1 -Population Teacher -Mode Tui`, der zweite zusätzlich mit `-DisabledFixture`. Rolle `SEC-A-ROL-Schule_PädagogischesTeam`, Kontotyp, Status und UTC-Login sichtbar. Leertaste, B beziehungsweise U und Enter, jeweils Y, Aktualisierung und Esc geprüft. Drei Aktionsvorschauen mit `WhatIf`, beide Aufrufe Exitcode 0 und `HasErrors = false` |
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich, 518 Tests, 517 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Discovery | 518 Tests erkannt, keine Discovery-Fehler |
| Parser und Analyzer | 51 PowerShell-Dateien in src, tests, config und Root-Einstiegen ohne Parserfehler. 0 ungeklärte Analyzer-Befunde, 18 erlaubte Warnungen für Terminalausgabe |
| Dokumentation und Git | README, Lehreranleitung, Changelog, Betrieb, Sicherheit, Support, Testplan, Personal-Spezifikation und Umsetzungsplan aktualisiert. Repository-Dateilinks in 16 Markdown-Dateien ohne fehlende Ziele. `git diff --check` ohne Befund |

Die Prüfung verwendete ausschließlich synthetische Konten und ersetzte Fernbefehle. Vorhandene lokale Änderungen auf `codex/0.3.0-personal` wurden erhalten. Keine produktiven Mandantenänderungen, kein Push und keine Veröffentlichung. Windows-Darstellung, der übersprungene Windows-Dateisperrtest und eine gesondert autorisierte Mandantenabnahme bleiben offen. Die folgenden Abschnitte dokumentieren frühere Ergänzungen.

## Ergänzung 0.3.0: Rollenanzeige und Gastaktionen, 30. September 2026

Die gemeinsame TUI zeigt im Detailbereich sämtliche erkannten `SEC-A-ROL-`-Gruppen oder `Keine Rollengruppe`. Der Personalmodus zeigt zusätzlich den Kontotyp. Die bereits geprüfte Anzeige der letzten erfolgreichen Anmeldung bleibt erhalten.

Die Personal-TUI kann nach Bestätigung aktive Gäste sperren und gesperrte Gäste löschen, einschließlich rollenloser Gäste, empfängerloser Konten und eindeutiger `MailUser`-/`GuestMailUser`-Empfänger. Die Gastfreigabe wird nur an die Sperr- und Löschprüfungen übergeben. Gastentsperrung, normale Profilpflege und der bisherige gesperrte-Konten-Einstieg behalten ihren Gastschutz. Geänderte Identität, Kontotyp, Mandant, geschützte oder ausgeschlossene Mitgliedschaften, lokale AD-Synchronisierung und unzulässige oder veränderte Empfänger blockieren die Aktion. Exportierte Gastzeilen erhalten weiterhin die Warnung vor einem schreibenden Import, auch bei einem gültigen Job.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Fokussierte Personal-/TUI-Tests | 133 bestanden, 0 fehlgeschlagen |
| Gastablauf | Markieren, Sperren mit Statusprüfung und Sitzungswiderruf, Neuladen, erneutes Markieren und Löschen synthetisch geprüft, jeweils ohne Empfänger sowie mit `UserMailbox`, `MailUser` und `GuestMailUser` |
| Frische Schutzprüfungen | Block und Delete nach neuem Adminschutz, ausgeschlossener Rolle, Verzeichnisrolle, Shared Mailbox, Identitätswechsel, Kontotypwechsel, lokaler AD-Synchronisierung und mehrdeutigen Empfängern jeweils ohne Mutation übersprungen |
| Rollenanzeige | Mehrere und unbekannte Rollengruppen, Lizenzgruppenfilter, rollenlose beziehungsweise ältere Ergebnisobjekte und beide TUI-Modi geprüft |
| Tatsächlicher Terminalaufruf | Zwei echte Personal-Einstiege mit `Invoke-SyntheticPersonalCli.ps1 -Mode Tui -GuestFixture`, der zweite zusätzlich mit `-DisabledFixture`. Sichtbare Rolle, `Kontotyp: Guest`, Login-Zeitstempel und `GuestMailUser`. Leertaste, B beziehungsweise Enter, Y und Esc geprüft. Beide lieferten `WhatIf` und `HasErrors = false`, Exitcode 0 |
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich, 492 Tests, 491 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Discovery | 492 Tests erkannt, keine Discovery-Fehler |
| Parser und Analyzer | 51 Dateien in src, tests, config und Root-Einstiegen ohne Parserfehler. 0 ungeklärte Analyzer-Befunde, 18 weiterhin erlaubte Warnungen für Terminalausgabe |
| Dokumentation und Git | README, Anleitungen, Betrieb, Sicherheit, Support, Spezifikation, Brainstorming, Plan, Testplan und Changelog aktualisiert. Interne Dateilinks in 16 Markdown-Dateien ohne fehlende Ziele, GitHub-relative Routen ausgenommen. `git diff --check` ohne Befund |

Alle Änderungen liegen lokal auf `codex/0.3.0-personal`. Vorhandene Änderungen wurden erhalten. Die Tests und Terminalaufrufe verwendeten ausschließlich synthetische Konten und ersetzte Fernbefehle. Es gab keine produktiven Mandantenänderungen, keinen Push und keine Veröffentlichung. Windows-Darstellung, der Windows-Dateisperrtest und eine gesondert autorisierte Mandantenprüfung bleiben offen. Die folgenden Abschnitte dokumentieren frühere Prüfstände.

## Ergänzung 0.3.0: Last Login in der TUI, 29. September 2026

`-ManageAccounts` zeigt im Detailbereich des aktuell mit den Pfeiltasten ausgewählten Kontos `Last Login Timestamp (UTC)`. Die gemeinsame TUI zeigt denselben Wert auch im bisherigen gesperrte-Konten-Einstieg. Der Wert stammt aus dem bereits geladenen Inventar. Bekannte Anmeldungen erscheinen als UTC-Zeitstempel, fehlende Werte als `Kein Wert` und nicht abrufbare Anmeldedaten als `Nicht abrufbar`. Ältere Ergebnisobjekte ohne Login-Felder bleiben darstellbar. Die Seitengröße reserviert eine zusätzliche Zeile für diese Anzeige.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Fokussierte TUI-/Personaltests | 109 bestanden, 0 fehlgeschlagen |
| Neue Anzeigefälle | UTC-Konvertierung und Wechsel zwischen zwei ausgewählten Konten in Personal- und bisherigem Modus, nicht abrufbare Daten, ältere Objekte ohne Login-Felder und leerer Bestand geprüft |
| Tatsächliche TUI | Echter Personal-Einstieg über `Invoke-SyntheticPersonalCli.ps1 -Mode Tui` im macOS-Pseudoterminal geöffnet. Sichtbare Anzeige `Last Login Timestamp (UTC): 2026-01-02T10:30:00Z`, mit Esc beendet. Ergebnis `Mode = Manage`, leere Aktionen, `HasErrors = false`. Ausschließlich synthetischer Bestand, keine Mutation |
| Vollständige lokale Verifikation | 468 Tests, 467 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Discovery | 468 Tests erkannt |
| Analyzer | 0 ungeklärte Befunde, weiterhin 18 erlaubte Warnungen für bestehende Terminalausgabe |

README, TUI-Anleitungen, Spezifikation, Testplan und Changelog sind ergänzt. `git diff --check` und interne Dokumentationslinks sind ohne Befund. Es wurden keine produktiven Daten oder Mandanten verwendet, keine Änderungen gepusht und nichts veröffentlicht. Windows-Darstellung und eine separat autorisierte Mandantenprüfung bleiben offen. Der folgende Abschnitt dokumentiert die zuvor vollständig geprüfte Timestamp-Erweiterung in allen drei Excel-Exporten.

## Ergänzung 0.3.0: Last Login Timestamp, 29. September 2026

Alle drei `-ListToExcel`-Exporte hängen `Last Login Timestamp` rechts an. Schüler haben jetzt zehn Spalten (Timestamp in J), Lehrer und Personal neun (Timestamp in I). Der gemeinsame Excel-Schreiber verwendet die letzte erfolgreiche Anmeldung aus `signInActivity.lastSuccessfulSignInDateTime`, formatiert als UTC-Text mit Sekunden, beispielsweise `2026-01-02T10:30:00Z`. Namen, Job/Klasse, Passwort- und Identitätsspalten behalten ihre Positionen.

`Get-EntraLastLogin` wird von der bisherigen Kontenliste sowie den neuen Rollenexporten gemeinsam verwendet. Ohne Wert bleibt die Zelle leer. Fehler beim Lesen der Anmeldedaten führen zu einer Warnung und einem erneuten Benutzerabruf ohne `signInActivity`, anschließend zu leeren Timestamp-Zellen. Auch der Personalexport gibt seine Inventarwarnung aus. Schüler-/Lehrerexporte fordern zusätzlich `AuditLog.Read.All` an. Der Zusatzwert ist informativ, wird vom Import nicht als Pflichtfeld behandelt und löst keine Sync-Aktion aus.

Die Bedeutung der erfolgreichen Anmeldung und die Berechtigungs-/Lizenzanforderung wurden gegen die offizielle [signInActivity-Dokumentation](https://learn.microsoft.com/en-us/graph/api/resources/signinactivity?view=graph-rest-1.0) und [Benutzerabfrage](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0#example-11-get-users-including-their-last-sign-in-time) geprüft. Eine leere Zelle belegt nicht, dass ein Konto nie verwendet wurde.

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Fokussierte Tests | 109 Export-, Personal- und Kontenlistentests bestanden, 0 fehlgeschlagen |
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich, 465 Tests, 464 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Discovery | 465 Tests in separatem PowerShell-Prozess erkannt |
| Neue Abdeckung | UTC-Konvertierung, bekannte und fehlende Werte für alle drei Exportarten, keine Übernahme neuerer bloßer Anmeldeversuche, Wiederabruf bei nicht verfügbaren Anmeldedaten und sichtbare Personalexport-Warnung |
| Tatsächliche CLI und XLSX | Timestamp-Zellen über alle drei echten Einstiege mit vollständig ersetzten Fernbefehlen geprüft. Tabellenbereiche mit zehn beziehungsweise neun Spalten, leere Passwörter und unveränderte Identitätsspalten geprüft. Alle drei Exportarten mit vorhandenen Importadaptern wieder eingelesen |
| Parser und Analyzer | 52 PowerShell-Dateien ohne Parserfehler. 0 ungeklärte Analyzer-Befunde, weiterhin 18 erlaubte Warnungen für bestehende Terminalausgabe |
| Dokumentation und Git | README, Anleitungen, Betrieb, Testplan, Sicherheit, Support, Mitwirkung, Spezifikation, Plan und Changelog angepasst. Interne Links in 16 Markdown-Dateien ohne fehlende Ziele. `git diff --check` ohne Befund |

Die Prüfungen nutzten ausschließlich synthetische Daten und dieselben temporären Module wie die vorherigen lokalen Läufe. Keine produktiven Arbeitsmappen oder Mandanten wurden geöffnet oder geändert, kein Push und keine Veröffentlichung. Windows und eine separat autorisierte Mandantenprüfung der tatsächlichen Anmeldedaten, Berechtigungen und Lizenz bleiben offen. Die folgenden Abschnitte dokumentieren die vorherigen Exportstände mit ihren damaligen Spaltenzahlen.

## Ergänzung 0.3.0: Schüler- und Lehrerexport, 29. September 2026

`Sync-SchuelerEntra.ps1` und `Sync-LehrerEntra.ps1` unterstützen jetzt `-ListToExcel`, ebenso ihre öffentlichen Modulfunktionen. Ohne `-File` gelten `Schueler.xlsx` beziehungsweise `Lehrer.xlsx` im Repository-Root. Der Export ist ein eigener Parametersatz, widersprüchliche Aktionsschalter werden vor einer Verbindung abgewiesen. Die bisherigen Sync-Parameter, Aliase und Regeln bleiben erhalten. Der gemeinsame Schreiber `Export-AccountWorkbook` wird von allen drei Exporten verwendet.

Exportiert werden aktive und gesperrte direkte Mitglieder der jeweiligen konfigurierten Rollen, ohne Shared Mailboxes. Lehrer-/Ganztagkonten werden bei Mehrfachrollen nur einmal ausgegeben. Schüler erhalten neun Spalten einschließlich direkter Klasse und Manager-UPN, Lehrer acht mit dem passenden vorhandenen Job. Fehlende oder widersprüchliche Werte bleiben leer und werden gemeldet. Die vorhandenen Schüler-/Lehrer-Updates können enthaltene gesperrte Konten reaktivieren, dieser Unterschied zum Personalmodus ist in den Anleitungen und Exportwarnungen ausdrücklich beschrieben.

Alle Exporte haben leere Passwortzellen, verifizierte Textwerte und Blatt/Tabelle `Tabelle1` mit `TableStyleMedium14`. Vorhandene Ziele werden nicht überschrieben. `-WhatIf` erstellt weder Datei noch fehlende Ordner. Schüler-/Lehrerexporte fordern ausschließlich `User.Read.All` und `Group.Read.All` sowie eine passende lesende Exchange-Verbindung an. Die lokale Prüfung nutzte dieselben isolierten Module wie die Personalabnahme darunter.

### Tatsächlich geprüfte Ergänzung

| Prüfung | Ergebnis |
|---|---|
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich, 459 Tests, 458 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Discovery | 459 Tests in separatem PowerShell-Prozess erkannt |
| Nativer Parser | 52 PowerShell-Dateien einschließlich Konfigurationen, Tests und Analyzer-Einstellungen, 0 Fehler |
| Fokussierter kombinierter Lauf | Zunächst 79 Export- und Personaltests erfolgreich nach Korrektur der Testisolation, danach Gesamtsuite einschließlich vier weiterer Exporttests und Dokumentationsvertrag erfolgreich |
| Neue Abdeckung | 19 Fälle in `WorkbookExport.Tests.ps1` und ein zusätzlicher Dokumentationsvertrag |
| Tatsächliche CLI-Aufrufe | Beide echten Sync-Einstiege über vollständig ersetzte Fernbefehle, neues Ziel, `-WhatIf`, Ablehnung vorhandener Datei mit unverändertem Hash und Ablehnung widersprüchlicher Parameter geprüft |
| XLSX und Importadapter | Neun Schülerspalten, acht Lehrerspalten, Klasse, Klassenlehrer, Job, IDs, leere Passwörter, kanonische Jobs, leerer Bestand und Formelzeichen als Text geprüft, beide Schemata wieder eingelesen |
| Bestand und Fehlerfälle | Gesperrte Konten, Ganztag, konkurrierende Rollen, fehlende Klassen, Managerlesefehler, verschachtelte Nichtbenutzer, SharedMailbox hinter doppelter Empfänger-ID, fehlende/unpassende Tenant-ID, fehlende Rollen, unbekannte Mitglieder, Exchange-Lesefehler und abgeschlossener Fortschritt nach Fehler geprüft |
| Analyzer | 0 ungeklärte Befunde, weiterhin 18 erlaubte Warnungen für bestehende Terminalausgabe |
| Git und Datenschutz | `git diff --check` ohne Befund, Branch `codex/0.3.0-personal`, weiterhin ausschließlich die zwei vorhandenen synthetischen XLSX-Fixtures verfolgt |
| Dokumentation | README, Bedienungsanleitungen, Betrieb, Testplan, Support, Sicherheit, Mitwirkung, Plan, Spezifikation und Changelog aktualisiert. 16 Markdown-Dateien mit internen Links geprüft, keine fehlenden Ziele |

Im ersten kombinierten Testversuch erwartete eine Prüfung einen englischen Parameterfehler, obwohl PowerShell einen deutschen Text ausgab. Außerdem ließ ein unvollständiger neuer Exchange-Stub bestehende Personaltests scheitern. Beide Testprobleme wurden korrigiert und der kombinierte Lauf wiederholt. Der erste Gesamtlauf wurde vor Pester wegen vier Analyzer-Befunden gestoppt. Zwei automatische Variablennamen wurden ersetzt. Zwei Delegationen der `ShouldProcess`-Prüfung haben gezielte Attribute mit Begründung erhalten, die vorhandenen Bestätigungs- und Vorschaufunktionen werden weitergereicht und geprüft. Der anschließende Gesamtlauf bestand.

Die bestehenden Änderungen zur Versionierung und Personalumsetzung bleiben erhalten. Keine produktiven Arbeitsmappen wurden für Tests verwendet, keine produktiven Verbindungen oder Änderungen ausgeführt, kein Push und keine Veröffentlichung. Windows und eine separat autorisierte Mandantenprüfung bleiben offen. Die folgenden Abschnitte dokumentieren die vorherigen Prüfläufe.

## Übrige Personalkonten 0.3.0, lokale Abnahme am 29. September 2026

Die freigegebene Architektur und Spezifikation sind auf `codex/0.3.0-personal` implementiert. Entwicklungsstand ist 0.3.0, die letzte veröffentlichte Version bleibt 0.2.0. Es wurden keine produktiven Entra-/Exchange-Verbindungen geöffnet, keine Mandantenänderungen ausgeführt und keine Änderungen gepusht oder veröffentlicht.

### Umgesetzter Umfang

- Eigener Personal-Einstieg mit Einzelanlage/-pflege, optionalem Sync, Remove und Exchange-Reparatur. Gemeinsame Passwort-, Identitäts-, Lizenz-, Rollen-, Excel-, Exchange- und Berichtsfunktionen werden wiederverwendet.
- Sieben freigegebene Profile mit sprechenden Jobwerten. ExterneAdmins und GlobalAdministrators sind sichtbar und vor Änderungen geschützt. Lehrer, CO-L, PA, Ganztag, Schüler und Shared Mailboxes sind einschließlich transitiver Ausschlussmitgliedschaften ausgenommen.
- Ein gemeinsamer aktiver/gesperrter Bestand für Liste, Ergebnisobjekt, Export und TUI, einschließlich rollenloser Konten und sichtbar geschützter Sonderkonten.
- `-Update` allein erstellt und aktualisiert ohne Abgangssperrung. Gesperrte Bestandskonten bleiben gesperrt. Der ausdrückliche Wiederanlauf einer passenden Teilanlage verwendet `-ResumeCreation` und Objekt-ID.
- TUI mit `B`, `U` und Enter, Namensbestätigung, passendem Statusfilter, frischer Schutzprüfung, Teilergebnissen, abgeschlossenen Fortschrittsanzeigen und neu geladenem Bestand.
- Neue XLSX-Datei mit den acht Lehrerspalten, Blatt und Tabelle `Tabelle1`, `TableStyleMedium14`, leeren Passwortzellen und verifizierten Textwerten. Bestehende Ziele bleiben erhalten.
- Vorhandene und ausstehende Lizenzen bleiben erhalten. Lizenzzustand und tatsächliche Lizenzgruppe werden vor einem nötigen Standardbeitritt frisch geprüft. Eine nicht konfigurierte ABP wird ausgelassen und nicht entfernt.
- Bestehende Schüler-, Lehrer- und Kontenverwaltungseinstiege bleiben kompatibel. Dokumentation, README, Changelog, Spezifikation und Umsetzungsplan sind aktualisiert.

### Tatsächlich ausgeführte Prüfungen

| Prüfung | Ergebnis |
|---|---|
| PowerShell | 7.6.6 auf macOS |
| Vollständige lokale Verifikation | `tests/Invoke-LocalVerification.ps1` erfolgreich |
| Pester-Gesamtsuite | 439 Tests, 438 bestanden, 0 fehlgeschlagen, 1 vorhandener Windows-Dateisperrtest übersprungen |
| Frische Pester-Discovery | 439 Tests in einem separaten Prozess ohne vorher importiertes Anwendungsmodul erkannt |
| Nativer Parser | 50 PowerShell-Quelldateien, Einstiege, Konfigurationen und Testdateien, 0 Fehler |
| Modulmanifest | Version 0.3.0, fünf öffentliche Einstiege einschließlich `Invoke-PersonalSync` |
| PSScriptAnalyzer | 0 ungeklärte Befunde, 18 durch die bestehende Prüfroutine erlaubte `PSAvoidUsingWriteHost`-Warnungen für Terminalausgabe |
| Personaltests | 64 synthetische Tests für Profile, Inventar, Schutz, Sync, Lizenzierung, Excel, TUI und reale CLI-Prozesse bestanden |
| Tatsächliche CLI-Aufrufe | Liste, PassThru, Export, Vergleich, Update-Vorschau, Add, Remove, beide UPN-Aliase, Exchange und Fehlerstatus über den echten Einstieg mit vollständig ersetzten Fernbefehlen geprüft |
| Native XLSX-Verifikation | Tabellenstruktur/-stil, acht Spalten, Objekt-IDs, Jobwerte, leere Passwörter, Formelzeichen als Text, leerer Export und kein Dateiziel unter `-WhatIf` geprüft |
| TUI im macOS-Pseudoterminal | Aktives synthetisches Konto mit Leertaste, `B`, `Y`, `Esc` geprüft. Gesperrtes synthetisches Konto mit Leertaste, `U`, `Y`, anschließend Leertaste, Enter, `Y`, `Esc` geprüft. Alle Aktionen unter `-WhatIf`, Ergebnisse `WhatIf`, Bestand neu geladen, Markierung zurückgesetzt, keine Fernmutation |
| TUI-Aktionsregressionen | Synthetische Sperrung, Entsperrung und Löschung, gemischte Auswahl, Nein/Esc, Teilfehler beim Widerruf, neue Ausschlussmitgliedschaft, SharedMailbox-/Identitäts-/Tenantwechsel und fehlgeschlagene Besitzabfrage geprüft |
| Git-Prüfung | `git diff --check` ohne Befund, Branch bestätigt. Unter Git weiterhin ausschließlich die zwei vorhandenen synthetischen XLSX-Fixtures |
| Dokumentationslinks | Interne Dateilinks in README, Changelog und `docs/*.md` ohne fehlende Ziele |

Die lokale Umgebung verwendete Pester 5.7.1, PSScriptAnalyzer 1.23.0, ImportExcel 7.8.10 sowie die Graph-Module 2.31.0. Sie wurden temporär unter `/private/tmp/entra-personal-v030-check/modules` bereitgestellt. `PSModulePath` und `XDG_DATA_HOME` waren ausschließlich für die Prüfprozesse gesetzt. Der erste Versuch ohne isolierten Benutzermodulpfad scheiterte mit `Access to the path '/Users/andre/.local/share/powershell/Modules' is denied`. Die isolierte Modulbereitstellung und expliziten Imports behoben das Umgebungsproblem. Es wurden keine Projektabhängigkeiten geändert.

ImportExcel meldete lokal die bekannte AutoSize-Warnung zu `System.Drawing`. Die Implementierung benutzt kein AutoSize, die tatsächliche Dateierzeugung und Wiedereinlesung bestanden. Das Verhalten der geladenen EPPlus-Tabellen wurde anhand des gespeicherten `StyleName = TableStyleMedium14` geprüft.

Die gesamte Suite enthält weiterhin die bisherigen Schüler-, Lehrer-, Exchange-, Excel- und Kontenverwaltungsregressionen. Synthetische Neuanlagen prüfen die Reihenfolge deaktivierte Anlage, früher Checkpoint, Gruppenprüfung und Aktivierung. Checkpointfehler erhalten den Sperrstatus. Eine zwischen Vergleich und Beitritt hinzugekommene Lizenz verhindert einen Standardbeitritt. Eine inzwischen fehlende Lizenz der Standardgruppe blockiert den Beitritt. Zwischen Mailbox- und CAS-Schreiben hinzugekommene Ausschlüsse verhindern den zweiten Schreibzugriff.

### Offene, separate Integrationsprüfungen

- PowerShell-7-Gesamtsuite unter Windows, Windows-Dateisperrverhalten und echte Windows-Tastatur-/Konsolendarstellung.
- Autorisierte kontrollierte Mandantenabnahme ausschließlich mit isolierten synthetischen Konten. Dazu gehören tatsächliche Gruppen-/Schutz-IDs, A1Faculty-SKU, Verzeichnisrollen, Exchange-Richtlinien, Adresslistenfilter, Auditaufbewahrung, Provisionierung und Sitzungswiderruf.
- Die tatsächliche GlobalAdministrators-ID bleibt unbestätigt. Implementiert ist die eindeutige exakte Namensauflösung mit Schreibsperre bei fehlendem oder mehrdeutigem Schutz.

Diese offenen Integrationsprüfungen sind keine ausgeführten Ergebnisse. Die lokale Implementierung und Verifikation sind abgeschlossen. Die historischen Abschnitte darunter bleiben unverändert als frühere Nachweise erhalten.


## Personalabgleich 0.2.0

Die Erweiterung benötigt eine eigene kontrollierte Windows- und Mandantenabnahme mit synthetischen Konten. Zu prüfen sind beide direkten Personalrollen, die A3-/A1-Gruppen-SKUs, Lizenzbestandsschutz bei Jobwechsel, ABP und Adresslistenfilter, Exchange-Richtlinien, die tatsächliche Purview-Auditaufbewahrung und der Erstwechsel bei genau zwölf Passwortzeichen. Ein voller Abgangstest darf nur mit einer vollständig isolierten Personalpopulation erfolgen. Die folgenden 0.1.1-Zahlen dokumentieren den früheren Schülerstand und sind keine Ergebnisse für 0.2.0.

Die lokalen Ergebnisse für 0.2.0 stehen unten in einem eigenen Abschnitt. Windows-Tests wurden für 0.2.0 nicht ausgeführt. Ein vom Betreiber ausgeführter lesender Mandantenvergleich meldete zunächst 28 `Preflight.Identity`-Fehler. Nach der lokalen Korrektur meldete der Betreiber keine Fehler mehr und zeigte eine Abgangsliste mit 14 Konten. Deren Aktivierungszustände wurden nicht übermittelt. Der Vergleich wurde nach der Änderung der Abgangslogik noch nicht wiederholt.

### Lokal ausgeführte Prüfungen für 0.2.0 am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| PowerShell | 7.6.6 auf macOS |
| Pester-Gesamtsuite | 285 erkannt, 284 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Frische Pester-Discovery | 285 Tests in einem neuen PowerShell-Prozess erkannt, keine Testausführung |
| Nativer Parser | 35 `.ps1`-, `.psm1`- und `.psd1`-Dateien geprüft, 0 Parserfehler |
| `Test-ModuleManifest` | Version `0.2.0`, Exporte `Invoke-SchuelerSync` und `Invoke-LehrerSync` |
| PSScriptAnalyzer | 0 Befunde für beide Einstiegsskripte und `src`, mit `PSScriptAnalyzerSettings.psd1` |
| Git-Prüfung | `git diff --check` ohne Befund, nur synthetische XLSX-Fixtures ausgenommen |
| Lehrerarbeitsmappe | Nach der gemeldeten Sperrfehlermeldung ließ sich `Lehrer2026-27.xlsx` lokal exklusiv zum Lesen und Schreiben öffnen. Die Datei wurde dabei nicht verändert |
| Mandant und Windows | Betreiber meldete nach der Lizenzkorrektur einen lesenden Vergleich ohne Fehler und mit 14 Abgängen. Aktivierungszustände offen. Kein Mandantenlauf und kein Windows-Test durch Codex |

Die lokalen Tests nutzten Pester 5.7.1, ImportExcel 7.8.10 und PSScriptAnalyzer 1.25.0 aus einem temporären Modulverzeichnis. Die Graph- und Exchange-Integrationen wurden in den Tests gemockt. Die produktiven Gruppen, Lizenz-SKUs und Richtlinien bleiben bis zur kontrollierten Mandantenabnahme unbestätigt.

Die zusätzlichen Tests prüfen die Lehrer-Fortschrittsphasen, den Vergleich vor Kontoaktionen, den vollständigen `-OutputFile`-Bericht sowie die Schreibprüfung bei gesperrter Arbeitsmappe mit und ohne geplante Rückschreibung. Die neue Fortschrittsanzeige wurde lokal mit Mocks geprüft, nicht in einer interaktiven Windows-Konsole oder im Mandanten.

Die Schülerregression wurde ebenfalls lokal geprüft: Bereits deaktivierte Konten ohne Excel-Zeile fehlen in der Abgangsliste. Ein eindeutig zugeordnetes deaktiviertes Konto mit Excel-Zeile erscheint als Änderung und wird nach erfolgreicher Update-Prüfung aktiviert. Der Integrationstest bestätigt zudem, dass ein bereits abgegangenes und gesperrtes Konto im folgenden lesenden Vergleich nicht erneut erscheint. Der vom Betreiber gezeigte Mandantenfall wurde nicht erneut live ausgeführt.

Ein zusätzlicher Berichtstest reproduzierte den mehrzeiligen Umbruch der Lehrer-Kategorie „Bestehende“ mit langen technischen Spalten. Nach der kompakten Spaltenauswahl enthält der Bericht für diesen synthetischen Eintrag eine Datenzeile. Die Live-Ausgabe mit 30 Personen wurde nicht erneut ausgeführt.

Die reale `Lehrer2026-27.xlsx` wurde lokal nur lesend importiert: 30 Zeilen, einschließlich `L, PA` und `PA, JAS`, ohne Importfehler. Eine synthetische Kopie bestätigte die Rückschreibung von Passwort, Objekt-ID, UPN und Mail bei erhaltenem Mehrfachjob. In die reale Datei wurde nicht zurückgeschrieben. Der vom Betreiber gemeldete erste lesende Mandantenvergleich führte wegen 28 Lizenz-Vorprüfungsfehlern zu keiner Personeneinteilung. Ein fokussierter Test reproduzierte den Fehler mit `State = Active` und `Error = None`. Nach dessen Korrektur meldete der Betreiber keine Fehler und 14 Abgänge, darunter ein Funktionskonto. Die lokale Abgangslogik wurde daraufhin so angepasst, dass gesperrte Konten ohne Excel-Zeile nicht als Abgänge erscheinen und passende gesperrte Bestandskonten als Reaktivierung geplant werden. Diese Änderung wurde noch nicht im Mandanten geprüft.

### Zusätzliche lokale TUI-Prüfung am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester-Gesamtsuite nach TUI-Erweiterung | 318 erkannt, 317 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| PSScriptAnalyzer und Parser | 0 Befunde für TUI-Modul und Einstiegsskript, 0 Parserfehler |
| Modulmanifest | Version `0.2.0`, vier Exporte einschließlich `Invoke-DisabledAccountTui` |
| Synthetische TUI in einem Pseudoterminal | Leertaste, Enter und Ja führten mit `-WhatIf` zu `WhatIf` ohne Löschaufruf. Ohne `-WhatIf` wurde ausschließlich ein gemocktes `Remove-MgUser` für die synthetische Objekt-ID aufgerufen und `Deleted` angezeigt |
| Lesender CLI-Modus | `-List` und `-PassThru` lieferten mit synthetischen Graph- und Exchange-Antworten die bisherige Prüfliste |
| Git-Prüfung | `git diff --check` ohne Befund |

Die Tests verwendeten keine produktive Verbindung. Die echte Kontolöschung, die Microsoft-Graph-Berechtigung im Mandanten, eine Windows-Konsole und die manuelle Prüfung von Personenzuordnung und Sperrfrist sind nicht live verifiziert.

### Alle gesperrten Entra-Konten in der Prüfliste

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester nach Erweiterung des Inventars | 334 erkannt, 333 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Gezielte Inventar- und TUI-Tests | 45 bestanden, 0 fehlgeschlagen. Zwei weitere Schutztests für fehlende Exchange-Empfänger bestanden danach im Gesamtlauf. |
| Synthetischer `-List`-Aufruf | 3 gesperrte Konten angezeigt: `UserMailbox`, `SharedMailbox` und ein Konto ohne Exchange-Empfänger. Kopfzeile: 2 ohne persönliches Postfach. |
| Synthetische TUI im Pseudoterminal mit `-WhatIf` | Vor der späteren Ausblendung von `SharedMailbox` waren alle 3 Konten sichtbar. Leertaste beim Konto ohne Exchange-Empfänger ließ `Markiert: 0` stehen und zeigte den Hinweis „nur zur Prüfung sichtbar“. Esc beendete die TUI ohne Aktion. |
| PSScriptAnalyzer | 0 Fehler oder Warnungen außer der bestehenden Stilregel für `Write-Host` bei farbiger Konsolenausgabe |
| Git-Prüfung | `git diff --check` ohne Befund |

Die neuen Inventar- und TUI-Fälle wurden nur mit synthetischen Graph- und Exchange-Antworten geprüft. Ein erneuter Lauf im produktiven Mandanten fand nicht statt.

### Entsperren von Konten mit Besitzobjekten

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regressionstest vor der Korrektur | Ein gesperrtes Konto mit sieben Besitzobjekten wurde als `Skipped` gemeldet. 28 von 29 gezielten Tests bestanden. |
| Gezielte TUI-Tests nach der Korrektur | 29 von 29 bestanden. Der Entsperrpfad fragte Besitzobjekte nicht erneut ab und setzte bei gemocktem Graph `AccountEnabled = true`. Der Löschpfad übersprang Eigentümerkonten weiterhin. |
| Pester-Gesamtsuite | 335 erkannt, 334 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI mit `-WhatIf` | Ein gesperrtes `UserMailbox`-Konto mit einem Besitzobjekt war markierbar. Nach `U` und Ja erschien `WhatIf`, ohne Änderung am Konto. |

Im produktiven Mandanten wurden keine Konten entsperrt. Die Wirkung der Korrektur für die vom Betreiber genannten Konten wurde nicht live geprüft.

### Freigegebene Postfächer aus der TUI ausgeblendet

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Gezielte TUI-Tests | 30 von 30 bestanden. `SharedMailbox` fehlte in der übergebenen TUI-Auswahl und konnte nicht in eine Entsperraktion gelangen. |
| Pester-Gesamtsuite | 336 erkannt, 335 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI mit `-WhatIf` | Vor der späteren Freigabe empfängerloser Konten zeigte die TUI 2 von 3 gesperrten Entra-Konten. `SharedMailbox` fehlte, das Konto ohne Exchange-Empfänger blieb als „Nur prüfen“ sichtbar. Esc beendete die Auswahl ohne Aktion. |
| Synthetisches `-List` | Alle 3 Konten blieben in der lesenden Inventarliste, einschließlich `SharedMailbox`. |

Im produktiven Mandanten wurde die TUI nach dieser Änderung nicht ausgeführt.

### Konten ohne Exchange-Empfänger auswählen

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regressionstest vor der Änderung | 52 gezielte Tests, davon 7 fehlgeschlagen. `RecipientStatus` fehlte, und die TUI sowie der Löschpfad übersprangen ein Konto ohne Empfänger. |
| Gezielte Check- und TUI-Tests nach der Änderung | 54 bestanden, 0 fehlgeschlagen. Der Check unterscheidet `None`, `Unique`, `Ambiguous` und `Unreadable`. Nur `None` und persönliche Benutzerpostfächer sind auswählbar. |
| Schutz vor `SharedMailbox` | Eine nach der Auswahl im frischen Exchange-Bestand gefundene `SharedMailbox` führte zu `Skipped`, auch mit großgeschriebener und von Leerzeichen umgebener Objekt-ID. Uneindeutige Zuordnung und Fehler beim erneuten Bestandsabruf führten ebenfalls zu `Skipped`. `Remove-MgUser` wurde in diesen Fällen nicht aufgerufen. |
| Pester-Gesamtsuite | 341 erkannt, 340 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI mit `-WhatIf` | Das Konto ohne Exchange-Empfänger war markierbar. Nach Enter und Ja meldete der erneute Bestandsabgleich `WhatIf`. `SharedMailbox` blieb ausgeblendet. Es gab keinen Löschaufruf. |

Ein echter Lösch- oder Entsperrlauf im produktiven Mandanten fand nicht statt. Die TUI kann nicht selbst belegen, ob ein empfängerloses Konto einer realen Person gehört oder seit mindestens 90 Tagen gesperrt ist.

### Fortschrittsanzeige bei mehreren ausgewählten Konten

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Gezielte Check- und TUI-Tests | 56 bestanden, 0 fehlgeschlagen. Für Löschen und Entsperren wurden die Schritte „1 von 2“ und „2 von 2“ sowie der Abschluss des Fortschrittsbalkens geprüft. Der Abschluss wurde auch bei übersprungenen Konten und Aktionsfehlern geprüft. |
| Pester-Gesamtsuite | 343 erkannt, 342 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI im Pseudoterminal | Je zwei Konten wurden für Löschen und Entsperren markiert. Beide Bestätigungen listeten die zwei Namen. Beide `-WhatIf`-Läufe meldeten zweimal `WhatIf` und zeigten während der Verarbeitung den Fortschritt. `SharedMailbox` blieb ausgeblendet. |
| PSScriptAnalyzer | 0 Fehler oder Warnungen in den drei betroffenen PowerShell-Dateien, abgesehen von der bestehenden Stilregel für `Write-Host` bei farbiger Konsolenausgabe |
| Git-Prüfung | `git diff --check` ohne Befund |

Die synthetischen Terminalaufrufe verwendeten keine produktive Verbindung und führten weder Löschung noch Entsperrung aus. Die Darstellung in einer Windows-Konsole und die echte Graph-Aktion bleiben ungeprüft.

### Löschung trotz Besitzobjekten nach Warnung

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Regressionstest vor der Änderung | 40 gezielte TUI-Tests, davon 38 bestanden und 2 fehlgeschlagen. Der Bestätigungsdialog zeigte kein Besitzobjekt, und der Löschpfad meldete `Skipped`. |
| Gezielte TUI-Tests nach der Änderung | 40 bestanden, 0 fehlgeschlagen. Die Bestätigung nennt Typ, Name und ID eines Besitzobjekts. Ein erneut bestätigter Eigentümer wird nach Warnung gelöscht. Ein Fehler bei der erneuten Besitzabfrage führt weiter zu `Skipped`. |
| Pester-Gesamtsuite | 346 erkannt, 345 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI im Pseudoterminal | Ein Eigentümerkonto war markierbar. Der Dialog zeigte `group: Schulteam (group-1)` und warnte vor fehlender Eigentümerschaft. Nach Ja erschien die Warnung vor dem `-WhatIf`-Löschaufruf. Es gab keinen Schreibzugriff. |
| PSScriptAnalyzer | 0 Fehler oder Warnungen im geänderten TUI-Modul, abgesehen von der ausgeschlossenen bestehenden Stilregel für `Write-Host` |
| Git-Prüfung | `git diff --check` ohne Befund |

Die aktuelle Regel wurde mit gemocktem Graph und synthetischen Konten geprüft. Eine echte Löschung im Mandanten wurde nicht durchgeführt.

### Lokale Prüfung der Entsperroption am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester-Gesamtsuite | 329 erkannt, 328 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Synthetische TUI im Pseudoterminal | Leertaste, `U` und Ja führten mit `-WhatIf` zu `WhatIf` ohne Änderung |
| Entsperrpfad mit gemocktem Graph | `Update-MgUser` erhielt ausschließlich die ausgewählte Objekt-ID und `AccountEnabled = true`. Aktive Konten, geänderte Identität und Shared Mailboxes wurden übersprungen. Graph-Fehler wurden als `Failed` gemeldet. |
| Berechtigung und Syntax | `User.EnableDisableAccount.All` wurde für die bestätigte Entsperraktion angefordert, kein Schreibscope unter `-WhatIf`. PowerShell-Parser ohne Befund. |
| PSScriptAnalyzer | 0 Befunde im TUI-Modul. Im Einstiegsskript bestehen 8 Stilwarnungen für `Write-Host`, das die farbigen Konsolenüberschriften ausgibt. Keine Fehler. |
| Git-Prüfung | `git diff --check` ohne Befund |

Es gab keine produktive Verbindung oder Entsperrung. Die echte Graph-Berechtigung und die Wirkung im Mandanten sind nicht live verifiziert.

### Lokale Prüfung der Exchange-Bestätigungen am 27. September 2026

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Pester-Gesamtsuite | 331 erkannt, 330 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| Lehrer- und Schülerpostfächer | Fokussierte Tests bestätigen `Compliant` ohne Konfigurationsaufruf bei leeren Exchange-Abweichungen. Abweichende Postfächer erreichen weiter den bestätigungspflichtigen Pfad. |
| Schüler-Integration | Ein konformes Bestands-Postfach wurde im vollständigen Update nicht erneut konfiguriert. Zwei abweichende Postfächer wurden weiterhin verarbeitet. |
| Parser und PSScriptAnalyzer | Geänderte PowerShell-Dateien ohne Parserfehler. PSScriptAnalyzer meldet 2 bestehende `Write-Host`-Stilwarnungen in der Lehrerorchestrierung und 1 Information an anderer Stelle im Exchange-Modul, keine Fehler. |
| Git-Prüfung | `git diff --check` ohne Befund |

Der vom Betreiber gezeigte Lauf und die Exchange-Einstellungen im Live-Mandanten wurden nach der Korrektur nicht erneut geprüft.

## Historischer Schülerabgleich 0.1.1

Stand der Release-Prüfung: 12. September 2026 auf `codex/v0.1.1-recovery-fix` vor dem Merge für `v0.1.1`.

Die automatisierten PowerShell-Prüfungen wurden auf macOS erfolgreich ausgeführt. Ein lesender Lauf gegen den vorgesehenen Mandanten wurde vom Betreiber durchgeführt und lieferte die erwarteten Vergleichstabellen. Produktive Schreibaktionen wurden im Rahmen dieser lokalen Prüfung nicht unabhängig ausgeführt oder bestätigt. Die kontrollierte Windows- und Mandantenabnahme bleibt deshalb getrennt dokumentiert.

## Lokal ausgeführte Prüfungen

| Prüfung | Ergebnis |
|---|---|
| PowerShell | PowerShell 7.6.6 auf macOS |
| Pester | 228 Tests erkannt, 227 bestanden, 0 fehlgeschlagen, 1 Windows-spezifischer Test übersprungen |
| PSScriptAnalyzer | Keine Befunde für Einstiegsskript und Modul mit `PSScriptAnalyzerSettings.psd1` |
| Nativer PowerShell-Parser | 29 versionierte `.ps1`-, `.psm1`- und `.psd1`-Dateien ohne Parserfehler |
| `Test-ModuleManifest` | Erfolgreich, Modulversion `0.1.1`, PowerShell-Mindestversion `7.0` |
| Manueller Wiederanlauf | Unit- und Integrationstests bestätigen `-Add -EntraObjectId` für ein passendes, noch nicht gruppiertes Teilkonto sowie die erneute Graph-Prüfung vor jeder Mutationsphase. Normaler Abgleich, Fremdkonten und ein zwischenzeitlich verändertes Konto bleiben blockiert |
| `git diff --check` | Keine Whitespace-Fehler |
| `git check-ignore -v Schueler.xlsx` und synthetische Zielpfade | XLSX-Regeln greifen im Root und in Unterordnern, auch für Sicherungen, temporäre Kopien und Großschreibung der Endung |
| `git ls-files '*.xlsx'` | Ausschließlich `tests/fixtures/Schueler-Testdaten.xlsx` verfolgt |
| Modulinventar | Graph SDK 2.39.0, ExchangeOnlineManagement 3.10.1, ImportExcel 7.8.10, Pester 5.7.1 und PSScriptAnalyzer 1.25.0 |
| Lesender Mandantenvergleich | Vom Betreiber erfolgreich ausgeführt, keine lokale Bestätigung produktiver Schreibaktionen |
| Passwort-Auswahlraum, unabhängig aus dem Wortkatalog berechnet | 847 eindeutige Wörter, 220.505 eindeutige Zehnbuchstaben-Präfixe, 19.845.450 mögliche Ausgaben, rund 24,24 Bit |

## Reproduzierbare VM-Prüfung

Verwende PowerShell 7 im Repository-Root. Installiere zuerst die in [ENTRA-SCHUELER-SYNC.md](ENTRA-SCHUELER-SYNC.md) aufgeführten Module. Das Repository enthält keine getestete Versionssperre. Dokumentiere die tatsächlich geladenen Versionen und den geprüften Commit, statt ungeprüfte Versionsnummern zu übernehmen.

```powershell
$ErrorActionPreference = 'Stop'
$PSVersionTable
git rev-parse HEAD

$required = @(
  'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users',
  'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Groups',
  'ExchangeOnlineManagement', 'ImportExcel', 'Pester', 'PSScriptAnalyzer'
)
Import-Module Pester -MinimumVersion 5.0 -Force
Import-Module PSScriptAnalyzer -Force
Get-Module -ListAvailable $required |
  Sort-Object Name,Version -Descending |
  Select-Object Name,Version,Path

$parseErrors = foreach ($relativePath in git ls-files '*.ps1' '*.psm1' '*.psd1') {
  $tokens = $null
  $errors = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PWD $relativePath), [ref]$tokens, [ref]$errors
  )
  $errors
}
$parseErrors | Format-List
if ($parseErrors) { throw 'Native PowerShell-Parserprüfung fehlgeschlagen.' }

Test-ModuleManifest .\src\SchuelerSync\SchuelerSync.psd1
$pesterResult = Invoke-Pester .\tests -Output Detailed -PassThru
if ($pesterResult.Result -ne 'Passed' -or $pesterResult.TotalCount -eq 0) {
  throw 'Pester nicht erfolgreich oder keine Tests ausgeführt.'
}

$analyzerFindings = @(
  Invoke-ScriptAnalyzer -Path '.\Sync-SchuelerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\Sync-LehrerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\src' -Recurse -Settings '.\PSScriptAnalyzerSettings.psd1'
)
$analyzerFindings | Format-Table RuleName,Severity,ScriptName,Line,Message -Wrap
if ($analyzerFindings.Count -gt 0) { throw 'Analyzer-Befunde sind vor der Freigabe zu klären.' }
Get-Module $required | Select-Object Name,Version
```

Die Live-Befehle und Sollwerte stehen vollständig in [TESTING.md](TESTING.md). Erstelle dafür eine Arbeitskopie der erfundenen Fixture. Nutze einen vollständig isolierten Mandanten oder eine konfigurierte Schüler-Rollengruppe mit ausschließlich synthetischen Mitgliedern. Eine Testklasse in der produktiven Schüler-Rollengruppe reicht nicht aus, weil Abgänge für die gesamte konfigurierte Rollenpopulation berechnet werden.

## Abnahme-Checkliste

Jede Position benötigt Datum, tatsächliches Ergebnis und gegebenenfalls eine bereinigte Fehlermeldung. Schreibe keine Passwörter oder Tokens in das Protokoll.

- [x] Native Parserprüfung aller 29 versionierten PowerShell-Dateien
- [x] `Test-ModuleManifest` auf macOS erfolgreich
- [x] Gesamte Pester-Suite auf macOS erfolgreich, tatsächliche Anzahl Passed/Failed/Skipped dokumentiert
- [x] PSScriptAnalyzer ohne ungeklärte Befunde
- [ ] Standardvergleich zeigt alle Tabellen, unveränderter Excel-Hash, keine Graph-/EXO-Schreiboperationen
- [ ] Vollständiges und selektives `-WhatIf` ohne Passworterzeugung, Backup, Dateiänderung, Mandantenänderung oder EXO-Wartezeit
- [ ] Isolierter Neuzugang bleibt bis zur verifizierten Excel-Rückschreibung deaktiviert
- [ ] Initialpasswort exakt 12 Zeichen, kontrollierte Erstanmeldung ohne erzwungenen Passwortwechsel
- [ ] Alle Entra-Attribute, Manager und drei Pflichtgruppen nachgelesen
- [ ] UPN-Kollision über Graph-/EXO-Adressen erkannt, alternatives Präfix und numerischer Fallback geprüft
- [ ] Klassen- und Lehrerwechsel geprüft, Zielgruppen vor Entfernung alter Gruppen nachgelesen, bestehendes Passwort unverändert
- [ ] Synthetischer Abgang deaktiviert und Sitzungswiderruf separat bestätigt
- [ ] EXO-Sollwerte vollständig nachgelesen, einschließlich Richtlinien, AuditDelegate, AuditOwner, beider AuditAdmin-Varianten und sämtlicher CAS-Werte
- [ ] Sofortiges Postfach und verzögerte Bereitstellung geprüft, höchstens sechs Prüfungen und fünf Batch-Wartezeiten von jeweils 60 Sekunden
- [ ] EXO-Schreibfehler wird sofort gemeldet und nicht als Bereitstellungsproblem wiederholt
- [ ] `-ConfigureExchangeOnlineOnly -Mail <synthetischer-UPN>` einschließlich Aliase und Idempotenz geprüft
- [ ] Zweiter lesender Vergleich ohne Änderungen für konforme aktive Schüler, deaktivierte Konten ohne Excel-Zeile erscheinen nicht erneut als Abgänge
- [ ] Wiederherstellung nach gesperrter oder konkurrierend veränderter Excel-Datei geprüft, neue Konten bleiben deaktiviert, erhaltene Datei- und Backup-Varianten zugeordnet
- [ ] Windows-Dateisperr-, Umbenennungs- und Wiederherstellungsverhalten aus `WorkbookIdentitySafety.Tests.ps1` bestätigt
- [ ] Ausgabeprüfung aller Kanäle mit ausschließlich erfundenen Passwortwerten, einschließlich Passwort-Retry-Fehlern
- [ ] Git-Hygiene nach VM-Test bestätigt und ausschließlich dokumentierte synthetische Testobjekte kontrolliert bereinigt

## Ergebnis eintragen

| Feld | Eintrag |
|---|---|
| Datum / Prüfer | offen |
| Geprüfter Commit | offen |
| Windows / PowerShell | offen |
| Tatsächlich geladene Modulversionen | offen |
| Isolierter Testmandant / Rollenpopulation | offen |
| Pester Passed / Failed / Skipped | offen |
| Analyzer / nativer Parser / Manifest | offen |
| Lesender Vergleich / WhatIf | offen |
| Synthetischer Lebenszyklus / EXO / Wiederherstellung | offen |
| Verbleibende Abweichungen | offen |
| Produktive Freigabe | ausstehend |

## Auditkorrekturen vom 27.09.2026, Version 0.2.0

Die folgenden Ergebnisse wurden nach den Korrekturen lokal unter macOS mit PowerShell 7.6.6, Pester 5.7.1 und PSScriptAnalyzer 1.23.0 ermittelt. Die bereits vorhandenen Arbeitsänderungen blieben erhalten. Es wurden keine produktiven Konten verändert und keine Änderungen gepusht oder veröffentlicht.

| Prüfung | Tatsächliches Ergebnis |
| --- | --- |
| `tests/Invoke-LocalVerification.ps1` mit lokalen Pester-/Analyzer-Manifesten | Erfolgreich, 374 Tests, 373 bestanden, 0 fehlgeschlagen, 1 übersprungen, Laufzeit 29,19 Sekunden |
| Discovery in frischer PowerShell-Sitzung | 374 Tests gefunden, keine Discovery-Fehler |
| PowerShell-Syntax der Quellen, Einstiegsskripte und Konfiguration | Keine Parserfehler |
| PSScriptAnalyzer für diese Dateien | Keine Fehler und keine ungeklärten Warnungen. 12 bekannte `PSAvoidUsingWriteHost`-Hinweise für farbige Überschriften |
| Windows-Dateisperrprüfung | Unter macOS übersprungen, weiterhin auf Windows auszuführen |
| Echte CLI-Aufrufe beider Sync-Skripte mit Excel als Berichtsziel | Abgewiesen, Exitcode 1, Quellhash unverändert, keine Cloud-Anmeldung |
| Echter Schüler-CLI-Aufruf mit ausschließlich Excel-Überschriften und `-Update` | Abgewiesen, Exitcode 1, Quellhash unverändert, keine Cloud-Anmeldung |
| Audit-Regressionen | Leerschutz, Mandantenabgleich, Rollen-/Lizenzänderungen, Wiederanlauf, Bestätigungslogik, Berichtsschutz und manuelles Add bestanden |
| Neues Schülerpasswortformat | Genau zwölf Zeichen, alle drei Zeichenklassen, mehr als 64 Bit Auswahlraum, Einmaligkeit und begrenzte Wiederholungen geprüft |
| Abschließender Secret-Musterscan | 74 aktuelle Textdateien geprüft, keine Treffer für private Schlüssel, GitHub-/AWS-Token oder feste Client-Secrets |
| Git-Schutz und Diff | Produktive Arbeitsmappen weiterhin ignoriert, `git diff --check` ohne Befund |
| Dokumentation nach dem abschließenden Abgleich | 11 Dokumentationstests bestanden, 0 fehlgeschlagen |
| Workflow-YAML | Lokal erfolgreich geparst |
| GitHub-Workflow | Vorbereitet, nicht ausgeführt, da kein Push erfolgt ist |

Die frühen Regressionstests reproduzierten zunächst fünf Auditfehler. Zwischenläufe fanden unvollständige Test-Mocks, eine unerwünschte Bestätigungsweitergabe durch `ForEach-Object -MemberName` sowie Kodierungswarnungen in zwei Konfigurationsdateien. Diese Punkte wurden korrigiert. Ein früher Exchange-Reparaturtest versuchte wegen eines fehlenden Mocks eine Graph-Anmeldung. Der Browseraufruf scheiterte, der Prozess wurde abgebrochen. Es wurde keine Anmeldung abgeschlossen. Der reine Schüler-Exchange-Reparaturmodus verwendet weiterhin ausschließlich Exchange. Die abschließenden Prüfungen liefen ohne Anmeldung und ausschließlich mit synthetischen Testdaten.

Die frühere Messung von rund 24,24 Bit in diesem Dokument beschreibt das historische Zweiwortformat. Ab diesem Korrekturstand gilt für neue Schülerkonten das zufällige Zeichenformat. Bestehende Passwörter sowie die Erstwechselregel bleiben unverändert.

## Abschließender lokaler Prüflauf vor Git-Freigabe 0.2.0

Am 27.09.2026 wurde der vollständige lokale Prüflauf mit PowerShell 7.6.6, Pester 5.7.1 und PSScriptAnalyzer 1.23.0 erneut ausgeführt:

- 374 Tests, davon 373 bestanden, 0 fehlgeschlagen und 1 Windows-spezifischer Test unter macOS übersprungen.
- Laufzeit der Pester-Tests: 26,65 Sekunden.
- Keine PowerShell-Parserfehler und keine ungeklärten Analyzer-Befunde.
- Secret-Musterscan über 70 vorhandene, für Git vorgesehene Textdateien ohne Treffer. Produktive Arbeitsmappen bleiben ignoriert. Die neue Lehrer-Fixture enthält ausschließlich erfundene Personen.
- `git diff --check` ohne Befund.

Dieser lokale Prüflauf führte keine produktiven Microsoft-365-Änderungen aus. Die Windows- und macOS-Jobs des neuen GitHub-Workflows laufen nach dem Push. Ihre Ergebnisse sind separat in GitHub Actions zu prüfen und ersetzen keine kontrollierte Mandantenabnahme.
