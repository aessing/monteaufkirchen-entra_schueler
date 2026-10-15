# Test- und Abnahmeplan für die Parallels-VM

## Lesbare Personaljobs in Export und Import

Die Personaltests prüfen die sieben freigegebenen Jobprofile sowie den geschützten Job `ExterneAdmins`. Native XLSX-Exporte müssen Unterstriche und CamelCase durch Wortabstände ersetzen, insbesondere `Kinderhaus Pädagogisches Team` mit Umlaut. Der Import muss lesbare Namen, unterschiedliche Groß-/Kleinschreibung, mehrfache Leerzeichen und bisherige interne Werte auf dasselbe Profil zurückführen. Interne Profilwerte, Rollen und Exchange-Profil behalten ihre kanonischen Werte. Das tatsächliche Personal-`JobTitle` verwendet bei Neuanlage und Profilpflege den lesbaren Namen. Bei den drei Kinderhausprofilen entfällt darin der Präfix `Kinderhaus`, während `CompanyName` und die eindeutigen Excel-Jobnamen erhalten bleiben. Regressionen prüfen alle sieben freigegebenen Profile, echte synthetische Graph-Schreibparameter, die Korrektur technischer und bisher lesbarer Bestandstitel mit Kinderhauspräfix sowie einen zweiten Vergleich ohne erneute Jobtiteländerung. Bestandsliste und Vergleich zeigen lesbare Jobnamen, während interne Jobschlüssel erhalten bleiben. Konkurrierende Jobs, unbekannte Namen und mehrdeutige konfigurierte Anzeigenamen müssen abgewiesen werden. `Externe Admins` bleibt im schreibenden Import gesperrt und im lesenden Vergleich eine kanonisch identifizierte Prüfzeile. Leere Jobzellen bleiben leer. Der tatsächliche isolierte Personal-CLI-Test prüft die exportierte Jobzelle sowie anschließenden Compare und Update-WhatIf mit unveränderter Quelldatei. Lehrerexport und Lehrerjobkürzel werden separat geprüft.

## Löschtasten in allen Konten-TUIs

Backspace und Entf müssen nach Markierung dieselbe Löschbestätigung öffnen. Enter darf in keiner Kontenliste eine Aktion oder Bestätigung starten. Die Zielauswahl für Klasse und Klassenlehrer verwendet weiterhin Enter. Die fokussierten Tests prüfen beide Löschpfade, alle drei verwalteten Populationen, erhaltene Auswahl nach Nein und den Sperrhinweis für aktive Gäste. Im synthetischen Terminal zusätzlich beide Tasten prüfen, mit N abbrechen und mit Esc zurückkehren. Anschließend einen bestätigten Löschvorgang ausschließlich mit `-WhatIf` prüfen.

## Letzte TUI-Zeile mit Tastenbefehlen

`DisabledAccountTui.Tests.ps1` erfasst die Reihenfolge der Terminalausgabe und prüft, dass die letzte Ausgabezeile die passenden Tastenbefehle enthält. Abgedeckt sind normale und leere Kontenlisten, Personal-/Lehreransicht, Schüleransicht mit K/L, Klassen-/Klassenlehrer-Zielauswahl sowie Sperr-/Entsperr-/Löschbestätigungen mit übersprungenen Konten.

## Adressregeln bei Neuanlagen, 10. Oktober 2026

Der Vertrag steht in [ADRESSREGELN.md](ADRESSREGELN.md). `CreationIdentity.Tests.ps1` prüft die neuen Domains, alle Schulteamjobs, OGTS ohne Alias, Personalnamen, `Externe*`-Namen und gemeinsame Reservierung von Lehrer-Hauptadresse und Alias. Kollisionsfälle umfassen Graph und Exchange, Vornamenspräfixe und Zahlenfallback. `Comparison.Tests.ps1` prüft abweichende und fehlende Bestandswerte. `EntraMutations.Tests.ps1` belegt, dass der Schüler-Schreibpfad auch bei ausdrücklich angelieferten Adressdifferenzen keine UPN-, Mail- oder Mail-Nickname-Änderung schreibt. `ExchangeOnline.Tests.ps1` prüft neue SMTP-Adressen, zusätzliche SMTP-/X500-Adressen, Vorschau, Rücklesefehler, Idempotenz und gewöhnliche Reparatur ohne Adressänderungen.

Die TDD-Regressionen wurden vor der Umsetzung ausgeführt und schlugen fehl. Alle Fernaufrufe der lokalen Tests sind synthetisch. Für eine separat autorisierte Mandantenabnahme nach Neuanlage `UserPrincipalName`, `Mail`, `PrimarySmtpAddress` und `EmailAddresses` prüfen. Ein Lehrer muss den passenden `smtp:`-Alias erhalten, bestehende Konten müssen ihre Adressen behalten. Domainfreigabe, reale Exchange-Replikation und Windows-Dateisperren sind lokal nicht belegt.

## Schüler-TUI 0.3.0

Die fokussierten Personal-/TUI-Tests prüfen die Schülerpopulation, aktuelle CLS-Gruppen und Manager-Anzeige. K und L durchlaufen Markierung, Zielauswahl, Namensbestätigung und Aktualisierung. Eine neue Klasse muss zuerst bestätigt werden, bevor alte CLS-Mitgliedschaften entfernt werden. Es dürfen ausschließlich `Department` und `OfficeLocation` als Benutzerattribute geschrieben werden. Schülerrolle, andere Gruppen, Lizenzen und ein gesperrter Status bleiben erhalten. Die erste Klasse ohne bisherige CLS-Zuordnung und bereits passende Zuordnungen werden gesondert geprüft.

Frisch verlorene Schülerrolle, neuer Adminschutz, neue Shared Mailbox, veränderte Ausgangsklasse beziehungsweise Manager-Zuordnung und geänderte Ziele müssen beide Aktionen vor dem Schreiben blockieren. Dynamische und lizenzierende Klassen sind geschützt. Fehlende Bestätigung der neuen Mitgliedschaft, fehlerhafte Attribute und nachträglich auftretender Adminschutz müssen Teilfehler melden und alte Gruppen erhalten. Vorschau, Esc und Nein dürfen nicht schreiben. Der gemeinsame Renderer muss Klasse, Jahrgang, Lehrer, K/L und UTC-Login zeigen. Tatsächliche CLI-Prozesse prüfen den eigenen Parametersatz und die Ablehnung umgeleiteter Eingabe.

Für die echte Tastaturprüfung ausschließlich synthetischer Daten:

```powershell
pwsh -NoProfile -File ./tests/Invoke-SyntheticPersonalCli.ps1 -Population Student -Mode Tui
```

Leertaste, K, Pfeil nach unten, Enter, Y prüfen. Danach erneut Leertaste, L, Pfeil nach unten, Enter, Y und Esc. Der Harness erzwingt `-WhatIf` und ersetzt sämtliche Fernbefehle. Windows-Darstellung und echte Mandantenänderungen sind gesonderte Integrationsprüfungen. Die TUI pflegt keine Quelldatei, eine veraltete Exceldatei kann Zuordnungen zurücksetzen und gesperrte Schüler beim bisherigen Sync reaktivieren. Ergebnisse stehen im [Abnahmeprotokoll](ABNAHME.md).

## Lehrer-Kontenverwaltung 0.3.0

Die fokussierten Personal-/TUI-Tests verwenden dieselben synthetischen Benutzer und Fernbefehle für die Lehrerpopulation. Sie prüfen aktive und gesperrte direkte Schulteam-/Ganztagsmitglieder, Ausschluss von Schülern, Shared Mailboxes, rein indirekten Mitgliedern und übrigen Personalkonten. B, U und Backspace oder Entf müssen die bestehenden bestätigten Aktionen ausführen. Frisch verlorene direkte Lehrerrolle, neue Admin-/Schülerrolle oder Shared-Mailbox-Zuordnung blockieren jede Aktion. Nach einer Sperrung muss der geladene Bestand im Lehrermodus bleiben. Rollen-, Status- und Login-Anzeige verwenden den gemeinsamen Renderer mit eigener Lehrerüberschrift.

Die tatsächlichen CLI-Tests prüfen den neuen Parametersatz in einem separaten Prozess. Ein umgeleitetes Terminal muss mit der Interaktivitätsmeldung abbrechen. `-ManageAccounts` mit `-Update`, `-ListToExcel` oder `-File` muss vor einer Verbindung abgewiesen werden. Für die echte Tastaturprüfung ausschließlich synthetischer Daten:

```powershell
pwsh -NoLogo -NoProfile -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Population Teacher -Mode Tui
pwsh -NoLogo -NoProfile -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Population Teacher -Mode Tui -DisabledFixture
```

Im ersten Aufruf Leertaste, B, Y und Esc prüfen. Im zweiten Leertaste, U, Y und nach der Aktualisierung erneut Leertaste, Backspace oder Entf, Y, Esc prüfen. Der Harness erzwingt `-WhatIf`, echte Verbindungen und Mutationen sind verboten. Beide Ansichten müssen `Lehrer- und Ganztagskonten`, Rolle und Login-Zeitstempel anzeigen. Die vollständige lokale Suite prüft anschließend die bisherigen Sync-, Export- und Personal-/Kontenverwaltungseinstiege. Windows und eine gesondert autorisierte Mandantenprüfung bleiben eigene Integrationsschritte.

## Schüler- und Lehrerexport 0.3.0

Die TUI-Anzeigetests prüfen `Last Login Timestamp (UTC)` im Detailbereich, Cursorwechsel zwischen Konten, UTC-Konvertierung, fehlende/nicht abrufbare Werte, ältere Ergebnisobjekte ohne Login-Felder sowie einen leeren Bestand. Personal- und bisheriger TUI-Modus verwenden dieselbe Anzeige. Der tatsächliche Personal-Einstieg wird zusätzlich im Pseudoterminal mit der synthetischen CLI-Fixture geöffnet und mit Esc verlassen.

Für `Last Login Timestamp` prüfen synthetische Tests alle drei Exportarten, UTC-Konvertierung, unveränderte Passwort-/Identitätsspalten, fehlende Werte, leeren Bestand und erfolgreiche Importe trotz Zusatzspalte. Die Rollenbestände verwenden ausschließlich die letzte erfolgreiche Anmeldung, keinen neueren fehlgeschlagenen Anmeldeversuch. Fehler beim Lesen von `signInActivity` müssen eine Warnung und einen erneuten Abruf ohne diese Eigenschaft auslösen, für Personal einschließlich des tatsächlichen Exportaufrufs. Die CLI-Fixture enthält einen synthetischen Login-Zeitstempel.

`WorkbookExport.Tests.ps1` prüft direkte Rollenpopulationen, gesperrte Konten, Shared-Mailbox-Ausschluss, Klassen und Manager-UPNs, kanonische Lehrerjobs, Ganztag, Mehrfachrollen, fehlende Zuordnungen und Bestandsfehler. Native XLSX-Tests prüfen zehn Schüler- beziehungsweise neun Lehrerspalten, den vorhandenen Importadapter, leere Passwörter und Formelzeichen als unveränderte Texte. Bestehende Personalexport- und Dateischutztests bleiben Teil der Gesamtsuite.

Der isolierte Harness `Invoke-SyntheticPersonalCli.ps1` unterstützt zusätzlich `-Population Student` und `-Population Teacher` im Exportmodus. Tatsächliche Prozesse prüfen beide Sync-Einstiege, neue Ziele, Ablehnung bestehender Dateien, unveränderte Dateihashes und `-WhatIf` ohne Datei oder Ordner. Die CLI-Prüfung weist widersprüchliche Selektoren vor einer Verbindung zurück. Alle Fernbefehle sind ersetzt, echte Verbindungen und Mutationen schlagen im Harness fehl.

```powershell
pwsh -NoLogo -NoProfile -NonInteractive -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Mode Export -Population Student -File "$([IO.Path]::GetTempPath())Schueler-Synthetic-Neu.xlsx"
pwsh -NoLogo -NoProfile -NonInteractive -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Mode Export -Population Teacher -File "$([IO.Path]::GetTempPath())Lehrer-Synthetic-Neu.xlsx" -Preview
```

Auch synthetische Ziele müssen neu sein. Wiederhole die Gesamtsuite unter Windows. Bei einer separat autorisierten Mandantenprüfung die direkte Rollenpopulation, Klassengruppen, Manager und Jobs kontrollieren. Ein Reimport mit Update kann im bisherigen Schüler-/Lehrerpfad enthaltene gesperrte Konten aktivieren. Lokale Ergebnisse stehen im [Abnahmeprotokoll](ABNAHME.md).

Für den Lehrer-/Ganztagsabgleich seit 0.2.0 die [Lehrerregeln](ENTRA-LEHRER-SYNC.md) zusätzlich prüfen. Benutze eine Arbeitskopie der synthetischen `tests/fixtures/Lehrer-Testdaten.xlsx` in einem isolierten Mandanten oder mit vollständig isolierten Personalrollen. Teste L, CO-L, PA und OGTS, den Wechsel zwischen Schul- und Ganztagsrolle, bestehende Direkt- und Gruppenlizenzen, die zwei Exchange-Profile und die Passwortregel mit genau zwölf Zeichen und erzwungenem Erstwechsel. Prüfe `-Remove` auf Deaktivierung und Sitzungswiderruf bei erhaltenem Konto, Postfach, Gruppen und Lizenzen. `-WhatIf` darf keine Mandanten- oder Excel-Mutation auslösen. Halte den Personaltest getrennt vom historischen Schülerprotokoll fest.

Der [Check gesperrter Konten](GESPERRTE-KONTEN-VERWALTUNG.md) lädt zunächst nur lesend. Die Pester-Tests prüfen aktive Konten, die Aufnahme von freigegebenen und Ressourcenpostfächern sowie Konten ohne Exchange-Empfänger, die eindeutige Zuordnung über Objekt-IDs, den Abbruch bei einem Exchange-Bestandsfehler, `UserMailbox` und `RemoteUserMailbox`, Gäste, synchronisierte Konten, Verzeichnisrollen, fehlende Anmeldedaten, besessene Entra-Objekte mit Typ, Name und ID aus Graph-`AdditionalProperties` und den abgeschlossenen Fortschrittsbalken. Die TUI-Tests prüfen das Ausblenden von `SharedMailbox`, Pfeile, Markierung mit Leertaste für persönliche Benutzerpostfächer und eindeutig empfängerlose Konten, Nein und Esc ohne Aktion, Backspace oder Entf zum Löschen, `U` zum Entsperren, Ja nur für markierte Konten sowie `-WhatIf`. Bei mehreren ausgewählten Konten müssen Löschen und Entsperren jeden Schritt im Fortschrittsbalken anzeigen und den Balken auch nach `Skipped` oder `Failed` abschließen. Beide Schreibpfade müssen vor jeder Aktion Identität, Sperrstatus, Exchange-Zuordnung und Verzeichnisrollen erneut lesen. Vor der Aktion bei einem empfängerlosen Konto muss ein neuer vollständiger Exchange-Bestand ohne Treffer für seine Objekt-ID geladen werden. Eine inzwischen zugeordnete `SharedMailbox`, eine unklare Zuordnung oder ein Exchange-Fehler muss die Aktion überspringen. Der Löschpfad nennt bekannte Besitzobjekte in der Bestätigung, fragt sie unmittelbar vor dem Löschen erneut ab und warnt bei Treffern. Eigentümerkonten dürfen nach Bestätigung gelöscht werden. Fehler der erneuten Besitzabfrage führen weiterhin zu `Skipped`. Der Entsperrpfad aktiviert Eigentümerkonten, ohne eine Besitzabfrage zu benötigen. Aktive Konten und freigegebene Postfächer werden bei beiden Aktionen übersprungen. Prüfe bei einem synthetischen Skriptaufruf, dass die TUI nur in einem interaktiven Terminal startet und `-List` und `-PassThru` lesend bleiben. Für einen kontrollierten Mandantentest zuerst ausschließlich `-List` oder `-WhatIf` verwenden und die Tenant-ID prüfen. Das Skript kann weder eine 90-Tage-Frist noch eine echte Person automatisch belegen. Vor einer echten Löschung muss der Betreiber beides außerhalb der TUI anhand verlässlicher Daten prüfen.

Für die Personalabnahme zuerst die Mitglieder beider Rollen und alle Abgänge im lesenden Vergleich kontrollieren. Prüfe die tatsächlichen SKUs der Lizenzgruppen und die Exchange-Richtlinien im Mandanten. Verwende für einen Test mit einer unvollständigen Personalliste ausschließlich `-Update -CreateNewUsers -UpdateUsers`. Vergleiche den Dateihash vor und nach `-WhatIf`. Führe den vollständigen Lebenszyklus erst mit vollständig isolierten Personalrollen aus und dokumentiere die Ergebnisse getrennt nach Graph, Exchange und Excel. Der lokale Testlauf ersetzt weder die Windows-Dateisperrprüfung noch den Mandantennachweis.

Prüfe bei Schülern und Personal einen vollständigen `-Update`-Lauf mit einem bereits konformen und einem abweichenden bestehenden Postfach. Für das konforme Postfach darf keine Konfigurationsbestätigung erscheinen und das Ergebnis muss `Compliant` sein. Für die Abweichung muss die Bestätigung vor der Exchange-Änderung erhalten bleiben. Auch eine erst nach dem Vergleich entstandene Abweichung muss bei der erneuten Prüfung berücksichtigt werden.

Prüfe zusätzlich `L, PA` und `PA, JAS` aus einer einzigen Excel-Zeile. Beide müssen mit dem Schulprofil verarbeitet werden, ohne ein zweites Konto oder eine zweite Standardlizenz anzulegen. `JAS` allein und eine Kombination aus Schul- und Ganztagsjob müssen vor jeder Mutation scheitern. Kontrolliere nach einer synthetischen Neuanlage die Rückschreibung von Passwort, Objekt-ID, UPN und Mail, während die Jobzelle erhalten bleibt.

Die Tests verwenden ausschließlich erfundene Konten. Führe keine Live-Mutation mit echten Schülerdaten aus.

Der aktuelle lokale Prüfstand und die auszufüllende Abnahme-Checkliste stehen in [ABNAHME.md](ABNAHME.md).

> [!CAUTION]
> Abgänge sind alle aktiven, nicht zugeordneten Mitglieder der konfigurierten Gruppe `SEC-A-ROL-Schule_Schüler`. Die Erkennung ist nicht auf eine Testklasse oder einen Arbeitsmappenausschnitt begrenzt. Für jede Abgangsaktion muss die Excel-Datei die vollständige konfigurierte Schülerpopulation enthalten. Eine Testklasse in einer gemeinsam genutzten produktiven Rollengruppe ist nicht isoliert und kann reale Schüler als Abgänge markieren. Live-Tests mit `-Update`, `-DisableUsers` oder `-RevokeSessions` sind deshalb nur in einem vollständig isolierten Mandanten oder mit einer konfigurierten Schüler-Rollengruppe erlaubt, deren Mitglieder ausschließlich synthetische Konten sind.

## Personal 0.3.0: lokaler und späterer Integrationsnachweis

`-IncompleteNameFixture` ergänzt sechs bestehende Mitgliedskonten mit dem gültigen Job `Kinderhaus Pädagogisches Team`, bestätigter Personalrolle und fehlenden Vor-/Nachnamen. Mit `-ReviewFixture -ProtectedFixture` prüfen tatsächliche isolierte Export-, Compare- und Update-WhatIf-Aufrufe weiterhin genau eine zulässige Jobtiteländerung und acht Prüfhinweise (Admin, rollenloses Konto, sechs Konten ohne Namen). Der Gast bleibt ohne Datenwarnung. Synthetische Updates prüfen die tatsächlichen Graph-Schreibparameter des vollständigen Kontos und keine Mutation der unvollständigen Konten. Native Exceltests prüfen Objekt-ID, UPN, kombinierte Identität sowie fehlenden Vor-, Nach- oder Anzeigenamen. Unbekannte oder widersprüchliche Identitäten, unbekannte oder fehlende Jobs verwalteter Konten, rollenlose Übernahme ohne Namen, mehrere Personalrollen, unlesbare Mitgliedschaften, unvollständiger Schutzkontext und unbekannter Kontotyp bleiben Fehler. Quelldateihashes bleiben unverändert. Dateien ausschließlich mit bestätigten Bestandskonten ohne vollständige Namensdaten dürfen keine Abgänge oder Aktionen erzeugen.

Der Global-Administrator-Gruppenname ist `SEC-A-ADM-GlobalAdministrators`. Die synthetischen Fixtures verwenden diesen Namen. Zusätzliche Schutztests prüfen direkte Mitglieder in Personal-, Lehrer- und Schülerpopulation, einen ausschließlich vorhandenen veralteten Rollengruppennamen und einen mehrdeutigen aktuellen Namen. Ein fehlender oder mehrdeutiger aktueller Name muss Schreibaktionen sperren.

Die Regression nach einem Gesamtexport wird mit `Invoke-SyntheticPersonalCli.ps1 -ReviewFixture` geprüft. Der Harness exportiert ein zulässiges Personalkonto, ExterneAdmins, ein rollenloses Konto ohne Namensbestandteile und einen Gast mit freigegebenem Job. Ein tatsächlicher anschließender Compare-Aufruf muss das zulässige Konto vergleichen und zwei `Workbook.Review`-Warnungen für Admin- und rollenlose Zeilen liefern. Der Gast hat unvollständige Namen und wird ohne Datenwarnung übersprungen. Mit zusätzlichem `-ProtectedFixture` enthält das zulässige Konto einen technischen Jobtitel. Compare und Update-WhatIf mit derselben Datei müssen die Jobtiteländerung anzeigen beziehungsweise einplanen. Das rollenlose Mitgliedskonto mit leerem Job wird als `Workbook.Review` übersprungen, der Dateihash muss unverändert bleiben. Native Exceltests prüfen Objekt-ID, UPN, kombinierte Identität und vollständig leere Pflichtzellen. Die Ausnahme für Konten ohne Job und Personalrolle gilt nicht für bestehende Personalrollen, unvollständigen Schutzkontext, unlesbare Mitgliedschaften, unbekannte Kontotypen, unbekannte oder widersprüchliche Identitäten oder einen ausdrücklich eingetragenen Job. Für solche Zeilen gelten die reguläre Pflichtdatenprüfung und gegebenenfalls die gesonderte Prüfung bestehender Personalkonten ohne vollständige Namensdaten. Synthetische Updates prüfen die tatsächlichen Graph-Schreibparameter des gültigen Personalkontos und das Ausbleiben von Änderungen am rollenlosen Konto. Dateien ausschließlich mit rollenlosen Konten ohne Job dürfen keine Abgänge oder Aktionen auslösen. Adaptertests prüfen die bewusst aktivierte Vergleichstoleranz, den weiterhin strikten Standardimport und die Ablehnung dieses Modus für Lehrer. Vergleichstests prüfen Adminschutz, lokale AD-Synchronisierung, Shared Mailboxes und den Ausschluss bekannter Prüfzeilen aus Abgängen. Gastregressionen prüfen Objekt-ID, UPN und kombinierte Identität, auch bei vollständig leeren Pflichtzellen. Interne, unbekannte, widersprüchliche und nicht identifizierte Konten behalten die Validierung. Tatsächliche Compare- und Update-WhatIf-Aufrufe mit einer Gastdatei dürfen weder Datenwarnungen noch Fehler oder Aktionen liefern, ihr Dateihash bleibt unverändert. Gemischte Listen bearbeiten interne Konten weiterhin. Dateien ausschließlich mit Gästen dürfen keine Abgänge interner Konten auslösen. `-ProtectedFixture` enthält ein zulässiges Konto mit technischem Jobtitel und einen bestätigten externen Admin. Der tatsächliche Update-WhatIf-Aufruf muss die Jobtiteländerung als Update einplanen und das Adminkonto mit Prüfhinweis überspringen. Synthetische echte Updates prüfen die Graph-Schreibparameter und anschließende Änderungsfreiheit. Native Exceltests prüfen ExterneAdmins, GlobalAdministrators und Verzeichnisrollen, ungeschützte Adminjob-Zeilen sowie unbekannte oder widersprüchliche Identitäten. Reine Adminkonto-Dateien dürfen keine Abgänge interner Konten auslösen.

Die [Personal-Anleitung](ENTRA-PERSONAL-SYNC.md) und [Spezifikation](PERSONAL-0.3.0-SPEC.md) gelten für den neuen Einstieg. `PersonalSync.Tests.ps1` prüft alle sieben schreibbaren Profile, die beiden geschützten Adminrollen, transitive Ausschlüsse, rollenlose aktive und gesperrte Konten, Gäste, lokal synchronisierte Benutzer, Ressourcen- und Shared Mailboxes, doppelte Empfänger und Fehler beim vollständigen Lesen.

Die Sync-Tests prüfen die explizite Übernahme mit Objekt-ID, Erhalt des Sperrstatus, ausdrückliches `-ResumeCreation`, Rollenwechsel bei erhaltenen fremden Gruppen, vorhandene und ausstehende Lizenzen, eine nach dem Vergleich hinzugekommene Lizenz, Neuanlage mit frühem Checkpoint, Checkpointfehler und Wiederanlaufbefehl. Personal-`-Update` allein darf keine Abgänge sperren. Explizites `-DisableUsers` betrifft nur verwaltete direkte Rollenmitglieder. Auch zwischen Mailbox- und CAS-Schreiben wird der Personal-Schutz erneut geprüft.

Native Exporttests prüfen neun Spalten, Tabelle und Blatt `Tabelle1`, `TableStyleMedium14`, Spaltenbreite, IDs, leere Passwörter, rollenlose Zeilen, Formelzeichen als Text, einen leeren Bestand und kein Dateiziel unter `-WhatIf`. Der vorhandene Excel-Dateischutz bleibt über die gesamte Suite geprüft. ABP-Leerwerte werden im tatsächlichen Vergleich und in Schreibparametern ausgelassen.

`Invoke-SyntheticPersonalCli.ps1` ersetzt jeden Fernbefehl. Verbindung und Mutation werfen bei versehentlichem Aufruf einen Fehler. Damit prüft die Suite tatsächliche separate CLI-Prozesse für Liste, Objekt, Export, Vergleich, Update-Vorschau, Add, Remove, Exchange, UPN-Aliase, Fehlerstatus und nicht interaktive TUI-Ablehnung. Für eine manuelle Tastaturprüfung ausschließlich synthetischer Daten:

```powershell
pwsh -NoLogo -NoProfile -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Mode Tui
pwsh -NoLogo -NoProfile -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Mode Tui -DisabledFixture
```

Im ersten Aufruf Leertaste, `B`, `Y`, danach `Esc` prüfen. Im zweiten Leertaste, `U` beziehungsweise Backspace oder Entf, `Y`, anschließend `Esc`. Diese echten TUI-Aufrufe sind ausschließlich `-WhatIf`. Pester prüft zusätzlich synthetische Schreibresultate, gemischte Auswahl, Nein, Esc, Teilfehler und fehlgeschlagene Besitzabfrage. Wiederhole die Tastaturprüfung unter Windows. Windows-Dateisperren und echte Graph-/Exchange-Provisionierung bleiben eigene offene Integrationsprüfungen.

Für den Gast-Löschpfad einen aktiven Gast markieren und Backspace oder Entf drücken. Erwartung: konkreter Hinweis zum vorherigen Sperren mit B, erhaltene Markierung und Cursor, kein neuer Bestandsabruf. B muss für diese Auswahl sofort nutzbar sein. Bei N oder Esc in der Bestätigung bleibt die Auswahl ebenfalls erhalten. Mit synthetischen Schreibstubs anschließend B, Y, erneutes Markieren, Backspace oder Entf und Y prüfen. Der Gast muss genau einmal gesperrt und gelöscht werden. Für einen bereits gesperrten synthetischen Gast muss Backspace oder Entf sofort die Löschbestätigung öffnen. Die tatsächliche Terminalprüfung verwendet weiterhin ausschließlich WhatIf.

Bei einer später ausdrücklich autorisierten Mandantenabnahme ausschließlich isolierte synthetische Konten verwenden. Prüfe Namen und IDs aller Ausschluss- und Schutzgruppen, die A1Faculty-SKU, Richtlinien und Adresslistenfilter. Teste die gesamte Lebensdauer mit einem aktiven, einem gesperrten und einem rollenlosen Konto. Ein Export darf erst nach bereinigter Profilzuordnung als Eingabe für einen schreibenden Sync dienen. Vollständige Abgänge benötigen eine ausschließlich synthetische, vollständig isolierte Rollenpopulation.

Die zusätzlichen TUI-Tests prüfen die Rollenanzeige mit mehreren und unbekannten Rollengruppen, ohne Rollengruppe und in beiden Verwaltungsmodi. Gasttests durchlaufen Markieren, Sperren, Bestandsaktualisierung, erneutes Markieren und Löschen für empfängerlose Gäste sowie `UserMailbox`, `MailUser` und `GuestMailUser`. Frische Änderungen von Administrator-/Ausschluss-/Verzeichnisrollen, Identität, Kontotyp, lokaler AD-Synchronisierung, Shared Mailbox und mehrdeutigen Empfängern müssen beide Aktionen blockieren. Gastentsperrung, normale Profilprüfung und der bisherige Löschpfad bleiben gesperrt.

Für die tatsächliche Gast-Tastaturprüfung mit synthetischen Daten:

```powershell
pwsh -NoLogo -NoProfile -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Mode Tui -GuestFixture
pwsh -NoLogo -NoProfile -File .\tests\Invoke-SyntheticPersonalCli.ps1 -Mode Tui -GuestFixture -DisabledFixture
```

Im ersten Aufruf Leertaste, `B`, `Y`, `Esc` prüfen. Im zweiten Leertaste, Backspace oder Entf, `Y`, `Esc`. Beide Aufrufe erzwingen `-WhatIf`, der Harness verbietet alle Fernänderungen. Die synthetischen Pester-Mutationen prüfen zusätzlich den tatsächlichen Statuswechsel.

## Freigabekriterien

Die Implementierung ist erst für den produktiven Einsatz freigegeben, wenn:

- alle Pester-Tests in PowerShell 7 unter Windows bestehen
- PSScriptAnalyzer keine ungeklärten Fehler meldet
- alle PowerShell-Dateien ohne Parserfehler sind
- Standardvergleich und `-WhatIf` nachweislich nichts schreiben
- Neuzugang, Änderung, Abgang und Exchange-Reparatur mit synthetischen Konten geprüft wurden
- keine produktive XLSX oder Passwortdatei von Git verfolgt wird

## 1. VM und Module dokumentieren

```powershell
$PSVersionTable

$required = @(
  'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users',
  'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Groups',
  'ExchangeOnlineManagement', 'ImportExcel', 'Pester', 'PSScriptAnalyzer'
)
Get-Module -ListAvailable $required |
  Sort-Object Name,Version -Descending |
  Select-Object Name,Version,Path |
  Format-Table -AutoSize
```

Dokumentiere Windows-Version, PowerShell-Version und die höchste gefundene Version jedes Moduls im Abnahmeprotokoll.

## 2. Repository-Tests

Im Repository-Root:

```powershell
pwsh -NoLogo -NoProfile -NonInteractive -File .\tests\Test-PesterDiscovery.ps1
Invoke-Pester .\tests -Output Detailed

$analyzerFindings = @(
  Invoke-ScriptAnalyzer -Path '.\Sync-SchuelerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\Sync-LehrerEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\Sync-PersonalEntra.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\Verwalte-GesperrteEntraKonten.ps1' -Settings '.\PSScriptAnalyzerSettings.psd1'
  Invoke-ScriptAnalyzer -Path '.\src' -Recurse -Settings '.\PSScriptAnalyzerSettings.psd1'
)
$analyzerFindings | Format-Table RuleName,Severity,ScriptName,Line,Message -Wrap
$unexpectedFindings = @($analyzerFindings | Where-Object RuleName -ne PSAvoidUsingWriteHost)
if ($unexpectedFindings) { throw 'Ungeklärte PSScriptAnalyzer-Befunde gefunden.' }

$parseErrors = foreach ($file in Get-ChildItem . -Recurse -File | Where-Object Extension -in '.ps1','.psm1','.psd1') {
  $tokens = $null
  $errors = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile(
    $file.FullName,
    [ref]$tokens,
    [ref]$errors
  )
  $errors
}
$parseErrors | Format-List
if ($parseErrors) { throw 'PowerShell-Parserfehler gefunden.' }
```

Erwartung: Pester erfolgreich, Analyzer ohne ungeklärte Fehler, `$parseErrors` leer.

Die separate Discovery-Prüfung startet ohne vorab importiertes Anwendungsmodul. Tests mit `InModuleScope` importieren das Modul ausdrücklich in `BeforeDiscovery`. Fixturepfade, Zustände und Mocks entstehen erst in der Run-Phase in `BeforeAll` beziehungsweise `BeforeEach`. Discovery-Blöcke verwenden keine dort erst entstehenden Werte. Testdateien ersetzen das Modul nicht mit `Import-Module -Force`, da das bereits gebundene Scriptblöcke ungültig machen könnte. Der Discovery-Test prüft diese Struktur und startet zusätzlich selbst eine frische PowerShell-Sitzung.

## 3. Git- und Datenschutzprüfung

```powershell
git status --short
git ls-files '*.xlsx'
git check-ignore .\Schueler.xlsx
git grep -n -I -E 'Passwort|PasswordProfile' -- ':!docs/**' ':!tests/**'
```

Erwartung:

- `Schueler.xlsx` wird ignoriert
- eigene Root- und Unterordner-Arbeitsmappen sowie konkrete Sicherungs- und temporäre Pfade werden ignoriert
- unter Git liegen ausschließlich die synthetischen Fixtures `tests/fixtures/Schueler-Testdaten.xlsx` und `tests/fixtures/Lehrer-Testdaten.xlsx`
- keine produktiven Namen, UPNs oder Passwörter sind versioniert
- Quellcode-Treffer für Passwortlogik enthalten keine echten Geheimnisse

Die Git-Regressionen verwenden temporäre Test-Repositories. Sie prüfen auch den Fehlerfall, dass nur die Quelle ignoriert ist. Dann darf weder eine vertrauliche Kopie noch eine Sicherung entstehen. Der Passworttest prüft das Format, die Einmaligkeit im Lauf und mehr als 64 Bit tatsächlich erreichbaren Auswahlraum.

## 4. Synthetische Testdaten

Verwende erfundene Personen, zum Beispiel:

| Fall | Vorname | Nachname | Klasse | Klassenlehrer |
|---|---|---|---|---|
| Neuzugang | Lina | Testwald | `JK1-3g2_1` | UPN eines Test-Managers |
| UPN-Kollision | Luis | Testwald | `JK1-3g2_1` | UPN eines Test-Managers |
| Klassenwechsel | Mira | Beispielstern | `JK4-6m2_4` | UPN eines Test-Managers |
| Abgang | Theo | Demoklang | nur im Mandanten | UPN eines Test-Managers |

Alle Testkonten und Gruppen müssen eindeutig als synthetisch gekennzeichnet und nach der Abnahme kontrolliert bereinigt werden. Nutze keine Namen existierender Kinder oder Beschäftigter. Erstelle diese Live-Testpopulation nur im vollständig isolierten Mandanten oder in der ausschließlich synthetischen konfigurierten Schüler-Rollengruppe. Eine bloße Testklasse innerhalb der produktiven Rollengruppe reicht nicht aus.

Erzeuge eine Arbeitskopie außerhalb des Repository-Roots:

```powershell
$testRoot = 'C:\Temp\EntraSchuelerSync-Abnahme'
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
Copy-Item .\tests\fixtures\Schueler-Testdaten.xlsx (Join-Path $testRoot 'Schueler-Abnahme.xlsx')
$testFile = Join-Path $testRoot 'Schueler-Abnahme.xlsx'
```

Passe nur die erfundenen Zeilen und den vorgesehenen Test-Manager an.

## 5. Lesenden Standardmodus beweisen

```powershell
$before = Get-FileHash $testFile -Algorithm SHA256
$result = .\Sync-SchuelerEntra.ps1 -File $testFile
$after = Get-FileHash $testFile -Algorithm SHA256

$before.Hash -eq $after.Hash
$result.Mode
$result.HasErrors
```

Erwartung:

- Hash bleibt identisch
- `Mode` ist `Compare`
- keine Graph- oder Exchange-Schreiboperation
- kein Warten auf fehlende Postfächer
- Tabellen für Neuzugänge, Abgänge, Änderungen und Bestehende

## 6. WhatIf beweisen

```powershell
$before = Get-FileHash $testFile -Algorithm SHA256
$result = .\Sync-SchuelerEntra.ps1 -File $testFile -Update -WhatIf
$after = Get-FileHash $testFile -Algorithm SHA256

$before.Hash -eq $after.Hash
$result.Actions | Format-Table Phase,Status,UserPrincipalName
```

Erwartung:

- Excel-Hash bleibt identisch
- keine neuen Konten, Attribute, Manager oder Gruppen
- keine Passworterzeugung, kein Backup und keine Berichtsdatei, auch mit `-OutputFile`
- keine Exchange-Schreiboperation
- keine 60-Sekunden-Wartezeit
- geplante Phasen tragen Status `WhatIf`

Wiederhole mit selektiven Schaltern:

```powershell
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -CreateNewUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -UpdateUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -DisableUsers -RevokeSessions -WhatIf
```

## 7. Neuzugang live testen

Dieser Schritt verändert den Testmandanten. Prüfe zuerst Tenant-ID, UPN und Zielgruppen.

```powershell
$result = .\Sync-SchuelerEntra.ps1 -File $testFile -Update -CreateNewUsers
$result.Actions | Format-Table Phase,Status,UserId,UserPrincipalName,RecoveryCommand -Wrap
```

Prüfe für den erfundenen Neuzugang:

```powershell
$row = Import-Excel $testFile | Where-Object Vorname -eq 'Lina'
$row.Passwort.Length
$user = Get-MgUser -UserId $row.EntraObjectId -Property id,displayName,givenName,surname,userPrincipalName,mail,mailNickname,department,officeLocation,companyName,employeeType,usageLocation,ageGroup,consentProvidedForMinor,legalAgeGroupClassification,accountEnabled
$user | Format-List
Get-MgUserManager -UserId $user.Id
Get-MgUserMemberOfAsGroup -UserId $user.Id -All | Select-Object Id,DisplayName
```

Erwartung:

- Passwortlänge genau 12
- zwölf zufällige, gut unterscheidbare ASCII-Buchstaben und Ziffern mit mindestens einem Zeichen jeder Klasse
- `ForceChangePasswordNextSignIn = false`, über das erstellte PasswordProfile oder einen kontrollierten Anmeldetest bestätigt
- Name bleibt mit Umlauten korrekt, UPN ist normalisiert
- Konto erst nach erfolgreicher Excel-Rückschreibung aktiviert
- Gruppen `SEC-A-LIC-O365A1Student`, `SEC-A-ROL-Schule_Schüler` und `SEC-A-CLS-JK1-3g2_1`
- Office Location `G2`
- `ageGroup = Minor`, `consentProvidedForMinor = Granted`, `usageLocation = DE`
- `legalAgeGroupClassification = MinorWithParentalConsent` wird nur gelesen
- genau eine Backup-Datei enthält den vorherigen Arbeitsmappenstand

Melde das Testkonto einmal mit dem Initialpasswort an. Es darf kein erzwungener Passwortwechsel erscheinen.

### Manuellen Add-Modus testen

Verwende ausschließlich ein erfundenes Testkonto und starte ohne Transcript:

```powershell
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin' -WhatIf
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin'
```

Erwartung bei der Neuanlage:

- keine Excel-Datei wird gelesen oder verändert
- das Passwort hat genau 12 Zeichen und erscheint genau einmal im Terminal, auch bei einem simulierten Teilfehler nach der Kontoerstellung
- Ergebnisobjekt und `-OutputFile` enthalten das Passwort nicht
- Konto, Pflichtattribute, Manager und Gruppen werden vor der Aktivierung verifiziert
- Exchange wird sofort und danach höchstens fünfmal mit jeweils 60 Sekunden Wartezeit geprüft

Führe denselben Befehl danach erneut mit einer geänderten Klasse aus. Erwartung: Der vorhandene Schüler wird aktualisiert. Es entsteht kein weiteres Konto und kein neues Passwort. Der vorhandene UPN bleibt unverändert. Deaktiviere das synthetische Konto und wiederhole `-Add`. Das Konto darf erst nach geprüftem Pflichtzustand aktiviert werden. Simuliere außerdem einen Fehler nach der Kontoerstellung und prüfe den ausgegebenen Wiederanlauf mit `-EntraObjectId`.

## 8. UPN-Kollision testen

Reserviere den ersten UPN-Kandidaten mit einem synthetischen Benutzer oder Exchange-Empfänger. Führe den Vergleich aus und prüfe:

- Warnung enthält den kollidierenden Kandidaten und den gewählten Ersatz
- Präfix wächst vom ersten zum zweiten Buchstaben und danach weiter
- nach vollständigem Vornamen beginnt der numerische Fallback bei `2`
- Kollisionsquellen aus UPN, Mail, SMTP-Proxyadresse und Exchange-Empfänger werden erkannt
- eine Adresse desselben Objekt-IDs wird wiederverwendet

Nutze den tatsächlich erzeugten UPN für alle weiteren Prüfungen.

## 9. Änderung und Gruppennormalisierung testen

Ändere die Klasse eines synthetischen bestehenden Schülers von `JK1-3g2_1` auf `JK4-6m2_4`. Füge ihm außerdem direkt eine fremde statische `SEC-A-ROL-...`-Gruppe und eine alte `SEC-A-CLS-...`-Gruppe hinzu.

```powershell
.\Sync-SchuelerEntra.ps1 -File $testFile
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -UpdateUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -UpdateUsers
```

Erwartung:

- `department = JK4-6m2_4`
- `officeLocation = M2`
- neue Klassengruppe wurde vor Entfernung der alten Gruppe bestätigt
- nur `SEC-A-ROL-Schule_Schüler` bleibt als direkte Rollengruppe
- nur `SEC-A-CLS-JK4-6m2_4` bleibt als direkte Klassengruppe
- Lizenzgruppe bleibt vorhanden
- Passwortspalte bleibt bytegenau unverändert
- dynamische oder geerbte Testgruppe wird nicht entfernt und als Warnung ausgegeben

## 10. Abgang testen

Entferne nur die erfundene Abgangszeile aus der Testarbeitsmappe. Mindestens eine andere gültige Zeile muss verbleiben, eine leere Schülerliste wird abgewiesen. Prüfe zuerst Compare und WhatIf, danach Sperren und Sitzungswiderruf im selben Lauf:

```powershell
.\Sync-SchuelerEntra.ps1 -File $testFile
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -DisableUsers -RevokeSessions -WhatIf
.\Sync-SchuelerEntra.ps1 -File $testFile -Update -DisableUsers -RevokeSessions
```

Erwartung:

- Konto ist deaktiviert
- Sitzungswiderruf wurde von Graph bestätigt
- Konto, Lizenz und Gruppen wurden nicht gelöscht
- ein erneuter Vergleich zeigt das deaktivierte Konto ohne Excel-Zeile nicht mehr als Abgang, auch wenn es weiterhin Mitglied der Schüler-Rollengruppe ist
- ein erneuter lesender Lauf führt keine Mutation aus

### Manuellen Remove-Modus testen

Verwende einen erfundenen UPN, der direkt Mitglied der Schüler-Rollengruppe ist:

```powershell
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com'
```

Erwartung: Genau dieses Konto wird deaktiviert und seine Sitzungen werden widerrufen. Der Benutzer wird nicht gelöscht. Gruppen und Lizenzen bleiben erhalten. Wiederhole den Test mit einem Benutzer außerhalb der direkten Schüler-Rollengruppe. Dieser Lauf muss vor jeder Mutation fehlschlagen.

## 11. Exchange Online testen

### Sofort vorhandenes Postfach

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com'
```

Lies danach `Get-Mailbox` und `Get-CASMailbox` aus und gleiche jeden Wert mit [ENTRA-SCHUELER-SYNC.md](ENTRA-SCHUELER-SYNC.md) ab. Prüfe beide AuditAdmin-Varianten, einmal ohne und einmal mit `CommunicationsCompliance`.

### Verzögerte Bereitstellung

Verwende einen synthetischen Neuzugang, dessen Postfach bei der ersten Prüfung noch fehlt. Miss Zeit und Aufrufe. Erwartung:

- eine sofortige Prüfung
- höchstens fünf weitere Prüfungen
- 60 Sekunden zwischen den Wiederholungen
- nur eine Wartezeit pro Batch-Runde
- gefundene Postfächer werden aus dem Batch entfernt
- nach sechs erfolglosen Prüfungen Status `Pending` mit fertigem Reparaturbefehl

### Schreibfehler

Entziehe in einer kontrollierten Testkonstellation einen benötigten Exchange-Parameter oder simuliere ihn ausschließlich im Pester-Test. Erwartung: `Set-Mailbox` oder `Set-CASMailbox` meldet sofort `Failed`. Der Schreibfehler wird nicht als vermeintlich fehlendes Postfach wiederholt.

### Lehrerbereitstellung und Wiederanlauf

Die Lehrerregressionen verwenden synthetische Empfänger und ersetzen Wartezeiten. Prüfe die sofortige Prüfung und höchstens drei Wiederholungen mit jeweils 150 Sekunden, nur eine Wartezeit je Batch-Runde. Fehlende Mailbox- und CAS-Objekte sowie eindeutige Exchange-Empfängerfehler mit generischen Ausnahmeinformationen müssen als Bereitstellungsverzögerung gelten. Auch ein solcher Fehler bei `Set-Mailbox`, `Set-CASMailbox` oder Rückleseprüfung muss im Lehrerpfad wiederholbar bleiben. Ein nach vier Prüfungen fehlendes Postfach erhält `Pending` mit `-ConfigureExchangeOnlineOnly -UPN` als Reparaturbefehl. Parameterfehler, Verbindungsprobleme und verbleibende Konfigurationsabweichungen bleiben Fehler. WhatIf darf weder warten noch schreiben.

Der isolierte CLI-Harness ruft `Invoke-SyntheticPersonalCli.ps1 -Population Teacher -Mode Exchange` auf und prüft den echten Lehrereinstieg mit einem bestehenden synthetischen Konto unter WhatIf. Jeder produktive Verbindungs- und Schreibbefehl ist ersetzt oder gesperrt. Reale 150-Sekunden-Wartezeiten und Exchange-Replikation sind durch diese lokalen Tests nicht belegt. Schüler- und übrige Personal-Konfiguration behalten fünf Wiederholungen mit jeweils 60 Sekunden.

## 12. Idempotenz

Nach einem vollständig erfolgreichen Testlauf:

```powershell
$before = Get-FileHash $testFile -Algorithm SHA256
$result = .\Sync-SchuelerEntra.ps1 -File $testFile
$after = Get-FileHash $testFile -Algorithm SHA256

$result.Comparison.ChangedStudents.Count
$before.Hash -eq $after.Hash
```

Erwartung: keine Änderungen für aktive, vollständig konforme Schüler und identischer Excel-Hash. Ein deaktiviertes Konto ohne Excel-Zeile darf nicht erneut als Abgang erscheinen. Ein deaktiviertes Konto mit eindeutig passender Excel-Zeile muss dagegen als Reaktivierung unter Änderungen erscheinen.

## 13. Aufräumen

Vor dem Löschen erfundener Konten sichere das Abnahmeprotokoll ohne Passwörter. Entferne danach ausschließlich die dokumentierten synthetischen Objekte. Lösche keine produktiven Benutzer oder Gruppen.

```powershell
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false

Get-ChildItem $testRoot -Force
```

Entferne die lokale Testarbeitsmappe, Backups und temporären Dateien nach der vereinbarten sicheren Löschregel. Prüfe abschließend:

```powershell
git status --short
git ls-files '*.xlsx'
```

## Abnahmeprotokoll

Halte fest:

- Datum, VM, Windows- und PowerShell-Version
- getestete Modulversionen
- Tenant-ID und Testbereich, ohne Tokens oder Passwörter
- Commit-ID des getesteten Stands
- Ergebnis jeder Sektion 1 bis 13
- IDs ausschließlich der erfundenen Testobjekte
- offene Abweichungen und verantwortliche Freigabe

## Lokale Gesamtprüfung und CI

Die Personalregression `preserves skipped Excel rows during real personnel ... writes` verwendet die tatsächlichen XLSX-Lese- und Schreibfunktionen zusammen mit synthetischen Graph-/Exchange-Befehlen. Eine gemischte Datei enthält zwei Neuanlagen mit `Externe Benutzer`, eine bestehende Personalzeile ohne vollständige gespeicherte Identität, `Externe Admins` in Zeile 5, einen Gast, ein rollenloses Konto und ein bestehendes Personalkonto mit fehlendem Vornamen. Neuanlage und Identitätsnachtrag müssen ihre Checkpoints schreiben, die Aktivierung darf erst nach rückgelesenem Passwort und Objekt-ID erfolgen. Alle übersprungenen Zellen bleiben unverändert. Direkte Rückschreibungen auf diese Zeilen werden vor dem ersten Schreiben des gesamten Batches abgewiesen. WhatIf erhält den Dateihash und erzeugt keine Konten oder Rückschreibungen.

```powershell
./tests/Test-PesterDiscovery.ps1
./tests/Invoke-LocalVerification.ps1
```

Der gemeinsame Prüflauf parst die PowerShell-Quelldateien, prüft sie mit PSScriptAnalyzer und führt Pester aus. Nur die bewusst eingesetzten farbigen Überschriften (`PSAvoidUsingWriteHost`) sind als Analyzer-Ausnahme zugelassen. Jeder andere Analyzer-Befund und jeder fehlgeschlagene Test lässt den Lauf scheitern.

`.github/workflows/powershell-tests.yml` führt Discovery und Gesamtprüfung auf Windows und macOS aus. Die Modulversionen sind festgelegt. Der Workflow hat ausschließlich Leserechte auf Repository-Inhalte und benötigt keine Mandantenzugangsdaten. Er wird lokal vorbereitet, ein erfolgreicher lokaler Lauf belegt keinen ausgeführten GitHub-Workflow.

Die Audit-Regressionen prüfen leere Imports, einen fremden oder mehrdeutigen Exchange-Mandanten, erneute Mandantenprüfung vor der TUI-Aktion, nicht mehr gültige Rollenmitgliedschaften, nachträglich lizenzierende Rollen, gesperrte Wiederanlaufkonten, fehlgeschlagene Excel-Sicherung, abgelehnte Lehrerupdates, manuelles Add ohne falsche Abgänge und vertrauliche Berichtspfade. Der Berichtsschutz wird auch gegen bereits vorhandene Dateien und nicht existierende Unterverzeichnisse in temporären Git-Repositories geprüft.
