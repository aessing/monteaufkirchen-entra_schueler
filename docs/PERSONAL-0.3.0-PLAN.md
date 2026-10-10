# Personal 0.3.0: Umsetzungsplan

Ergänzung vom 10. Oktober 2026: Die beauftragten [Adressregeln](ADRESSREGELN.md) verwenden die gemeinsamen Normalisierungs- und Vergleichsfunktionen. Neue Lehrer-/Personaladressen erhalten `montessori-erding.de`, neue Schulteamkonten zusätzlich den passenden Alias unter `monteaufkirchen.com`. Freigegebene `Externe*`-Rollen verwenden `vorname.nachname`. Hauptadresse und Alias werden gemeinsam reserviert. Bestehende Adressfelder bleiben erhalten, der Schüler-Schreibpfad lässt diese Felder aus. Exchange-Adresswerte werden ausschließlich für Neuanlagen übergeben und verifiziert. Die TDD-Prüfung umfasst den bisherigen Gesamtprüflauf und die neuen Adressregressionen. Ergebnisse stehen im [Abnahmeprotokoll](ABNAHME.md).

Stand: 10. Oktober 2026. Grundlage: [Spezifikation](PERSONAL-0.3.0-SPEC.md) und [Variantenentscheidung](PERSONAL-0.3.0-BRAINSTORMING.md).

## Aktueller Stand

- [x] Ausgangsstand 0.2.0 und vorhandene Implementierung geprüft.
- [x] Eigenen Branch `codex/0.3.0-personal` von sauberem `main` bei `c5de9c2` angelegt.
- [x] Entwicklungsstand im Modul, README und Changelog auf 0.3.0 gesetzt.
- [x] Brainstorming, Spezifikation und Umsetzungsplan erstellt.
- [x] Lehrerdatei nur auf Struktur untersucht, keine Personendaten übernommen.
- [x] Fachentscheidungen zu Adminrollen, Sync-Verhalten und neuen Jobwerten bestätigt.
- [x] Neue Personalfunktionen implementiert und lokal mit synthetischen Daten geprüft.
- [x] Adressregeln für Neuanlagen und Bestandsschutz umgesetzt.
- [x] Gemeinsame Konten-TUIs mit letzter Befehlszeile sowie Backspace/Entf zum Löschen geprüft.
- [x] Lesbare Personaljobs im Export, in Bestandsliste und Vergleich sowie im tatsächlichen Entra-`JobTitle` geprüft. Kinderhauspräfix nur aus dem Entra-Jobtitel entfernt, Firma und eindeutige Excel-Jobnamen erhalten. Rückzuordnung und Kompatibilität bisheriger Schreibweisen erhalten.
- [x] Eindeutig identifizierte Entra-Gäste aus Pflichtdatenwarnungen und Fehlern in Compare und Update ausgenommen. Interne Datenprüfung sowie Gastschutz und TUI-Aktionen erhalten.
- [x] Bedienungsanleitung, bestehende Dokumentation, README, Changelog und Abnahme aktualisiert.
- [ ] Separate Windows- und kontrollierte Mandantenabnahme (nicht Bestandteil der lokalen Umsetzung).

Die lokale Umsetzung ist abgeschlossen. Historische Abnahmestände bleiben erhalten. Der aktuelle Prüfnachweis steht in [ABNAHME.md](ABNAHME.md), der ursprüngliche Planungsnachweis am Ende dieses Dokuments. Die folgenden sieben Schritte sind umgesetzt.

## 1. Konfiguration und gemeinsame Rollenregeln

Status: lokal umgesetzt und geprüft.

**Ziel:** Profile und Ausschlüsse eindeutig beschreiben, ohne Konfiguration zu duplizieren.

- `config/PersonalSync.psd1` mit den gelieferten acht Rollenprofilen ergänzen. Adminschutz entsprechend Fachentscheidung abbilden.
- Lehrer-/Ganztagsrollen aus `config/LehrerSync.psd1` und Schülerrolle über `Get-ConfiguredStudentRole` beziehen.
- Profilauflösung aus `Get-LehrerProfile` so erweitern oder gezielt herauslösen, dass der neue Kontext dieselbe Struktur für Attribute, Lizenz und Exchange erhält. Lehrerjobs und neue Personaljobs getrennt validieren.
- Rolle `GlobalAdministrators` anhand des exakten eindeutigen Namens für den Schutz auflösen, später die verifizierte ID festhalten. Kein Zielprofil erfinden.
- Die Verwaltung auf explizite Rollen-IDs begrenzen. Ausschlüsse gelten transitiv, bearbeitbare Rollenmitgliedschaften müssen direkt sein.

**Prüfung zuerst:** Tabellengetriebene Tests für alle Profile, Ausschlüsse, Mehrfachrollen, unbekannte Jobs, dynamische beziehungsweise lizenzierende Rollen und Adminschutz.

## 2. Gemeinsamen Code nutzbar machen

Status: lokal umgesetzt und geprüft.

| Vorhandene Stelle | Wiederverwendung und nötige Änderung |
|---|---|
| `Private/Normalization.ps1`, `Private/Comparison.ps1` | Namensnormalisierung, UPN-Kollisionen und feldgenaue Unterschiede übernehmen |
| `Private/LehrerComparison.ps1` | Profil-, Identitäts- und Lizenzlogik nutzen. Feste Schul-/Ganztagsreferenzen durch expliziten Kontext ersetzen, soweit beide Aufrufer die Funktion brauchen |
| `Private/LehrerUpdate.ps1` | Passwort, deaktivierte Neuanlage, Attributprüfung, Lizenzentscheidung und frühen Excel-Checkpoint gemeinsam nutzen. Rollenentfernung für Personal auf verwaltete IDs begrenzen |
| `Private/Excel.ps1` | Vier Pflicht- und vier verwaltete Spalten beibehalten. Personal-Jobvalidierung über Kontext ermöglichen. Hash-, Git-, Sicherungs- und Rückschreibelogik verwenden |
| `Private/ExchangeOnline.ps1` | Gemeinsame Vergleichs-/Schreibfunktionen erweitern, damit eine nicht konfigurierte ABP in allen Phasen ausgelassen wird |
| `Private/LehrerOrchestration.ps1` | Fortschritt, Mailbox-Wiederholungen, Bericht und Ergebnistypen verwenden. Skriptname und verwaltete Rollen für Wiederanlauf und Überschriften übergeben |
| `Private/DisabledAccountCheck.ps1` | Bestandsaufnahme, Exchange-Zuordnung, Anmeldedaten und Besitzobjekte nutzen. Aktive Konten für Personal zusätzlich erfassen |
| `Private/DisabledAccountTui.ps1` | Tasten, Auswahl, Bestätigung und vorhandene Aktionsprüfungen nutzen. Statusabhängige Sperraktion und frische Personal-Ausschlussprüfung ergänzen |
| `Private/Safety.ps1` | Mandanten-, Rollen- und Dateischutz verwenden und fehlende Personal-Zielprüfung ergänzen |

Neue fachliche Orchestrierung kann in `Private/PersonalOrchestration.ps1` liegen. Weitere Dateien nur aufteilen, wenn daraus zusammenhängende Verantwortlichkeiten entstehen. Bestehende Funktionen nicht vollständig kopieren. Keine vorsorgliche Umbenennung des Moduls und keine neuen Laufzeitabhängigkeiten.

**Abnahme dieses Schritts:** Fokussierte Lehrer-, Schüler-, Exchange- und Excel-Regressionen bestehen nach jedem gemeinsamen Umbau. Standardwerte der alten Aufrufer bleiben erhalten.

## 3. Personalbestand, Liste und Schutzprüfungen

Status: lokal umgesetzt und geprüft.

- Alle Benutzer laden und die gemeinsame Auswahl aus der Spezifikation umsetzen.
- Ein Inventarergebnis für `-List`, `-PassThru`, Export und TUI bereitstellen.
- Schutz-/Prüfgründe von der Sichtbarkeit trennen. Rollenlose Konten werden sichtbar, aber erhalten kein geratenes Profil.
- Einzelzielprüfung für aktuelle Identität, Ausschlüsse, Kontostatus, Empfänger und Adminschutz umsetzen und in jeden schreibenden Personalpfad einbauen.

**Tests:** Aktive/gesperrte Konten, rollenlose Konten, Gäste, lokal synchronisierte Konten, persönliche/Shared-/Ressourcenpostfächer, fehlende oder doppelte Empfänger, indirekte Ausschlussmitgliedschaft, tatsächliche Verzeichnisrollen und Fehler beim vollständigen Laden. Dieselben IDs müssen in Liste, Export und TUI ankommen.

## 4. Einzelanlage und optionaler Sync

Status: lokal umgesetzt und geprüft.

- Dünnen Einstieg `Sync-PersonalEntra.ps1` mit den Parametersätzen aus der Spezifikation und öffentlichem `Invoke-PersonalSync` ergänzen.
- `-Add`, `-Remove` und `-ConfigureExchangeOnlineOnly` mit vorhandenen Parametern und UPN-Aliasen implementieren.
- `Personal.xlsx` lesen, vergleichen und beim gewählten Update sicher zurückschreiben.
- Standardauswahl von `-Update` sowie Erhalt gesperrter Bestandskonten nach der abgestimmten Personalregel umsetzen.
- Wiederanlauf einer echten Teilanlage von der Übernahme eines bestehenden rollenlosen Kontos unterscheiden.
- Fortschrittsphasen, Vergleich vor Aktionen, `Compliant` ohne Bestätigung und vollständigen Textbericht übernehmen.

**Tests:** Tatsächliche Skriptaufrufe mit synthetischen Mocks. Teilliste ohne fremde Sperrung, explizite Abgänge, Reimport gesperrter Konten, explizite Übernahme rollenloser Konten, Neuanlage mit Checkpointfehler, Identitätswechsel zwischen Vergleich und Mutation, vorhandene/ausstehende/fehlerhafte Lizenzen und `-WhatIf` ohne Passwort oder Mutation.

## 5. Excel-Export

Status: lokal umgesetzt und geprüft.

- `-ListToExcel` auf dem gemeinsamen Inventar aufbauen.
- Blatt, Tabelle, Stil und acht Basisspalten der strukturell geprüften Lehrerdatei nachbilden, anschließend die beauftragte Spalte `Last Login Timestamp` rechts ergänzen. Keine produktive Vorlage mit Daten kopieren.
- Vorhandenes ImportExcel/EPPlus zum Erstellen verwenden. Ziel und temporäre Datei mit den vorhandenen Git- und Pfadschutzfunktionen absichern.
- Passwörter leer lassen, IDs und Mailwerte unverändert exportieren, unbekannte Jobs erklären und Daten als Text schreiben.
- Personaljobs beim Export lesbar trennen und beim Einlesen eindeutig auf die konfigurierten internen Werte zurückführen. Bisherige Schreibweisen und Schutzregeln erhalten.
- Leeren Bestand mit Kopfzeile unterstützen. Vorhandene Zieldatei, Symlinks und `-WhatIf` prüfen.

**Tests:** Temporäre synthetische XLSX-Datei wieder einlesen. Spaltenreihenfolge, Tabellenstil, Blatt-/Tabellenname, Zelltypen, leere Passwortspalte, Anzahl und Objekt-IDs prüfen. Keine Personen aus der lokalen Jahresdatei in Fixtures übernehmen. Ein Import des Exports muss ausgeschlossene beziehungsweise ungeklärte Schreibziele zuverlässig blockieren.

## 6. TUI zum Sperren, Entsperren und Löschen

Status: lokal umgesetzt und geprüft.

- `-ManageAccounts` an die bestehende Auswahloberfläche anschließen.
- `B`, `U` und `Backspace` oder `Entf` mit eigener Namensbestätigung und passender Statusauswahl verbinden.
- Nur frisch geprüfte Ziele bearbeiten, Teilergebnisse getrennt anzeigen, Balken im `finally` abschließen und Auswahl aktualisieren.
- Die alte gesperrte-Konten-TUI mit ihrer bisherigen Population betreiben. Backspace und Entf öffnen dort ebenfalls die Löschbestätigung, Enter hat keine Aktion in der Kontenliste.

**Tests:** Pfeile, Leertaste, Nein, Esc, gemischte aktive/gesperrte Auswahl, Teilfehler beim Sitzungswiderruf, Besitzobjekte, Adminschutz und Wechsel zu Lehrer/OGTS/Schüler/SharedMailbox nach Auswahl. Synthetische echte CLI-/TUI-Aufrufe im Pseudoterminal und später unter Windows, jeweils zunächst mit `-WhatIf`.

## 7. Dokumentation und Abnahme

Status: lokal umgesetzt und geprüft.

- README um ausführbare Personalbeispiele ergänzen, sobald sie vorhanden sind.
- Operatoranleitung `docs/ENTRA-PERSONAL-SYNC.md` ergänzen und vorhandene Lehrertexte eindeutig benennen.
- `docs/BETRIEB.md`, `docs/TESTING.md`, `docs/ABNAHME.md` und Changelog um das tatsächlich implementierte Verhalten erweitern.
- Nur tatsächlich ausgeführte Prüfungen dokumentieren. Lokale Tests, Windows-TUI, Dateisperren und Mandantentest getrennt ausweisen.

Abschließende lokale Prüfungen:

```powershell
./tests/Test-PesterDiscovery.ps1
./tests/Invoke-LocalVerification.ps1
git diff --check
```

Neue Fälle gehören in fokussierte Pester-Dateien für Personalprofile/Sync, Inventar/TUI und Export. Die gesamte bestehende Suite ist vor Abschluss erforderlich, weil gemeinsam genutzte Funktionen geändert werden. Kontrollierte Mandantentests verwenden ausschließlich synthetische Konten in isolierten Rollen und benötigen einen eigenen ausdrücklich autorisierten Schritt.

## Historischer Prüfstand der ursprünglichen Planungsänderung

Lokal auf macOS am 29. September 2026 geprüft:

| Prüfung | Ergebnis |
|---|---|
| Pester 5.7.1, `tests/Documentation.Tests.ps1` und `tests/Module.Tests.ps1` | 13 bestanden, 0 fehlgeschlagen, 0 übersprungen |
| PowerShell-Parser für `src`, Root-Einstiege, Konfigurationen und geänderten Dokumentationstest | 22 Dateien ohne Parserfehler |
| `Test-ModuleManifest ./src/SchuelerSync/SchuelerSync.psd1` | Erfolgreich, Version 0.3.0 |
| `git diff --check` | Ohne Befund |
| Interne Dateilinks in README und den drei neuen Dokumenten | Alle Ziele vorhanden |
| XLSX-Strukturprüfung | Überschriften, Blatt, Tabelle, Tabellenstil und Spaltenbreiten aus der lokalen Lehrerdatei gelesen, keine Datenzeilen ausgegeben oder verändert |

Pester war im normalen Modulpfad nicht verfügbar. Für die fokussierte Prüfung wurde die im Repository festgelegte Version 5.7.1 temporär aus der PowerShell Gallery unter `/private/tmp/entra-personal-v030-check` bereitgestellt. Der Testlauf importierte sie über den konkreten Manifestpfad. Es wurden keine Projektabhängigkeiten geändert.

Diese Zahlen beschreiben ausschließlich den früheren Planungsschritt. Die vollständige lokale Suite und PSScriptAnalyzer wurden inzwischen für die Implementierung ausgeführt. Ergebnisse stehen in [ABNAHME.md](ABNAHME.md). Windows und produktive Mandanten wurden in dieser Umsetzung nicht verwendet.

## Umsetzungsergebnis

Die anschließende Export-/Vergleichsregression wurde mit einem vollständigen synthetischen Personalexport reproduziert. Der vorhandene Exceladapter kann Prüfzeilen ausdrücklich für den lesenden Personalvergleich erhalten. Der Vergleich meldet sie als `Workbook.Review`, gleicht zulässige Zeilen weiter ab und verhindert Abgangskandidaten aus bekannten Prüfidentitäten. Die Option wird nur im Personalmodus Compare gesetzt. Strikte Update-Vorprüfung und Schreibschutz bleiben unverändert. Tatsächliche CLI-Aufrufe und Dateihashes prüfen die komplette Export-/Compare-/Update-WhatIf-Folge.

Die zusätzlich beauftragte Schüler-TUI verwendet `-ManageAccounts` am bisherigen Schüler-Skript- und Moduleinstieg. Der gemeinsame Inventarfilter erhält die Schülerpopulation. Klasse, Jahrgang und Manager ergänzen den bestehenden Detailbereich. `StudentTui.ps1` verwendet die vorhandenen Attribut-, Manager-, Normalisierungs- und Schutzhelfer für K und L. Zielauswahl, Bestätigung, Fortschritt, Teilfehler und Aktualisierung sind synthetisch geprüft. Die tatsächlichen Ergebnisse und offene Windows-/Mandantenprüfung stehen im [Abnahmeprotokoll](ABNAHME.md).

Die anschließend beauftragte Lehrer-Kontenverwaltung ergänzt `-ManageAccounts` am vorhandenen Skript- und Moduleinstieg. `Get-PersonalInventory` erhält einen begrenzten Populationsparameter, die Schutzprüfung bestätigt dafür direkte Lehrer-/Ganztagsrollen vor jeder Aktion. Der gemeinsame Renderer, Bestätigungen und Aktionsfunktionen bleiben wiederverwendet. Das Neuladen berücksichtigt dieselbe Population. Fokussierte Tests und der echte synthetische CLI-/TUI-Harness prüfen diese Ergänzung, Ergebnisse stehen im [Abnahmeprotokoll](ABNAHME.md).

Die Ergänzung vom 30. September 2026 zeigt Rollengruppen und den Kontotyp im Personal-TUI-Detailbereich. Gastfreigaben werden nur an den Sperr- und Löschprüfpfad übergeben. Inventar und Auswahl unterstützen dafür auch `MailUser`/`GuestMailUser`, während normale Profilpflege und Gastentsperrung gesperrt bleiben. Fokussierte Tests prüfen den vollständigen Gastablauf vom Sperren bis zum Löschen, frische Schutzänderungen und die gemeinsame Rollenanzeige. Der synthetische CLI-Harness hat dafür `-GuestFixture`. Aktuelle Prüfergebnisse stehen im [Abnahmeprotokoll](ABNAHME.md).

Die anschließend beauftragte Spalte `Last Login Timestamp` ist rechts in allen drei Exporten ergänzt. Die Kontenliste und Rollenexporte verwenden `Get-EntraLastLogin` für die letzte erfolgreiche Anmeldung. Der Excel-Schreiber formatiert den Wert als UTC-Text. Fehlende Werte bleiben leer, nicht abrufbare Anmeldedaten lösen eine Warnung und den bestehenden Wiederabruf ohne `signInActivity` aus. Schüler-/Lehrerexporte fordern zusätzlich `AuditLog.Read.All` an. Pflicht- und verwaltete Importspalten bleiben unverändert.

Die zusätzlich beauftragte Erweiterung von `-ListToExcel` auf Schüler und Lehrer verwendet denselben `Export-AccountWorkbook`. `WorkbookExport.ps1` übernimmt ausschließlich die lesende Rollenbestandsaufnahme und Export-Orchestrierung. Schülerschema, Manager-UPN und Klasse beziehungsweise kanonische Lehrerjobs werden über bestehende Helfer verarbeitet. Der neue Parametersatz verändert die alten Sync-Modi nicht. `WorkbookExport.Tests.ps1` und der erweiterte isolierte CLI-Harness prüfen diese Ergänzung, das [Abnahmeprotokoll](ABNAHME.md) hält den aktuellen Prüflauf fest.

Der neue Einstieg nutzt `Invoke-PersonalSync` und den gemeinsamen `Invoke-PersonnelSyncCore`. `PersonalInventory.ps1` bündelt Bestandsauswahl und frische Zielprüfung, `PersonalActions.ps1` verbindet die bestehende TUI mit den Personalaktionen. Der Excel-Adapter unterstützt `Personnel`. Die bestehenden Konto-, Lizenz-, Passwort-, Exchange-, Berichts- und Dateischutzfunktionen bleiben gemeinsam genutzt. `-ResumeCreation` kennzeichnet ausdrücklich eine Teilanlage.

Die Tests in `PersonalSync.Tests.ps1` prüfen Profile, Ausschlüsse, Adminschutz, rollenlose Konten, Bestandsübernahme, Sperrstatus, Lizenzbestandsschutz, Rollenwechsel, Neuanlage und Checkpoint, frische Mutationsprüfung, TUI und native XLSX-Dateien. `Invoke-SyntheticPersonalCli.ps1` isoliert alle Fernbefehle für tatsächliche Skriptaufrufe und den Pseudoterminaltest. Die gesamte bestehende Suite bleibt Bestandteil der Abschlussprüfung.
