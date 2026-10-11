# Personal 0.3.0: Varianten und Empfehlungen

Ergänzung vom 10. Oktober 2026: Die [Adressregeln](ADRESSREGELN.md) konkretisieren die Namensbildung für Neuanlagen und den Bestandsschutz. Sie ergänzen diese ursprüngliche Variantenentscheidung. Die gemeinsame Umsetzung bleibt erhalten, getrennte Skriptkopien sind weiterhin nicht erforderlich.

Stand: 29. September 2026. Grundlage sind der Code aus 0.2.0, die mitgelieferten Rollenzuordnungen und der Screenshot. Die Empfehlung wurde vom Betreiber freigegeben und umgesetzt. Die folgenden Varianten halten die Entscheidung und den damaligen Ausgangsstand fest.

Der anschließend beauftragte Schüler-/Lehrerexport ist ebenfalls umgesetzt. Alle drei Einstiege unterstützen `-ListToExcel` und verwenden denselben verifizierten Excel-Schreiber. Die bestehenden Populationen und Sync-Regeln bleiben erhalten. Bedienung und Importhinweise stehen in den jeweiligen [Schüler-](ENTRA-SCHUELER-SYNC.md), [Lehrer-](ENTRA-LEHRER-SYNC.md) und [Personalregeln](ENTRA-PERSONAL-SYNC.md).

## Empfehlung zur Aufteilung

Ein eigener Einstieg **`Sync-PersonalEntra.ps1`** passt am besten. Lehrer, Co-Lehrer, Schulpädagogen und OGTS behalten `Sync-LehrerEntra.ps1`. Beide verwenden dieselben passenden Funktionen im vorhandenen Modul `SchuelerSync`.

| Variante | Vorteil | Nachteil | Bewertung |
|---|---|---|---|
| Lehrerskript in allgemeines Personalskript umbenennen | Ein öffentlicher Einstieg | Vorhandene Aufrufe ändern sich. Vollständige Lehrerlisten und selten gepflegte Personallisten brauchen unterschiedliche Regeln für Abgänge. | Für diesen Ausbau nicht empfohlen |
| Neues Personalskript mit gemeinsamem Modul | Bekannte Aufrufe bleiben erhalten. Rollen und Bestandsauswahl lassen sich gezielt abgrenzen. | Einige bisher lehrerspezifische Funktionen müssen einen expliziten Konfigurationskontext erhalten. | Empfohlen |
| Vollständige Kopie des Lehrerskripts | Zunächst geringer Umbau | Lizenz-, Passwort-, Excel- und Exchange-Korrekturen müssten mehrfach gepflegt werden. | Verwerfen |

Ein neuer Einstieg bedeutet nur eine kleine Parameterhülle. Profilauflösung, Lizenzentscheidung, Passworterzeugung, Kontoänderungen, Exchange-Abgleich, Excel-Schutz und Berichte werden soweit passend gemeinsam verwendet. Ein allgemeines Plugin- oder Workflow-System wird dafür nicht benötigt.

## Welche Konten gehören dazu?

Für Liste, Export und TUI lautet die Ausgangsmenge: alle Entra-Benutzer außer Mitgliedern der bisherigen Lehrer-/Ganztagsrollen, Mitgliedern der Schülerrolle und `SharedMailbox`. Dadurch erscheinen auch Konten ohne Rollengruppe. Der Ausschluss von Ganztag gilt aus der übergeordneten Anforderung auch für die TUI.

Für automatische Attribut-, Rollen- und Lizenzänderungen braucht ein Konto zusätzlich ein eindeutiges freigegebenes Profil. Sichtbarkeit allein begründet keine automatische Rollenzuordnung. Gäste, aus lokalem AD synchronisierte Konten, Ressourcenpostfächer und administrativ geschützte Konten bleiben mit Prüfgrund sichtbar.

Ergänzung vom 30. September 2026: Für Gäste ist ausdrücklich das Sperren und Löschen über die Personal-TUI freigegeben, einschließlich rollenloser Gäste und eindeutiger `MailUser`-/`GuestMailUser`-Empfänger. Gastentsperrung und automatische Profilpflege bleiben ausgeschlossen. Administrator-, Rollen- und Empfängerschutz gelten weiterhin. Die TUI zeigt tatsächliche Rollengruppen und bei Personal den Kontotyp neben der letzten Anmeldung.

Das Kinderhaus-Pädagogikprofil gehört zum neuen Skript. `SEC-A-ROL-Schule_PädagogischesTeam` gehört vollständig zum Lehrerskript, einschließlich CO-L und PA.

## Bedienung

- `-Add` wird der häufigste Weg. Die Parameter `-Vorname`, `-Nachname`, `-Job` und `-EntraObjectId` bleiben bekannt.
- Ein optionaler Sync liest `Personal.xlsx` mit demselben Spaltenvertrag wie die Lehrerdatei.
- `-List` zeigt den gesamten passenden Bestand mit derselben Informationstiefe wie die bisherige Kontenliste.
- `-ListToExcel` erzeugt eine neue `Personal.xlsx`.
- `-ManageAccounts` startet die TUI im Personalskript. So bleibt der parameterlose Aufruf wie bisher ein lesender Excel-Vergleich.
- `-Remove -UPN` behält seine Bedeutung: sperren und Sitzungen widerrufen. Löschen erfolgt über die ausdrücklich bestätigte TUI.
- Fortschrittsbalken, farbige Überschriften, Leerzeilen, `-WhatIf`, `-Confirm`, `-OutputFile` und die UPN-Aliase folgen den bestehenden Skripten.

Ein separates viertes Verwaltungsskript bringt hier keinen zusätzlichen Nutzen. Die bestehende TUI wird intern erweitert, ihr alter Einstieg behält seine bisherige Auswahl.

## Zwei bestätigte Unterschiede zum Lehrersync

1. **Keine impliziten Abgänge bei `-Update`.** Beim neuen Skript bedeutet `-Update` ohne weitere Aktionsschalter Erstellen und Aktualisieren. `-DisableUsers` und `-RevokeSessions` bleiben ausdrücklich wählbar. Fehlende Einträge in einer selten gepflegten Exceldatei reichen dadurch nicht versehentlich für eine Sperrung.
2. **Gesperrte Bestandskonten bleiben beim Sync gesperrt.** Die neue Bestandsliste und ihr Export enthalten auch gesperrte Konten. Ein unveränderter Reimport soll diese nicht aktivieren. Entsperren erfolgt gezielt mit `U` in der TUI. Neue Konten werden nach erfolgreicher Pflichtkonfiguration wie bisher aktiviert.

Diese Unterschiede betreffen nur den neuen Personalmodus. Sie werden in Hilfe, Bericht und Tests ausdrücklich beschrieben. Ein vollständiger Abgangslauf umfasst nur aktive Mitglieder der verwalteten Personalrollen. Rollenlose Konten gelten nie allein wegen einer fehlenden Excel-Zeile als Abgang.

## Rollen, Jobs und Lizenzen

Empfehlung für die neuen `Job`-Werte: den eindeutigen Rollensuffix verwenden, beispielsweise `Schule_Sekretariat` und `Kinderhaus_Sekretariat`. Ein neues Kürzelsystem wäre ohne bestehende Vorgaben schwerer zuzuordnen. Die Spalte bleibt `Job`, und `JobTitle` enthält den gewählten Wert. Schreibweise wird anhand der Konfiguration normalisiert.

Eine Person bekommt in 0.3.0 genau ein verwaltetes Zielprofil. Mehrere konkurrierende Personalrollen werden sichtbar gemeldet. Die Reihenfolge der Gruppen darf keine Firmen-, ABP- oder Mitarbeitertypentscheidung treffen. Zusätzliche Gruppen außerhalb der verwalteten Rollen bleiben erhalten.

Die Standardlizenzgruppe ist `SEC-A-LIC-O365A1Faculty`. Jede vorhandene Lizenz oder bereits ausstehende gruppenbasierte Zuweisung verhindert eine zusätzliche Standardlizenz. Die vorhandene Prüfung von `AssignedLicenses`, `LicenseAssignmentStates` und Lizenzgruppen wird wiederverwendet.

## Adminrollen: bestätigter Schutz

Die Vorlage enthält `ExterneAdmins` mit bekannten Profilwerten. Der Screenshot enthält zusätzlich `GlobalAdministrators`, ohne Gruppen-ID und ohne Fachprofil.

Der Betreiber hat die Variante bestätigt, beide Gruppen in Liste, Export und TUI anzuzeigen und vor Änderungen zu schützen. Für `GlobalAdministrators` werden keine Firma, Lizenz oder ABP erfunden. Die Gruppe kann für den Schutz anhand ihres exakten eindeutigen Namens aufgelöst werden. Die eindeutige Namensauflösung ist implementiert. Eine künftige feste ID benötigt eine verifizierte Mandanteninformation.

## Ausgangsstand und umgesetzte Wiederverwendung

- `Get-LehrerSnapshot`, Wiederanlauf und Berichtskopf referenzieren heute fest `Roles.School` und `Roles.Ganztag`. Sie benötigen einen konkreten Kontext für die jeweils verwalteten Rollen.
- `Read-StudentWorkbook -Kind Teacher` lädt fest die Lehrerkonfiguration. Derselbe Spaltenvertrag braucht für Personal eine andere Jobvalidierung.
- `Compare-LehrerDirectory` und `Sync-LehrerGroup` behandeln bisher alle anderen `SEC-A-ROL-*` als konkurrierend. Der neue Personalpfad darf nur explizit verwaltete Zielrollen ersetzen.
- Die Exchange-Vorprüfung erwartet überall eine ABP. Eine fehlende ABP muss wie im manuellen Skript übersprungen werden. Auch Vergleich, Parameterbau und Nachprüfung müssen dieselbe Regel verwenden.
- Die bestehende Kontenprüfung liest nur gesperrte Benutzer. Der neue Bestand umfasst aktive und gesperrte Konten und braucht dieselben Ausschlussregeln in allen drei Ansichten.
- Im Ausgangsstand 0.2.0 gab es keinen `-ListToExcel`-Modus. Dieser ist jetzt implementiert. Wiederverwendbar sind der Spaltenvertrag und die sichere Dateiverarbeitung.

Die genauen Regeln stehen in der [Spezifikation](PERSONAL-0.3.0-SPEC.md), die Reihenfolge in der Umsetzung im [Plan](PERSONAL-0.3.0-PLAN.md).
