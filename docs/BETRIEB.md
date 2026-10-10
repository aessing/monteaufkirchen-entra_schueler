# Betrieb und Wiederanlauf

## Adressen bei Neuanlagen (0.3.0)

Schüler erhalten UPN und Mail unter `monteaufkirchen.com`. Schulteam, OGTS und übriges Personal verwenden `montessori-erding.de`. Nur das Schulteam (`L`, `CO-L`, `PA`, auch mit `JAS`) erhält zusätzlich den gleichnamigen Alias unter `monteaufkirchen.com`. Angestellte verwenden Vornamensinitiale plus Nachname mit wachsendem Präfix bei Belegung. Freigegebene Rollen `Externe*` verwenden `vorname.nachname` mit Zahl ab `2` bei Belegung.

Bestehende UPNs, Mailadressen, Mail-Nicknames und Aliase bleiben erhalten. Ein normaler Exchange-Reparaturlauf ändert keine Adressen. Prüfe nach einer fehlgeschlagenen Lehrer-Neuanlage Hauptadresse und Alias separat. Die [Adressregeln](ADRESSREGELN.md) beschreiben den gezielten Abschluss eines fehlenden Alias, ohne eine Bestandsmigration auszulösen.

## Schülerkonten verwalten (0.3.0)

`Sync-SchuelerEntra.ps1 -ManageAccounts -WhatIf` öffnet die gemeinsame Schüler-TUI als Vorschau. Ohne `-WhatIf` können bestätigte Aktionen schreiben. Leertaste markiert Schüler, `K` wählt eine vorhandene CLS-Gruppe für Klasse/Jahrgang, `L` einen aktiven internen direkten Angehörigen des Schul-Pädagogikteams als Klassenlehrer. B sperrt mit Sitzungswiderruf, U entsperrt gezielt, Backspace oder Entf löscht gesperrte Konten. Rolle, Klasse, Jahrgang, Klassenlehrer und letzte Anmeldung stehen im Detailbereich. Lehrer, Ganztag, Shared Mailboxes und geschützte Administratoren können nicht geändert werden.

Beim Klassenwechsel werden neue Mitgliedschaft und Klassenattribute vor der Entfernung alter CLS-Zuordnungen bestätigt. Bei einem Teilfehler alte und neue Gruppen sowie `Department`/`OfficeLocation` im neu geladenen Bestand prüfen. Es gibt keinen automatischen Rollback. Ein gesperrtes Konto bleibt beim Klassen- oder Lehrerwechsel gesperrt. Excel wird nicht aktualisiert, übernimm Änderungen vor dem nächsten Sync in die Quelle oder gleiche einen neuen `-ListToExcel`-Export ab. Eine alte Quelle kann Zuordnungen zurücksetzen und gesperrte Schüler reaktivieren. [Bedienung und Berechtigungen](ENTRA-SCHUELER-SYNC.md#schüler-tui-manageaccounts).

## Lehrer- und Ganztagskonten verwalten (0.3.0)

`Sync-LehrerEntra.ps1 -ManageAccounts` startet dieselbe TUI wie die Personalverwaltung, beschränkt auf aktive und gesperrte direkte Schulteam-/Ganztagsmitglieder. Zuerst mit `-WhatIf` prüfen. B sperrt einschließlich Sitzungswiderruf, U entsperrt Mitgliedskonten, Backspace oder Entf löscht gesperrte Konten nach Namensbestätigung. Rolle, Kontotyp und letzte Anmeldung stehen im Detailbereich. Vor jeder Aktion wird die direkte Rollenmitgliedschaft frisch geprüft, ein inzwischen ausgetretenes Konto wird übersprungen. Schüler und Shared Mailboxes sind ausgeschlossen, Adminschutz und Empfängerprüfung gelten weiter. Der Bestand bleibt nach einer Aktion auf dieselben Lehrerrollen begrenzt. Der Modus benötigt keine Exceldatei und verändert keine Gruppen, Lizenzzuweisungen oder Passwörter separat. [Aufrufe, Grenzen und Berechtigungen](ENTRA-LEHRER-SYNC.md).

## Bestandsdateien erzeugen (0.3.0)

Alle drei Sync-Einstiege unterstützen `-ListToExcel`. Ohne `-File` erzeugen sie eine neue `Schueler.xlsx`, `Lehrer.xlsx` beziehungsweise `Personal.xlsx` im Repository-Root. Wenn das Ziel existiert, wähle einen neuen Dateinamen. Vorhandene Dateien werden nicht ersetzt.

Alle Exporte hängen `Last Login Timestamp` rechts an, als UTC-Zeitstempel der letzten erfolgreichen Anmeldung. Schüler haben damit zehn, Lehrer und Personal neun Spalten. Die Schüler-/Lehrerexporte fordern nun zusätzlich `AuditLog.Read.All` an. Ohne verfügbaren Wert bleibt die Zelle leer, bei einem Fehler der Anmeldedatenabfrage läuft der Export nach einer Warnung ohne diese Daten weiter. Die neue Spalte ist rein informativ und wird vom Sync nicht ausgewertet.

```powershell
.\Sync-SchuelerEntra.ps1 -ListToExcel -File '.\Schueler-Bestand.xlsx'
.\Sync-LehrerEntra.ps1 -ListToExcel -File '.\Lehrer-Bestand.xlsx'
.\Sync-PersonalEntra.ps1 -ListToExcel -File '.\Personal-Bestand.xlsx'
```

`-WhatIf` lädt denselben Bestand und erzeugt keine Datei oder Ordner. Alle Exportpasswörter bleiben leer. Schüler- und Lehrerexporte enthalten auch gesperrte direkte Rollenmitglieder. Vor der Wiederverwendung Klasse, Klassenlehrer, Job und gewünschte Population prüfen. Der bisherige Schüler-/Lehrerupdate kann gelistete gesperrte Konten aktivieren. Die Personalregel zum Erhalt des Sperrstatus gilt ausschließlich für den Personal-Einstieg. Ein Export ersetzt diese Prüfung nicht. Schema, Rollen, Warnungen und Berechtigungen stehen in den [Schülerregeln](ENTRA-SCHUELER-SYNC.md), [Lehrerregeln](ENTRA-LEHRER-SYNC.md) und [Personalregeln](ENTRA-PERSONAL-SYNC.md).

## Bestehende Verwaltungsabläufe

Ein unveränderter Personal-Gesamtexport kann ohne `-Update` mit `Sync-PersonalEntra.ps1 -File '.\Personal.xlsx'` verglichen werden. Nicht importierbare oder nicht schreibbare Zeilen stehen als `Workbook.Review` mit Zeilennummer und Identität in den Warnungen. Die übrigen Zeilen werden weiter abgeglichen. Vor Update einschließlich WhatIf die Arbeitskopie bereinigen, gültige Profile nur zulässigen Konten zuweisen und geschützte/nicht schreibbare Zeilen entfernen. Eine geänderte Jobzelle hebt Schutzrollen nicht auf. Die vollständige aktive verwaltete Population muss bei ausdrücklich gewählten Abgangsaktionen weiterhin enthalten sein.

Der Lehrer-/Ganztagsabgleich hat einen eigenen Einstiegspunkt und ein rollenabhängiges Exchange-Profil. Bedienung, Lizenzbestandsschutz und Wiederanlauf stehen in [ENTRA-LEHRER-SYNC.md](ENTRA-LEHRER-SYNC.md). Ein vollständiger Lehrer-Update-Lauf benötigt beide Lehrer-/Ganztagsrollen vollständig in der Arbeitsmappe. `-Remove` deaktiviert und widerruft Sitzungen, ohne Konto oder Lizenz zu löschen.

Für die übrigen Personalkonten ab 0.3.0 gibt es [Sync-PersonalEntra.ps1](ENTRA-PERSONAL-SYNC.md). `-List`, `-PassThru` und `-ListToExcel` zeigen denselben Bestand einschließlich rollenloser Konten. Lehrer, CO-L, PA, OGTS, Schüler und Shared Mailboxes fehlen. ExterneAdmins und GlobalAdministrators bleiben sichtbar und geschützt. Personal-`-Update` allein erstellt und aktualisiert, gesperrte Bestandskonten bleiben gesperrt. Personal-`JobTitle` wird dabei lesbar geschrieben, etwa `Kinderhaus Pädagogisches Team`. Technische Bestandstitel erscheinen zunächst im lesenden Vergleich als Änderung. Bestandsliste und Vergleich zeigen dieselben lesbaren Jobnamen. Abgänge werden erst mit `-Update -DisableUsers` beziehungsweise `-Update -RevokeSessions` bearbeitet. Dafür ist eine vollständige Liste der verwalteten aktiven Personalpopulation erforderlich.

`-ManageAccounts` startet die Personal-TUI mit `B`, `U` und Backspace oder Entf. Bei jeder Aktion werden die Schutzkriterien erneut gelesen. Nach Bestätigung zeigt sie Einzelresultate, lädt den Bestand neu und entfernt Markierungen. Ein Teilfehler beim Sitzungswiderruf hebt eine erfolgte Sperrung nicht auf. `-Remove -UPN` bleibt eine gezielte Sperrung einschließlich Widerruf, es löscht nicht.

Die Personal-TUI zeigt Rolle und Kontotyp im Detailbereich. Auch zulässige Gäste können mit `B` gesperrt und nach erneuter Markierung mit Backspace oder Entf gelöscht werden. Zum Löschen ist der Sperrstatus erforderlich. Gäste mit eindeutigem `MailUser`/`GuestMailUser`, persönlichem Postfach oder ohne Exchange-Empfänger sind zulässig. Adminschutz, ausgeschlossene Rollen und alle frischen Zielprüfungen gelten weiter. Gäste können mit `U` nicht entsperrt werden. Die Gastfreigabe gilt ausschließlich für die Personal-TUI.

Der Export braucht eine neue ignorierte XLSX-Datei. Personaljobs werden mit Leerzeichen statt Unterstrichen und CamelCase exportiert. Import und Einzelpflege erkennen die lesbaren Namen sowie die bisherigen internen Jobwerte, unbekannte oder mehrdeutige Zuordnungen werden abgewiesen. Ein unbearbeiteter Gesamtexport mit leeren oder geschützten Jobs blockiert einen schreibenden Import. Die [Personal-Anleitung](ENTRA-PERSONAL-SYNC.md) erklärt die Übernahme rollenloser Konten mit Objekt-ID und den ausdrücklichen Wiederanlauf einer Teilanlage über `-ResumeCreation`. Für den neuen Einstieg `ExpectedTenantId` in `config/PersonalSync.psd1` prüfen.

Für eine Bestandsaufnahme aller gesperrten Entra-Konten gibt es [Verwalte-GesperrteEntraKonten.ps1](GESPERRTE-KONTEN-VERWALTUNG.md). `-List` zeigt die lesende Ausgabe einschließlich freigegebener Postfächer. Ohne Schalter öffnet eine TUI mit Markierung, Backspace oder Entf zur Löschbestätigung und `U` zur Entsperrbestätigung. `SharedMailbox` wird in der TUI ausgeblendet. Konten ohne Exchange-Empfänger sind auswählbar, wenn die Inventarzuordnung eindeutig fehlt. Ressourcenpostfächer und unklare Empfänger bleiben nur zur Prüfung sichtbar. Besitzobjekte erscheinen mit Typ, Name und ID. Kann der Exchange-Bestand insgesamt nicht geladen werden, bricht der Check ab. Ein Sperrdatum und die 90-Tage-Frist werden weiterhin nicht automatisch geprüft.

Besitzobjekte blockieren weder Entsperren noch eine bestätigte Löschung. Die Löschbestätigung nennt die Objekte und warnt vor möglicherweise besitzerlosen Gruppen oder Anwendungen. Unmittelbar vor dem Löschen fragt das Skript den Besitz erneut ab und warnt bei Treffern. Kann es den Besitz nicht prüfen, überspringt es das Konto.

Nach der Ja-Bestätigung zeigen Löschen und Entsperren den Fortschritt je ausgewähltem Konto an. Übersprungene Konten und Fehler erscheinen anschließend einzeln in den Aktionsergebnissen.

Diese Anleitung gilt für PowerShell 7 auf Windows und macOS. Für Windows-spezifische Dateisperren und kontrollierte Mandantentests ist die Windows-Parallels-VM vorgesehen. Die fachlichen Regeln stehen in [ENTRA-SCHUELER-SYNC.md](ENTRA-SCHUELER-SYNC.md). Verwende in Tests ausschließlich erfundene Schüler.

## 1. Arbeitsumgebung vorbereiten

Öffne PowerShell 7 im Repository-Root und installiere die benötigten Module:

```powershell
$PSVersionTable.PSVersion

Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Users.Actions,Microsoft.Graph.Groups -Scope CurrentUser
Install-Module ExchangeOnlineManagement,ImportExcel -Scope CurrentUser

Get-Module -ListAvailable Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Users.Actions,Microsoft.Graph.Groups,ExchangeOnlineManagement,ImportExcel |
  Sort-Object Name,Version -Descending |
  Select-Object Name,Version,Path
```

Schließe die Arbeitsmappe in Excel vor einem Update. Kontrolliere die erwartete Tenant-ID in `config/SchuelerSync.psd1`. Ist `ExpectedTenantId = $null`, musst du die vom Skript ausgegebene Tenant-ID vor jeder Änderung manuell prüfen.

Graph und Exchange müssen bei kombinierten Läufen denselben Mandanten verwenden. Prüfe aktive Exchange-Verbindungen mit `Get-ConnectionInformation`. Bei mehreren Verbindungen oder einem falschen Mandanten die Verbindungen bewusst korrigieren und den Vergleich neu starten. Das Skript wechselt keine bestehende Sitzung stillschweigend. Die verwendete Tenant-ID wird von [Get-ConnectionInformation](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-connectioninformation?view=exchange-ps) bereitgestellt.

Berichte benötigen neue `.txt`-Dateien. Innerhalb eines Git-Repositories muss das Ziel ignoriert sein. Nutze etwa `-OutputFile "./Berichte/Abgleich-$(Get-Date -Format yyyyMMdd-HHmmss).txt"`. Vorhandene Dateien und symbolische Links werden nicht überschrieben. Für die Pfadprüfung muss Git verfügbar sein.

## 2. Arbeitsmappe prüfen

Pflichtspalten sind `Name mit Rufname`, `Vorname`, `Nachname`, `Klassen` und `Klassenlehrer`. Genau ein Arbeitsblatt muss alle fünf Überschriften enthalten. Bestehende Zusatzspalten `Passwort`, `EntraObjectId`, `UPN` und `Mail` dürfen nicht manuell umsortiert oder geleert werden, während ein Lauf aktiv ist.

Die Schule muss keinen UPN liefern. Sind `EntraObjectId` und `UPN` leer, gleicht das Skript die Excel-Felder `Vorname` (Rufname) und `Nachname` mit `givenName` und `surname` der direkten Mitglieder von `SEC-A-ROL-Schule_Schüler` ab. Der Treffer muss innerhalb dieser Gruppe eindeutig sein.

Prüfe den Pfad und den Git-Schutz:

```powershell
$studentFile = Resolve-Path '.\Schueler.xlsx'
$studentFile
git check-ignore -- .\Schueler.xlsx
git ls-files --error-unmatch -- .\Schueler.xlsx
```

`git check-ignore` muss bei einer Datei innerhalb des Repositories erfolgreich sein. `git ls-files` darf die Datei nicht als versioniert ausgeben. Die mitgelieferten Regeln schützen XLSX-Dateien im Root und in Unterordnern, einschließlich Sicherungen und temporärer Kopien. In einem fremden Repository muss dessen `.gitignore` ebenfalls Quelle und beide Artefaktmuster abdecken. Das Skript prüft alle konkret gewählten Pfade vor der ersten Kopie und bricht andernfalls ab.

## 3. Vergleich ausführen

Der Vergleich liest Excel, Entra und den aktuellen Exchange-Zustand. Er schreibt nichts und wartet nicht auf fehlende Postfächer:

```powershell
.\Sync-SchuelerEntra.ps1
.\Sync-SchuelerEntra.ps1 -File 'C:\GeschuetzteDaten\Schueler-2026.xlsx'
```

Prüfe nacheinander:

1. Tenant-ID und angemeldetes Konto
2. `Neuzugänge`
3. `Abgänge`
4. jede Zeile in `Änderungen`
5. UPN-Kollisionswarnungen
6. Manager- und Gruppenfehler
7. `EXO-Konfiguration ausstehend`

Ein Eintrag unter `Fehler` blockiert den schreibenden Lauf. Korrigiere zuerst Excel, Gruppen, Manager oder Berechtigungen und wiederhole den Vergleich.

## 4. WhatIf vor jedem Update

```powershell
.\Sync-SchuelerEntra.ps1 -Update -WhatIf
.\Sync-SchuelerEntra.ps1 -File 'C:\GeschuetzteDaten\Schueler-2026.xlsx' -Update -WhatIf
```

`-WhatIf` erzeugt keine Passwörter, schreibt weder Excel noch Berichte, Graph oder Exchange und wartet nicht auf Postfächer. Es zeigt die geplanten Phasen als `WhatIf`.

Für eine Teilaktion:

```powershell
.\Sync-SchuelerEntra.ps1 -Update -CreateNewUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -Update -UpdateUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -Update -DisableUsers -RevokeSessions -WhatIf
```

## 5. Autorisierten Lauf starten

> [!CAUTION]
> Abgänge sind alle aktiven, nicht zugeordneten Mitglieder der konfigurierten Gruppe `SEC-A-ROL-Schule_Schüler`. Die Erkennung ist nicht auf eine Klasse oder einen Arbeitsmappenausschnitt begrenzt. Für jede Aktion mit `-DisableUsers` oder `-RevokeSessions` muss die Excel-Datei deshalb die vollständige konfigurierte Schülerpopulation enthalten. Eine einzelne Testklasse ist in einer gemeinsam genutzten produktiven Rollengruppe nicht isoliert. Verwende für destruktive Tests einen vollständig isolierten Mandanten oder eine Rollengruppe, deren Mitglieder ausschließlich synthetische Testkonten sind.

Vollständiger Lauf:

```powershell
.\Sync-SchuelerEntra.ps1 -Update
```

Lauf mit eigener Datei:

```powershell
.\Sync-SchuelerEntra.ps1 -File 'C:\GeschuetzteDaten\Schueler-2026.xlsx' -Update
```

Selektive Läufe:

```powershell
.\Sync-SchuelerEntra.ps1 -Update -CreateNewUsers
.\Sync-SchuelerEntra.ps1 -Update -UpdateUsers
.\Sync-SchuelerEntra.ps1 -Update -DisableUsers
.\Sync-SchuelerEntra.ps1 -Update -RevokeSessions
```

Ein vollständiger Lauf konfiguriert Exchange für alle aktiven Schüler aus Excel. Ein selektiver Create- oder Update-Lauf konfiguriert nur erfolgreich erstellte oder aktualisierte Konten. Disable und Revoke allein starten keine Exchange-Konfiguration.

## 6. Ergebnisse kontrollieren

Prüfe nach einem Neuzugang:

- Konto wurde zuerst deaktiviert erstellt
- Attribute, Manager und drei Pflichtgruppen stimmen
- `Passwort`, `EntraObjectId`, `UPN` und `Mail` stehen in der richtigen Excel-Zeile
- Passwort ist exakt 12 Zeichen lang
- `ForceChangePasswordNextSignIn` ist `false`
- Konto wurde erst danach aktiviert
- eine `.backup-YYYYMMDD-HHmmss.xlsx` enthält den vorherigen Stand

Prüfe nach einer Änderung:

- nur gelistete Abweichungen wurden geschrieben
- die stabile Objekt-ID blieb erhalten
- `EntraObjectId`, der tatsächlich vorhandene `UPN` und `Mail` stehen in Excel
- ein manuell vergebener bestehender UPN blieb unverändert
- Zielgruppen wurden hinzugefügt und bestätigt, bevor alte direkte Rollen- oder Klassengruppen entfernt wurden
- dynamische und geerbte Mitgliedschaften wurden nur gewarnt
- die Passwortspalte blieb unverändert

Prüfe nach einem Abgang:

- `accountEnabled` ist `false`, wenn Disable gewählt war
- Graph hat den Sitzungswiderruf bestätigt, wenn Revoke gewählt war
- Benutzer, Gruppen und Lizenzen wurden nicht gelöscht

## Einzelnen Schüler ohne Excel bearbeiten

Prüfe eine manuelle Neuanlage oder Aktualisierung zuerst mit `-WhatIf`:

```powershell
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin' -WhatIf
.\Sync-SchuelerEntra.ps1 -Add -Vorname 'Mia' -Nachname 'Muster' -Klasse 'JK1-3g2_1' -Klassenlehrer 'Lea Lehrerin'
```

Bei einer erfolgreichen Neuanlage erscheint das Passwort nach verifizierter Aktivierung genau einmal im Terminal. Scheitert ein später Schritt nach der Kontoerstellung, erscheint es ebenfalls einmal für den Wiederanlauf. Der Bericht nennt dann einen unbekannten Kontostatus, die Objekt-ID und einen `-EntraObjectId`-Wiederanlaufbefehl. Dieser ausdrückliche Wiederanlauf darf das exakt bezeichnete Konto auch vor seiner Aufnahme in die Schüler-Rollengruppe erreichen. Unmittelbar vor jeder Mutationsphase liest er das Konto erneut aus Graph und verlangt weiterhin den passenden normalisierten Vor- und Nachnamen, `CompanyName` und `EmployeeType`. Prüfe die Objekt-ID vor der Ausführung. Starte dafür kein Transcript und sichere das Passwort unmittelbar. Ist der Schüler bereits eindeutig in der direkten Schüler-Rollengruppe vorhanden, aktualisiert das Skript nur seine Abweichungen. UPN und Passwort bleiben unverändert. Ein deaktiviertes Bestandskonto wird nach verifiziertem Pflichtzustand aktiviert.

Zum Deaktivieren eines einzelnen Schülers und Widerrufen seiner Sitzungen:

```powershell
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com'
```

Prüfe vor dem zweiten Befehl den exakten UPN und die angezeigte Tenant-ID. `-Remove` akzeptiert nur direkte Mitglieder von `SEC-A-ROL-Schule_Schüler`. Konto, Gruppen und Lizenzen werden nicht gelöscht.

## Exchange-Reparatur

Ein neues Schüler- oder übriges Personalpostfach wird sofort und danach bis zu fünfmal im Abstand von 60 Sekunden geprüft. Beim Lehrer-/Ganztagspfad folgen auf die sofortige Prüfung bis zu drei Wiederholungen im Abstand von 150 Sekunden, insgesamt höchstens 7,5 Minuten Wartezeit. Das Warten geschieht einmal pro Batch-Runde. Nach dem letzten Versuch bleibt der Entra-Zustand erhalten und das Ergebnis meldet `Pending` mit `EXO-Konfiguration ausstehend` und einem Reparaturbefehl.

Ist beim Lehrerimport nur Exchange fehlgeschlagen, setze den vorhandenen UPN mit `Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN` fort. Dieser Modus erzeugt kein Konto und setzt kein Passwort zurück. Die [Lehreranleitung](ENTRA-LEHRER-SYNC.md#sicherheitsprüfungen-und-wiederanlauf) enthält die Vorschau, den Schreibaufruf und den selektiven Excel-Wiederanlauf für frühere Entra-Fehler. Eindeutige Empfänger-nicht-gefunden-Fehler bei Exchange-Lese-, Schreib- und Rückleseoperationen werden dort erneut geprüft. Andere Schreibfehler bleiben sofortige Fehler.

Ein einzelnes Postfach nachziehen:

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com'
```

Mehrere Postfächer gemeinsam nachziehen:

```powershell
$upns = @(
  'test1@monteaufkirchen.com'
  'test2@monteaufkirchen.com'
)
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail $upns
```

Vorprüfung ohne Schreiben und ohne Wartezeit:

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail $upns -WhatIf
```

Der Schüler-Exchange-Modus akzeptiert auch `-UPN` und `-UserPrincipalName`. Er liest keine Excel-Datei und benötigt keine Graph-Verbindung. Er prüft genau eine aktive Exchange-Verbindung mit bekannter Tenant-ID und gegebenenfalls `ExpectedTenantId`. Er ändert weder Entra-Attribute noch Gruppen. Der Lehrer-Exchange-Modus liest zusätzlich Graph, um das Personalprofil zu bestimmen, und verlangt deshalb übereinstimmende Graph- und Exchange-Mandanten.

## Wiederanlauf nach Fehlern

### Vorprüfungsfehler

Es wurde nichts geschrieben. Behebe die gemeldete Ursache und starte denselben Vergleich erneut. Typische Ursachen sind eine falsche Tenant-ID, mehrdeutige Identität, ein nicht eindeutiger Klassenlehrer, eine ungültige Klasse, eine fehlende Zielgruppe oder eine veränderte Arbeitsmappe.

### Neuzugang bleibt deaktiviert

Lösche das Konto nicht vorschnell. Prüfe es anhand der Objekt-ID. Kontrolliere Attribute, Manager, Gruppen und die Excel-Zeile. Stelle bei Bedarf den passenden Arbeitsmappenstand wieder her. Starte danach den lesenden Vergleich. Das Skript ordnet den Benutzer über `EntraObjectId`, gespeicherten UPN oder eindeutigen Namen erneut zu.

Eine manuelle Aktivierung darf erst erfolgen, wenn Pflichtzustand und Excel-Zuordnung vollständig geprüft wurden.

### Excel-Rückschreibung scheitert

Die Fehlermeldung nennt die Quelldatei, die Backup-Datei und gegebenenfalls eine vertrauliche temporäre Datei. Gehe so vor:

1. Beende Excel und alle Prozesse mit Zugriff auf die Datei.
2. Kopiere keine Variante über eine andere.
3. Vergleiche Zeitstempel, SHA-256 und Inhalt jeder genannten Datei.
4. Wähle den richtigen Stand und stelle ihn unter dem ursprünglichen Pfad wieder her.
5. Schütze oder lösche verbleibende Passwortkopien erst nach der Wiederherstellung.
6. Starte zuerst einen lesenden Vergleich.

Hilfsbefehle:

```powershell
Get-ChildItem 'C:\GeschuetzteDaten' -Filter '*.xlsx' |
  Select-Object FullName,Length,LastWriteTime

Get-FileHash 'C:\GeschuetzteDaten\*.xlsx' -Algorithm SHA256
```

### UPN-Änderung ist nur teilweise erfolgreich

Verwende die gespeicherte `EntraObjectId`. Lies den aktuellen UPN direkt aus Entra, korrigiere keine Excel-Zelle nach Vermutung und starte den Vergleich neu. Der Orchestrator versucht bei einer Umbenennung vor und nach der Graph-Mutation einen Identitätscheckpoint zu speichern.

### Set-Mailbox oder Set-CASMailbox schlägt fehl

Diese Fehler werden nicht wiederholt. Prüfe den exakten Cmdlet-Fehler, Exchange RBAC, vorhandene Richtliniennamen und die installierte Modulversion. Starte danach nur den Reparaturmodus:

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail 'test1@monteaufkirchen.com'
```

## Sichere Aufbewahrung

- Bewahre Quelle, Backup und temporäre Dateien verschlüsselt und zugriffsbeschränkt auf.
- Kopiere produktive Dateien nicht in `tests/fixtures/`.
- Poste keine Vergleichstabellen mit realen Namen in Tickets oder Chats.
- Lösche Sicherungen nur gemäß der gültigen Aufbewahrungsregel und erst nach einer geprüften Rücksicherung.
- Verwende keine Transcripts oder Debug-Ausgaben, die Passwörter oder vollständige Schülerdaten erfassen könnten.
- Prüfe vor jedem Commit mit `git status --short`, dass keine produktive XLSX oder Wiederherstellungsdatei vorgemerkt ist.

## Abmeldung nach dem Lauf

```powershell
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
```

## Änderungen nach der Codeprüfung vom 27.09.2026

Leere Schülerdateien werden vor Graph-/Exchange-Zugriffen abgewiesen. Die direkte Schülerrolle wird vor Bestandsänderungen und Abgangsaktionen erneut geprüft. Im Lehrerabgleich werden Rollenmitgliedschaften und die aktuellen Lizenzzuweisungen der verwalteten Rollen vor Änderungen gelesen.

Bei Lehrer-Neuanlagen speichert Excel die Zugangsdaten und Objekt-ID direkt nach der gesperrten Anlage. Wiederanlauf mit bereits gesicherter Zeile und der Sonderfall einer fehlgeschlagenen Excel-Sicherung sind in der [Personalanleitung](ENTRA-LEHRER-SYNC.md#sicherheitsprüfungen-und-wiederanlauf) beschrieben. Abgelehnte Lehrerupdates lösen keine Identitätsrückschreibung aus.

Neue Schülerpasswörter bestehen aus zwölf zufälligen, gut lesbaren Buchstaben und Ziffern. Die bisherigen Zweiwortpasswörter werden nicht geändert. Schüler müssen ihr Passwort weiterhin nicht bei der ersten Anmeldung wechseln.
