# Betrieb und Wiederanlauf

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

`-WhatIf` erzeugt keine Passwörter, schreibt weder Excel noch Graph oder Exchange und wartet nicht auf Postfächer. Es zeigt die geplanten Phasen als `WhatIf`.

Für eine Teilaktion:

```powershell
.\Sync-SchuelerEntra.ps1 -Update -CreateNewUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -Update -UpdateUsers -WhatIf
.\Sync-SchuelerEntra.ps1 -Update -DisableUsers -RevokeSessions -WhatIf
```

## 5. Autorisierten Lauf starten

> [!CAUTION]
> Abgänge sind alle nicht zugeordneten Mitglieder der konfigurierten Gruppe `SEC-A-ROL-Schule_Schüler`. Die Erkennung ist nicht auf eine Klasse oder einen Arbeitsmappenausschnitt begrenzt. Für jede Aktion mit `-DisableUsers` oder `-RevokeSessions` muss die Excel-Datei deshalb die vollständige konfigurierte Schülerpopulation enthalten. Eine einzelne Testklasse ist in einer gemeinsam genutzten produktiven Rollengruppe nicht isoliert. Verwende für destruktive Tests einen vollständig isolierten Mandanten oder eine Rollengruppe, deren Mitglieder ausschließlich synthetische Testkonten sind.

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

Bei einer erfolgreichen Neuanlage erscheint das Passwort nach verifizierter Aktivierung genau einmal im Terminal. Scheitert ein später Schritt nach der Kontoerstellung, erscheint es ebenfalls einmal für den Wiederanlauf. Der Bericht nennt dann einen unbekannten Kontostatus, die Objekt-ID und einen `-EntraObjectId`-Wiederanlaufbefehl. Prüfe die Objekt-ID vor der Ausführung. Starte dafür kein Transcript und sichere das Passwort unmittelbar. Ist der Schüler bereits eindeutig in der direkten Schüler-Rollengruppe vorhanden, aktualisiert das Skript nur seine Abweichungen. UPN und Passwort bleiben unverändert. Ein deaktiviertes Bestandskonto wird nach verifiziertem Pflichtzustand aktiviert.

Zum Deaktivieren eines einzelnen Schülers und Widerrufen seiner Sitzungen:

```powershell
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com' -WhatIf
.\Sync-SchuelerEntra.ps1 -Remove -UPN 'mmuster@monteaufkirchen.com'
```

Prüfe vor dem zweiten Befehl den exakten UPN und die angezeigte Tenant-ID. `-Remove` akzeptiert nur direkte Mitglieder von `SEC-A-ROL-Schule_Schüler`. Konto, Gruppen und Lizenzen werden nicht gelöscht.

## Exchange-Reparatur

Ein neues Postfach wird sofort und danach bis zu fünfmal im Abstand von 60 Sekunden geprüft. Das Warten geschieht einmal pro Batch-Runde. Nach dem letzten Versuch bleibt der Entra-Zustand erhalten und das Ergebnis meldet `EXO-Konfiguration ausstehend`.

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

Der Exchange-Modus akzeptiert auch `-UPN` und `-UserPrincipalName`. Er liest keine Excel-Datei und ändert weder Entra-Attribute noch Gruppen.

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
