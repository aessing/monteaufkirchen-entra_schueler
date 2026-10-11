# UPN und Mail bei Neuanlagen

Diese Regeln gelten für neue Konten im Excel-Sync und bei `-Add`.

| Kontotyp | UPN und primäre Mailadresse | Zusätzlicher Alias |
|---|---|---|
| Schüler | `<name>@monteaufkirchen.com` | keiner |
| Schulteam mit `L`, `CO-L` oder `PA`, auch mit Zusatzjob `JAS` | `<name>@montessori-erding.de` | `<name>@monteaufkirchen.com` |
| OGTS und übriges Personal | `<name>@montessori-erding.de` | keiner |
| Freigegebene Personalrolle `Externe*` | `<vorname>.<nachname>@montessori-erding.de` | keiner |

Der Lehrer-Alias gehört zum Profil `School`. OGTS gehört zum Profil `Ganztag`. `ExterneAdmins` bleibt geschützt und kann weiterhin nicht angelegt oder gepflegt werden. Die Namensregel hebt diesen Schutz nicht auf.

## Namensbildung und Kollisionen

Für Schüler und Angestellte beginnt `<name>` mit dem ersten Buchstaben des normalisierten Vornamens und dem vollständigen Nachnamen. Das Muster für Oliver Winkels ist `owinkels`. Ist die Adresse belegt, folgt `olwinkels`, danach wächst das Vornamenspräfix weiter. Sind alle Präfixe belegt, folgt `oliverwinkels2`, dann `oliverwinkels3`.

Externe Benutzer erhalten den vollständigen normalisierten Vornamen, einen Punkt und den vollständigen Nachnamen. Bei einer Kollision folgt unmittelbar eine Zahl ab `2`, beispielsweise `lea.testwald2@montessori-erding.de`. Der Vorname wird dabei nicht verkürzt.

Umlaute werden umgeschrieben, Akzente entfernt und andere Sonderzeichen aus den Namen entfernt. Namen in `DisplayName`, `GivenName` und `Surname` behalten ihre Schreibweise.

Die Prüfung berücksichtigt Entra-UPNs, Mailadressen, SMTP-Proxyadressen und Exchange-Empfänger. Beim Schulteam werden Hauptadresse und Alias gemeinsam geprüft und für den laufenden Batch reserviert. Ist nur der Alias belegt, wächst das Präfix für beide Adressen.

## Bestandsschutz

Bestehende UPNs, Mailadressen, Mail-Nicknames und SMTP-Aliase bleiben unverändert. Das gilt auch bei Namens- und Rollenwechseln sowie für fehlende Mailwerte oder Mail-Nicknames. Die Skripte stellen bestehende Konten nicht auf die neuen Domains oder Namen um und ergänzen keinen Lehrer-Alias im Bestand.

Die übrige Profil-, Gruppen-, Manager-, Lizenz- und Postfachpflege behält ihre bisherigen Regeln. Ein normaler `-ConfigureExchangeOnlineOnly`-Aufruf repariert diese Postfacheinstellungen und verändert keine Adressen. Exporte übernehmen die tatsächlich vorhandenen Identitäten.

## Exchange und Teilfehler

Bei einer Neuanlage erhalten UPN und Entra-`Mail` dieselbe Hauptadresse. Die anschließende Exchange-Konfiguration prüft die primäre SMTP-Adresse und beim Schulteam zusätzlich den Alias. Beim Schreiben bleiben andere SMTP-Aliase und Nicht-SMTP-Adressen wie `X500` erhalten. Ein erneutes Lesen bestätigt den Zustand. Ein fehlender Alias darf nicht als erfolgreiche Konfiguration erscheinen.

Das Skript wartet im bisherigen Rahmen auf das Postfach. Bleibt es ausstehend oder schlägt die Konfiguration fehl, erscheint `Pending` beziehungsweise `Failed`. Das bereits angelegte Konto wird nicht erneut erstellt.

Wenn nur der Alias der gerade erfolgten Lehreranlage fehlt, prüfe die tatsächliche Hauptadresse und die Belegung des vorgesehenen Alias. Ergänze ihn danach gezielt in Exchange. Der normale Reparaturmodus ergänzt wegen des Bestandsschutzes keine Adressen. Beispiel ausschließlich für ein geprüftes synthetisches Neukonto:

```powershell
Set-Mailbox -Identity 'ltestwald@montessori-erding.de' -EmailAddresses @{ Add = 'smtp:ltestwald@monteaufkirchen.com' } -WhatIf
# Nach Prüfung denselben Aufruf ohne -WhatIf ausführen.
Get-Mailbox -Identity 'ltestwald@montessori-erding.de' | Select-Object PrimarySmtpAddress,EmailAddresses
```

Microsoft Graph erlaubt keine direkte Änderung von `proxyAddresses`. Exchange kennzeichnet die primäre Adresse mit `SMTP:` und Aliase mit `smtp:`. Beim Setzen der vollständigen Adressliste müssen alle vorhandenen Adressen enthalten bleiben. Quellen: [Microsoft Graph: Benutzer](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0), [Microsoft Exchange: Postfachadressen](https://learn.microsoft.com/en-us/exchange/recipients/user-mailboxes/email-addresses).

## Lokale Prüfung

`CreationIdentity.Tests.ps1` prüft Domains, Rollen, externe Namen, Präfix- und Zahlenkollisionen sowie die gemeinsame Reservierung von Hauptadresse und Lehrer-Alias. Die Schülervergleichstests prüfen abweichende und fehlende Bestandsadressen. Die Exchange-Tests prüfen Vorschau, Erhalt zusätzlicher Adressen, Rückleseprüfung und Idempotenz. Alle Fernzugriffe sind synthetisch ersetzt. Diese Tests belegen keine reale Domainfreigabe oder Postfachbereitstellung im Mandanten.
