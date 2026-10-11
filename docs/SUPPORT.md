# Support Information

New-account domains and naming are described in the [address policy](ADRESSREGELN.md). Existing UPNs, mail addresses, mail nicknames and aliases are preserved. Ordinary Exchange repair does not add or change addresses. If a newly created school staff mailbox is missing its school alias after a provisioning failure, follow the targeted completion in that policy after checking the actual primary address and alias ownership. This does not require another account or password.

## Operating guides and troubleshooting

- [Student synchronization](ENTRA-SCHUELER-SYNC.md)
- [Teacher and Ganztag synchronization](ENTRA-LEHRER-SYNC.md)
- [Remaining personnel accounts, inventory, Excel export and TUI](ENTRA-PERSONAL-SYNC.md)
- [Disabled-account management](GESPERRTE-KONTEN-VERWALTUNG.md)
- [Operations and recovery](BETRIEB.md)
- [Testing](TESTING.md) and [recorded acceptance results](ABNAHME.md)

For a bug report, include the script version, PowerShell and module versions, selected mode, anonymized error text and a reproduction with synthetic accounts. Remove real names, UPNs, account IDs, passwords and tokens from screenshots and logs. Report security vulnerabilities through the private process in [SECURITY.md](SECURITY.md).

For a teacher creation where only Exchange failed, use `Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN '<existing UPN>' -WhatIf`, then repeat without `-WhatIf` to apply the missing mailbox settings after confirmation. The account and password remain intact. Teacher mailbox provisioning is checked immediately, then retried up to three times at 150-second intervals. Recipient-not-found errors during reads, writes or verification remain retryable. Other errors are reported immediately. If Entra steps also failed, follow the selective Excel recovery in the [teacher guide](ENTRA-LEHRER-SYNC.md#sicherheitsprüfungen-und-wiederanlauf), retaining the saved identity and password cells.

All three synchronization scripts support `-ListToExcel`. Use `-File` for a new destination if the default workbook exists. Warnings about missing class, manager or job mappings require row review before a writing import. They do not imply that the export failed. Student and teacher exports also report disabled accounts because their existing update behavior can reactivate listed accounts. See the linked guides for the exact schemas and population rules.

The trailing `Last Login Timestamp` column uses the last successful sign-in in UTC. Empty cells mean that no value was returned. If sign-in activity cannot be read, the export warns and retries without it. Student and teacher exports now also request `AuditLog.Read.All`. Review sign-in permissions and licensing in the operator guides when that warning appears.

Personnel `-ManageAccounts` shows the selected account's role groups and account type. `Keine Rollengruppe` means no role group was found in the loaded memberships. Use `B` to block an eligible active guest, then select it again and press Backspace or Delete to delete it after the inventory refresh. Guests with no Exchange recipient or a unique personal mailbox, `MailUser` or `GuestMailUser` are supported. Guests cannot be unlocked here. Protected administrators, excluded roles and unclear recipients remain blocked. Ordinary sync and the legacy disabled-account entry retain their guest restrictions.

If Backspace or Delete is pressed for a marked active guest, deletion is skipped with an explicit instruction to block it using `B` first. The current selection and cursor remain intact. Entirely skipped actions and cancelled confirmations do not reload the inventory. After confirmed blocking, mark the refreshed disabled guest again, press Backspace or Delete and confirm deletion using `Y` or `J`.

Teacher `-ManageAccounts` uses the same TUI for active and disabled direct school-team and Ganztag members. It displays role groups, account type and last successful sign-in. Direct membership is checked again before blocking, unlocking or deleting, and refresh keeps the teacher scope. Pupils, shared mailboxes and protected administrators cannot be acted on. Run `Sync-LehrerEntra.ps1 -ManageAccounts -WhatIf` first. This mode requires an interactive terminal and cannot be combined with workbook or sync parameters.

Student `-ManageAccounts` adds K for class/year and L for the class teacher, alongside B, U and Backspace or Delete for account lifecycle actions. Class targets are existing unique writable CLS groups, teacher targets are active internal direct members of the school pedagogical team. Run `Sync-SchuelerEntra.ps1 -ManageAccounts -WhatIf` first. A class change confirms new membership and attributes before removing old CLS memberships. Disabled students stay disabled during assignment changes. If an action reports a possible partial change, inspect the refreshed class groups and attributes before retrying. Excel sources are not updated and must be aligned before the next sync. See the [student TUI guide](ENTRA-SCHUELER-SYNC.md#schüler-tui-manageaccounts).

If a personnel comparison after ListToExcel reports a protected job such as ExterneAdmins as a preflight failure, use the current 0.3.0 code. A read-only personnel comparison now reports incomplete, unclassified or protected rows as `Workbook.Review` with row number and identity and continues with valid rows. Explicitly identified non-writable accounts are also review-only. Confirmed Entra guests are skipped without incomplete-data or import-review warnings in Compare and Update, including WhatIf. Guests are identified through object ID or UPN, and both must match when supplied together. Missing names or jobs do not block the run for confirmed guests. Unknown or conflicting identities retain validation. Guests stay in inventories and exports and never become profile updates, new users or departures. A guest-only workbook produces no sync actions or member departures. Known member accounts with a blank Job and no managed personnel role are also skipped with a `Workbook.Review` warning, even when names are missing. This requires a complete protection context and freshly read, fully identifiable transitive Entra memberships. When object ID and UPN are both supplied, they must identify the same account. Missing Jobs on managed personnel accounts, unknown or conflicting identities, and incomplete safety information retain strict validation. Assigning a Job invokes ordinary validation and explicit object ID requirements for taking over an account. A workbook containing only skipped accounts never generates member departures. Known member accounts with valid Jobs and missing name fields are skipped with `Workbook.Review: Unvollständige Namensdaten, nicht abgeglichen.` when a complete protection context and freshly read, fully identifiable transitive memberships confirm exactly one managed personnel role. These accounts remain unchanged and do not become creation or departure candidates. Valid personnel updates continue. Add the actual name fields to process those accounts normally. Unknown or conflicting identities, invalid Jobs, multiple personnel roles and incomplete safety information retain strict validation. Update remains strict for other rows and requires a reviewed source. An exported `Externe Admins` row no longer aborts Update when its object ID or UPN identifies a confirmed protected administrator. Protected group and directory role memberships are checked against Entra. These rows remain `Workbook.Review` warnings and are skipped in Compare and Update, including WhatIf. Valid member changes are applied independently, and unknown or conflicting identities still block invalid imports. Do not change a protected account's Job value to bypass its protections.

All account TUIs open deletion confirmation with Backspace or Delete. Enter does nothing in account lists and still selects a class or teacher target. Confirm with Y/J or cancel with N/Esc.

Personnel exports display readable job names such as `Kinderhaus Pädagogisches Team` and `Externe Benutzer`. Workbook imports and `-Add -Job` resolve these labels and the previous internal names to the configured canonical profile. Case and repeated whitespace are ignored. Unknown or ambiguous jobs are rejected. `Externe Admins` remains protected, including on writable imports. Personnel JobTitle is written as the readable label during creation and profile maintenance. For the three Kinderhaus profiles, the Entra title omits the `Kinderhaus` prefix, yielding `Pädagogisches Team`, `Sekretariat` or `Vertretungskräfte`. CompanyName remains `Montessori Kinderhaus Erding`. Workbook job labels retain the prefix for unambiguous profile mapping. Existing technical titles or titles with the old Kinderhaus prefix appear as differences and are corrected only by an explicitly started update or single-account maintenance. Inventory and comparison reports show readable labels too. Internal profile and role identifiers keep their canonical values.

## PowerShell confirmation prompts

To apply a reviewed personnel update without individual PowerShell Yes/No prompts, run this command in PowerShell:

```powershell
./Sync-PersonalEntra.ps1 -File ./Personal.xlsx -Update -Confirm:$false
```

The common `-Confirm:$false` parameter is forwarded to the selected write operations. It does not enable extra actions or bypass preflight, identity, role or administrator protections. Personnel `-Update` alone creates and updates accounts. Departures require explicit action switches. `-WhatIf` previews the planned actions without applying them. The TUI's own account-name confirmation remains required even with `-Confirm:$false`.

## :cop: Code of conduct

My [Code of Conduct](CODE_OF_CONDUCT.md) outlines expectations for participation in the community, as well as steps for reporting unacceptable behavior. I am committed to provide a welcoming and inspiring community for all. People violating this code of conduct may be banned from the community and the repository.

## :construction_worker: Contributor's Guide

In my [Contributor's Guide](CONTRIBUTING.md) you can find some guidance for how to report issues, propose new features, and submit contributions via Pull Requests. Please read it, before you start working. I am looking forward to your contributions!

## :scroll: License

You can find the license information for this repository in its [license file](../LICENSE).

## :detective: Security Policy

I take the security of my software and services serious, which includes all my repositories.
If you believe you have found a security vulnerability in any of my repositories, please report it as described in my [Security Policy](SECURITY.md).

## :us: Preferred Language

I prefer all communications to be in English. This helps to ensure that vulnerabilities are understood and can be addressed quickly.
