# Security Policy

The [address policy](ADRESSREGELN.md) applies only to new accounts. Existing address fields and SMTP aliases must not be migrated during synchronization or ordinary Exchange repair. New school staff primary and alias addresses are reserved together against Graph and Exchange ownership. The external naming rule does not lift the protection for external administrators.

I take the security of my software and services serious, which includes all my repositories.

If you believe you have found a security vulnerability in any of my repositories, please report it as described below.

## Repository-specific data handling

This project processes student and personnel identities, account status, group memberships and initial passwords. Use synthetic data for reports and reproductions. Do not attach production workbooks, backups, account tables, transcripts, access tokens or passwords, including to private reports.

Production workbooks and their temporary copies are excluded from Git. Reports created with `-OutputFile` must use a new `.txt` file and an ignored path inside any Git repository. These checks do not replace access controls and secure storage. Existing reports are never overwritten, and `-WhatIf` does not write a report.

The account-management TUI can delete accounts after confirmation. It protects shared mailboxes but cannot establish that an account belongs to a real person or has been disabled for 90 days. See the [account-management guide](GESPERRTE-KONTEN-VERWALTUNG.md) for the exact boundaries and [operations guide](BETRIEB.md) for recovery procedures.

The separate personnel entry point in 0.3.0 excludes teachers, Ganztag, pupils and shared mailboxes, including transitive role membership. External administrator and GlobalAdministrators group members and directory-role holders remain visible but protected. Missing or ambiguous protection groups block personnel writes. Identity, tenant, membership and recipient checks are repeated before mutations. The legacy disabled-account entry retains its existing scope.

The personnel `-ManageAccounts` TUI explicitly permits blocking guests and deleting already disabled guests after confirmation. It also accepts uniquely identified `MailUser` and `GuestMailUser` recipients for guests. Roleless guests are included. Guest unlocking and ordinary personnel profile maintenance remain blocked. Account-type changes and recipient changes after selection block the action. Guest permission does not bypass administrator, excluded-role, shared-mailbox, local AD or tenant protections. The detail view displays actual role groups and account type.

Teacher `-ManageAccounts` reuses these action checks for direct members of the configured school-team and Ganztag roles. Every action confirms current direct role membership again. Pupils and shared mailboxes are excluded, protected administrators and directory-role holders remain read-only. Refresh preserves the teacher population. The teacher configuration supplies roles and expected tenant, the personnel configuration supplies the shared administrator protections. Management is a separate parameter set, incompatible with sync, export and individual maintenance. `-WhatIf` does not request additional write scopes or mutate accounts.

Personnel exports contain nine columns with empty passwords and literal text cells. Existing destinations and symbolic links are rejected. Both temporary and final workbooks must be ignored in a Git repository. A roleless export row never acquires a guessed profile. Ordinary personnel updates preserve disabled accounts and do not implicitly disable departures. Explicit partial-creation recovery is separate from normal account maintenance.

Student and teacher `-ListToExcel` use the same verified file writer and request `User.Read.All`, `Group.Read.All` and `AuditLog.Read.All` for sign-in activity, plus a matching read connection to Exchange. They do not change tenant data. Student exports have ten columns, teacher exports nine, always with empty passwords. Both include disabled direct role members. Review these rows before reimport because the existing student and teacher update paths can reactivate them. Missing or conflicting mappings remain empty with warnings. No guessed class, manager or job is assigned.

See the [personnel guide](ENTRA-PERSONAL-SYNC.md) and [recorded verification](ABNAHME.md) for tested behavior and remaining Windows and tenant checks.

All exports append `Last Login Timestamp` as UTC text from the last successful sign-in. Missing values stay empty. Unavailable sign-in data produces a warning and a retry without that property. The additional column is informational and is ignored by the import. An empty timestamp does not prove that the account was never used.

Student `-ManageAccounts` is restricted to direct pupil-role members. Teacher/Ganztag membership and shared mailboxes are excluded. Guests, locally synchronized identities, protected administrator groups and directory roles cannot be mutated. K selects an existing unique, non-dynamic and non-licensing CLS group. The new membership and class attributes are verified before old CLS memberships are removed. L selects an active internal direct school-team member and changes only the manager reference. Account identity, status, memberships, tenant, recipient and the selected source and target are rechecked before writes. Neither assignment reactivates a disabled pupil. Class changes request `User.ReadWrite.All` and `GroupMember.ReadWrite.All`, teacher changes request `User.ReadWrite.All`. Preview does not request extra write scopes. Partial changes are reported and are not automatically rolled back. Workbooks are not updated, an outdated workbook can reverse these assignments during a later sync.

Read-only personnel comparison tolerates inventory-only rows from a complete export. Missing mandatory values, unknown/protected jobs and explicitly identified non-writable accounts produce row-specific `Workbook.Review` warnings and no desired mutations. Known review identities are excluded from departure candidates. This tolerance is enabled only for personnel Compare. Update (including WhatIf), individual maintenance and teacher/student imports retain strict validation. Review warnings never authorize a protected account write.

## :mega: Reporting Security Issues

> Please do not report security vulnerabilities through public GitHub issues. This is a security risk itself, as it could allow malicious users to exploit the vulnerability before it is fixed.

Instead, please open a [Draft Security Advisory](../../../security/advisories/new).

Please include the requested information listed below (as much as you can provide) to help me better understand the nature and scope of the vulnerability:

- Type of issue (e.g., buffer overflow, SQL injection, cross-site scripting, credentials stored in code, etc.).
- Full paths of source file(s) related to the manifestation of the issue.
- The location of the affected source code (tag/branch/commit or direct URL).
- Any special configuration required to reproduce the issue.
- Step-by-step instructions to reproduce the issue.
- Proof-of-concept or exploit code (if possible).
- Impact of the issue, including how an attacker might exploit the issue.

This information will help me to understand your report more quickly.

## :us: Preferred Language

I prefer all communications to be in English. This helps to ensure that vulnerabilities are understood and can be addressed quickly.
