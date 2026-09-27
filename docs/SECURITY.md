# Security Policy

I take the security of my software and services serious, which includes all my repositories.

If you believe you have found a security vulnerability in any of my repositories, please report it as described below.

## Repository-specific data handling

This project processes student and personnel identities, account status, group memberships and initial passwords. Use synthetic data for reports and reproductions. Do not attach production workbooks, backups, account tables, transcripts, access tokens or passwords, including to private reports.

Production workbooks and their temporary copies are excluded from Git. Reports created with `-OutputFile` must use a new `.txt` file and an ignored path inside any Git repository. These checks do not replace access controls and secure storage. Existing reports are never overwritten, and `-WhatIf` does not write a report.

The account-management TUI can delete accounts after confirmation. It protects shared mailboxes but cannot establish that an account belongs to a real person or has been disabled for 90 days. See the [account-management guide](GESPERRTE-KONTEN-VERWALTUNG.md) for the exact boundaries and [operations guide](BETRIEB.md) for recovery procedures.

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
