# Entra-Schülersynchronisation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ein PowerShell-Werkzeug erstellt einen standardmäßig lesenden Soll-Ist-Vergleich zwischen einer frei wählbaren Schüler-Exceldatei, Microsoft Entra ID und Exchange Online und führt nur ausdrücklich autorisierte, verifizierte Änderungen aus.

**Architecture:** `Sync-SchuelerEntra.ps1` bleibt ein dünner Einstiegspunkt. Ein internes PowerShell-Modul trennt reine Normalisierungs- und Vergleichslogik von Excel-, Graph- und Exchange-Adaptern, damit Pester die Fachlogik mit synthetischen Daten und Mocks prüfen kann. Schreibende Abläufe führen zuerst eine vollständige Vorprüfung aus, konfigurieren neue Benutzer deaktiviert, sichern die Exceldatei atomar und verifizieren jeden externen Endzustand.

**Tech Stack:** PowerShell 7, Microsoft Graph PowerShell SDK v1.0, ExchangeOnlineManagement, ImportExcel, Pester 5, PSScriptAnalyzer, Git, PNG-Asset für die README

**Spec:** `docs/superpowers/specs/2026-09-11-entra-schueler-sync-design.md`

## Global Constraints

- Ohne `-Update` ist der Synchronisationsmodus vollständig lesend.
- `-ConfigureExchangeOnlineOnly` ist ein eigener Parameter-Satz und verändert nur Exchange Online.
- Neue Schüler werden deaktiviert erstellt und erst nach erfolgreicher Pflichtkonfiguration und Excel-Rückschreibung aktiviert.
- Produktive Passwörter sind exakt zwölf Zeichen lang, werden nur für Neuzugänge erzeugt und erscheinen niemals in Konsole, Logs oder Fixtures. Tests verwenden ausschließlich synthetische Werte in temporären Arbeitsmappen.
- `Schueler.xlsx` ist der Standard. `-File` akzeptiert jeden relativen oder absoluten `.xlsx`-Pfad.
- Die produktive Exceldatei und Root-Sicherungen werden nicht versioniert.
- UPN, Mail, Gruppen und Manager werden nach jeder Änderung erneut gelesen.
- Andere direkte `SEC-A-ROL-*`- und `SEC-A-CLS-*`-Gruppen werden erst entfernt, nachdem die jeweilige Zielgruppe nachweislich hinzugefügt wurde.
- Fehlende Postfächer werden einmal sofort und danach höchstens fünfmal mit 60 Sekunden Abstand gesammelt geprüft.
- Live-Tenant-Tests und die Passwortanmeldung führt der Benutzer in einer Windows-Parallels-VM mit einem synthetischen Testkonto aus.
- Keine Beta-Graph-Endpunkte, keine Benutzerlöschung und kein automatischer tenantweiter Rollback.

---

## File Responsibility Map

| Datei | Verantwortung |
|---|---|
| `Sync-SchuelerEntra.ps1` | Öffentliche Parameter, Parameter-Sätze, Modulimport und Aufruf von `Invoke-SchuelerSync` |
| `config/SchuelerSync.psd1` | Nicht geheime Domain-, Gruppen-, Rollen-, Attribut- und Exchange-Sollwerte |
| `src/SchuelerSync/SchuelerSync.psd1` | Modulmanifest und einziger öffentlicher Export `Invoke-SchuelerSync` |
| `src/SchuelerSync/SchuelerSync.psm1` | Dot-Sourcing der Private-Dateien und Orchestrierung |
| `src/SchuelerSync/Private/Normalization.ps1` | UPN-Normalisierung, UPN-Auswahl, Office Location und Passwortgenerator |
| `src/SchuelerSync/Private/Excel.ps1` | Excel-Import, Schema-Prüfung, Git-Schutz, Sicherung und atomare Rückschreibung |
| `src/SchuelerSync/Private/Entra.ps1` | Graph-Verbindung, Snapshot, Manager, Gruppen und Entra-Schreiboperationen |
| `src/SchuelerSync/Private/Comparison.ps1` | Stabile Identitätszuordnung, Sollzustände, Kategorien und Differenzen |
| `src/SchuelerSync/Private/ExchangeOnline.ps1` | EXO-Snapshot, Soll-Ist-Vergleich, Konfiguration und gesammelte Wiederholung |
| `src/SchuelerSync/Private/Update.ps1` | Auswahl und Ausführung autorisierter Mutationen |
| `src/SchuelerSync/Private/Reporting.ps1` | Konsolentabellen, Zusammenfassung, Warnungen und Reparaturbefehle |
| `tests/*.Tests.ps1` | Synthetische Unit-, Adapter- und Orchestrierungstests |
| `tests/fixtures/Schueler-Testdaten.xlsx` | Arbeitsmappe mit ausschließlich erfundenen Personen |
| `docs/ENTRA-SCHUELER-SYNC.md` | Installation, Bedienung, Berechtigungen, Fehlerbehebung und VM-Testlauf |
| `docs/assets/entra-schueler-sync-hero.png` | Generisches breites 3D-Comic-Hero-Bild |
| `README.md` | Projektübersicht und Schnellstart |

## Shared Object Contracts

`Read-StudentWorkbook` liefert:

```powershell
[pscustomobject]@{
    Path          = [string]
    WorksheetName = [string]
    Headers       = [string[]]
    Students      = [pscustomobject[]]
}
```

Jeder Student besitzt `RowNumber`, `NameMitRufname`, `GivenName`, `Surname`, `ClassName`, `Teacher`, `Password`, `EntraObjectId`, `StoredUpn` und `StoredMail`.

`Get-EntraSnapshot` liefert:

```powershell
[pscustomobject]@{
    UsersById             = [hashtable]
    UsersByUpn            = [hashtable]
    StudentRoleMemberIds  = [System.Collections.Generic.HashSet[string]]
    GroupsById            = [hashtable]
    GroupsByDisplayName   = [hashtable]
    ReservedAddresses     = [System.Collections.Generic.HashSet[string]]
    AddressOwners         = [hashtable]
    DirectGroupsByUserId  = [hashtable]
    TransitiveGroupsByUserId = [hashtable]
    ManagerByUserId       = [hashtable]
}
```

`AddressOwners` maps a normalized address to the owning Entra object IDs. Exchange recipient owners are merged by `ExternalDirectoryObjectId`. `Compare-StudentDirectory` liefert `NewStudents`, `Departures`, `ChangedStudents`, `ExistingStudents`, `Warnings` und `Errors` als Arrays. Jede Differenz besitzt `Area`, `Field`, `Current`, `Desired` und `Action` mit `Add`, `Remove` oder `Set`.

---

### Task 1: Repository-Schutz, Konfiguration und Modulgerüst

**Files:**
- Modify: `.gitignore`
- Create: `config/SchuelerSync.psd1`
- Create: `src/SchuelerSync/SchuelerSync.psd1`
- Create: `src/SchuelerSync/SchuelerSync.psm1`
- Create: `Sync-SchuelerEntra.ps1`
- Create: `tests/Module.Tests.ps1`

**Interfaces:**
- Consumes: Keine produktiven Schnittstellen.
- Produces: `Invoke-SchuelerSync` als einzigen öffentlichen Export und den Konfigurationsvertrag aus `config/SchuelerSync.psd1`.

- [ ] **Step 1: Write the failing module contract test**

Create `tests/Module.Tests.ps1`:

```powershell
BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $manifest = Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1'
    $configPath = Join-Path $repoRoot 'config/SchuelerSync.psd1'
}

Describe 'SchuelerSync module contract' {
    It 'imports and exports only Invoke-SchuelerSync' {
        Import-Module $manifest -Force
        (Get-Command Invoke-SchuelerSync -Module SchuelerSync).Name |
            Should -Be 'Invoke-SchuelerSync'
        @(Get-Command -Module SchuelerSync).Name |
            Should -Be @('Invoke-SchuelerSync')
    }

    It 'contains the immutable student configuration' {
        $config = Import-PowerShellDataFile $configPath
        $config.Domain | Should -Be 'monteaufkirchen.com'
        $config.StudentRoleGroup.Id |
            Should -Be 'cebc1326-1174-4126-ba84-7a8960850e0a'
        $config.PasswordLength | Should -Be 12
        $config.Exchange.MaxMailboxRetries | Should -Be 5
        $config.Exchange.RetryDelaySeconds | Should -Be 60
    }
}
```

- [ ] **Step 2: Run the test and confirm the missing manifest failure**

```powershell
pwsh -NoLogo -NoProfile -Command "Invoke-Pester tests/Module.Tests.ps1 -Output Detailed"
```

Expected: FAIL because the manifest and configuration do not exist.

- [ ] **Step 3: Add the workbook ignore rule**

Append to `.gitignore`:

```gitignore

# ---------------------- Schüler-Importdaten ---------------------- #
# Contains personal data and generated initial passwords.
/*.xlsx
```

Run `git check-ignore -v Schueler.xlsx` and require the new rule to match.

- [ ] **Step 4: Create the exact non-secret configuration**

Create `config/SchuelerSync.psd1`:

```powershell
@{
    Domain = 'monteaufkirchen.com'
    ExpectedTenantId = $null
    CompanyName = 'Montessori Schule Aufkirchen'
    EmployeeType = 'Schüler'
    AgeGroup = 'Minor'
    ConsentProvidedForMinor = 'Granted'
    LegalAgeGroupClassification = 'MinorWithParentalConsent'
    UsageLocation = 'DE'
    PasswordLength = 12
    StudentRoleGroup = @{
        Name = 'SEC-A-ROL-Schule_Schüler'
        Id = 'cebc1326-1174-4126-ba84-7a8960850e0a'
    }
    LicenseGroupName = 'SEC-A-LIC-O365A1Student'
    RoleGroupPrefix = 'SEC-A-ROL-'
    ClassGroupPrefix = 'SEC-A-CLS-'
    KnownRoleGroups = @(
        @{ Name = 'SEC-A-ROL-Schule_PädagogischesTeam'; Id = '103a1c4c-036b-40bc-bbe8-e3f7a866cb01' }
        @{ Name = 'SEC-A-ROL-Schule_Schüler'; Id = 'cebc1326-1174-4126-ba84-7a8960850e0a' }
        @{ Name = 'SEC-A-ROL-Schule_Sekretariat'; Id = '534f94ed-420d-4794-a374-20c572def8cb' }
        @{ Name = 'SEC-A-ROL-Schule_Vertretungskräfte'; Id = '4ed4f87b-374f-4679-ac11-5e51212e9a91' }
        @{ Name = 'SEC-A-ROL-Kinderhaus_PädagogischesTeam'; Id = 'cfc6d358-875d-441b-b45a-0d9734b7d788' }
        @{ Name = 'SEC-A-ROL-Kinderhaus_Sekretariat'; Id = '43165f2c-3ce5-471e-a7be-e16b66c03406' }
        @{ Name = 'SEC-A-ROL-Kinderhaus_Vertretungskräfte'; Id = 'd2e7e424-83d2-481d-9226-bf529bc9640a' }
        @{ Name = 'SEC-A-ROL-Ganztag'; Id = '3997db35-2914-4b0f-92b4-08038ef8f84a' }
        @{ Name = 'SEC-A-ROL-Unterstützung'; Id = 'a4a00442-bc04-48ee-b659-cb3bde194630' }
        @{ Name = 'SEC-A-ROL-ExterneBenutzer'; Id = '9c620843-0071-4d44-ba23-318e19cd09a0' }
        @{ Name = 'SEC-A-ROL-ExterneAdmins'; Id = 'a246dfde-eccf-489b-ac1c-8d92f67145ca' }
    )
    Exchange = @{
        AddressBookPolicy = 'MON-EXO-ABP-Schule_Schüler'
        CustomAttribute1 = 'Montessori Schule Aufkirchen - Schüler'
        AuditLogAgeLimitDays = 365
        RetainDeletedItemsForDays = 30
        RoleAssignmentPolicy = 'MON-EXO-UserRoles-Default'
        SharingPolicy = 'MON-EXO-Sharing-Default'
        RetentionPolicy = 'MON-EXO-Retention-Default'
        OwaMailboxPolicy = 'MON-EXO-OWA-Default'
        MaxMailboxRetries = 5
        RetryDelaySeconds = 60
    }
}
```

- [ ] **Step 5: Create the module manifest and loader**

Create `src/SchuelerSync/SchuelerSync.psd1` with `RootModule = 'SchuelerSync.psm1'`, `ModuleVersion = '0.1.0'`, `PowerShellVersion = '7.0'` and only `FunctionsToExport = @('Invoke-SchuelerSync')`.

Create `src/SchuelerSync/SchuelerSync.psm1`:

```powershell
Set-StrictMode -Version Latest

Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    Sort-Object Name |
    ForEach-Object { . $_.FullName }

function Invoke-SchuelerSync {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Sync', Mandatory)][string] $File,
        [Parameter(ParameterSetName = 'Sync')][switch] $Update,
        [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
        [switch] $ConfigureExchangeOnlineOnly,
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
        [Alias('UPN', 'UserPrincipalName')][string[]] $Mail
    )
    throw 'OrchestrationNotImplemented'
}

Export-ModuleMember -Function Invoke-SchuelerSync
```

Create `Sync-SchuelerEntra.ps1` with the same parameter sets. Default `-File` to `Join-Path $PSScriptRoot 'Schueler.xlsx'`, set strict mode and `$ErrorActionPreference = 'Stop'`, then import the manifest. In the `Sync` parameter set, copy `$PSBoundParameters` and explicitly add `File = $File` when the caller omitted it so the internal mandatory parameter receives the default. Forward the resulting hashtable, including common `WhatIf`, `Confirm` and `Verbose` values, to `Invoke-SchuelerSync`.

- [ ] **Step 6: Run the focused test**

Run `Invoke-Pester tests/Module.Tests.ps1 -Output Detailed`. Expected: PASS.

- [ ] **Step 7: Commit the protected scaffold**

```bash
git add .gitignore Sync-SchuelerEntra.ps1 config/SchuelerSync.psd1 src/SchuelerSync tests/Module.Tests.ps1
git commit -m "build: scaffold safe student sync command"
```

---

### Task 2: Normalisierung, Office Location, UPN-Auswahl und Passwortgenerator

**Files:**
- Create: `src/SchuelerSync/Private/Normalization.ps1`
- Create: `tests/Normalization.Tests.ps1`
- Create: `tests/Password.Tests.ps1`

**Interfaces:**
- Consumes: Domain and Passwortlänge aus der Konfiguration.
- Produces: `ConvertTo-UpnToken`, `Get-OfficeLocation`, `Get-UpnCandidates`, `Select-AvailableUpn` and `New-StudentPassword`.

- [ ] **Step 1: Write failing normalization tests**

Create `tests/Normalization.Tests.ps1` and import the module in `BeforeAll`. Add:

```powershell
Describe 'Student identity normalization' {
    InModuleScope SchuelerSync {
        It 'transliterates German characters' {
            ConvertTo-UpnToken 'Änne Weiß' | Should -Be 'aenneweiss'
            ConvertTo-UpnToken 'JÖRG-MÜLLER' | Should -Be 'joergmueller'
        }

        It 'derives allowed office locations' -ForEach @(
            @{ Class = 'JK1-3g2_1'; Expected = 'G2' }
            @{ Class = 'JK4-6m2_4'; Expected = 'M2' }
            @{ Class = 'JK7-9O1_2'; Expected = 'O1' }
            @{ Class = 'JK7-9a2_3'; Expected = 'A2' }
        ) {
            Get-OfficeLocation $Class | Should -Be $Expected
        }

        It 'rejects an unsupported office token' {
            { Get-OfficeLocation 'JK1-3x9_1' } | Should -Throw '*Office Location*'
        }

        It 'generates prefix candidates in order' {
            $actual = @(Get-UpnCandidates -GivenName 'Maria' -Surname 'Müller' -Domain 'monteaufkirchen.com')
            $actual | Should -Be @(
                'mmueller@monteaufkirchen.com'
                'mamueller@monteaufkirchen.com'
                'marmueller@monteaufkirchen.com'
                'marimueller@monteaufkirchen.com'
                'mariamueller@monteaufkirchen.com'
            )
        }
    }
}
```

- [ ] **Step 2: Write failing password tests**

Create `tests/Password.Tests.ps1`:

```powershell
Describe 'Student password generation' {
    InModuleScope SchuelerSync {
        It 'creates exactly twelve friendly characters' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $password = New-StudentPassword -UsedPasswords $used
            $password.Length | Should -Be 12
            $password | Should -Match '^[A-Z][a-z]+[A-Z][a-z]+[0-9]{2}$'
        }

        It 'does not return a password already reserved in the run' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $first = New-StudentPassword -UsedPasswords $used
            $second = New-StudentPassword -UsedPasswords $used
            $first | Should -Not -Be $second
            $used.Count | Should -Be 2
        }
    }
}
```

- [ ] **Step 3: Run the tests and confirm missing function failures**

Run `Invoke-Pester tests/Normalization.Tests.ps1,tests/Password.Tests.ps1 -Output Detailed`. Expected: FAIL because the functions do not exist.

- [ ] **Step 4: Implement normalization and office parsing**

Create `Normalization.ps1`:

```powershell
function ConvertTo-UpnToken {
    param([Parameter(Mandatory)][string] $Value)
    $mapped = $Value.Trim().ToLowerInvariant().
        Replace('ä', 'ae').Replace('ö', 'oe').Replace('ü', 'ue').Replace('ß', 'ss')
    $decomposed = $mapped.Normalize([Text.NormalizationForm]::FormD)
    $builder = [Text.StringBuilder]::new()
    foreach ($character in $decomposed.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($character) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void] $builder.Append($character)
        }
    }
    return ($builder.ToString().Normalize([Text.NormalizationForm]::FormC) -replace '[^a-z0-9]', '')
}

function Get-OfficeLocation {
    param([Parameter(Mandatory)][string] $ClassName)
    $matches = [regex]::Matches(
        $ClassName,
        '(?i)(?<office>g[1-4]|m[1-3]|o[1-2]|a[1-2])(?=_|$)'
    )
    if ($matches.Count -ne 1) {
        throw "Klasse '$ClassName' enthält keine eindeutige erlaubte Office Location."
    }
    return $matches[0].Groups['office'].Value.ToUpperInvariant()
}
```

- [ ] **Step 5: Implement UPN candidates and collision fallback**

`Get-UpnCandidates` normalizes both names, rejects an empty normalized token and emits prefixes from one character through the complete first name. `Select-AvailableUpn` accepts a case-insensitive `HashSet[string]`, an address-owner map and optional `CurrentObjectId`, reserves the first available candidate and returns:

```powershell
[pscustomobject]@{
    Upn = $candidate
    WasFallback = [bool]
    CollisionCount = [int]
    Collisions = [string[]]
}
```

When every prefix candidate is occupied, test `<fullGivenName><surname>2@domain` and increasing suffixes. Treat an address as available to a matched user only when every known Graph or Exchange owner ID is that same `CurrentObjectId`. Add tests for collisions caused by UPN, mail, proxy address, Exchange recipient and a UPN reserved earlier in the same run.

- [ ] **Step 6: Implement the exact-length password generator**

Use word lists grouped by length 4, 5 and 6. Build only two-word pairs whose letters sum to ten. Include at least the approved words `Lego`, `Wald`, `Mond`, `Tiger`, `Wiese`, `Wolke`, `Zebra`, `Biene`, `Pizza`, `Sonne`, `Blume`, `Apfel`, `Panda`, `Koala`, `Garten`, `Rakete`, `Schule`, `Pferde`, `Bienen` and `Kuchen`.

`New-StudentPassword` takes a `HashSet[string]` and optional random-index scriptblock. Select the word pair and a number from 10 through 99 with `RandomNumberGenerator.GetInt32`. Return only a unique string of exactly twelve characters. Stop with a clear exception after 100 unsuccessful uniqueness attempts.

- [ ] **Step 7: Run and commit**

Run both focused test files. Expected: PASS.

```bash
git add src/SchuelerSync/Private/Normalization.ps1 tests/Normalization.Tests.ps1 tests/Password.Tests.ps1
git commit -m "feat: add deterministic student identity rules"
```

---

### Task 3: Excel-Adapter und sichere Rückschreibung

**Files:**
- Create: `src/SchuelerSync/Private/Excel.ps1`
- Create: `tests/Excel.Tests.ps1`
- Create: `tests/fixtures/Schueler-Testdaten.xlsx`

**Interfaces:**
- Consumes: ImportExcel and row updates with `RowNumber`, `Password`, `EntraObjectId`, `UPN` and `Mail`.
- Produces: `Resolve-StudentWorkbookPath`, `Read-StudentWorkbook`, `Assert-WorkbookSafeForPasswordWrite` and `Write-StudentWorkbookUpdates`.

- [ ] **Step 1: Create a synthetic fixture**

Generate `tests/fixtures/Schueler-Testdaten.xlsx` with ImportExcel and exactly two invented rows:

```powershell
@(
    [pscustomobject]@{
        'Name mit Rufname' = 'Muster, Mia'
        Vorname = 'Mia'
        Nachname = 'Muster'
        Klassen = 'JK1-3g2_1'
        Klassenlehrer = 'Lara Lehrer'
    }
    [pscustomobject]@{
        'Name mit Rufname' = 'Beispiel, Ben'
        Vorname = 'Ben'
        Nachname = 'Beispiel'
        Klassen = 'JK4-6m2_4'
        Klassenlehrer = 'Lars Lehrer'
    }
) | Export-Excel -Path tests/fixtures/Schueler-Testdaten.xlsx -WorksheetName Tabelle1 -TableName Schueler
```

- [ ] **Step 2: Write failing Excel tests**

Create `tests/Excel.Tests.ps1`. Copy the fixture to `TestDrive` with the name `Eigener-Dateiname.xlsx` and assert:

```powershell
$context = Read-StudentWorkbook -Path $copy
$context.WorksheetName | Should -Be 'Tabelle1'
$context.Students.Count | Should -Be 2
$context.Students[0].RowNumber | Should -Be 2
$context.Students[0].ClassName | Should -Be 'JK1-3g2_1'

{ Resolve-StudentWorkbookPath -Path (Join-Path $TestDrive 'alt.xls') } |
    Should -Throw '*.xlsx*'
```

Add a write test that passes `-SkipGitSafetyCheck`, writes `TigerWiese56`, a synthetic GUID, `mmuster@monteaufkirchen.com` as UPN and the current mail address to row 2, then reimports the workbook and checks all four values plus the existence of a backup.

- [ ] **Step 3: Run the tests**

Run `Invoke-Pester tests/Excel.Tests.ps1 -Output Detailed`. Expected: FAIL because the adapter functions do not exist.

- [ ] **Step 4: Implement path, sheet and schema validation**

`Resolve-StudentWorkbookPath` accepts only an existing leaf with extension `.xlsx`. `Read-StudentWorkbook` opens with `Open-ExcelPackage` and selects exactly one worksheet containing the five required headers. It maps rows to the shared student contract and trims text. A row is empty only when all five mandatory cells are empty. Ignore those rows. Any partially filled row or any nonempty row missing `Name mit Rufname`, `Vorname`, `Nachname`, `Klassen` or `Klassenlehrer` is a preflight error with worksheet and row number. Zero or multiple matching worksheets throw.

- [ ] **Step 5: Implement Git and lock protection**

`Assert-WorkbookSafeForPasswordWrite` opens the file with `FileShare.None` to detect an Excel lock. When Git is available and the file is within a worktree, require `git check-ignore -q -- <relative path>` to return zero and require `git ls-files --error-unmatch -- <relative path>` to return nonzero. Throw before creating any backup when either check fails.

- [ ] **Step 6: Implement backup, temporary write and read-back**

`Write-StudentWorkbookUpdates`:

1. Run safety checks.
2. Create `<base>.backup-YYYYMMDD-HHmmss.xlsx`.
3. Copy the source to `.<base>.<guid>.tmp.xlsx`.
4. Add missing `Passwort`, `EntraObjectId`, `UPN` and `Mail` headers at the right.
5. Write only supplied row updates. Preserve existing passwords when `Password` is empty.
6. Close and reopen the temporary workbook.
7. Verify object ID, UPN, mail and any new password for every row.
8. Replace the source with the verified temporary file.
9. Delete only the temporary file after a failure and retain the backup.
10. Return `Path` and `BackupPath`.

- [ ] **Step 7: Run and commit**

Run `Invoke-Pester tests/Excel.Tests.ps1 -Output Detailed`. Expected: PASS.

```bash
git add src/SchuelerSync/Private/Excel.ps1 tests/Excel.Tests.ps1 tests/fixtures/Schueler-Testdaten.xlsx
git commit -m "feat: add protected Excel student source"
```

---

### Task 4: Microsoft-Graph-Inventar und sichere Auflösung

**Files:**
- Create: `src/SchuelerSync/Private/Entra.ps1`
- Create: `tests/Entra.Tests.ps1`

**Interfaces:**
- Consumes: Microsoft Graph PowerShell SDK, Konfiguration und die UPN-Domain.
- Produces: `Connect-SchuelerGraph`, `Get-EntraSnapshot`, `Get-UserDirectGroups`, `Get-UserTransitiveGroups`, `Get-UserManagerId` und `Resolve-StudentManager`.

- [ ] **Step 1: Write failing Graph snapshot tests**

Create `tests/Entra.Tests.ps1`. Import the module and mock `Connect-MgGraph`, `Get-MgContext`, `Get-MgUser`, `Get-MgGroup`, `Get-MgGroupMember`, `Get-MgUserMemberOfAsGroup`, `Get-MgUserTransitiveMemberOfAsGroup` and `Get-MgUserManager`. Assert:

```powershell
InModuleScope SchuelerSync {
    It 'requests the complete user property set once' {
        Get-EntraSnapshot -Config $config
        Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter {
            $All -and $Property -contains 'mail' -and
            $Property -contains 'proxyAddresses' -and
            $Property -contains 'accountEnabled' -and
            $Property -contains 'consentProvidedForMinor'
        }
    }

    It 'reserves UPN and mail values case-insensitively' {
        $snapshot = Get-EntraSnapshot -Config $config
        $snapshot.ReservedAddresses.Contains(
            'mmuster@monteaufkirchen.com'
        ) | Should -BeTrue
    }
}
```

Add tests where `Resolve-StudentManager` finds exactly one teacher by UPN, mail or display name. Zero or multiple matches must throw. Use only invented users and GUIDs. Include one assigned and one dynamic group in the mocks.

- [ ] **Step 2: Run the tests**

Run `Invoke-Pester tests/Entra.Tests.ps1 -Output Detailed`. Expected: FAIL because the Graph adapter does not exist.

- [ ] **Step 3: Implement connection validation**

`Connect-SchuelerGraph` requires:

```powershell
$requiredScopes = @(
    'User.ReadWrite.All'
    'User-Mail.ReadWrite.All'
    'Group.Read.All'
    'GroupMember.ReadWrite.All'
    'User.RevokeSessions.All'
)
```

Reuse an existing context only when all scopes are present. Otherwise call `Connect-MgGraph -Scopes $requiredScopes -NoWelcome`. Re-read `Get-MgContext.Scopes` case-insensitively and throw with the missing scope names if any are absent.

- [ ] **Step 4: Implement one bounded tenant snapshot**

`Get-EntraSnapshot` calls `Get-MgUser -All` exactly once with:

```text
id,displayName,givenName,surname,userPrincipalName,mail,proxyAddresses,department,
officeLocation,companyName,employeeType,usageLocation,ageGroup,
consentProvidedForMinor,legalAgeGroupClassification,accountEnabled,userType
```

It then:

1. Reads all groups once with `Get-MgGroup -All -Property id,displayName,groupTypes,membershipRule,membershipRuleProcessingState`.
2. Validates the configured role group ID and unique names for the license and required class groups.
3. Reads the student-role membership with `Get-MgGroupMember -GroupId <id> -All`.
4. Builds case-insensitive dictionaries for ID and UPN plus `ReservedAddresses` and `AddressOwners` from every nonempty tenant UPN, mail value and SMTP proxy address. Strip the case-insensitive `smtp:` prefix before normalization.
5. Marks a group dynamic when `GroupTypes` contains `DynamicMembership` or `MembershipRule` is nonempty.
6. Retrieves direct groups, transitive groups and manager lazily per matched user with `Get-UserDirectGroups`, `Get-UserTransitiveGroups` and `Get-UserManagerId`. Cache results in the snapshot. Mark transitive groups absent from the direct set as inherited and therefore non-removable.

- [ ] **Step 5: Implement deterministic teacher resolution**

`Resolve-StudentManager` trims the Excel value and searches tenant users in this order:

1. Exact UPN or mail when the cell contains `@`.
2. Exact, case-insensitive display name.

Return one user only. Zero or multiple results are hard validation errors for that student.

- [ ] **Step 6: Run and commit**

Run `Invoke-Pester tests/Entra.Tests.ps1 -Output Detailed`. Expected: PASS.

```bash
git add src/SchuelerSync/Private/Entra.ps1 tests/Entra.Tests.ps1
git commit -m "feat: add Microsoft Graph inventory"
```

---

### Task 5: Stabile Zuordnung und Vergleichsmodell

**Files:**
- Create: `src/SchuelerSync/Private/Comparison.ps1`
- Create: `tests/Comparison.Tests.ps1`

**Interfaces:**
- Consumes: normalisierte Excelzeilen, Entra-Snapshot, Manager- und Gruppenmetadaten.
- Produces: `Resolve-StudentIdentity`, `New-StudentDesiredState`, `New-StateDifference`, `Compare-StudentDirectory` und `Add-ExchangeComparison`.

- [ ] **Step 1: Write the failing identity tests**

Create `tests/Comparison.Tests.ps1` with table-driven cases for the mandatory precedence:

```powershell
It 'matches by object ID before stored UPN and name' {
    $match = Resolve-StudentIdentity -Student $student -Snapshot $snapshot
    $match.Method | Should -Be 'EntraObjectId'
    $match.User.Id | Should -Be 'object-id-wins'
}

It 'uses stored UPN when no object ID exists' {
    $student.EntraObjectId = $null
    $match = Resolve-StudentIdentity -Student $student -Snapshot $snapshot
    $match.Method | Should -Be 'StoredUpn'
}

It 'uses a unique normalized name only inside the student role group' {
    $match = Resolve-StudentIdentity -Student $nameOnlyStudent -Snapshot $snapshot
    $match.Method | Should -Be 'UniqueStudentName'
}
```

Add failing cases for stale object IDs, duplicate Excel identities, duplicate normalized student names, one address owned by multiple directory objects, an address owned by a non-student tenant user, and two Excel rows resolving to one Entra ID. Every ambiguous case must land in `Errors` and must not appear as a departure or new student.

- [ ] **Step 2: Write the failing desired-state and category tests**

Build synthetic desired and current users. Assert exact values for:

- `DisplayName = '<Vorname> <Nachname>'` while keeping umlauts in `GivenName` and `Surname`.
- `Mail = UserPrincipalName`.
- `MailNickname` is the local part of `UserPrincipalName`. Existing students keep their current UPN and mail nickname.
- UPN candidates are calculated only for new students. Existing students keep their current UPN even when their name changes or an earlier collision disappears.
- `Department = Klassen` and `OfficeLocation` from `Get-OfficeLocation`.
- Company, employee type, usage location, age group and minor consent from config.
- `LegalAgeGroupClassification` is comparison-only because Graph computes it from age and consent. It is never sent to `Update-MgUser`.
- Manager ID is the resolved teacher.
- Required role, license and current class groups.
- Removal candidates are only direct, static, assigned memberships whose names start with `SEC-A-ROL-` or `SEC-A-CLS-` and are not desired.
- Dynamic or inherited memberships create warnings, not removal operations.

Assert `NewStudents`, `Departures`, `ChangedStudents` and `ExistingStudents` are mutually exclusive. Any attribute, manager, group or EXO difference makes a matched user changed.

Add a test for `Add-ExchangeComparison`. A mailbox difference moves a matched user from `ExistingStudents` to `ChangedStudents` and appends field-level `Area = 'Exchange'` differences. A missing mailbox appends an `EXO-Konfiguration ausstehend` warning without inventing current mailbox values.

- [ ] **Step 3: Run the tests**

Run `Invoke-Pester tests/Comparison.Tests.ps1 -Output Detailed`. Expected: FAIL because the comparison functions do not exist.

- [ ] **Step 4: Implement stable identity matching**

`Resolve-StudentIdentity` applies exactly:

1. Nonempty `EntraObjectId` must match an existing tenant user.
2. Otherwise nonempty `StoredUpn` must match a unique tenant user.
3. Otherwise match a unique normalized `GivenName + Surname` among current members of `SEC-A-ROL-Schule_Schüler`.

Return `Method`, `User` and `Warnings`. Never fall through from a present but invalid object ID or UPN to a weaker match. `Compare-StudentDirectory` detects collisions before category construction and suppresses destructive actions whenever ownership is ambiguous.

The only opt-in exception is the manual `-Add -EntraObjectId` recovery path. It may resolve the exact partially created object before role membership exists, but orchestration must verify normalized given name and surname, `CompanyName` and `EmployeeType` before any mutation. Workbook matching, stored UPN matching and name matching never receive this exception.

- [ ] **Step 5: Implement desired state and explicit differences**

`New-StudentDesiredState` creates:

```powershell
[pscustomobject]@{
    DisplayName = "$($student.GivenName) $($student.Surname)"
    GivenName = $student.GivenName
    Surname = $student.Surname
    UserPrincipalName = $selectedUpn
    Mail = $selectedUpn
    MailNickname = ($selectedUpn -split '@')[0]
    Department = $student.ClassName
    OfficeLocation = Get-OfficeLocation $student.ClassName
    CompanyName = $config.CompanyName
    EmployeeType = $config.EmployeeType
    UsageLocation = $config.UsageLocation
    AgeGroup = $config.AgeGroup
    ConsentProvidedForMinor = $config.ConsentProvidedForMinor
    LegalAgeGroupClassification = $config.LegalAgeGroupClassification
    ManagerId = $manager.Id
    RequiredGroupNames = @(
        $config.StudentRoleGroup.Name
        $config.LicenseGroupName
        "$($config.ClassGroupPrefix)$($student.ClassName)"
    )
}
```

`New-StateDifference` emits one object per field and never concatenates differences into an opaque string. Compare text ordinally after trimming, UPN/mail and group names case-insensitively, and IDs exactly. `Select-AvailableUpn` may ignore an occupied address only when every recorded owner ID equals the already matched Entra object ID.

- [ ] **Step 6: Implement safe categories and warnings**

`Compare-StudentDirectory`:

1. Validates duplicate Excel object IDs, stored UPNs and normalized name keys.
2. Resolves all Excel rows before calculating departures.
3. Merges Graph UPNs, mail values and proxy addresses with the Exchange recipient-address map before selecting any new UPN.
4. Keeps every matched user's current UPN, including manually assigned values and name changes.
5. Calculates UPNs only for new students and adds selected new UPNs to the shared reservation maps in workbook order.
6. Creates warnings for UPN fallback levels, missing `mail` values, dynamic managed groups and comparison-only legal-age discrepancies.
7. Adds manager lookup failures, missing or ambiguous groups, dynamic mandatory target groups and identity conflicts to `Errors`.
8. Calculates departures only from student-role members not claimed by an unambiguous Excel row.

If `Errors` is nonempty, comparison output is still shown, but every update mode stops before the first write.

`Add-ExchangeComparison` merges the attempt-0 mailbox result after `Compare-StudentDirectory`. It preserves category exclusivity and never adds an Exchange difference when the mailbox does not yet exist.

- [ ] **Step 7: Run and commit**

Run `Invoke-Pester tests/Comparison.Tests.ps1 -Output Detailed`. Expected: PASS.

```bash
git add src/SchuelerSync/Private/Comparison.ps1 tests/Comparison.Tests.ps1
git commit -m "feat: compare workbook and Entra students"
```

---

### Task 6: Verifizierte Entra-Mutationen

**Files:**
- Modify: `src/SchuelerSync/Private/Entra.ps1`
- Create: `tests/EntraMutations.Tests.ps1`

**Interfaces:**
- Consumes: Vergleichselemente, generierte Kennwörter und autorisierte Aktionsauswahl.
- Produces: `New-DisabledEntraStudent`, `Set-EntraStudentAttributes`, `Set-EntraStudentManager`, `Sync-EntraStudentGroups`, `Enable-EntraStudent`, `Disable-EntraStudent`, `Revoke-EntraStudentSessions` und `New-StudentWithPasswordRetry`.

- [ ] **Step 1: Write failing creation and password tests**

Mock Graph writes and assert:

```powershell
It 'creates a new student disabled without forcing a password change' {
    New-DisabledEntraStudent -Desired $desired -Password $securePassword

    Should -Invoke New-MgUser -Times 1 -Exactly -ParameterFilter {
        -not $BodyParameter.AccountEnabled -and
        $BodyParameter.PasswordProfile.ForceChangePasswordNextSignIn -eq $false -and
        $BodyParameter.PasswordProfile.Password -eq $securePassword
    }
}
```

Add a test where `New-MgUser` returns a password-policy rejection five times. `New-StudentWithPasswordRetry` must generate a fresh password after each rejection and may make at most six total create attempts, one initial password plus five replacements. A conflict, permission error or any non-password-policy failure is rethrown immediately.

- [ ] **Step 2: Write failing lifecycle and group tests**

Assert the exact order with an event list:

```text
create-disabled -> verify-attributes -> set-manager -> verify-manager
-> add-role -> add-license -> add-class -> verify-target-groups
-> remove-other-role -> remove-other-class -> verify-groups
-> write-excel -> verify-excel -> enable -> verify-enabled
```

Also assert:

- The desired role/class membership is verified before any competing direct static membership is removed.
- `Set-EntraStudentAttributes` never includes `legalAgeGroupClassification` in its Graph body.
- Manager updates use `Set-MgUserManagerByRef` with `@odata.id = https://graph.microsoft.com/v1.0/users/<id>`.
- Departures call `Update-MgUser -AccountEnabled:$false`.
- Session revocation calls `Revoke-MgUserSignInSession` only when selected.

- [ ] **Step 3: Run the tests**

Run `Invoke-Pester tests/EntraMutations.Tests.ps1 -Output Detailed`. Expected: FAIL because mutation functions are absent.

- [ ] **Step 4: Implement attribute and manager writes**

`New-DisabledEntraStudent` calls `New-MgUser -BodyParameter` with:

```powershell
$body = @{
    AccountEnabled = $false
    DisplayName = $desired.DisplayName
    GivenName = $desired.GivenName
    Surname = $desired.Surname
    UserPrincipalName = $desired.UserPrincipalName
    MailNickname = ($desired.UserPrincipalName -split '@')[0]
    Mail = $desired.Mail
    Department = $desired.Department
    OfficeLocation = $desired.OfficeLocation
    CompanyName = $desired.CompanyName
    EmployeeType = $desired.EmployeeType
    UsageLocation = $desired.UsageLocation
    AgeGroup = $desired.AgeGroup
    ConsentProvidedForMinor = $desired.ConsentProvidedForMinor
    PasswordProfile = @{
        Password = $password
        ForceChangePasswordNextSignIn = $false
    }
}
```

`Set-EntraStudentAttributes` sends only fields with `Area = 'Entra'` and `Action = 'Set'`. Explicitly exclude the computed legal-age field. `Set-EntraStudentManager` changes the manager only when the desired manager differs.

- [ ] **Step 5: Implement add-before-remove group reconciliation**

Use `New-MgGroupMemberByRef -GroupId <id> -BodyParameter @{ '@odata.id' = 'https://graph.microsoft.com/v1.0/directoryObjects/<user-id>' }` and `Remove-MgGroupMemberDirectoryObjectByRef`.

`Sync-EntraStudentGroups` performs:

1. Add missing student role, license and current class groups.
2. Re-read direct groups.
3. Verify all mandatory target IDs.
4. Remove only competing direct static `SEC-A-ROL-*` and `SEC-A-CLS-*` groups.
5. Re-read and verify exact managed-group state.

If any required add or verification fails, do not remove old groups. Dynamic and inherited memberships are never passed to the remove cmdlet.

- [ ] **Step 6: Implement activation, departure and read-back**

`Enable-EntraStudent` is callable only after the orchestrator confirms Excel ID/UPN/password write-back and successful mandatory group, manager and attribute verification. Re-read the user after enabling. `Disable-EntraStudent` re-reads `accountEnabled` after the write. `Revoke-EntraStudentSessions` returns the Graph result for reporting.

- [ ] **Step 7: Run and commit**

Run `Invoke-Pester tests/EntraMutations.Tests.ps1 -Output Detailed`. Expected: PASS.

```bash
git add src/SchuelerSync/Private/Entra.ps1 tests/EntraMutations.Tests.ps1
git commit -m "feat: add verified Entra lifecycle updates"
```

---

### Task 7: Exchange-Online-Abgleich und gebündelte Wiederholung

**Files:**
- Create: `src/SchuelerSync/Private/ExchangeOnline.ps1`
- Create: `tests/ExchangeOnline.Tests.ps1`

**Interfaces:**
- Consumes: ExchangeOnlineManagement, aktive Schüler-UPNs und die Exchange-Konfiguration.
- Produces: `Connect-SchuelerExchangeOnline`, `Get-ExchangeRecipientAddresses`, `Get-StudentMailboxState`, `Compare-StudentMailboxState`, `Set-StudentMailboxConfiguration` und `Wait-StudentMailboxes`.

- [ ] **Step 1: Write failing mailbox-state tests**

Mock `Get-ConnectionInformation`, `Connect-ExchangeOnline`, `Get-Recipient`, `Get-Mailbox`, `Get-CASMailbox`, `Set-Mailbox` and `Set-CASMailbox`. Assert `Get-ExchangeRecipientAddresses` returns every primary and proxy address with `ExternalDirectoryObjectId` as owner when present. Assert `Compare-StudentMailboxState` reports individual fields for:

- Address Book Policy, `CustomAttribute1`, audit status and 365-day audit retention.
- Deleted-item retention of 30 days, role assignment policy, sharing policy and retention policy.
- Delegate and owner audit action sets.
- Admin audit actions with `MailItemsAccessed` only when `PersistedCapabilities` contains `CommunicationsCompliance`.
- CAS values for ActiveSync, IMAP, MAPI, OWA, OWA for devices, OWA policy, POP and SMTP client authentication.

- [ ] **Step 2: Write failing retry and idempotency tests**

Use three absent mailboxes that become available in different rounds. Assert:

```powershell
$result = Wait-StudentMailboxes `
    -UserPrincipalName $upns `
    -MaxRetries 5 `
    -RetryDelaySeconds 60 `
    -SleepAction { param($seconds) $script:Sleeps += $seconds }

$script:Sleeps | Should -Be @(60, 60)
```

The test proves one immediate attempt plus at most five retry attempts. Sleep occurs once per pending batch between attempts, never once per user. Add cases for all-ready immediately, still-missing after six total attempts, and `-WhatIf` where no sleep or write cmdlet runs.

Add an idempotency test. A compliant mailbox must not call `Set-Mailbox` or `Set-CASMailbox`. An update error is reported immediately and does not re-enter the mailbox-availability retry loop.

- [ ] **Step 3: Run the tests**

Run `Invoke-Pester tests/ExchangeOnline.Tests.ps1 -Output Detailed`. Expected: FAIL because the Exchange adapter is absent.

- [ ] **Step 4: Implement connection and mailbox snapshots**

`Connect-SchuelerExchangeOnline` reuses a connected `ExchangeOnline` session returned by `Get-ConnectionInformation`. Otherwise it runs `Connect-ExchangeOnline -ShowBanner:$false` and verifies an active connection.

`Get-ExchangeRecipientAddresses` calls `Get-Recipient -ResultSize Unlimited` once, extracts `PrimarySmtpAddress` and all SMTP `EmailAddresses`, normalizes them case-insensitively and returns address-to-owner entries. Merge these entries into the Graph `ReservedAddresses` and `AddressOwners` structures before UPN selection.

`Get-StudentMailboxState` calls `Get-Mailbox -Identity <UPN> -ErrorAction Stop` once and `Get-CASMailbox` once. Treat only the documented recipient-not-found condition as `MailboxNotReady`. Authentication, authorization, throttling and transport errors are operational errors.

- [ ] **Step 5: Implement exact Exchange desired state**

`Set-StudentMailboxConfiguration` calculates differences first and issues the minimum writes. `Set-Mailbox` covers:

```powershell
$mailboxSettings = @{
    CustomAttribute1 = 'Montessori Schule Aufkirchen - Schüler'
    AddressBookPolicy = 'MON-EXO-ABP-Schule_Schüler'
    AuditEnabled = $true
    AuditLogAgeLimit = [TimeSpan]::FromDays(365)
    RetainDeletedItemsFor = [TimeSpan]::FromDays(30)
    RoleAssignmentPolicy = 'MON-EXO-UserRoles-Default'
    SharingPolicy = 'MON-EXO-Sharing-Default'
    RetentionPolicy = 'MON-EXO-Retention-Default'
}
```

Required audit actions:

```powershell
$auditDelegate = @(
    'Create', 'FolderBind', 'HardDelete', 'Move', 'MoveToDeletedItems',
    'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
    'UpdateFolderPermissions', 'UpdateInboxRules'
)
$auditOwner = @(
    'Create', 'HardDelete', 'Move', 'MailboxLogin', 'MoveToDeletedItems',
    'SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateInboxRules',
    'UpdateCalendarDelegation'
)
$auditAdmin = @(
    'Copy', 'Create', 'FolderBind', 'HardDelete', 'Move',
    'MoveToDeletedItems', 'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
    'UpdateFolderPermissions', 'UpdateInboxRules',
    'UpdateCalendarDelegation'
)
```

Append `MailItemsAccessed` to `AuditAdmin` for mailboxes with `CommunicationsCompliance`. Add missing required actions with `@{ Add = ... }` and preserve additional existing actions.

`Set-CASMailbox` enforces:

```powershell
@{
    ActiveSyncEnabled = $false
    ImapEnabled = $false
    MAPIEnabled = $true
    OWAEnabled = $true
    OWAforDevicesEnabled = $false
    OwaMailboxPolicy = 'MON-EXO-OWA-Default'
    PopEnabled = $false
    SmtpClientAuthenticationDisabled = $true
}
```

Re-read mailbox and CAS state after changes. Report the user as failed if any requested value or required audit action is still missing. If the installed ExchangeOnlineManagement version rejects a parameter, preserve and report the exact cmdlet error. Do not silently skip, rename or substitute the setting.

- [ ] **Step 6: Implement the batch retry contract**

`Wait-StudentMailboxes` makes attempt 0 immediately. For attempts 1 through 5 it sleeps once for 60 seconds and checks only still-pending UPNs. It returns `Ready` and `Missing` arrays plus an `AttemptsByUpn` map. It never retries `Set-Mailbox` or `Set-CASMailbox` failures.

Under `-WhatIf`, connect and perform attempt 0 so the plan reflects current mailbox availability. Return a planned Exchange action immediately afterward. Do not wait or call any Exchange write cmdlet.

- [ ] **Step 7: Run and commit**

Run `Invoke-Pester tests/ExchangeOnline.Tests.ps1 -Output Detailed`. Expected: PASS.

```bash
git add src/SchuelerSync/Private/ExchangeOnline.ps1 tests/ExchangeOnline.Tests.ps1
git commit -m "feat: reconcile Exchange Online mailboxes"
```

---

### Task 8: Aktionsauswahl, Orchestrierung und Bericht

**Files:**
- Create: `src/SchuelerSync/Private/Update.ps1`
- Create: `src/SchuelerSync/Private/Reporting.ps1`
- Modify: `src/SchuelerSync/SchuelerSync.psm1`
- Modify: `Sync-SchuelerEntra.ps1`
- Create: `tests/UpdateSelection.Tests.ps1`
- Create: `tests/Orchestration.Tests.ps1`

**Interfaces:**
- Consumes: Excel-, Graph-, Vergleichs- und Exchange-Adapter sowie öffentliche Parameter.
- Produces: `Resolve-UpdateSelection`, `Invoke-StudentCreateBatch`, `Invoke-StudentUpdates`, `Invoke-StudentDepartures`, `Write-StudentComparisonReport` und die vollständige `Invoke-SchuelerSync`-Orchestrierung.

- [ ] **Step 1: Write failing parameter and selection tests**

In `tests/UpdateSelection.Tests.ps1` invoke the public script with mocks and assert:

```powershell
It 'selects every action for plain Update' {
    $selection = Resolve-UpdateSelection -Update
    $selection.CreateNewUsers | Should -BeTrue
    $selection.UpdateUsers | Should -BeTrue
    $selection.DisableUsers | Should -BeTrue
    $selection.RevokeSessions | Should -BeTrue
    $selection.ConfigureAllActiveExchange | Should -BeTrue
}

It 'limits a selective run to the named actions' {
    $selection = Resolve-UpdateSelection -Update -CreateNewUsers
    $selection.CreateNewUsers | Should -BeTrue
    $selection.UpdateUsers | Should -BeFalse
    $selection.DisableUsers | Should -BeFalse
    $selection.RevokeSessions | Should -BeFalse
    $selection.ConfigureAllActiveExchange | Should -BeFalse
}
```

Add tests that action switches without `-Update` throw, `-File` defaults to `<repo>/Schueler.xlsx`, relative paths resolve from the current PowerShell directory, and `-ConfigureExchangeOnlineOnly` rejects `-File`, `-Update` and every Graph action switch. Verify `-Mail` is a mandatory string array with aliases `UPN` and `UserPrincipalName`.

- [ ] **Step 2: Write failing orchestration tests**

Use only mocks for adapters. Assert:

- Default compare reads Excel, Graph and each matched mailbox once, prints all categories and never calls any write function.
- Default compare does not sleep when a mailbox is absent.
- `-Update` stops before the first mutation when preflight `Errors` is nonempty.
- Plain `-Update` executes create, update, disable, revoke and Exchange for every active workbook student.
- Selective `-CreateNewUsers` configures Exchange only for successfully created users.
- Selective `-UpdateUsers` configures Exchange only for successfully updated users.
- `-DisableUsers` and `-RevokeSessions` alone never configure active-user mailboxes.
- Exchange-only mode never opens Excel and never connects to Graph.
- A failure for one user is recorded and independent users continue.
- Password values never reach report objects, `Write-Host`, `Write-Output`, `Write-Warning`, `Write-Verbose` or `Write-Error` mocks.

- [ ] **Step 3: Write failing ShouldProcess and WhatIf tests**

Call the public entry point with `-Update -WhatIf` and with `-ConfigureExchangeOnlineOnly -Mail test@monteaufkirchen.com -WhatIf`. Assert:

- No password generator, Excel backup/write, Graph write, Exchange write or sleep is called.
- Graph and Exchange read functions still run.
- Exchange performs only its immediate availability read.
- Planned actions appear in the report with status `WhatIf`.

Every state-changing helper is an advanced function with `SupportsShouldProcess` and calls its own `$PSCmdlet.ShouldProcess(target, action)` immediately before each external write. The caller's `WhatIfPreference` and `ConfirmPreference` flow into nested calls.

- [ ] **Step 4: Run the tests**

Run:

```powershell
Invoke-Pester tests/UpdateSelection.Tests.ps1, tests/Orchestration.Tests.ps1 -Output Detailed
```

Expected: FAIL because orchestration and reporting functions do not exist.

- [ ] **Step 5: Implement the public parameter sets**

`Sync-SchuelerEntra.ps1` is an advanced script:

```powershell
[CmdletBinding(
    DefaultParameterSetName = 'Sync',
    SupportsShouldProcess,
    ConfirmImpact = 'High'
)]
param(
    [Parameter(ParameterSetName = 'Sync')]
    [ValidateScript({ [IO.Path]::GetExtension($_) -eq '.xlsx' })]
    [string] $File = (Join-Path $PSScriptRoot 'Schueler.xlsx'),

    [Parameter(ParameterSetName = 'Sync')][switch] $Update,
    [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,

    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [switch] $ConfigureExchangeOnlineOnly,

    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [Alias('UPN', 'UserPrincipalName')]
    [ValidateNotNullOrEmpty()]
    [string[]] $Mail
)
```

It imports the manifest, injects the default `File` value into a copy of `$PSBoundParameters` when needed and forwards that hashtable, including common `WhatIf`, `Confirm` and `Verbose` parameters. Runtime validation rejects action selectors when `Update` is false.

- [ ] **Step 6: Implement update selection and preflight**

`Resolve-UpdateSelection` returns five booleans. No selectors means all four Graph actions plus all-active EXO. One or more selectors means exactly those Graph actions. In selective mode, EXO targets are only users successfully processed by `CreateNewUsers` or `UpdateUsers`.

Before any authorized write:

1. Resolve and validate the workbook.
2. Connect and verify Graph tenant context. If `ExpectedTenantId` is configured, require an exact match. Otherwise prominently display tenant ID and account before updates.
3. Read Graph users, groups, student-role members, proxy addresses, managers and direct groups required for comparison.
4. Connect to Exchange and collect all recipient addresses for collision detection.
5. Resolve every teacher and required target group.
6. Build the directory comparison, perform one mailbox availability/state read for every matched active student and merge it with `Add-ExchangeComparison`.
7. Print the complete comparison.
8. Stop all writes when any preflight error exists.

- [ ] **Step 7: Implement fail-safe new-user batching**

`Invoke-StudentCreateBatch` processes each new student independently until Excel persistence:

1. Under `ShouldProcess`, generate an unused twelve-character password in memory.
2. Create the Entra user disabled. On a password-policy rejection, generate another password up to five times.
3. Re-read and verify attributes.
4. Set and verify manager.
5. Add and verify mandatory role, license and class groups. Remove competing managed groups only after target verification.
6. Store row number, password, object ID and actual UPN in an in-memory pending-write object that is never passed to reporting.

After all candidates have reached step 6, call `Write-StudentWorkbookUpdates` once for the successful candidates. Re-read the workbook. Enable only candidates whose password, object ID and UPN are verified in Excel. If the workbook write fails, every newly created account remains disabled.

- [ ] **Step 8: Implement existing-user and departure updates**

`Invoke-StudentUpdates` applies only listed differences, preserves the Excel password field, and verifies UPN, mail, manager and managed groups after every relevant write. `Invoke-StudentDepartures` disables and revokes independently according to selection. It never deletes users or removes their historical groups.

Catch errors per user after preflight. Record `UserId`, `UserPrincipalName`, failed phase, exact exception text and a safe rerun command. Do not include passwords in success or error objects.

- [ ] **Step 9: Implement Exchange orchestration**

For default comparison, perform attempt 0 only and add mailbox differences to each matched student. For full update, target every active Excel student. For selective runs, target only successful creates and updates. For Exchange-only mode, normalize and deduplicate `Mail`, perform the immediate lookup plus five retry rounds, and apply the complete mailbox configuration without reading Excel or Graph.

- [ ] **Step 10: Implement five tables and final status**

`Write-StudentComparisonReport` emits:

1. `Neuzugänge` with Excel row, `NameMitRufname`, display name, class and proposed UPN.
2. `Abgänge` with Entra ID, display name, UPN and enabled status.
3. `Änderungen` with `NameMitRufname` and one row per `Area` and `Field` difference, current value, desired value and action.
4. `Bestehende` with `NameMitRufname`, ID, display name, UPN and class.
5. `Warnungen und Fehler` with code, student, message and safe recovery command.

Always print counts and an action result table. For a missing mailbox, produce:

```powershell
.\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail '<UPN>'
```

Return a structured result object for tests and automation. Set a nonzero process exit code only in the script wrapper when preflight errors or failed actions exist.

- [ ] **Step 11: Run and commit**

Run:

```powershell
Invoke-Pester tests/UpdateSelection.Tests.ps1, tests/Orchestration.Tests.ps1 -Output Detailed
```

Expected: PASS.

```bash
git add Sync-SchuelerEntra.ps1 src/SchuelerSync tests/UpdateSelection.Tests.ps1 tests/Orchestration.Tests.ps1
git commit -m "feat: orchestrate safe student synchronization"
```

---

### Task 9: Zustandsbehafteter Integrationstest und statische Analyse

**Files:**
- Create: `tests/Integration.Tests.ps1`
- Create: `PSScriptAnalyzerSettings.psd1`

**Interfaces:**
- Consumes: gesamtes Modul mit zustandsbehafteten Excel-, Graph- und Exchange-Mocks.
- Produces: Regressionstest für den vollständigen Ablauf und Analyzer-Regeln.

- [ ] **Step 1: Write a failing stateful integration test**

Create in-memory stores for users, groups, managers, mailbox values and CAS values. Copy the synthetic workbook to `TestDrive`. The initial state contains one new student, one changed student, one unchanged student and one departure.

Invoke:

```powershell
$first = Invoke-SchuelerSync -File $copy -Update -Confirm:$false
$second = Invoke-SchuelerSync -File $copy
```

Assert the first run creates one disabled user, writes exactly one new twelve-character password, applies attributes and manager, adds all mandatory groups before removing managed competitors, enables the new account, updates the changed user, disables and revokes the departure, and configures active mailboxes.

Assert the second run has no new, departure or changed entries and issues no write. The departure remains absent from the active Excel set and disabled in the mock tenant.

- [ ] **Step 2: Run the integration test**

Run `Invoke-Pester tests/Integration.Tests.ps1 -Output Detailed`. Expected: FAIL until any orchestration contract gaps are corrected.

- [ ] **Step 3: Fix only contract gaps exposed by the test**

Keep the production interfaces from Tasks 1 through 8. Correct ordering, state refresh or mock-boundary defects without weakening assertions. Run the integration test after each focused fix until it passes.

- [ ] **Step 4: Add strict analyzer settings**

Create `PSScriptAnalyzerSettings.psd1` enabling all default rules and specifically requiring:

```powershell
@{
    Severity = @('Error', 'Warning')
    IncludeDefaultRules = $true
    Rules = @{
        PSAvoidUsingPlainTextForPassword = @{ Enable = $true }
        PSAvoidUsingWriteHost = @{ Enable = $true }
        PSUseShouldProcessForStateChangingFunctions = @{ Enable = $true }
    }
}
```

Use an in-memory plain `string` only at the Graph SDK boundary because `New-MgUser` requires it. Add the narrowest inline analyzer suppression there with a comment explaining that the value is not logged and is persisted only to the protected workbook.

- [ ] **Step 5: Run all automated checks**

```powershell
Invoke-Pester tests -Output Detailed
Invoke-ScriptAnalyzer `
    -Path Sync-SchuelerEntra.ps1, src/SchuelerSync `
    -Recurse `
    -Settings PSScriptAnalyzerSettings.psd1
```

Expected: all Pester tests pass and the analyzer returns no errors or warnings.

- [ ] **Step 6: Commit**

```bash
git add tests/Integration.Tests.ps1 PSScriptAnalyzerSettings.psd1 src/SchuelerSync Sync-SchuelerEntra.ps1
git commit -m "test: cover end-to-end student synchronization"
```

---

### Task 10: README, Betriebsdokumentation und Hero-Bild

**Files:**
- Replace: `README.md`
- Create: `docs/ENTRA-SCHUELER-SYNC.md`
- Create: `docs/assets/entra-schueler-sync-hero.png`
- Create: `tests/Documentation.Tests.ps1`

**Interfaces:**
- Consumes: implementierte Parameter und Konfiguration.
- Produces: Schnellstart, Betriebsanleitung, VM-Testplan und rein generisches Bildmaterial.

- [ ] **Step 1: Write failing documentation contract tests**

Create `tests/Documentation.Tests.ps1`. Parse the public command and both Markdown files. Assert:

```powershell
$command = Get-Command (Join-Path $repoRoot 'Sync-SchuelerEntra.ps1')
$command.Parameters.Keys | Should -Contain 'File'
$command.Parameters.Keys | Should -Contain 'ConfigureExchangeOnlineOnly'
$command.Parameters['Mail'].Aliases | Should -Contain 'UPN'

Test-Path (Join-Path $repoRoot 'docs/assets/entra-schueler-sync-hero.png') |
    Should -BeTrue
```

Also require README examples for compare, full update, selective update, `-WhatIf` and Exchange-only repair. Require the long-form document to name all modules, Graph scopes, expected admin roles, five-plus-one mailbox checks, privacy rules and all Exchange settings.

- [ ] **Step 2: Run the documentation test**

Run `Invoke-Pester tests/Documentation.Tests.ps1 -Output Detailed`. Expected: FAIL because the final documentation and asset do not exist.

- [ ] **Step 3: Generate the generic hero asset**

Use the `imagegen` skill and generate one wide PNG with this prompt:

```text
Wide 3D comic-art hero illustration for an open-source PowerShell school
identity synchronization tool. Friendly abstract student avatars and a
teacher avatar move from a clean spreadsheet grid through glowing secure
identity nodes into a cloud directory and mailbox. Montessori-inspired warm
wood colors, Microsoft-like blue accents without logos, polished clay-rendered
characters, soft studio lighting, optimistic and professional, ample negative
space, no readable text, no real people, no personal data, 21:9 composition.
```

Save the generated result as `docs/assets/entra-schueler-sync-hero.png`. Inspect it and regenerate if it contains readable text, branding, visual artifacts or an unsuitable crop.

- [ ] **Step 4: Write the README**

Place the hero at the top with a relative Markdown link and a concise German project summary. Include prerequisites, install commands, default read-only comparison, `-File` examples, full and selective update examples, `-WhatIf`, Exchange-only repair, security warnings and a link to the full documentation.

- [ ] **Step 5: Write the operating guide**

`docs/ENTRA-SCHUELER-SYNC.md` must document:

- Supported PowerShell and module versions and installation commands.
- Excel schema, optional output columns, first matching worksheet behavior and `.xlsx` restriction.
- Git safety for root workbooks, custom repository paths and files outside a Git worktree.
- Exact UPN normalization and collision sequence, including proxy addresses and Exchange recipients.
- Matching precedence and why stored object ID is authoritative.
- All parameter sets and selector semantics.
- Graph scopes, tenant display/check and administrator-role prerequisites.
- Every Entra attribute, role/license/class group rule and manager resolution rule.
- Exact twelve-character password rules and safe storage behavior.
- Complete mailbox, audit and CAS values from the supplied scripts.
- Immediate mailbox check plus five retries, 60 seconds apart and batch behavior.
- Fail-safe new-user lifecycle, read-back checks, errors and rerun commands.
- Parallels VM test matrix with invented accounts and cleanup steps.
- Known limitation that `legalAgeGroupClassification` is read-only and validated rather than written directly.

- [ ] **Step 6: Run and commit**

Run `Invoke-Pester tests/Documentation.Tests.ps1 -Output Detailed`. Expected: PASS.

```bash
git add README.md docs/ENTRA-SCHUELER-SYNC.md docs/assets/entra-schueler-sync-hero.png tests/Documentation.Tests.ps1
git commit -m "docs: add student sync operating guide"
```

---

### Task 11: Abschlussprüfung und Parallels-VM-Übergabe

**Files:**
- Modify only when a verification failure requires a scoped fix.
- Verify: all production, test and documentation files.

**Interfaces:**
- Consumes: completed repository and the user's Windows test tenant.
- Produces: reproducible local evidence plus a live-test checklist. No production deployment.

- [ ] **Step 1: Run the complete local PowerShell suite**

```powershell
$ErrorActionPreference = 'Stop'
Invoke-Pester tests -Output Detailed
Invoke-ScriptAnalyzer `
    -Path Sync-SchuelerEntra.ps1, src/SchuelerSync `
    -Recurse `
    -Settings PSScriptAnalyzerSettings.psd1
Test-ModuleManifest src/SchuelerSync/SchuelerSync.psd1
```

Expected: zero failed tests and zero analyzer findings. If PowerShell cannot execute on the development host, record the exact error and run these commands in the Parallels VM before calling the implementation verified.

- [ ] **Step 2: Verify repository hygiene**

```bash
git check-ignore -v Schueler.xlsx
git ls-files '*.xlsx'
git diff --check
git status --short
rg -n -i 'password|passwort' . \
  -g '!docs/**' \
  -g '!tests/**' \
  -g '!*.xlsx'
```

Expected: the root workbook is ignored, only the synthetic fixture is tracked, the diff is clean, and no literal password value is present. Review every password-related code hit for logging or transcript leakage.

- [ ] **Step 3: Run a read-only VM smoke test**

Install pinned module versions, copy an invented workbook, authenticate to the test tenant and run:

```powershell
.\Sync-SchuelerEntra.ps1 -File .\tests\fixtures\Schueler-Testdaten.xlsx
.\Sync-SchuelerEntra.ps1 -File .\tests\fixtures\Schueler-Testdaten.xlsx -Update -WhatIf
```

Confirm both runs show the intended tenant, all five tables, no workbook mutation and no Graph or Exchange writes.

- [ ] **Step 4: Run the live synthetic lifecycle test**

With dedicated invented test users and groups:

1. Create a new-student row and run selective creation.
2. Verify the account was disabled until Excel persistence and then enabled.
3. Verify `Passwort` has exactly twelve characters and first sign-in does not require a password change.
4. Verify company, employee type, usage location, age group, minor consent, computed legal classification, department, office location and manager.
5. Verify exactly the desired direct managed role and class group plus the license group.
6. Verify every mailbox, audit and CAS setting.
7. Change class and teacher, run `-Update -UpdateUsers` and verify add-before-remove group order from verbose output.
8. Remove the row, run `-Update -DisableUsers -RevokeSessions` and verify disabled state plus successful session revocation.
9. Run `-ConfigureExchangeOnlineOnly -Mail <synthetic-UPN>` and verify idempotency.

- [ ] **Step 5: Record evidence and final limitations**

Capture commands, module versions, tenant ID, timestamps and sanitized pass/fail results without passwords or real student data. State separately:

- Automated tests run and their counts.
- Analyzer result.
- Read-only tenant comparison result.
- Live synthetic lifecycle result.
- Any unverified permission, policy or mailbox-provisioning behavior.

Do not claim the live tenant workflow is verified until the user reports the Parallels VM results.

- [ ] **Step 6: Commit verification-only corrections**

If verification required changes, rerun the complete relevant suite and commit only those fixes:

```bash
git add <exact-files-fixed>
git commit -m "fix: address student sync verification findings"
```

If verification required no changes, create no empty commit.

---

## Execution Notes

- Work through tasks in order. Do not combine commits across task boundaries.
- Keep `Schueler.xlsx` untracked and do not inspect or print real row data during implementation.
- Use only `tests/fixtures/Schueler-Testdaten.xlsx` for automated tests.
- Never run `-Update` against the production tenant from the development machine.
- Stop at the Parallels VM handoff. The user owns live test execution and cleanup.
