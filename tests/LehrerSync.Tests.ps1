BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Personal profiles and license protection' {
    InModuleScope SchuelerSync {
        BeforeAll {
            $script:LehrerConfig = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
        }

        It 'maps <Job> to the existing employee type and role' -ForEach @(
            @{ Job = 'L'; Type = 'Pädagogisches Team'; Role = 'School' }
            @{ Job = 'CO-L'; Type = 'Pädagogisches Team'; Role = 'School' }
            @{ Job = 'PA'; Type = 'Pädagogisches Team'; Role = 'School' }
            @{ Job = 'OGTS'; Type = 'Ganztag'; Role = 'Ganztag' }
        ) {
            $profile = Get-LehrerProfile -Job $Job -Config $script:LehrerConfig
            $profile.Job | Should -Be $Job
            $profile.EmployeeType | Should -Be $Type
            $profile.Key | Should -Be $Role
            $profile.Exchange.CustomAttribute1 | Should -Be "$($profile.CompanyName) - $Type"
        }

        It 'normalizes the job code and rejects unknown jobs' {
            (Get-LehrerProfile -Job ' co-l ' -Config $script:LehrerConfig).Job | Should -Be 'CO-L'
            { Get-LehrerProfile -Job 'L/PA' -Config $script:LehrerConfig } | Should -Throw
        }

        It 'uses one school profile for multiple school jobs and keeps every code in JobTitle' {
            $profile = Get-LehrerProfile -Job ' pa, l ' -Config $script:LehrerConfig
            $profile.Job | Should -Be 'L, PA'
            $profile.Key | Should -Be 'School'
            $profile.EmployeeType | Should -Be 'Pädagogisches Team'
            $profile.Role.Id | Should -Be $script:LehrerConfig.Roles.School.Id
        }

        It 'treats JAS as an additional code without its own profile' {
            $profile = Get-LehrerProfile -Job 'jas, pa' -Config $script:LehrerConfig
            $profile.Job | Should -Be 'PA, JAS'
            $profile.Key | Should -Be 'School'
            { Get-LehrerProfile -Job 'JAS' -Config $script:LehrerConfig } | Should -Throw '*Zusatz*'
        }

        It 'rejects jobs from both personnel profiles in one account' {
            { Get-LehrerProfile -Job 'L, OGTS' -Config $script:LehrerConfig } | Should -Throw '*unterschiedlichen Profilen*'
        }

        It 'preserves any existing direct license' {
            $user = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = @([pscustomobject]@{ SkuId = 'other' }); LicenseAssignmentStates = @() }
            $decision = Test-LehrerLicenseState -User $user -DirectGroups @()
            $decision.Add | Should -BeFalse
            $decision.Reason | Should -Match 'erhalten'
        }

        It 'preserves an active license when Graph reports Error None' {
            $user = [pscustomobject]@{
                Id = 'teacher-1'
                AssignedLicenses = @([pscustomobject]@{ SkuId = 'other' })
                LicenseAssignmentStates = @([pscustomobject]@{ SkuId = 'other'; State = 'Active'; Error = 'None' })
            }
            $decision = Test-LehrerLicenseState -User $user -DirectGroups @()
            $decision.Add | Should -BeFalse
            $decision.Reason | Should -Match 'erhalten'
        }

        It 'preserves a group license even when assignment has not propagated' {
            $user = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = @(); LicenseAssignmentStates = @() }
            $group = [pscustomobject]@{ DisplayName = 'SEC-A-LIC-Other'; AssignedLicenses = @([pscustomobject]@{ SkuId = 'other' }) }
            (Test-LehrerLicenseState -User $user -DirectGroups @($group)).Add | Should -BeFalse
        }

        It 'adds the standard group only when an existing account is provably unlicensed' {
            $user = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = @(); LicenseAssignmentStates = @() }
            (Test-LehrerLicenseState -User $user -DirectGroups @()).Add | Should -BeTrue
        }

        It 'blocks incomplete or erroneous license data' {
            $missing = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = $null; LicenseAssignmentStates = @() }
            { Test-LehrerLicenseState -User $missing -DirectGroups @() } | Should -Throw '*unvollständig*'
            $failed = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = @(); LicenseAssignmentStates = @([pscustomobject]@{ State = 'Error'; Error = 'CountViolation' }) }
            { Test-LehrerLicenseState -User $failed -DirectGroups @() } | Should -Throw '*Fehler*'
            $failedState = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = @(); LicenseAssignmentStates = @([pscustomobject]@{ State = 'ActiveWithError'; Error = 'None' }) }
            { Test-LehrerLicenseState -User $failedState -DirectGroups @() } | Should -Throw '*Fehler*'
            $failedAssignment = [pscustomobject]@{ Id = 'teacher-1'; AssignedLicenses = @(); LicenseAssignmentStates = @([pscustomobject]@{ State = 'Active'; Error = 'Other' }) }
            { Test-LehrerLicenseState -User $failedAssignment -DirectGroups @() } | Should -Throw '*Fehler*'
        }

        It 'generates twelve random characters including all required classes' {
            $password = New-LehrerPassword -UsedPasswords ([Collections.Generic.HashSet[string]]::new())
            $password.Length | Should -Be 12
            $password | Should -Match '[A-Z]'
            $password | Should -Match '[a-z]'
            $password | Should -Match '[0-9]'
            $password | Should -Match '[!#%+?@]'
        }

        It 'creates personnel disabled and requires a first sign-in password change' {
            $desired = [pscustomobject]@{
                DisplayName = 'Lea Testwald'; GivenName = 'Lea'; Surname = 'Testwald'
                UserPrincipalName = 'ltestwald@monteaufkirchen.com'; Mail = 'ltestwald@monteaufkirchen.com'
                MailNickname = 'ltestwald'; JobTitle = 'CO-L'; CompanyName = 'Montessori Schule Aufkirchen'
                EmployeeType = 'Pädagogisches Team'; UsageLocation = 'DE'; AgeGroup = 'Adult'; ConsentProvidedForMinor = 'NotRequired'
            }
            Mock New-MgUser { [pscustomobject]@{ Id = 'created-1' } }
            $null = New-DisabledEntraLehrer -Desired $desired -Password 'Abcdef123!45' -Confirm:$false
            Should -Invoke New-MgUser -Exactly 1 -ParameterFilter {
                $BodyParameter.AccountEnabled -eq $false -and $BodyParameter.JobTitle -eq 'CO-L' -and
                $BodyParameter.EmployeeType -eq 'Pädagogisches Team' -and
                $BodyParameter.PasswordProfile.Password.Length -eq 12 -and
                $BodyParameter.PasswordProfile.ForceChangePasswordNextSignIn -eq $true
            }
        }

        It 'never creates a personnel account or password under WhatIf' {
            $desired = [pscustomobject]@{
                DisplayName = 'Lea Testwald'; GivenName = 'Lea'; Surname = 'Testwald'
                UserPrincipalName = 'ltestwald@monteaufkirchen.com'; Mail = 'ltestwald@monteaufkirchen.com'
                MailNickname = 'ltestwald'; JobTitle = 'L'; CompanyName = 'Montessori Schule Aufkirchen'
                EmployeeType = 'Pädagogisches Team'; UsageLocation = 'DE'; AgeGroup = 'Adult'; ConsentProvidedForMinor = 'NotRequired'
            }
            Mock New-MgUser { throw 'Graph write under WhatIf' }
            $null = New-DisabledEntraLehrer -Desired $desired -Password 'Abcdef123!45' -WhatIf
            Should -Invoke New-MgUser -Exactly 0
        }
    }
}

Describe 'Personnel workbook and command' {
    InModuleScope SchuelerSync {
        BeforeAll { $script:LehrerFixture = Join-Path $script:SchuelerSyncRepositoryRoot 'tests/fixtures/Lehrer-Testdaten.xlsx' }

        It 'reads four required columns without class or manager' {
            $path = Join-Path $TestDrive 'Personal.xlsx'
            Copy-Item $script:LehrerFixture $path
            $data = Read-StudentWorkbook -Path $path -Kind Teacher
            $data.Students.Count | Should -Be 2
            $data.Students[0].Job | Should -Be 'L'
            $data.Students[1].Job | Should -Be 'OGTS'
        }

        It 'reads combined school jobs and the additional JAS code from Excel' {
            $path = Join-Path $TestDrive 'CombinedJobs.xlsx'
            Copy-Item $script:LehrerFixture $path
            $package = Open-StudentWorkbookPackage -Path $path
            try {
                $package.Workbook.Worksheets['Personal'].Cells[2,4].Value = 'L, PA'
                $package.Workbook.Worksheets['Personal'].Cells[3,4].Value = 'PA, JAS'
                $package.Save()
            } finally { $package.Dispose() }
            $data = Read-StudentWorkbook -Path $path -Kind Teacher
            $data.Students[0].Job | Should -Be 'L, PA'
            $data.Students[1].Job | Should -Be 'PA, JAS'
        }

        It 'writes personnel identity and password while retaining the source backup' {
            $path = Join-Path $TestDrive 'Personal.xlsx'
            Copy-Item $script:LehrerFixture $path
            $package = Open-StudentWorkbookPackage -Path $path
            try { $package.Workbook.Worksheets['Personal'].Cells[2,4].Value = 'PA, JAS'; $package.Save() }
            finally { $package.Dispose() }
            $hash = (Get-FileHash $path -Algorithm SHA256).Hash
            $update = [pscustomobject]@{ RowNumber = 2; EntraObjectId = [guid]::NewGuid().Guid
                UPN = 'ltestwald@monteaufkirchen.com'; Mail = 'ltestwald@monteaufkirchen.com'; Password = 'Abcdef123!45' }
            $written = Write-StudentWorkbookUpdate -Path $path -Kind Teacher -ExpectedSourceHash $hash -Updates @($update) -SkipGitSafetyCheck -Confirm:$false
            Test-Path $written.BackupPath | Should -BeTrue
            $data = Read-StudentWorkbook -Path $path -Kind Teacher
            $data.Students[0].Password | Should -Be 'Abcdef123!45'
            $data.Students[0].EntraObjectId | Should -Be $update.EntraObjectId
            $data.Students[0].StoredUpn | Should -Be $update.UPN
            $data.Students[0].StoredMail | Should -Be $update.Mail
            $data.Students[0].Job | Should -Be 'PA, JAS'
        }

        It 'rejects xls and unknown jobs' {
            { Resolve-StudentWorkbookPath -Path (Join-Path $TestDrive 'Alt.xls') } | Should -Throw '*.xlsx*'
            $path = Join-Path $TestDrive 'Personal.xlsx'
            Copy-Item $script:LehrerFixture $path
            $package = Open-StudentWorkbookPackage -Path $path
            try { $package.Workbook.Worksheets['Personal'].Cells[2,4].Value = 'L/PA'; $package.Save() }
            finally { $package.Dispose() }
            { Read-StudentWorkbook -Path $path -Kind Teacher } | Should -Throw '*Ungültiger Job*'
        }

        It 'rejects a duplicate required personnel column' {
            $path = Join-Path $TestDrive 'DuplicateHeader.xlsx'
            Copy-Item $script:LehrerFixture $path
            $package = Open-StudentWorkbookPackage -Path $path
            try { $package.Workbook.Worksheets['Personal'].Cells[1,8].Value = 'Job'; $package.Save() }
            finally { $package.Dispose() }
            { Read-StudentWorkbook -Path $path -Kind Teacher } | Should -Throw '*Doppelte Pflichtüberschrift*'
        }

        It 'rejects two matching personnel worksheets' {
            $path = Join-Path $TestDrive 'TwoSheets.xlsx'
            Copy-Item $script:LehrerFixture $path
            $package = Open-StudentWorkbookPackage -Path $path
            try {
                $other = $package.Workbook.Worksheets.Add('Another')
                @('Name mit Rufname','Vorname','Nachname','Job') | ForEach-Object -Begin { $column = 1 } -Process {
                    $other.Cells[1,$column].Value = $_; $column++
                }
                $package.Save()
            } finally { $package.Dispose() }
            { Read-StudentWorkbook -Path $path -Kind Teacher } | Should -Throw '*Mehrere Arbeitsblätter*'
        }

        It 'rejects a partially filled personnel row' {
            $path = Join-Path $TestDrive 'PartialRow.xlsx'
            Copy-Item $script:LehrerFixture $path
            $package = Open-StudentWorkbookPackage -Path $path
            try { $package.Workbook.Worksheets['Personal'].Cells[2,4].Value = $null; $package.Save() }
            finally { $package.Dispose() }
            { Read-StudentWorkbook -Path $path -Kind Teacher } | Should -Throw '*Unvollständige Personaldaten*'
        }
    }
}

Describe 'Personnel comparison and safe update selection' {
    InModuleScope SchuelerSync {
        BeforeAll {
            $script:LehrerConfig = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $script:LehrerFixture = Join-Path $script:SchuelerSyncRepositoryRoot 'tests/fixtures/Lehrer-Testdaten.xlsx'
            function New-LehrerTestSnapshot {
                param([object[]] $Users = @(), [string] $Role = 'Ganztag')
                $school = [pscustomobject]@{ Id = $script:LehrerConfig.Roles.School.Id; DisplayName = $script:LehrerConfig.Roles.School.Name
                    AssignedLicenses = @(); GroupTypes = @(); MembershipRule = $null }
                $ganztag = [pscustomobject]@{ Id = $script:LehrerConfig.Roles.Ganztag.Id; DisplayName = $script:LehrerConfig.Roles.Ganztag.Name
                    AssignedLicenses = @(); GroupTypes = @(); MembershipRule = $null }
                $a3 = [pscustomobject]@{ Id = 'a3-group'; DisplayName = 'SEC-A-LIC-M365A3Faculty'
                    AssignedLicenses = @([pscustomobject]@{ SkuId = 'a3' }); GroupTypes = @(); MembershipRule = $null }
                $a1 = [pscustomobject]@{ Id = 'a1-group'; DisplayName = 'SEC-A-LIC-O365A1Faculty'
                    AssignedLicenses = @([pscustomobject]@{ SkuId = 'a1' }); GroupTypes = @(); MembershipRule = $null }
                $byId = New-CaseInsensitiveHashtable
                $byUpn = New-CaseInsensitiveHashtable
                $groups = New-CaseInsensitiveHashtable
                foreach ($group in @($school,$ganztag,$a3,$a1)) { $groups[[string]$group.Id] = $group }
                $memberIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $rolesByUserId = New-CaseInsensitiveHashtable
                $direct = New-CaseInsensitiveHashtable
                foreach ($user in $Users) {
                    $byId[[string]$user.Id] = $user
                    $byUpn[[string]$user.UserPrincipalName] = $user
                    [void]$memberIds.Add([string]$user.Id)
                    $sourceRole = if ($Role -eq 'Ganztag') { $ganztag } else { $school }
                    $direct[[string]$user.Id] = @($sourceRole)
                    $rolesByUserId[[string]$user.Id] = @($sourceRole.DisplayName)
                }
                [pscustomobject]@{
                    UsersById = $byId; UsersByUpn = $byUpn; GroupsById = $groups
                    LicenseGroups = @{ 'SEC-A-LIC-M365A3Faculty' = $a3; 'SEC-A-LIC-O365A1Faculty' = $a1 }
                    PersonnelMemberIds = $memberIds; PersonnelRolesByUserId = $rolesByUserId
                    StudentMemberIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    DirectGroupsByUserId = $direct
                    ReservedAddresses = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    AddressOwners = New-CaseInsensitiveHashtable
                }
            }
            function New-LehrerTestPerson {
                param([string] $Job = 'L')
                [pscustomobject]@{ RowNumber = 2; NameMitRufname = 'Testwald, Lea'; GivenName = 'Lea'; Surname = 'Testwald'
                    Job = $Job; EntraObjectId = 'person-1'; StoredUpn = 'ltestwald@monteaufkirchen.com'; StoredMail = '' }
            }
            function New-LehrerTestUser {
                param([object[]] $Licenses = @([pscustomobject]@{ SkuId = 'a1' }))
                [pscustomobject]@{ Id = 'person-1'; DisplayName = 'Lea Testwald'; GivenName = 'Lea'; Surname = 'Testwald'
                    UserPrincipalName = 'ltestwald@monteaufkirchen.com'; Mail = 'ltestwald@monteaufkirchen.com'
                    JobTitle = 'OGTS'; CompanyName = 'Montessori Verein Landkreis Erding e.V.'; EmployeeType = 'Ganztag'
                    UsageLocation = 'DE'; AgeGroup = 'Adult'; ConsentProvidedForMinor = 'NotRequired'
                    LegalAgeGroupClassification = 'Adult'; AccountEnabled = $true; UserType = 'Member'
                    AssignedLicenses = $Licenses; LicenseAssignmentStates = @() }
            }
        }

        It 'preserves a licensed OGTS account while changing job and role to teacher' {
            $user = New-LehrerTestUser
            $user.LicenseAssignmentStates = @([pscustomobject]@{ SkuId = 'a1'; State = 'Active'; Error = 'None' })
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $comparison = Compare-LehrerDirectory -People @((New-LehrerTestPerson)) -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.Errors | Should -BeNullOrEmpty
            $comparison.Departures.Count | Should -Be 0
            $comparison.NewPeople.Count | Should -Be 0
            $comparison.ChangedPeople.Count | Should -Be 1
            $entry = $comparison.ChangedPeople[0]
            $entry.DesiredState.UserPrincipalName | Should -Be $user.UserPrincipalName
            $entry.DesiredState.AddLicenseGroup | Should -BeFalse
            $entry.DesiredState.JobTitle | Should -Be 'L'
            $entry.DesiredState.EmployeeType | Should -Be 'Pädagogisches Team'
            @($entry.Differences | Where-Object Area -eq LicenseGroup).Count | Should -Be 0
            @($entry.Differences | Where-Object { $_.Area -eq 'Group' -and $_.Action -eq 'Add' }).Count | Should -Be 1
            @($entry.Differences | Where-Object { $_.Area -eq 'Group' -and $_.Action -eq 'Remove' }).Count | Should -Be 1
        }

        It 'plans the standard license group only for an unlicensed existing account' {
            $user = New-LehrerTestUser -Licenses @()
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $entry = (Compare-LehrerDirectory -People @((New-LehrerTestPerson)) -Snapshot $snapshot -Config $script:LehrerConfig).ChangedPeople[0]
            $entry.DesiredState.AddLicenseGroup | Should -BeTrue
            @($entry.Differences | Where-Object Area -eq LicenseGroup).Count | Should -Be 1
        }

        It 'plans reactivation of a matching disabled account instead of creating another one' {
            $user = New-LehrerTestUser
            $user.AccountEnabled = $false
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $comparison = Compare-LehrerDirectory -People @((New-LehrerTestPerson -Job OGTS)) -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.Errors | Should -BeNullOrEmpty
            $comparison.NewPeople.Count | Should -Be 0
            $comparison.Departures.Count | Should -Be 0
            $comparison.ChangedPeople.Count | Should -Be 1
            @($comparison.ChangedPeople[0].Differences | Where-Object { $_.Field -eq 'AccountEnabled' -and $_.Desired -eq $true }).Count | Should -Be 1
        }

        It 'lists only active personnel accounts missing from Excel as departures' {
            $active = New-LehrerTestUser
            $disabled = New-LehrerTestUser
            $disabled.Id = 'person-2'; $disabled.UserPrincipalName = 'disabled@example.invalid'; $disabled.AccountEnabled = $false
            $snapshot = New-LehrerTestSnapshot -Users @($active,$disabled)
            $comparison = Compare-LehrerDirectory -People @() -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.Errors | Should -BeNullOrEmpty
            $comparison.Departures.Count | Should -Be 1
            $comparison.Departures[0].User.Id | Should -Be $active.Id
            $safe = ConvertTo-SafeLehrerComparison -Comparison $comparison
            $safe.Departures[0].Role | Should -Be $script:LehrerConfig.Roles.Ganztag.Name
            $safe.Departures[0].AccountEnabled | Should -BeTrue
        }

        It 'blocks departures when an account state is unknown' {
            $user = New-LehrerTestUser
            $user.AccountEnabled = $null
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $comparison = Compare-LehrerDirectory -People @() -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.Errors[0].Field | Should -Be 'Departure'
            $comparison.Departures.Count | Should -Be 0
        }

        It 'blocks new accounts before creation when the target role assigns a license' {
            $snapshot = New-LehrerTestSnapshot
            $snapshot.GroupsById[[string]$script:LehrerConfig.Roles.School.Id].AssignedLicenses = @([pscustomobject]@{ SkuId = 'other' })
            $person = New-LehrerTestPerson
            $person.EntraObjectId = ''; $person.StoredUpn = ''
            $comparison = Compare-LehrerDirectory -People @($person) -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.NewPeople.Count | Should -Be 0
            $comparison.Errors[0].Message | Should -Match 'Zielrolle.*Lizenzen'
        }

        It 'builds the personnel population as the union of both direct roles' {
            $snapshot = New-LehrerTestSnapshot
            $school = $snapshot.GroupsById[[string]$script:LehrerConfig.Roles.School.Id]
            $ganztag = $snapshot.GroupsById[[string]$script:LehrerConfig.Roles.Ganztag.Id]
            $a3 = $snapshot.LicenseGroups['SEC-A-LIC-M365A3Faculty']
            $a1 = $snapshot.LicenseGroups['SEC-A-LIC-O365A1Faculty']
            $users = @(
                [pscustomobject]@{ Id = 'school-1'; UserPrincipalName = 'school1@example.invalid'; Mail = ''; ProxyAddresses = @() }
                [pscustomobject]@{ Id = 'shared-1'; UserPrincipalName = 'shared1@example.invalid'; Mail = ''; ProxyAddresses = @() }
                [pscustomobject]@{ Id = 'ganztag-1'; UserPrincipalName = 'ganztag1@example.invalid'; Mail = ''; ProxyAddresses = @() }
            )
            Mock Get-MgGroup { @($school,$ganztag,$a3,$a1,[pscustomobject]@{Id=(Get-ConfiguredStudentRole).Id;DisplayName=(Get-ConfiguredStudentRole).Name}) }
            Mock Get-MgUser { $users }
            Mock Get-MgGroupMember {
                if ($GroupId -eq $school.Id) { @([pscustomobject]@{ Id = 'school-1' },[pscustomobject]@{ Id = 'shared-1' },[pscustomobject]@{ Id = 'service-principal-1' }) }
                else { @([pscustomobject]@{ Id = 'shared-1' },[pscustomobject]@{ Id = 'ganztag-1' }) }
            }
            $actual = Get-LehrerSnapshot -Config $script:LehrerConfig
            $actual.PersonnelMemberIds.Count | Should -Be 3
            $actual.PersonnelMemberIds.Contains('school-1') | Should -BeTrue
            $actual.PersonnelMemberIds.Contains('ganztag-1') | Should -BeTrue
            $actual.PersonnelRolesByUserId['shared-1'] | Should -Be @($school.DisplayName,$ganztag.DisplayName)
            $actual.LicenseGroups.Count | Should -Be 2
        }

        It 'fails closed when the configured student protection group is absent' {
            $snapshot = New-LehrerTestSnapshot
            Mock Get-MgGroup { @($snapshot.GroupsById.Values) }
            Mock Get-MgUser { @() }
            Mock Get-MgGroupMember { @() }
            { Get-LehrerSnapshot -Config $script:LehrerConfig } | Should -Throw '*Schülerrolle*'
        }

        It 'blocks a foreign account with the same name instead of creating a duplicate' {
            $user = New-LehrerTestUser
            $snapshot = New-LehrerTestSnapshot -Users @()
            $snapshot.UsersById[$user.Id] = $user
            $snapshot.UsersByUpn[$user.UserPrincipalName] = $user
            $person = New-LehrerTestPerson
            $person.EntraObjectId = ''; $person.StoredUpn = ''
            $comparison = Compare-LehrerDirectory -People @($person) -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.Errors.Count | Should -BeGreaterThan 0
            $comparison.NewPeople.Count | Should -Be 0
            $comparison.Departures.Count | Should -Be 0
        }

        It 'reports a second matching run as compliant without a new license group' {
            $user = New-LehrerTestUser
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $person = New-LehrerTestPerson -Job OGTS
            $comparison = Compare-LehrerDirectory -People @($person) -Snapshot $snapshot -Config $script:LehrerConfig
            $comparison.Errors | Should -BeNullOrEmpty
            $comparison.ChangedPeople.Count | Should -Be 0
            $comparison.ExistingPeople.Count | Should -Be 1
            $comparison.ExistingPeople[0].DesiredState.AddLicenseGroup | Should -BeFalse
        }

        It 'plans a full WhatIf without Graph or Exchange writes' {
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{ AddressOwners = @{} } }
            Mock Assert-LehrerExchangePolicy { }
            Mock New-MgUser { throw 'Graph write' }
            Mock Update-MgUser { throw 'Graph write' }
            Mock New-MgGroupMemberByRef { throw 'Graph write' }
            Mock Set-StudentMailboxConfiguration { throw 'Exchange write' }
            Mock Start-Sleep { throw 'Wait' }
            Mock Write-Progress { }
            $path = Join-Path $TestDrive 'Personal.xlsx'
            Copy-Item $script:LehrerFixture $path
            $before = (Get-FileHash $path -Algorithm SHA256).Hash
            $result = Invoke-LehrerSync -File $path -Update -WhatIf
            $result.HasErrors | Should -BeFalse
            $result.Mode | Should -Be 'Update'
            $result.Actions.Count | Should -BeGreaterThan 0
            (Get-FileHash $path -Algorithm SHA256).Hash | Should -Be $before
            Should -Invoke New-MgUser -Exactly 0
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
            Should -Invoke Set-StudentMailboxConfiguration -Exactly 0
            Should -Invoke Start-Sleep -Exactly 0
            Should -Invoke Write-Progress -ParameterFilter { $Id -eq 1 -and $Activity -eq 'Lehrerabgleich' -and $Status -eq 'Vergleiche Personaldaten ...' }
            Should -Invoke Write-Progress -ParameterFilter { $Id -eq 2 -and $ParentId -eq 1 -and $Activity -eq 'Lehrerdetails prüfen' -and $Status -like 'Prüfe Person 2 von 2:*' }
            Should -Invoke Write-Progress -Exactly 1 -ParameterFilter { $Id -eq 1 -and $Activity -eq 'Lehrerabgleich' -and $Completed }
            Should -Invoke Write-Progress -Exactly 1 -ParameterFilter { $Id -eq 2 -and $Activity -eq 'Lehrerdetails prüfen' -and $Completed }
        }

        It 'prints the update comparison before the first account action' {
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot
            $script:LehrerRunOrder = [Collections.Generic.List[string]]::new()
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{ AddressOwners = @{} } }
            Mock Assert-LehrerExchangePolicy { }
            Mock Write-LehrerReport {
                $script:LehrerRunOrder.Add($(if ($ActionsOnly) { 'actions' } else { 'comparison' }))
            }
            Mock Invoke-LehrerCreate {
                $script:LehrerRunOrder.Add('create')
                New-StudentActionResult -Phase Create -Status WhatIf
            }
            $path = Join-Path $TestDrive 'Personal.xlsx'
            Copy-Item $script:LehrerFixture $path
            $result = Invoke-LehrerSync -File $path -Update -WhatIf
            $result.HasErrors | Should -BeFalse
            @($script:LehrerRunOrder)[0] | Should -Be 'comparison'
            @($script:LehrerRunOrder)[-1] | Should -Be 'actions'
            @($script:LehrerRunOrder | Where-Object { $_ -eq 'create' }).Count | Should -BeGreaterThan 0
        }

        It 'does not require an unlocked workbook when no Excel value needs backfilling' {
            $user = New-LehrerTestUser
            $person = New-LehrerTestPerson -Job OGTS
            $person.StoredMail = $user.Mail
            $person | Add-Member -NotePropertyName Password -NotePropertyValue ''
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot -Users @($user)
            $desired = [pscustomobject]@{ DisplayName = $user.DisplayName; UserPrincipalName = $user.UserPrincipalName
                LicenseReason = 'Unverändert'; Role = $script:LehrerMockSnapshot.GroupsById[[string]$script:LehrerConfig.Roles.Ganztag.Id]
                Exchange = @{ AddressBookPolicy = 'test'; CustomAttribute1 = 'test' } }
            $entry = [pscustomobject]@{ Person = $person; User = $user; DesiredState = $desired; Differences = @() }
            Mock Read-StudentWorkbook { [pscustomobject]@{ Students = @($person); SourceHash = 'test-hash' } }
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{ AddressOwners = @{} } }
            Mock Compare-LehrerDirectory { [pscustomobject]@{ NewPeople = @(); Departures = @(); ChangedPeople = @(); ExistingPeople = @($entry); Warnings = @(); Errors = @() } }
            Mock Get-StudentMailboxState { [pscustomobject]@{ Exists = $true; Differences = @() } }
            Mock Assert-LehrerExchangePolicy { }
            Mock Assert-StudentWorkbookVersion { }
            Mock Assert-WorkbookSafeForPasswordWrite { throw 'Workbook is locked' }
            Mock Write-LehrerReport { }
            $result = Invoke-LehrerSync -File (Join-Path $TestDrive 'Personal.xlsx') -Update -UpdateUsers -Confirm:$false
            $result.HasErrors | Should -BeFalse
            Should -Invoke Assert-WorkbookSafeForPasswordWrite -Exactly 0
        }

        It 'requires workbook write access when new accounts would be created' {
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{ AddressOwners = @{} } }
            Mock Assert-LehrerExchangePolicy { }
            Mock Assert-WorkbookSafeForPasswordWrite { throw 'Workbook is locked' }
            Mock Invoke-LehrerCreate { throw 'Account creation must be blocked' }
            Mock Write-LehrerReport { }
            $path = Join-Path $TestDrive 'Personal.xlsx'
            Copy-Item $script:LehrerFixture $path
            $result = Invoke-LehrerSync -File $path -Update -CreateNewUsers -Confirm:$false
            $result.HasErrors | Should -BeTrue
            $result.Comparison.Errors[0].Message | Should -Be 'Workbook is locked'
            Should -Invoke Assert-WorkbookSafeForPasswordWrite -Exactly 1
            Should -Invoke Invoke-LehrerCreate -Exactly 0
        }

        It 'writes the complete report file when the terminal only shows actions' {
            $comparison = [pscustomobject]@{ NewPeople = @(); Departures = @(); ChangedPeople = @(); ExistingPeople = @()
                Warnings = @(); Errors = @() }
            $report = Join-Path $TestDrive 'Personnel-report.txt'
            Write-LehrerReport -Comparison $comparison -Actions @([pscustomobject]@{ Phase = 'Create'; Status = 'WhatIf' }) `
                -OutputFile $report -ActionsOnly
            $content = Get-Content -LiteralPath $report -Raw
            $content | Should -Match 'Neuzugänge \(0\)'
            $content | Should -Match 'Aktionsergebnisse \(1\)'
            $content | Should -Match 'Neuzugänge \(0\)\r?\n\r?\nAbgänge \(0\)'
            $content | Should -Match 'Warnungen und Fehler \(0\)\r?\n\r?\nAktionsergebnisse \(1\)'
            $content | Should -Not -Match "`e\["
        }

        It 'prints existing personnel with the student-style identity columns in one line' {
            $existing = [pscustomobject]@{
                RowNumber = 2; NameMitRufname = 'Testwald, Lea'; UserId = '11111111-2222-3333-4444-555555555555'
                DisplayName = 'Lea Testwald'; UserPrincipalName = 'ltestwald@monteaufkirchen.com'; Job = 'L, PA'
                LicenseReason = 'Bestehende Lizenz bleibt erhalten'; Role = 'SEC-A-ROL-Schule_PädagogischesTeam'
                AddressBookPolicy = 'MON-EXO-ABP-Schule_PädagogischesTeam'
                CustomAttribute1 = 'Montessori Schule Aufkirchen - Pädagogisches Team'
            }
            $comparison = [pscustomobject]@{ NewPeople = @(); Departures = @(); ChangedPeople = @()
                ExistingPeople = @($existing); Warnings = @(); Errors = @() }
            $report = Join-Path $TestDrive 'Personnel-existing.txt'

            Write-LehrerReport -Comparison $comparison -Actions @() -OutputFile $report

            $section = ((Get-Content -LiteralPath $report -Raw) -split 'Bestehende \(1\)' -split 'Warnungen und Fehler \(0\)')[1]
            $section | Should -Match 'NameMitRufname\s+UserId\s+DisplayName\s+UserPrincipalName\s+Job'
            $section | Should -Match 'Testwald, Lea\s+11111111-2222-3333-4444-555555555555\s+Lea Testwald\s+ltestwald@monteaufkirchen.com\s+L, PA'
            $section | Should -Not -Match 'RowNumber|LicenseReason|Role|AddressBookPolicy|CustomAttribute1'
            @($section.Trim() -split '\r?\n').Count | Should -Be 3
        }

        It 'plans manual Add under WhatIf without creating a password or account' {
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{ AddressOwners = @{} } }
            Mock Assert-LehrerExchangePolicy { }
            Mock New-LehrerPassword { throw 'Password generation under WhatIf' }
            Mock New-MgUser { throw 'Graph write under WhatIf' }
            Mock Write-LehrerReport { }
            $result = Invoke-LehrerSync -Add -Vorname Lea -Nachname Testwald -Job L -WhatIf
            $result.HasErrors | Should -BeFalse
            $result.Actions[0].Status | Should -Be 'WhatIf'
            Should -Invoke New-LehrerPassword -Exactly 0
            Should -Invoke New-MgUser -Exactly 0
        }

        It 'plans manual Remove under WhatIf without disabling or revoking' {
            $user = New-LehrerTestUser
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot -Users @($user) -Role School
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Disable-EntraStudent { throw 'Graph write under WhatIf' }
            Mock Revoke-EntraStudentSession { throw 'Graph write under WhatIf' }
            Mock Write-LehrerReport { }
            $result = Invoke-LehrerSync -Remove -UPN $user.UserPrincipalName -WhatIf
            $result.HasErrors | Should -BeFalse
            @($result.Actions | Where-Object Status -eq WhatIf).Count | Should -Be 2
            Should -Invoke Disable-EntraStudent -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }

        It 'plans Exchange-only under WhatIf without waiting or writing' {
            $user = New-LehrerTestUser
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot -Users @($user) -Role Ganztag
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Assert-LehrerExchangePolicy { }
            Mock Get-StudentMailboxState { [pscustomobject]@{ Exists = $true
                    Differences = @([pscustomobject]@{ Field = 'AddressBookPolicy' }) } }
            Mock Set-StudentMailboxConfiguration { [pscustomobject]@{ Status = 'Planned' } }
            Mock Start-Sleep { throw 'Wait under WhatIf' }
            Mock Write-LehrerReport { }
            $result = Invoke-LehrerSync -ConfigureExchangeOnlineOnly -UPN $user.UserPrincipalName -WhatIf
            $result.HasErrors | Should -BeFalse
            @($result.Actions | Where-Object Status -eq WhatIf).Count | Should -Be 1
            Should -Invoke Start-Sleep -Exactly 0
            Should -Invoke Set-StudentMailboxConfiguration -Exactly 1 -ParameterFilter { $WhatIf }
        }

        It 'does not report unrelated departures when manually adding one person' {
            $user = New-LehrerTestUser
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot -Users @($user)
            Mock Connect-SchuelerGraph { [pscustomobject]@{TenantId='test';Account='test@example.org'} }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{AddressOwners=@{}} }
            Mock Assert-LehrerExchangePolicy { }
            Mock Invoke-LehrerCreate { New-StudentActionResult -Phase Create -Status WhatIf }
            Mock Write-LehrerReport { }
            $result = Invoke-LehrerSync -Add -Vorname Synthetic -Nachname Newperson -Job L -WhatIf
            $result.HasErrors | Should -BeFalse
            $result.Comparison.NewPeople.Count | Should -Be 1
            $result.Comparison.Departures.Count | Should -Be 0
        }

        It 'does not backfill Excel identities for declined personnel updates' {
            $user = New-LehrerTestUser
            $person = New-LehrerTestPerson -Job OGTS
            $person.EntraObjectId = ''; $person.StoredUpn = ''; $person.StoredMail = ''
            $person | Add-Member -NotePropertyName Password -NotePropertyValue ''
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot -Users @($user)
            $desired = [pscustomobject]@{DisplayName=$user.DisplayName;UserPrincipalName=$user.UserPrincipalName;LicenseReason='Unverändert'
                Role=$script:LehrerMockSnapshot.GroupsById[[string]$script:LehrerConfig.Roles.Ganztag.Id];Exchange=@{AddressBookPolicy='test';CustomAttribute1='test'}}
            $entry = [pscustomobject]@{Person=$person;User=$user;DesiredState=$desired;Differences=@()}
            Mock Read-StudentWorkbook { [pscustomobject]@{Students=@($person);SourceHash='test'} }
            Mock Connect-SchuelerGraph { [pscustomobject]@{TenantId='test';Account='test@example.org'} }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Connect-SchuelerExchangeOnline { }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{AddressOwners=@{}} }
            Mock Compare-LehrerDirectory { [pscustomobject]@{NewPeople=@();Departures=@();ChangedPeople=@($entry);ExistingPeople=@();Warnings=@();Errors=@()} }
            Mock Get-StudentMailboxState { [pscustomobject]@{Exists=$true;Differences=@()} }
            Mock Assert-LehrerExchangePolicy { }
            Mock Assert-StudentWorkbookVersion { }
            Mock Assert-WorkbookSafeForPasswordWrite { }
            Mock Write-LehrerReport { }
            Mock Invoke-LehrerUpdate { New-StudentActionResult -UserId $user.Id -Phase Update -Status Skipped }
            Mock Write-StudentWorkbookUpdate { throw 'Declined identity must not be written' }
            $result = Invoke-LehrerSync -File 'synthetic.xlsx' -Update -UpdateUsers -Confirm:$true
            $result.HasErrors | Should -BeFalse
            $result.Actions[0].Status | Should -Be Skipped
            Should -Invoke Write-StudentWorkbookUpdate -Exactly 0
        }

        It 'adds the new role before removing the old role and keeps the license group' {
            $user = New-LehrerTestUser
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $school = $snapshot.GroupsById[[string]$script:LehrerConfig.Roles.School.Id]
            $ganztag = $snapshot.GroupsById[[string]$script:LehrerConfig.Roles.Ganztag.Id]
            $license = $snapshot.LicenseGroups['SEC-A-LIC-O365A1Faculty']
            Mock Get-MgGroup { $snapshot.GroupsById[$GroupId] }
            $script:DirectLehrerGroups = @($ganztag, $license)
            $script:LehrerMutationOrder = [Collections.Generic.List[string]]::new()
            Mock Get-FreshEntraUserDirectGroup { $script:DirectLehrerGroups }
            Mock New-MgGroupMemberByRef {
                $script:LehrerMutationOrder.Add('AddRole')
                $script:DirectLehrerGroups += $school
            } -ParameterFilter { $GroupId -eq $school.Id }
            Mock Remove-MgGroupMemberDirectoryObjectByRef {
                $script:LehrerMutationOrder.Add('RemoveRole')
                $script:DirectLehrerGroups = @($script:DirectLehrerGroups | Where-Object Id -ne $ganztag.Id)
            } -ParameterFilter { $GroupId -eq $ganztag.Id }
            $desired = [pscustomobject]@{ Role = $school }
            (Sync-LehrerGroup -UserId $user.Id -Desired $desired -Confirm:$false) | Should -BeTrue
            @($script:LehrerMutationOrder) | Should -Be @('AddRole','RemoveRole')
            @($script:DirectLehrerGroups | Where-Object Id -eq $license.Id).Count | Should -Be 1
        }

        It 'rechecks a new license before joining the standard group' {
            $user = New-LehrerTestUser
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $group = $snapshot.LicenseGroups['SEC-A-LIC-M365A3Faculty']
            Mock Get-FreshEntraUserDirectGroup { @() }
            Mock Get-MgUser { $user }
            Mock New-MgGroupMemberByRef { throw 'Existing license must be preserved' }
            (Add-LehrerLicenseGroup -UserId $user.Id -Desired ([pscustomobject]@{ LicenseGroup = $group }) -Confirm:$false) | Should -BeTrue
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }

        It 'reactivates the matching disabled account without creating or resetting a password' {
            $user = New-LehrerTestUser
            $user.AccountEnabled = $false
            $snapshot = New-LehrerTestSnapshot -Users @($user)
            $entry = (Compare-LehrerDirectory -People @((New-LehrerTestPerson -Job OGTS)) -Snapshot $snapshot -Config $script:LehrerConfig).ChangedPeople[0]
            Mock Assert-LehrerRecoveryIdentity { $user }
            Mock Set-LehrerAttribute { $true }
            Mock Sync-LehrerGroup { $true }
            Mock Enable-EntraStudent { [pscustomobject]@{ Verified = $true } }
            Mock New-LehrerPassword { throw 'Password must not be generated' }
            Mock New-MgUser { throw 'Account must not be recreated' }
            $result = Invoke-LehrerUpdate -Entry $entry -Config $script:LehrerConfig -Confirm:$false
            $result.Status | Should -Be 'Succeeded'
            Should -Invoke Enable-EntraStudent -Exactly 1
            Should -Invoke New-LehrerPassword -Exactly 0
            Should -Invoke New-MgUser -Exactly 0
        }

        It 'configures mixed Exchange profiles with one batch wait per retry' {
            $schoolProfile = (Get-LehrerProfile -Job L -Config $script:LehrerConfig).Exchange
            $ganztagProfile = (Get-LehrerProfile -Job OGTS -Config $script:LehrerConfig).Exchange
            $targets = New-CaseInsensitiveHashtable
            $targets['ltestwald@monteaufkirchen.com'] = $schoolProfile
            $targets['nbeispielstern@monteaufkirchen.com'] = $ganztagProfile
            $script:MailboxChecks = New-CaseInsensitiveHashtable
            $script:MailboxProfiles = New-CaseInsensitiveHashtable
            $script:MailboxSleeps = 0
            Mock Get-StudentMailboxState {
                $script:MailboxProfiles[$UserPrincipalName] = $Config.AddressBookPolicy
                $count = [int]$script:MailboxChecks[$UserPrincipalName] + 1
                $script:MailboxChecks[$UserPrincipalName] = $count
                [pscustomobject]@{ Exists = ($UserPrincipalName -like 'ltestwald*' -or $count -gt 1)
                    Differences = @([pscustomobject]@{ Field = 'AddressBookPolicy' }) }
            }
            Mock Set-StudentMailboxConfiguration { [pscustomobject]@{ Status = 'Compliant' } }
            Mock Write-Progress { }
            $batch = Wait-LehrerMailbox -Targets $targets -MaxRetries 2 -RetryDelaySeconds 1 -SleepAction {
                param($seconds) $script:MailboxSleeps++
            } -Confirm:$false
            $batch.Ready.Count | Should -Be 2
            $batch.Missing.Count | Should -Be 0
            $script:MailboxSleeps | Should -Be 1
            $script:MailboxProfiles['ltestwald@monteaufkirchen.com'] | Should -Be 'MON-EXO-ABP-Schule_PädagogischesTeam'
            $script:MailboxProfiles['nbeispielstern@monteaufkirchen.com'] | Should -Be 'MON-EXO-ABP-Ganztag'
            Should -Invoke Set-StudentMailboxConfiguration -Exactly 2
            Should -Invoke Write-Progress -ParameterFilter { $Id -eq 3 -and $ParentId -eq 1 -and $Activity -eq 'Personalpostfächer bereitstellen' -and $Status -like 'Warte 1 Sekunden*' }
            Should -Invoke Write-Progress -Exactly 1 -ParameterFilter { $Id -eq 3 -and $Activity -eq 'Personalpostfächer bereitstellen' -and $Completed }
        }

        It 'does not configure an existing personnel mailbox without Exchange differences' {
            $targets = New-CaseInsensitiveHashtable
            $targets['ltestwald@monteaufkirchen.com'] = (Get-LehrerProfile -Job L -Config $script:LehrerConfig).Exchange
            Mock Get-StudentMailboxState { [pscustomobject]@{ Exists = $true; Differences = @() } }
            Mock Set-StudentMailboxConfiguration { throw 'No Exchange change is required.' }
            Mock Write-Progress { }

            $batch = Wait-LehrerMailbox -Targets $targets -Confirm:$false

            $batch.Ready.Count | Should -Be 1
            $batch.Ready[0].Configuration.Status | Should -Be 'Compliant'
            Should -Invoke Set-StudentMailboxConfiguration -Exactly 0
        }

        It 'removes an account by disabling it and revoking sessions without deleting resources' {
            $user = New-LehrerTestUser
            $script:LehrerMockSnapshot = New-LehrerTestSnapshot -Users @($user) -Role School
            $script:RemoveRole = $script:LehrerMockSnapshot.GroupsById[[string]$script:LehrerConfig.Roles.School.Id]
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'tester@example.invalid' } }
            Mock Get-LehrerSnapshot { $script:LehrerMockSnapshot }
            Mock Get-FreshEntraUserDirectGroup { @($script:RemoveRole) }
            Mock Disable-EntraStudent { [pscustomobject]@{ Verified = $true } }
            Mock Revoke-EntraStudentSession { [pscustomobject]@{ Value = $true } }
            Mock Remove-MgUser { throw 'User deletion is forbidden' }
            Mock Remove-MgGroupMemberDirectoryObjectByRef { throw 'Group removal is forbidden' }
            Mock Set-MgUserLicense { throw 'License removal is forbidden' }
            $result = Invoke-LehrerSync -Remove -UPN $user.UserPrincipalName -Confirm:$false
            $result.HasErrors | Should -BeFalse
            @($result.Actions | Where-Object Status -eq Succeeded).Count | Should -Be 2
            Should -Invoke Disable-EntraStudent -Exactly 1
            Should -Invoke Revoke-EntraStudentSession -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
            Should -Invoke Set-MgUserLicense -Exactly 0
        }
    }
}

Describe 'Personnel CLI entry point' {
    BeforeAll {
        $entryPoint = Join-Path (Split-Path $PSScriptRoot -Parent) 'Sync-LehrerEntra.ps1'
        $harness = @'
function Import-Module { }
function Invoke-LehrerSync {
    [pscustomobject]@{ HasErrors = ($env:LEHRER_CLI_FAIL -eq '1'); Forwarded = ($args -join '|') }
}
$result = & '__ENTRY__' @args
if (-not $?) { exit 1 }
if ($LASTEXITCODE) { exit $LASTEXITCODE }
if ($null -ne $result) { $result.Forwarded }
'@.Replace('__ENTRY__', $entryPoint.Replace("'", "''"))
        $script:CliHarness = Join-Path $TestDrive 'LehrerCliHarness.ps1'
        Set-Content -LiteralPath $script:CliHarness -Value $harness -Encoding utf8
        $script:PwshExecutable = Join-Path $PSHOME $(if ($IsWindows) { 'pwsh.exe' } else { 'pwsh' })
    }

    It 'forwards Add parameters and WhatIf with a successful exit code' {
        $output = & $script:PwshExecutable -NoLogo -NoProfile -NonInteractive -File $script:CliHarness `
            -Add -Vorname Lea -Nachname Testwald -Job CO-L -WhatIf 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        ($output -join ' ') | Should -Match '(-Add:.*True)'
        ($output -join ' ') | Should -Match '(-Job:.*CO-L)'
        ($output -join ' ') | Should -Match '(-WhatIf:.*True)'
    }

    It 'accepts a combined job in the manual Add entry point' {
        $output = & $script:PwshExecutable -NoLogo -NoProfile -NonInteractive -File $script:CliHarness `
            -Add -Vorname Lea -Nachname Testwald -Job 'PA, JAS' -WhatIf 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        ($output -join ' ') | Should -Match 'PA, JAS'
    }

    It 'returns a nonzero exit code when the mocked sync reports errors' {
        $env:LEHRER_CLI_FAIL = '1'
        try {
            $null = & $script:PwshExecutable -NoLogo -NoProfile -NonInteractive -File $script:CliHarness `
                -Remove -UPN 'ltestwald@monteaufkirchen.com' 2>&1
            $LASTEXITCODE | Should -Be 1
        } finally { Remove-Item Env:LEHRER_CLI_FAIL -ErrorAction SilentlyContinue }
    }

    It 'rejects an unknown job before calling any integration' {
        $output = & $script:PwshExecutable -NoLogo -NoProfile -NonInteractive -File $script:CliHarness `
            -Add -Vorname Lea -Nachname Testwald -Job L/PA 2>&1
        $LASTEXITCODE | Should -Be 1
        ($output -join ' ') | Should -Match 'L/PA'
    }
}
