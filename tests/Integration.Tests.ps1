BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $global:IntegrationStubCommands = [Collections.Generic.List[string]]::new()
    if ($null -eq (Get-Command New-MgUser -ErrorAction SilentlyContinue)) {
        function global:New-MgUser {
            [CmdletBinding()]
            param([hashtable] $BodyParameter)
            throw 'Integration test stub must be mocked.'
        }
        $global:IntegrationStubCommands.Add('New-MgUser')
    }
    if ($null -eq (Get-Command Get-MgUserMemberOfAsGroup -ErrorAction SilentlyContinue)) {
        function global:Get-MgUserMemberOfAsGroup {
            [CmdletBinding()]
            param([string] $UserId, [switch] $All)
            throw 'Integration test stub must be mocked.'
        }
        $global:IntegrationStubCommands.Add('Get-MgUserMemberOfAsGroup')
    }
    if ($null -eq (Get-Command New-MgGroupMemberByRef -ErrorAction SilentlyContinue)) {
        function global:New-MgGroupMemberByRef {
            [CmdletBinding()]
            param([string] $GroupId, [hashtable] $BodyParameter)
            throw 'Integration test stub must be mocked.'
        }
        $global:IntegrationStubCommands.Add('New-MgGroupMemberByRef')
    }
    if ($null -eq (Get-Command Remove-MgGroupMemberByRef -ErrorAction SilentlyContinue)) {
        function global:Remove-MgGroupMemberByRef {
            [CmdletBinding()]
            param([string] $GroupId, [string] $DirectoryObjectId)
            throw 'Integration test stub must be mocked.'
        }
        $global:IntegrationStubCommands.Add('Remove-MgGroupMemberByRef')
    }
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop

    $global:IntegrationGroupIds = @{
        StudentRole = 'cebc1326-1174-4126-ba84-7a8960850e0a'
        OtherRole = '103a1c4c-036b-40bc-bbe8-e3f7a866cb01'
        License = 'group-license'
        ClassG1 = 'group-class-g1'
        ClassG2 = 'group-class-g2'
        ClassM2 = 'group-class-m2'
    }

    function global:New-IntegrationStudent {
        param(
            [Parameter(Mandatory)][int] $RowNumber,
            [Parameter(Mandatory)][string] $GivenName,
            [Parameter(Mandatory)][string] $Surname,
            [Parameter(Mandatory)][string] $ClassName,
            [Parameter(Mandatory)][string] $Teacher,
            [string] $EntraObjectId = '',
            [string] $StoredUpn = ''
        )
        [pscustomobject]@{
            RowNumber = $RowNumber
            NameMitRufname = "$Surname, $GivenName"
            GivenName = $GivenName
            Surname = $Surname
            ClassName = $ClassName
            Teacher = $Teacher
            Password = ''
            EntraObjectId = $EntraObjectId
            StoredUpn = $StoredUpn
        }
    }

    function global:New-IntegrationUser {
        param(
            [Parameter(Mandatory)][string] $Id,
            [Parameter(Mandatory)][string] $GivenName,
            [Parameter(Mandatory)][string] $Surname,
            [Parameter(Mandatory)][string] $UserPrincipalName,
            [string] $Department = '',
            [string] $OfficeLocation = '',
            [bool] $AccountEnabled = $true,
            [string] $CompanyName = 'Montessori Schule Aufkirchen',
            [string] $EmployeeType = 'Schüler'
        )
        [pscustomobject]@{
            Id = $Id
            DisplayName = "$GivenName $Surname"
            GivenName = $GivenName
            Surname = $Surname
            UserPrincipalName = $UserPrincipalName
            Mail = $UserPrincipalName
            MailNickname = ($UserPrincipalName -split '@')[0]
            ProxyAddresses = @("SMTP:$UserPrincipalName")
            Department = $Department
            OfficeLocation = $OfficeLocation
            CompanyName = $CompanyName
            EmployeeType = $EmployeeType
            UsageLocation = 'DE'
            AgeGroup = if ($EmployeeType -eq 'Schüler') { 'Minor' } else { 'Adult' }
            ConsentProvidedForMinor = if ($EmployeeType -eq 'Schüler') { 'Granted' } else { 'NotRequired' }
            LegalAgeGroupClassification = if ($EmployeeType -eq 'Schüler') { 'MinorWithParentalConsent' } else { 'Adult' }
            AccountEnabled = $AccountEnabled
            UserType = 'Member'
        }
    }

    function global:New-IntegrationGroup {
        param([string] $Id, [string] $DisplayName)
        [pscustomobject]@{
            Id = $Id
            DisplayName = $DisplayName
            GroupTypes = @()
            MembershipRule = $null
            MembershipRuleProcessingState = $null
            IsDynamic = $false
            IsInherited = $false
            IsRemovable = $true
        }
    }

    function global:New-IntegrationSnapshot {
        $state = $global:IntegrationState
        $usersById = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $usersByUpn = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $owners = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($user in $state.Users.Values) {
            $usersById[$user.Id] = $user
            $usersByUpn[$user.UserPrincipalName] = $user
            [void]$reserved.Add($user.UserPrincipalName)
            $owners[$user.UserPrincipalName] = @($user.Id)
        }

        $groupsById = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $groupsByName = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $groupMatches = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($group in $state.Groups.Values) {
            $groupsById[$group.Id] = $group
            $groupsByName[$group.DisplayName] = $group
            $groupMatches[$group.DisplayName] = @($group)
        }

        $roleMembers = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $directByUser = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $transitiveByUser = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $managers = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($userId in $state.Users.Keys) {
            $direct = @($state.DirectGroups[$userId] | ForEach-Object { $state.Groups[$_] })
            $directByUser[$userId] = $direct
            $transitiveByUser[$userId] = $direct
            if (@($state.DirectGroups[$userId]) -contains $global:IntegrationGroupIds.StudentRole) {
                [void]$roleMembers.Add($userId)
            }
            $managers[$userId] = $state.Managers[$userId]
        }

        [pscustomobject]@{
            UsersById = $usersById
            UsersByUpn = $usersByUpn
            StudentRoleMemberIds = $roleMembers
            GroupsById = $groupsById
            GroupsByDisplayName = $groupsByName
            GroupMatchesByDisplayName = $groupMatches
            ReservedAddresses = $reserved
            AddressOwners = $owners
            DirectGroupsByUserId = $directByUser
            TransitiveGroupsByUserId = $transitiveByUser
            ManagerByUserId = $managers
        }
    }
}

AfterAll {
    foreach ($commandName in @($global:IntegrationStubCommands)) {
        Remove-Item -LiteralPath "Function:\global:$commandName" -Force -ErrorAction SilentlyContinue
    }
    foreach ($name in 'IntegrationGroupIds', 'IntegrationState', 'IntegrationStubCommands') {
        Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    }
    foreach ($name in 'New-IntegrationStudent', 'New-IntegrationUser', 'New-IntegrationGroup', 'New-IntegrationSnapshot') {
        Remove-Item -LiteralPath "Function:\global:$name" -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Stateful student synchronization' {
    BeforeEach {
        $copy = Join-Path $TestDrive 'Schueler-Integration.xlsx'
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'fixtures/Schueler-Testdaten.xlsx') -Destination $copy

        $users = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $users['teacher-lara'] = New-IntegrationUser -Id teacher-lara -GivenName Lara -Surname Lehrer `
            -UserPrincipalName lara.lehrer@monteaufkirchen.com -EmployeeType 'Pädagogisches Team'
        $users['teacher-lars'] = New-IntegrationUser -Id teacher-lars -GivenName Lars -Surname Lehrer `
            -UserPrincipalName lars.lehrer@monteaufkirchen.com -EmployeeType 'Pädagogisches Team'
        $users['changed-id'] = New-IntegrationUser -Id changed-id -GivenName Ben -Surname Beispiel `
            -UserPrincipalName bbeispiel@monteaufkirchen.com -Department JK1-3g1_1 -OfficeLocation G1
        $users['existing-id'] = New-IntegrationUser -Id existing-id -GivenName Clara -Surname Bestand `
            -UserPrincipalName cbestand@monteaufkirchen.com -Department JK1-3g2_1 -OfficeLocation G2
        $users['departure-id'] = New-IntegrationUser -Id departure-id -GivenName David -Surname Alt `
            -UserPrincipalName dalt@monteaufkirchen.com -Department JK1-3g2_1 -OfficeLocation G2

        $groups = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $groups[$global:IntegrationGroupIds.StudentRole] = New-IntegrationGroup -Id $global:IntegrationGroupIds.StudentRole -DisplayName 'SEC-A-ROL-Schule_Schüler'
        $groups[$global:IntegrationGroupIds.OtherRole] = New-IntegrationGroup -Id $global:IntegrationGroupIds.OtherRole -DisplayName 'SEC-A-ROL-Schule_PädagogischesTeam'
        $groups[$global:IntegrationGroupIds.License] = New-IntegrationGroup -Id $global:IntegrationGroupIds.License -DisplayName 'SEC-A-LIC-O365A1Student'
        $groups[$global:IntegrationGroupIds.ClassG1] = New-IntegrationGroup -Id $global:IntegrationGroupIds.ClassG1 -DisplayName 'SEC-A-CLS-JK1-3g1_1'
        $groups[$global:IntegrationGroupIds.ClassG2] = New-IntegrationGroup -Id $global:IntegrationGroupIds.ClassG2 -DisplayName 'SEC-A-CLS-JK1-3g2_1'
        $groups[$global:IntegrationGroupIds.ClassM2] = New-IntegrationGroup -Id $global:IntegrationGroupIds.ClassM2 -DisplayName 'SEC-A-CLS-JK4-6m2_4'

        $directGroups = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $directGroups['teacher-lara'] = @()
        $directGroups['teacher-lars'] = @()
        $directGroups['changed-id'] = @(
            $global:IntegrationGroupIds.StudentRole,
            $global:IntegrationGroupIds.License,
            $global:IntegrationGroupIds.ClassG1,
            $global:IntegrationGroupIds.OtherRole
        )
        $directGroups['existing-id'] = @(
            $global:IntegrationGroupIds.StudentRole,
            $global:IntegrationGroupIds.License,
            $global:IntegrationGroupIds.ClassG2
        )
        $directGroups['departure-id'] = @(
            $global:IntegrationGroupIds.StudentRole,
            $global:IntegrationGroupIds.License,
            $global:IntegrationGroupIds.ClassG2
        )

        $managers = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $managers['teacher-lara'] = $null
        $managers['teacher-lars'] = $null
        $managers['changed-id'] = 'teacher-lara'
        $managers['existing-id'] = 'teacher-lara'
        $managers['departure-id'] = 'teacher-lara'

        $mailboxes = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
        $mailboxes['bbeispiel@monteaufkirchen.com'] = [pscustomobject]@{ Exists = $true; Configured = $false }
        $mailboxes['cbestand@monteaufkirchen.com'] = [pscustomobject]@{ Exists = $true; Configured = $true }
        $mailboxes['dalt@monteaufkirchen.com'] = [pscustomobject]@{ Exists = $true; Configured = $false }

        $global:IntegrationState = @{
            CopyPath = $copy
            SourceHash = 'source-1'
            Students = @(
                New-IntegrationStudent -RowNumber 2 -GivenName Mia -Surname Muster -ClassName JK1-3g2_1 -Teacher 'Lara Lehrer'
                New-IntegrationStudent -RowNumber 3 -GivenName Ben -Surname Beispiel -ClassName JK4-6m2_4 -Teacher 'Lars Lehrer' `
                    -EntraObjectId changed-id -StoredUpn bbeispiel@monteaufkirchen.com
                New-IntegrationStudent -RowNumber 4 -GivenName Clara -Surname Bestand -ClassName JK1-3g2_1 -Teacher 'Lara Lehrer' `
                    -EntraObjectId existing-id -StoredUpn cbestand@monteaufkirchen.com
            )
            Users = $users
            Groups = $groups
            DirectGroups = $directGroups
            Managers = $managers
            Mailboxes = $mailboxes
            CasMailboxes = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
            Events = [Collections.Generic.List[string]]::new()
            PasswordsGenerated = [Collections.Generic.List[string]]::new()
            PasswordsWritten = [Collections.Generic.List[string]]::new()
            CreateBodies = [Collections.Generic.List[object]]::new()
            GroupReads = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
            WorkbookWrites = 0
            SessionRevocations = 0
        }

        Mock Connect-SchuelerGraph -ModuleName SchuelerSync {
            [pscustomobject]@{ TenantId = 'integration-tenant'; Account = 'admin@integration.invalid' }
        }
        Mock Get-EntraSnapshot -ModuleName SchuelerSync { New-IntegrationSnapshot }
        Mock Connect-SchuelerExchangeOnline -ModuleName SchuelerSync { return }
        Mock Write-Information -ModuleName SchuelerSync { return }
        Mock Write-StudentComparisonReport -ModuleName SchuelerSync { return }
        Mock Start-Sleep -ModuleName SchuelerSync { throw 'Integration test must not wait.' }
        Mock Assert-WorkbookSafeForPasswordWrite -ModuleName SchuelerSync { return }
        Mock Assert-StudentWorkbookVersion -ModuleName SchuelerSync {
            if ($ExpectedSourceHash -cne $global:IntegrationState.SourceHash) {
                throw 'Workbook changed since preflight.'
            }
        }
        Mock Read-StudentWorkbook -ModuleName SchuelerSync {
            [pscustomobject]@{
                Path = $Path
                WorksheetName = 'Schueler'
                SourceHash = $global:IntegrationState.SourceHash
                Students = [object[]]@($global:IntegrationState.Students)
            }
        }
        Mock Write-StudentWorkbookUpdate -ModuleName SchuelerSync {
            if ($ExpectedSourceHash -cne $global:IntegrationState.SourceHash) {
                throw 'Workbook changed before write.'
            }
            foreach ($update in $Updates) {
                $row = @($global:IntegrationState.Students | Where-Object RowNumber -eq $update.RowNumber)
                if ($row.Count -ne 1) { throw "Unknown workbook row $($update.RowNumber)." }
                $row[0].EntraObjectId = [string]$update.EntraObjectId
                $row[0].StoredUpn = [string]$update.UPN
                if (-not [string]::IsNullOrWhiteSpace([string]$update.Password)) {
                    $row[0].Password = [string]$update.Password
                    $global:IntegrationState.PasswordsWritten.Add([string]$update.Password)
                }
            }
            $global:IntegrationState.WorkbookWrites++
            $global:IntegrationState.SourceHash = "source-$($global:IntegrationState.WorkbookWrites + 1)"
            $global:IntegrationState.Events.Add('excel-write')
            [pscustomobject]@{ Path = $Path; SourceHash = $global:IntegrationState.SourceHash }
        }
        Mock New-StudentPassword -ModuleName SchuelerSync {
            $password = 'TigerWiese56'
            $global:IntegrationState.PasswordsGenerated.Add($password)
            $password
        }
        Mock New-MgUser -ModuleName SchuelerSync {
            $id = 'new-id'
            $global:IntegrationState.CreateBodies.Add($BodyParameter)
            $user = New-IntegrationUser -Id $id -GivenName $BodyParameter.GivenName -Surname $BodyParameter.Surname `
                -UserPrincipalName $BodyParameter.UserPrincipalName -Department $BodyParameter.Department `
                -OfficeLocation $BodyParameter.OfficeLocation -AccountEnabled:$BodyParameter.AccountEnabled
            $global:IntegrationState.Users[$id] = $user
            $global:IntegrationState.DirectGroups[$id] = @()
            $global:IntegrationState.Managers[$id] = $null
            $global:IntegrationState.Mailboxes[$BodyParameter.UserPrincipalName] = [pscustomobject]@{ Exists = $true; Configured = $false }
            $global:IntegrationState.Events.Add('create-disabled')
            $user
        }
        Mock Assert-NewEntraStudentAttribute -ModuleName SchuelerSync {
            $user = $global:IntegrationState.Users[$UserId]
            $verified = $user.AccountEnabled -eq $false -and $user.DisplayName -ceq $Desired.DisplayName -and
                $user.Department -ceq $Desired.Department -and $user.OfficeLocation -ceq $Desired.OfficeLocation
            $global:IntegrationState.Events.Add('attributes-verified')
            return $verified
        }
        Mock Set-EntraStudentAttribute -ModuleName SchuelerSync {
            $user = $global:IntegrationState.Users[$UserId]
            foreach ($difference in @($Differences | Where-Object { $_.Area -eq 'Entra' -and $_.Action -eq 'Set' })) {
                $user.($difference.Field) = $Desired.($difference.Field)
            }
            $global:IntegrationState.Events.Add("attributes-set:$UserId")
            [pscustomobject]@{ Verified = $true }
        }
        Mock Set-EntraStudentManager -ModuleName SchuelerSync {
            $global:IntegrationState.Managers[$UserId] = $DesiredManagerId
            $global:IntegrationState.Events.Add("manager-set:$UserId")
            [pscustomobject]@{ Verified = $true; ManagerId = $DesiredManagerId }
        }
        Mock New-MgGroupMemberByRef -ModuleName SchuelerSync {
            $userId = ([string]$BodyParameter['@odata.id'] -split '/')[-1]
            $currentIds = @($global:IntegrationState.DirectGroups[$userId])
            if ($currentIds -notcontains $GroupId) {
                $global:IntegrationState.DirectGroups[$userId] = @($currentIds + $GroupId)
            }
            $name = [string]$global:IntegrationState.Groups[$GroupId].DisplayName
            $global:IntegrationState.Events.Add("group-add:${userId}:$name")
        }
        Mock Remove-MgGroupMemberByRef -ModuleName SchuelerSync {
            $global:IntegrationState.DirectGroups[$DirectoryObjectId] = @(
                $global:IntegrationState.DirectGroups[$DirectoryObjectId] | Where-Object { $_ -ne $GroupId }
            )
            $name = [string]$global:IntegrationState.Groups[$GroupId].DisplayName
            $global:IntegrationState.Events.Add("group-remove:${DirectoryObjectId}:$name")
        }
        Mock Get-MgUserMemberOfAsGroup -ModuleName SchuelerSync {
            $reads = if ($global:IntegrationState.GroupReads.ContainsKey($UserId)) {
                [int]$global:IntegrationState.GroupReads[$UserId] + 1
            } else { 1 }
            $global:IntegrationState.GroupReads[$UserId] = $reads
            $global:IntegrationState.Events.Add("group-read:${UserId}:$reads")
            foreach ($groupId in @($global:IntegrationState.DirectGroups[$UserId])) {
                $global:IntegrationState.Groups[$groupId]
            }
        }
        Mock Enable-EntraStudent -ModuleName SchuelerSync {
            $global:IntegrationState.Users[$UserId].AccountEnabled = $true
            $global:IntegrationState.Events.Add("enable:$UserId")
            [pscustomobject]@{ Verified = $true }
        }
        Mock Disable-EntraStudent -ModuleName SchuelerSync {
            $global:IntegrationState.Users[$UserId].AccountEnabled = $false
            $global:IntegrationState.Events.Add("disable:$UserId")
            [pscustomobject]@{ Verified = $true }
        }
        Mock Revoke-EntraStudentSession -ModuleName SchuelerSync {
            $global:IntegrationState.SessionRevocations++
            $global:IntegrationState.Events.Add("revoke:$UserId")
            return $true
        }
        Mock Get-EntraStudentCurrentIdentity -ModuleName SchuelerSync {
            $global:IntegrationState.Users[$UserId]
        }
        Mock Get-ExchangeRecipientAddress -ModuleName SchuelerSync {
            $owners = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($user in $global:IntegrationState.Users.Values) {
                $owners[$user.UserPrincipalName] = @($user.Id)
            }
            [pscustomobject]@{ ReservedAddresses = @($owners.Keys); AddressOwners = $owners }
        }
        Mock Get-StudentMailboxState -ModuleName SchuelerSync {
            $mailbox = $global:IntegrationState.Mailboxes[$UserPrincipalName]
            if ($null -eq $mailbox -or -not $mailbox.Exists) {
                return [pscustomobject]@{ Exists = $false; Status = 'MailboxNotReady'; Differences = @() }
            }
            $differences = if ($mailbox.Configured) { @() } else {
                @([pscustomobject]@{ Field = 'AddressBookPolicy'; Current = 'Wrong'; Desired = 'MON-EXO-ABP-Schule_Schüler'; Action = 'Set' })
            }
            [pscustomobject]@{
                Exists = $true
                Status = 'Ready'
                Mailbox = [pscustomobject]@{ UserPrincipalName = $UserPrincipalName }
                CasMailbox = [pscustomobject]@{ UserPrincipalName = $UserPrincipalName }
                Differences = [object[]]@($differences)
            }
        }
        Mock Set-StudentMailboxConfiguration -ModuleName SchuelerSync {
            $mailbox = $global:IntegrationState.Mailboxes[$UserPrincipalName]
            if ($null -eq $mailbox -or -not $mailbox.Exists) {
                return [pscustomobject]@{ UserPrincipalName = $UserPrincipalName; Status = 'MailboxNotReady'; Changed = $false; Error = $null }
            }
            $changed = -not $mailbox.Configured
            $mailbox.Configured = $true
            $global:IntegrationState.CasMailboxes[$UserPrincipalName] = [pscustomobject]@{ Configured = $true }
            $global:IntegrationState.Events.Add("exchange-configure:$UserPrincipalName")
            [pscustomobject]@{
                UserPrincipalName = $UserPrincipalName
                Status = if ($changed) { 'Configured' } else { 'Compliant' }
                Changed = $changed
                Differences = @()
                RemainingDifferences = @()
                Error = $null
            }
        }
    }

    It 'applies a complete update in order and leaves the next comparison read-only' {
        $first = Invoke-SchuelerSync -File $global:IntegrationState.CopyPath -Update -Confirm:$false

        $first.HasErrors | Should -BeFalse
        $first.Comparison.NewStudents.Count | Should -Be 1
        $first.Comparison.ChangedStudents.Count | Should -Be 1
        $first.Comparison.ExistingStudents.Count | Should -Be 1
        $first.Comparison.Departures.Count | Should -Be 1
        $global:IntegrationState.PasswordsGenerated | Should -HaveCount 1
        $global:IntegrationState.PasswordsWritten | Should -HaveCount 1
        $global:IntegrationState.PasswordsWritten[0].Length | Should -Be 12
        $global:IntegrationState.PasswordsWritten[0] | Should -Be $global:IntegrationState.PasswordsGenerated[0]
        $global:IntegrationState.CreateBodies | Should -HaveCount 1
        $createBody = $global:IntegrationState.CreateBodies[0]
        $createBody.AccountEnabled | Should -BeFalse
        $createBody.PasswordProfile.Password | Should -Be $global:IntegrationState.PasswordsGenerated[0]
        $createBody.PasswordProfile.ForceChangePasswordNextSignIn | Should -BeFalse
        Should -Invoke New-MgUser -ModuleName SchuelerSync -Times 1 -Exactly
        $global:IntegrationState.WorkbookWrites | Should -Be 1
        $global:IntegrationState.SourceHash | Should -Be 'source-2'
        $global:IntegrationState.Users['new-id'].AccountEnabled | Should -BeTrue
        $global:IntegrationState.Users['new-id'].DisplayName | Should -Be 'Mia Muster'
        $global:IntegrationState.Users['new-id'].UsageLocation | Should -Be 'DE'
        $global:IntegrationState.Users['new-id'].AgeGroup | Should -Be 'Minor'
        $global:IntegrationState.Users['new-id'].ConsentProvidedForMinor | Should -Be 'Granted'
        $global:IntegrationState.Users['changed-id'].Department | Should -Be 'JK4-6m2_4'
        $global:IntegrationState.Managers['new-id'] | Should -Be 'teacher-lara'
        $global:IntegrationState.Managers['changed-id'] | Should -Be 'teacher-lars'
        $global:IntegrationState.Users['departure-id'].AccountEnabled | Should -BeFalse
        $global:IntegrationState.SessionRevocations | Should -Be 1

        $events = @($global:IntegrationState.Events)
        [Array]::IndexOf($events, 'create-disabled') | Should -BeLessThan ([Array]::IndexOf($events, 'attributes-verified'))
        [Array]::IndexOf($events, 'attributes-verified') | Should -BeLessThan ([Array]::IndexOf($events, 'manager-set:new-id'))
        [Array]::IndexOf($events, 'manager-set:new-id') | Should -BeLessThan ([Array]::IndexOf($events, 'excel-write'))
        $newGroupVerification = [Array]::IndexOf($events, 'group-read:new-id:1')
        foreach ($requiredGroup in 'SEC-A-ROL-Schule_Schüler', 'SEC-A-LIC-O365A1Student', 'SEC-A-CLS-JK1-3g2_1') {
            $requiredAdd = [Array]::IndexOf($events, "group-add:new-id:$requiredGroup")
            $requiredAdd | Should -BeGreaterOrEqual 0
            $requiredAdd | Should -BeLessThan $newGroupVerification
        }
        $newGroupVerification | Should -BeLessThan ([Array]::IndexOf($events, 'group-read:new-id:2'))
        [Array]::IndexOf($events, 'group-read:new-id:2') | Should -BeLessThan ([Array]::IndexOf($events, 'excel-write'))
        [Array]::IndexOf($events, 'excel-write') | Should -BeLessThan ([Array]::IndexOf($events, 'enable:new-id'))
        $classAdd = [Array]::IndexOf($events, 'group-add:changed-id:SEC-A-CLS-JK4-6m2_4')
        $changedGroupVerification = [Array]::IndexOf($events, 'group-read:changed-id:1')
        $changedFinalVerification = [Array]::IndexOf($events, 'group-read:changed-id:2')
        $classRemove = [Array]::IndexOf($events, 'group-remove:changed-id:SEC-A-CLS-JK1-3g1_1')
        $roleRemove = [Array]::IndexOf($events, 'group-remove:changed-id:SEC-A-ROL-Schule_PädagogischesTeam')
        $classAdd | Should -BeGreaterOrEqual 0
        $classAdd | Should -BeLessThan $changedGroupVerification
        $changedGroupVerification | Should -BeLessThan $classRemove
        $changedGroupVerification | Should -BeLessThan $roleRemove
        $classRemove | Should -BeLessThan $changedFinalVerification
        $roleRemove | Should -BeLessThan $changedFinalVerification
        @($global:IntegrationState.DirectGroups['new-id']) | Should -Contain $global:IntegrationGroupIds.StudentRole
        @($global:IntegrationState.DirectGroups['new-id']) | Should -Contain $global:IntegrationGroupIds.License
        @($global:IntegrationState.DirectGroups['new-id']) | Should -Contain $global:IntegrationGroupIds.ClassG2
        $changedRoles = @($global:IntegrationState.DirectGroups['changed-id'] | Where-Object {
                $global:IntegrationState.Groups[$_].DisplayName -like 'SEC-A-ROL-*'
            })
        $changedClasses = @($global:IntegrationState.DirectGroups['changed-id'] | Where-Object {
                $global:IntegrationState.Groups[$_].DisplayName -like 'SEC-A-CLS-*'
            })
        $changedRoles | Should -HaveCount 1
        $changedRoles | Should -Contain $global:IntegrationGroupIds.StudentRole
        $changedClasses | Should -HaveCount 1
        $changedClasses | Should -Contain $global:IntegrationGroupIds.ClassM2
        Should -Invoke New-MgGroupMemberByRef -ModuleName SchuelerSync -Times 4 -Exactly
        Should -Invoke Remove-MgGroupMemberByRef -ModuleName SchuelerSync -Times 2 -Exactly
        Should -Invoke Get-MgUserMemberOfAsGroup -ModuleName SchuelerSync -Times 4 -Exactly
        [Array]::IndexOf($events, 'disable:departure-id') | Should -BeLessThan ([Array]::IndexOf($events, 'revoke:departure-id'))
        [Array]::IndexOf($events, 'revoke:departure-id') | Should -BeLessThan ([Array]::IndexOf($events, 'exchange-configure:mmuster@monteaufkirchen.com'))
        foreach ($upn in 'mmuster@monteaufkirchen.com', 'bbeispiel@monteaufkirchen.com', 'cbestand@monteaufkirchen.com') {
            $global:IntegrationState.Mailboxes[$upn].Configured | Should -BeTrue
            $global:IntegrationState.CasMailboxes[$upn].Configured | Should -BeTrue
        }
        @($first.Actions | Where-Object { $_.Phase -eq 'Create' -and $_.Status -eq 'Succeeded' }) | Should -HaveCount 1
        @($first.Actions | Where-Object { $_.Phase -eq 'Update' -and $_.Status -eq 'Succeeded' }) | Should -HaveCount 1
        @($first.Actions | Where-Object { $_.Phase -eq 'Disable' -and $_.Status -eq 'Succeeded' }) | Should -HaveCount 1
        @($first.Actions | Where-Object { $_.Phase -eq 'RevokeSessions' -and $_.Status -eq 'Succeeded' }) | Should -HaveCount 1

        $writesBeforeSecondComparison = $global:IntegrationState.Events.Count
        $second = Invoke-SchuelerSync -File $global:IntegrationState.CopyPath

        $second.HasErrors | Should -BeFalse
        $second.Mode | Should -Be 'Compare'
        $second.Comparison.NewStudents | Should -BeNullOrEmpty
        $second.Comparison.ChangedStudents | Should -BeNullOrEmpty
        $second.Comparison.ExistingStudents | Should -HaveCount 3
        $second.Comparison.Departures | Should -HaveCount 1
        $second.Comparison.Departures[0].UserId | Should -Be 'departure-id'
        $second.Comparison.Departures[0].AccountEnabled | Should -BeFalse
        $second.Actions | Should -BeNullOrEmpty
        $global:IntegrationState.Events.Count | Should -Be $writesBeforeSecondComparison
        $global:IntegrationState.WorkbookWrites | Should -Be 1
        $global:IntegrationState.SourceHash | Should -Be 'source-2'
        $global:IntegrationState.SessionRevocations | Should -Be 1
    }

    It 'keeps Update WhatIf free of writes waits and password generation' {
        $before = $global:IntegrationState.Events.Count

        $result = Invoke-SchuelerSync -File $global:IntegrationState.CopyPath -Update -WhatIf

        $result.HasErrors | Should -BeFalse
        $result.Actions.Status | Should -Contain 'WhatIf'
        $global:IntegrationState.Events.Count | Should -Be $before
        $global:IntegrationState.WorkbookWrites | Should -Be 0
        $global:IntegrationState.PasswordsGenerated | Should -BeNullOrEmpty
        $global:IntegrationState.PasswordsWritten | Should -BeNullOrEmpty
        $global:IntegrationState.SessionRevocations | Should -Be 0
        Should -Invoke New-StudentPassword -ModuleName SchuelerSync -Times 0 -Exactly
        Should -Invoke New-MgUser -ModuleName SchuelerSync -Times 0 -Exactly
        Should -Invoke New-MgGroupMemberByRef -ModuleName SchuelerSync -Times 0 -Exactly
        Should -Invoke Remove-MgGroupMemberByRef -ModuleName SchuelerSync -Times 0 -Exactly
        Should -Invoke Write-StudentWorkbookUpdate -ModuleName SchuelerSync -Times 0 -Exactly
        Should -Invoke Set-StudentMailboxConfiguration -ModuleName SchuelerSync -Times 0 -Exactly
        Should -Invoke Start-Sleep -ModuleName SchuelerSync -Times 0 -Exactly
    }
}

Describe 'Static analyzer contract' {
    It 'enables the required strict analyzer rules' {
        $settings = Import-PowerShellDataFile (Join-Path $repoRoot 'PSScriptAnalyzerSettings.psd1')

        ($settings.Severity -join ',') | Should -Be 'Error,Warning'
        $settings.IncludeDefaultRules | Should -BeTrue
        $settings.Rules.PSAvoidUsingPlainTextForPassword.Enable | Should -BeTrue
        $settings.Rules.PSAvoidUsingWriteHost.Enable | Should -BeTrue
        $settings.Rules.PSUseShouldProcessForStateChangingFunctions.Enable | Should -BeTrue
    }

    It 'documents the narrow plain-text password suppression at the Graph boundary' {
        $tokens = $null
        $parseErrors = $null
        $sourcePath = Join-Path $repoRoot 'src/SchuelerSync/Private/Entra.ps1'
        $ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$parseErrors)

        $parseErrors | Should -BeNullOrEmpty
        $function = $ast.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq 'New-DisabledEntraStudent'
            }, $true)
        $function | Should -Not -BeNullOrEmpty
        $passwordParameter = @($function.Body.ParamBlock.Parameters | Where-Object {
                $_.Name.VariablePath.UserPath -eq 'Password'
            })
        $passwordParameter | Should -HaveCount 1
        $passwordParameter[0].StaticType | Should -Be ([string])

        $suppressions = @($function.Body.ParamBlock.Attributes | Where-Object {
                $_.TypeName.FullName -match '(^|\.)SuppressMessageAttribute$'
            })
        $matchingSuppression = @($suppressions | Where-Object {
                $_.PositionalArguments.Count -ge 2 -and
                $_.PositionalArguments[0].SafeGetValue() -eq 'PSAvoidUsingPlainTextForPassword' -and
                $_.PositionalArguments[1].SafeGetValue() -eq 'Password'
            })
        $matchingSuppression | Should -HaveCount 1
        $justification = @($matchingSuppression[0].NamedArguments | Where-Object ArgumentName -eq 'Justification')
        $justification | Should -HaveCount 1
        $justification[0].Argument.SafeGetValue() | Should -Match 'New-MgUser requires an in-memory plain string'
        $justification[0].Argument.SafeGetValue() | Should -Match 'not logged and is persisted only to the protected workbook'

        $graphCreate = $function.Find({
                param($node)
                $node -is [Management.Automation.Language.CommandAst] -and
                    $node.GetCommandName() -eq 'New-MgUser'
            }, $true)
        $graphCreate | Should -Not -BeNullOrEmpty
    }
}
