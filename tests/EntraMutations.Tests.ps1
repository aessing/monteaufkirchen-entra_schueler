BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $global:EntraMutationDesired = [pscustomobject]@{
        DisplayName = 'Mia Beispiel'
        GivenName = 'Mia'
        Surname = 'Beispiel'
        UserPrincipalName = 'mbeispiel@schule.example'
        Mail = 'mbeispiel@schule.example'
        Department = 'G1'
        OfficeLocation = 'Grundschule'
        CompanyName = 'Beispielschule'
        EmployeeType = 'Schueler'
        UsageLocation = 'DE'
        AgeGroup = 'Minor'
        ConsentProvidedForMinor = 'Granted'
        LegalAgeGroupClassification = 'MinorWithParentalConsent'
    }
    $global:EntraMutationGroups = @{
        Role = [pscustomobject]@{
            Id = '11111111-1111-1111-1111-111111111111'
            DisplayName = 'SEC-A-ROL-Schule_Schueler'
            GroupTypes = @()
            MembershipRule = $null
        }
        License = [pscustomobject]@{
            Id = '22222222-2222-2222-2222-222222222222'
            DisplayName = 'SEC-A-LIC-O365A1Student'
            GroupTypes = @()
            MembershipRule = $null
        }
        Class = [pscustomobject]@{
            Id = '33333333-3333-3333-3333-333333333333'
            DisplayName = 'SEC-A-CLS-G1'
            GroupTypes = @()
            MembershipRule = $null
        }
    }
}

AfterAll {
    Remove-Variable -Name EntraMutationDesired -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name EntraMutationGroups -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name EntraMutationState -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Verified Entra student mutations' {
    BeforeEach {
        $global:EntraMutationState = @{
            Events = [Collections.Generic.List[string]]::new()
            CreateBodies = [Collections.Generic.List[object]]::new()
            CreateFailuresRemaining = 0
            CreateFailureMessage = ''
            CreateWriteErrorsRemaining = 0
            CreateWriteErrorMessage = ''
            DirectGroupReads = [Collections.Generic.Queue[object]]::new()
            GroupAddFailureId = ''
            UserReads = [Collections.Generic.Queue[object]]::new()
            ManagerReads = [Collections.Generic.Queue[object]]::new()
            UpdateBodies = [Collections.Generic.List[object]]::new()
        }

        Mock New-MgUser -ModuleName SchuelerSync -RemoveParameterType BodyParameter {
            $global:EntraMutationState.Events.Add('create-disabled')
            $global:EntraMutationState.CreateBodies.Add($BodyParameter)
            if ($global:EntraMutationState.CreateWriteErrorsRemaining -gt 0) {
                $global:EntraMutationState.CreateWriteErrorsRemaining--
                Write-Error $global:EntraMutationState.CreateWriteErrorMessage -ErrorAction Stop
            }
            if ($global:EntraMutationState.CreateFailuresRemaining -gt 0) {
                $global:EntraMutationState.CreateFailuresRemaining--
                throw [InvalidOperationException]::new($global:EntraMutationState.CreateFailureMessage)
            }
            [pscustomobject]@{
                Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                UserPrincipalName = $BodyParameter.UserPrincipalName
                AccountEnabled = $BodyParameter.AccountEnabled
            }
        }
        Mock Update-MgUser -ModuleName SchuelerSync {
            if ($null -ne $BodyParameter) {
                $global:EntraMutationState.UpdateBodies.Add($BodyParameter)
                $global:EntraMutationState.Events.Add('update-attributes')
            } else {
                $global:EntraMutationState.Events.Add("account-enabled-$AccountEnabled")
            }
        }
        Mock Get-MgUser -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add('read-user')
            if ($global:EntraMutationState.UserReads.Count -gt 0) {
                return $global:EntraMutationState.UserReads.Dequeue()
            }
            [pscustomobject]@{ Id = $UserId; AccountEnabled = $false }
        }
        Mock Set-MgUserManagerByRef -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add('set-manager')
        }
        Mock Get-MgUserManager -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add('read-manager')
            if ($global:EntraMutationState.ManagerReads.Count -gt 0) {
                return $global:EntraMutationState.ManagerReads.Dequeue()
            }
            [pscustomobject]@{ Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' }
        }
        Mock New-MgGroupMemberByRef -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add("add-$GroupId")
            if ([string]::Equals($GroupId, $global:EntraMutationState.GroupAddFailureId, [StringComparison]::OrdinalIgnoreCase)) {
                throw [InvalidOperationException]::new('Request failed while adding the mandatory group.')
            }
        }
        Mock Remove-MgGroupMemberByRef -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add("remove-$GroupId")
        }
        Mock Get-MgUserMemberOfAsGroup -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add('read-direct-groups')
            if ($global:EntraMutationState.DirectGroupReads.Count -gt 0) {
                return $global:EntraMutationState.DirectGroupReads.Dequeue()
            }
            @()
        }
        Mock Revoke-MgUserSignInSession -ModuleName SchuelerSync {
            $global:EntraMutationState.Events.Add('revoke-sessions')
            [pscustomobject]@{ Value = $true }
        }
    }

    InModuleScope SchuelerSync {
        Context 'disabled creation and password retry' {
            It 'creates a new student disabled with the exact supported attributes' {
                $password = 'LegoWolke42'

                $created = New-DisabledEntraStudent -Desired $global:EntraMutationDesired -Password $password -Confirm:$false

                $created.Id | Should -Be 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                $body = $global:EntraMutationState.CreateBodies[0]
                $body.Count | Should -Be 15
                $body.AccountEnabled | Should -BeFalse
                $body.DisplayName | Should -Be 'Mia Beispiel'
                $body.GivenName | Should -Be 'Mia'
                $body.Surname | Should -Be 'Beispiel'
                $body.UserPrincipalName | Should -Be 'mbeispiel@schule.example'
                $body.MailNickname | Should -Be 'mbeispiel'
                $body.Mail | Should -Be 'mbeispiel@schule.example'
                $body.Department | Should -Be 'G1'
                $body.OfficeLocation | Should -Be 'Grundschule'
                $body.CompanyName | Should -Be 'Beispielschule'
                $body.EmployeeType | Should -Be 'Schueler'
                $body.UsageLocation | Should -Be 'DE'
                $body.AgeGroup | Should -Be 'Minor'
                $body.ConsentProvidedForMinor | Should -Be 'Granted'
                $body.PasswordProfile.Count | Should -Be 2
                $body.PasswordProfile.Password | Should -Be $password
                $body.PasswordProfile.ForceChangePasswordNextSignIn | Should -BeFalse
                $body.ContainsKey('LegalAgeGroupClassification') | Should -BeFalse
                Should -Invoke New-MgUser -Times 1 -Exactly
                Should -Invoke New-MgUser -Times 1 -Exactly -ParameterFilter { $ErrorAction -eq 'Stop' }
            }

            It 'accepts an initially empty password collection on the first create attempt' {
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                $result = New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false

                $result.Attempts | Should -Be 1
                $usedPasswords.Count | Should -Be 1
                Should -Invoke New-MgUser -Times 1 -Exactly -ParameterFilter {
                    $ErrorAction -eq 'Stop'
                }
            }

            It 'turns a non-terminating Graph password-policy error into a caught retry' {
                $global:EntraMutationState.CreateWriteErrorsRemaining = 1
                $global:EntraMutationState.CreateWriteErrorMessage = 'Request_BadRequest: PasswordProfile.Password does not comply with password complexity requirements.'
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                $previousErrorActionPreference = $ErrorActionPreference
                try {
                    $ErrorActionPreference = 'Continue'
                    $result = New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false
                } finally {
                    $ErrorActionPreference = $previousErrorActionPreference
                }

                $result.Attempts | Should -Be 2
                $global:EntraMutationState.CreateBodies.Count | Should -Be 2
                @($global:EntraMutationState.CreateBodies | ForEach-Object { $_.PasswordProfile.Password } | Select-Object -Unique).Count | Should -Be 2
            }

            It 'uses at most five fresh replacement passwords after recognized policy rejections' {
                $global:EntraMutationState.CreateFailuresRemaining = 5
                $global:EntraMutationState.CreateFailureMessage = 'Request_BadRequest: PasswordProfile.Password does not comply with password complexity requirements.'
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                $result = New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false

                $result.User.Id | Should -Be 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                $result.Attempts | Should -Be 6
                $result.Password | Should -Be $global:EntraMutationState.CreateBodies[5].PasswordProfile.Password
                $global:EntraMutationState.CreateBodies.Count | Should -Be 6
                @($global:EntraMutationState.CreateBodies | ForEach-Object { $_.PasswordProfile.Password } | Select-Object -Unique).Count | Should -Be 6
                $usedPasswords.Count | Should -Be 6
            }

            It 'stops after six total password-policy rejections' {
                $global:EntraMutationState.CreateFailuresRemaining = 6
                $global:EntraMutationState.CreateFailureMessage = 'Request_BadRequest: PasswordProfile.Password does not comply with password policy requirements.'
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                { New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false } |
                    Should -Throw '*password policy*'

                $global:EntraMutationState.CreateBodies.Count | Should -Be 6
                $usedPasswords.Count | Should -Be 6
            }

            It 'rethrows permission failures without generating another password' {
                $global:EntraMutationState.CreateFailuresRemaining = 1
                $global:EntraMutationState.CreateFailureMessage = 'Authorization_RequestDenied: Insufficient privileges to complete the operation.'
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                { New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false } |
                    Should -Throw '*Insufficient privileges*'

                $global:EntraMutationState.CreateBodies.Count | Should -Be 1
                $usedPasswords.Count | Should -Be 0
            }

            It 'rethrows object conflicts without generating another password' {
                $global:EntraMutationState.CreateFailuresRemaining = 1
                $global:EntraMutationState.CreateFailureMessage = 'Request_ResourceConflict: Another object with the same value already exists.'
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                { New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false } |
                    Should -Throw '*same value already exists*'

                $global:EntraMutationState.CreateBodies.Count | Should -Be 1
                $usedPasswords.Count | Should -Be 0
            }

            It 'rethrows transport failures without generating another password' {
                $global:EntraMutationState.CreateFailuresRemaining = 1
                $global:EntraMutationState.CreateFailureMessage = 'The transport connection was interrupted.'
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                { New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -Confirm:$false } |
                    Should -Throw '*transport connection*'

                $global:EntraMutationState.CreateBodies.Count | Should -Be 1
                $usedPasswords.Count | Should -Be 0
            }

            It 'propagates WhatIf through the password retry wrapper without reserving a password' {
                $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

                $result = New-StudentWithPasswordRetry -Desired $global:EntraMutationDesired -InitialPassword 'LegoWolke42' -UsedPasswords $usedPasswords -WhatIf

                $result | Should -BeNullOrEmpty
                $usedPasswords.Count | Should -Be 0
                Should -Invoke New-MgUser -Times 0 -Exactly
            }
        }

        Context 'attribute and manager verification' {
            It 'writes only selected Entra set differences and never legal age classification' {
                [void]$global:EntraMutationState.UserReads.Enqueue([pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    DisplayName = 'Mia Beispiel'
                })
                $differences = @(
                    [pscustomobject]@{ Area = 'Entra'; Field = 'DisplayName'; Desired = 'Mia Beispiel'; Action = 'Set' }
                    [pscustomobject]@{ Area = 'Entra'; Field = 'LegalAgeGroupClassification'; Desired = 'MinorWithParentalConsent'; Action = 'Set' }
                    [pscustomobject]@{ Area = 'Entra'; Field = 'Department'; Desired = 'G1'; Action = 'Compare' }
                    [pscustomobject]@{ Area = 'Manager'; Field = 'ManagerId'; Desired = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; Action = 'Set' }
                    [pscustomobject]@{ Area = 'Exchange'; Field = 'CustomAttribute1'; Desired = 'Schueler'; Action = 'Set' }
                )

                $result = Set-EntraStudentAttribute -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Desired $global:EntraMutationDesired -Differences $differences -Confirm:$false

                $result.Verified | Should -BeTrue
                Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $BodyParameter.Count -eq 1 -and
                    $BodyParameter.DisplayName -eq 'Mia Beispiel' -and
                    -not $BodyParameter.ContainsKey('LegalAgeGroupClassification') -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $Property -contains 'DisplayName' -and
                    $ErrorAction -eq 'Stop'
                }
            }

            It 'fails attribute verification when Graph returns a different value' {
                [void]$global:EntraMutationState.UserReads.Enqueue([pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    DisplayName = 'Wrong Name'
                })
                $differences = @(
                    [pscustomobject]@{ Area = 'Entra'; Field = 'DisplayName'; Desired = 'Mia Beispiel'; Action = 'Set' }
                )

                { Set-EntraStudentAttribute -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Desired $global:EntraMutationDesired -Differences $differences -Confirm:$false } |
                    Should -Throw '*DisplayName*could not be verified*'

                Should -Invoke Update-MgUser -Times 1 -Exactly
                Should -Invoke Get-MgUser -Times 1 -Exactly
            }

            It 'uses the exact manager reference and verifies it by rereading' {
                [void]$global:EntraMutationState.ManagerReads.Enqueue([pscustomobject]@{ Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' })

                $result = Set-EntraStudentManager -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -CurrentManagerId 'cccccccc-cccc-cccc-cccc-cccccccccccc' -DesiredManagerId 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' -Confirm:$false

                $result.Verified | Should -BeTrue
                Should -Invoke Set-MgUserManagerByRef -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $BodyParameter.Count -eq 1 -and
                    $BodyParameter['@odata.id'] -eq 'https://graph.microsoft.com/v1.0/users/bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-MgUserManager -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $ErrorAction -eq 'Stop'
                }
                $global:EntraMutationState.Events | Should -Be @('set-manager', 'read-manager')
            }

            It 'does not rewrite an unchanged manager but still verifies it' {
                [void]$global:EntraMutationState.ManagerReads.Enqueue([pscustomobject]@{ Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' })

                $result = Set-EntraStudentManager -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -CurrentManagerId 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' -DesiredManagerId 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' -Confirm:$false

                $result.Changed | Should -BeFalse
                $result.Verified | Should -BeTrue
                Should -Invoke Set-MgUserManagerByRef -Times 0 -Exactly
                Should -Invoke Get-MgUserManager -Times 1 -Exactly
            }

            It 'fails manager verification when the reread differs' {
                [void]$global:EntraMutationState.ManagerReads.Enqueue([pscustomobject]@{ Id = 'dddddddd-dddd-dddd-dddd-dddddddddddd' })

                { Set-EntraStudentManager -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -CurrentManagerId 'cccccccc-cccc-cccc-cccc-cccccccccccc' -DesiredManagerId 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' -Confirm:$false } |
                    Should -Throw '*manager*could not be verified*'

                Should -Invoke Set-MgUserManagerByRef -Times 1 -Exactly
                Should -Invoke Get-MgUserManager -Times 1 -Exactly
            }
        }

        Context 'add-before-remove group reconciliation' {
            It 'adds role license and class, verifies targets, then removes competing static groups and rereads' {
                $oldRole = [pscustomobject]@{
                    Id = '44444444-4444-4444-4444-444444444444'
                    DisplayName = 'SEC-A-ROL-Schule_Alt'
                    GroupTypes = @()
                    MembershipRule = $null
                }
                $oldClass = [pscustomobject]@{
                    Id = '55555555-5555-5555-5555-555555555555'
                    DisplayName = 'SEC-A-CLS-G0'
                    GroupTypes = @()
                    MembershipRule = $null
                }
                [void]$global:EntraMutationState.DirectGroupReads.Enqueue(@(
                    $global:EntraMutationGroups.Role
                    $global:EntraMutationGroups.License
                    $global:EntraMutationGroups.Class
                    $oldRole
                    $oldClass
                ))
                [void]$global:EntraMutationState.DirectGroupReads.Enqueue(@(
                    $global:EntraMutationGroups.Role
                    $global:EntraMutationGroups.License
                    $global:EntraMutationGroups.Class
                ))

                $result = Sync-EntraStudentGroup -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -RoleGroup $global:EntraMutationGroups.Role -LicenseGroup $global:EntraMutationGroups.License -ClassGroup $global:EntraMutationGroups.Class -CurrentDirectGroups @() -Confirm:$false

                $result.Verified | Should -BeTrue
                $global:EntraMutationState.Events | Should -Be @(
                    'add-11111111-1111-1111-1111-111111111111'
                    'add-22222222-2222-2222-2222-222222222222'
                    'add-33333333-3333-3333-3333-333333333333'
                    'read-direct-groups'
                    'remove-44444444-4444-4444-4444-444444444444'
                    'remove-55555555-5555-5555-5555-555555555555'
                    'read-direct-groups'
                )
                Should -Invoke New-MgGroupMemberByRef -Times 3 -Exactly -ParameterFilter {
                    $BodyParameter.Count -eq 1 -and
                    $BodyParameter['@odata.id'] -eq 'https://graph.microsoft.com/v1.0/directoryObjects/aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Remove-MgGroupMemberByRef -Times 2 -Exactly -ParameterFilter {
                    $DirectoryObjectId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-MgUserMemberOfAsGroup -Times 2 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $All -and $ErrorAction -eq 'Stop'
                }
            }

            It 'does not remove any old group when a mandatory target cannot be verified' {
                [void]$global:EntraMutationState.DirectGroupReads.Enqueue(@(
                    $global:EntraMutationGroups.Role
                    $global:EntraMutationGroups.License
                ))

                { Sync-EntraStudentGroup -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -RoleGroup $global:EntraMutationGroups.Role -LicenseGroup $global:EntraMutationGroups.License -ClassGroup $global:EntraMutationGroups.Class -CurrentDirectGroups @() -Confirm:$false } |
                    Should -Throw '*mandatory target groups*'

                Should -Invoke Remove-MgGroupMemberByRef -Times 0 -Exactly
                $global:EntraMutationState.Events[-1] | Should -Be 'read-direct-groups'
            }

            It 'does not read or remove old groups when a mandatory add fails' {
                $global:EntraMutationState.GroupAddFailureId = $global:EntraMutationGroups.Class.Id

                { Sync-EntraStudentGroup -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -RoleGroup $global:EntraMutationGroups.Role -LicenseGroup $global:EntraMutationGroups.License -ClassGroup $global:EntraMutationGroups.Class -CurrentDirectGroups @() -Confirm:$false } |
                    Should -Throw '*mandatory group*'

                Should -Invoke Remove-MgGroupMemberByRef -Times 0 -Exactly
                Should -Invoke Get-MgUserMemberOfAsGroup -Times 0 -Exactly
                $global:EntraMutationState.Events | Should -Be @(
                    'add-11111111-1111-1111-1111-111111111111'
                    'add-22222222-2222-2222-2222-222222222222'
                    'add-33333333-3333-3333-3333-333333333333'
                )
            }

            It 'never removes dynamic or inherited competing managed groups' {
                $dynamicRole = [pscustomobject]@{
                    Id = '66666666-6666-6666-6666-666666666666'
                    DisplayName = 'SEC-A-ROL-Dynamic'
                    GroupTypes = @('DynamicMembership')
                    MembershipRule = '(user.department -eq "G1")'
                    IsInherited = $false
                }
                $inheritedClass = [pscustomobject]@{
                    Id = '77777777-7777-7777-7777-777777777777'
                    DisplayName = 'SEC-A-CLS-Inherited'
                    GroupTypes = @()
                    MembershipRule = $null
                    IsInherited = $true
                }
                $groupsWithProtectedCompetitors = @(
                    $global:EntraMutationGroups.Role
                    $global:EntraMutationGroups.License
                    $global:EntraMutationGroups.Class
                    $dynamicRole
                    $inheritedClass
                )
                [void]$global:EntraMutationState.DirectGroupReads.Enqueue($groupsWithProtectedCompetitors)
                [void]$global:EntraMutationState.DirectGroupReads.Enqueue($groupsWithProtectedCompetitors)

                { Sync-EntraStudentGroup -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -RoleGroup $global:EntraMutationGroups.Role -LicenseGroup $global:EntraMutationGroups.License -ClassGroup $global:EntraMutationGroups.Class -CurrentDirectGroups $groupsWithProtectedCompetitors -Confirm:$false } |
                    Should -Throw '*exact managed group state*'

                Should -Invoke Remove-MgGroupMemberByRef -Times 0 -Exactly -ParameterFilter {
                    $GroupId -in @(
                        '66666666-6666-6666-6666-666666666666'
                        '77777777-7777-7777-7777-777777777777'
                    )
                }
            }
        }

        Context 'activation, departure and session revocation' {
            It 'refuses activation until every orchestrator prerequisite is verified' {
                { Enable-EntraStudent -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -WorkbookVerified -GroupsVerified -ManagerVerified -Confirm:$false } |
                    Should -Throw '*verified prerequisites*'

                Should -Invoke Update-MgUser -Times 0 -Exactly
                Should -Invoke Get-MgUser -Times 0 -Exactly
            }

            It 'enables only after all prerequisites and verifies accountEnabled by rereading' {
                [void]$global:EntraMutationState.UserReads.Enqueue([pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    AccountEnabled = $true
                })

                $result = Enable-EntraStudent -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -WorkbookVerified -GroupsVerified -ManagerVerified -AttributesVerified -Confirm:$false

                $result.Verified | Should -BeTrue
                Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $AccountEnabled -eq $true -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $Property -contains 'accountEnabled' -and
                    $ErrorAction -eq 'Stop'
                }
                $global:EntraMutationState.Events | Should -Be @('account-enabled-True', 'read-user')
            }

            It 'fails activation verification when accountEnabled remains false' {
                [void]$global:EntraMutationState.UserReads.Enqueue([pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    AccountEnabled = $false
                })

                { Enable-EntraStudent -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -WorkbookVerified -GroupsVerified -ManagerVerified -AttributesVerified -Confirm:$false } |
                    Should -Throw '*Enabled state could not be verified*'

                Should -Invoke Update-MgUser -Times 1 -Exactly
                Should -Invoke Get-MgUser -Times 1 -Exactly
            }

            It 'disables a departure and verifies accountEnabled by rereading' {
                [void]$global:EntraMutationState.UserReads.Enqueue([pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    AccountEnabled = $false
                })

                $result = Disable-EntraStudent -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Confirm:$false

                $result.Verified | Should -BeTrue
                Should -Invoke Update-MgUser -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and
                    $AccountEnabled -eq $false -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter {
                    $ErrorAction -eq 'Stop'
                }
            }

            It 'fails departure verification when accountEnabled remains true' {
                [void]$global:EntraMutationState.UserReads.Enqueue([pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    AccountEnabled = $true
                })

                { Disable-EntraStudent -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Confirm:$false } |
                    Should -Throw '*Disabled state could not be verified*'

                Should -Invoke Update-MgUser -Times 1 -Exactly
                Should -Invoke Get-MgUser -Times 1 -Exactly
            }

            It 'revokes sessions only when the action was selected and returns the Graph result' {
                $notSelected = Revoke-EntraStudentSession -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Confirm:$false
                $selected = Revoke-EntraStudentSession -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Selected -Confirm:$false

                $notSelected | Should -BeNullOrEmpty
                $selected.Value | Should -BeTrue
                Should -Invoke Revoke-MgUserSignInSession -Times 1 -Exactly -ParameterFilter {
                    $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $ErrorAction -eq 'Stop'
                }
            }
        }

        It 'exposes ShouldProcess on every state-changing helper' {
            foreach ($name in @(
                    'New-DisabledEntraStudent'
                    'New-StudentWithPasswordRetry'
                    'Set-EntraStudentAttribute'
                    'Set-EntraStudentManager'
                    'Sync-EntraStudentGroup'
                    'Enable-EntraStudent'
                    'Disable-EntraStudent'
                    'Revoke-EntraStudentSession'
                )) {
                (Get-Command $name).Parameters.ContainsKey('WhatIf') | Should -BeTrue -Because "$name must support ShouldProcess"
                (Get-Command $name).Parameters.ContainsKey('Confirm') | Should -BeTrue -Because "$name must support ShouldProcess"
            }
        }

        It 'does not call Graph writes under WhatIf' {
            New-DisabledEntraStudent -Desired $global:EntraMutationDesired -Password 'LegoWolke42' -WhatIf

            Should -Invoke New-MgUser -Times 0 -Exactly
        }
    }
}
