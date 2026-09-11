BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $global:EntraTestConfig = @{
        Domain = 'monteaufkirchen.com'
        StudentRoleGroup = @{
            Id = '11111111-1111-1111-1111-111111111111'
            Name = 'SEC-A-ROL-Schule_Schüler'
        }
        LicenseGroupName = 'SEC-A-LIC-O365A1Student'
        ClassGroupPrefix = 'SEC-A-CLS-'
    }
}

AfterAll {
    Remove-Variable -Name EntraTestConfig -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name EntraTestGraphContexts -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Microsoft Graph inventory and manager resolution' {
    BeforeEach {
        Mock Connect-MgGraph -ModuleName SchuelerSync { return }
        $global:EntraTestGraphContexts = [Collections.Generic.Queue[object]]::new()
        Mock Get-MgContext -ModuleName SchuelerSync {
            if ($global:EntraTestGraphContexts.Count -gt 0) {
                return $global:EntraTestGraphContexts.Dequeue()
            }
            [pscustomobject]@{
                Scopes = @(
                    'User.ReadWrite.All'
                    'User-Mail.ReadWrite.All'
                    'Group.Read.All'
                    'GroupMember.ReadWrite.All'
                    'User.RevokeSessions.All'
                )
            }
        }
        Mock Get-MgUser -ModuleName SchuelerSync {
            @(
                [pscustomobject]@{
                    Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                    DisplayName = 'Max Muster'
                    GivenName = 'Max'
                    Surname = 'Muster'
                    UserPrincipalName = 'MMuster@monteaufkirchen.com'
                    Mail = 'Max.Muster@monteaufkirchen.com'
                    ProxyAddresses = @('SMTP:MMuster@monteaufkirchen.com', 'smtp:max.muster@schule.example')
                    AccountEnabled = $true
                }
                [pscustomobject]@{
                    Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
                    DisplayName = 'Lea Lehrerin'
                    GivenName = 'Lea'
                    Surname = 'Lehrerin'
                    UserPrincipalName = 'lea.lehrerin@monteaufkirchen.com'
                    Mail = 'lea.lehrerin@schule.example'
                    ProxyAddresses = @()
                    AccountEnabled = $true
                }
            )
        }
        Mock Get-MgGroup -ModuleName SchuelerSync {
            @(
                [pscustomobject]@{
                    Id = '11111111-1111-1111-1111-111111111111'
                    DisplayName = 'SEC-A-ROL-Schule_Schüler'
                    GroupTypes = @()
                    MembershipRule = $null
                    MembershipRuleProcessingState = $null
                }
                [pscustomobject]@{
                    Id = '22222222-2222-2222-2222-222222222222'
                    DisplayName = 'SEC-A-LIC-O365A1Student'
                    GroupTypes = @('DynamicMembership')
                    MembershipRule = '(user.department -eq "G1")'
                    MembershipRuleProcessingState = 'On'
                }
                [pscustomobject]@{
                    Id = '33333333-3333-3333-3333-333333333333'
                    DisplayName = 'SEC-A-CLS-G1'
                    GroupTypes = @()
                    MembershipRule = $null
                    MembershipRuleProcessingState = $null
                }
            )
        }
        Mock Get-MgGroupMember -ModuleName SchuelerSync {
            @([pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' })
        }
        Mock Get-MgUserMemberOfAsGroup -ModuleName SchuelerSync {
            @(
                [pscustomobject]@{
                    Id = '33333333-3333-3333-3333-333333333333'
                    DisplayName = 'SEC-A-CLS-G1'
                    GroupTypes = @()
                    MembershipRule = $null
                    MembershipRuleProcessingState = $null
                }
            )
        }
        Mock Get-MgUserTransitiveMemberOfAsGroup -ModuleName SchuelerSync {
            @(
                [pscustomobject]@{
                    Id = '33333333-3333-3333-3333-333333333333'
                    DisplayName = 'SEC-A-CLS-G1'
                    GroupTypes = @()
                    MembershipRule = $null
                    MembershipRuleProcessingState = $null
                }
                [pscustomobject]@{
                    Id = '22222222-2222-2222-2222-222222222222'
                    DisplayName = 'SEC-A-LIC-O365A1Student'
                    GroupTypes = @('DynamicMembership')
                    MembershipRule = '(user.department -eq "G1")'
                    MembershipRuleProcessingState = 'On'
                }
            )
        }
        Mock Get-MgUserManager -ModuleName SchuelerSync {
            [pscustomobject]@{ Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb' }
        }
    }

    InModuleScope SchuelerSync {
        It 'builds exact null-free group and address indexes from the real snapshot adapter' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig
            $snapshot.GroupsById.Count | Should -Be 3
            $snapshot.GroupsByDisplayName.Count | Should -Be 3
            $snapshot.GroupMatchesByDisplayName.Count | Should -Be 3
            foreach ($name in @('SEC-A-ROL-Schule_Schüler', 'SEC-A-LIC-O365A1Student', 'SEC-A-CLS-G1')) {
                $matches = @($snapshot.GroupMatchesByDisplayName[$name])
                $matches.Count | Should -Be 1
                $matches[0] | Should -Not -BeNullOrEmpty
                $snapshot.GroupsByDisplayName[$name].Id | Should -Be $matches[0].Id
            }
            $snapshot.ReservedAddresses.Count | Should -Be 5
            $snapshot.AddressOwners.Count | Should -Be 5
            foreach ($address in $snapshot.AddressOwners.Keys) {
                @($snapshot.AddressOwners[$address]).Count | Should -Be 1
                $snapshot.AddressOwners[$address][0] | Should -Not -BeNullOrEmpty
            }
            { Resolve-EntraSnapshotGroup -Snapshot $snapshot -DisplayName 'missing' } | Should -Throw '*found 0*'
        }

        It 'adds the first Entra address to empty accumulators and deduplicates its owner' {
            $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $owners = @{}
            Add-EntraReservedAddress -ReservedAddresses $reserved -AddressOwners $owners -Address 'FIRST@school.example' -OwnerId 'student-1'
            Add-EntraReservedAddress -ReservedAddresses $reserved -AddressOwners $owners -Address 'first@school.example' -OwnerId 'STUDENT-1'
            $reserved.Count | Should -Be 1
            @($owners['first@school.example']).Count | Should -Be 1
            $owners['first@school.example'][0] | Should -Be 'student-1'
        }

        It 'caches a missing manager for Graph exception shape <Shape>' -ForEach @(
            @{ Shape = 'StatusCode' }, @{ Shape = 'ResponseStatusCode' }, @{ Shape = 'Response' }
        ) {
            $script:missingManagerShape = $Shape
            Mock Get-MgUserManager {
                $exception = [Exception]::new('The requested manager reference does not exist.')
                if ($script:missingManagerShape -eq 'Response') {
                    $exception | Add-Member NoteProperty Response ([pscustomobject]@{ StatusCode = [Net.HttpStatusCode]::NotFound })
                } else {
                    $exception | Add-Member NoteProperty $script:missingManagerShape 404
                }
                $record = [Management.Automation.ErrorRecord]::new($exception, 'Request_ResourceNotFound', [Management.Automation.ErrorCategory]::ObjectNotFound, $UserId)
                $record.ErrorDetails = [Management.Automation.ErrorDetails]::new('{"error":{"code":"Request_ResourceNotFound","message":"No manager reference exists."}}')
                throw $record
            }
            $snapshot = [pscustomobject]@{ ManagerByUserId = @{} }
            Get-UserManagerId -Snapshot $snapshot -UserId 'student-1' | Should -BeNullOrEmpty
            Get-UserManagerId -Snapshot $snapshot -UserId 'student-1' | Should -BeNullOrEmpty
            $snapshot.ManagerByUserId.ContainsKey('student-1') | Should -BeTrue
            Should -Invoke Get-MgUserManager -Times 1 -Exactly -ParameterFilter { $ErrorAction -eq 'Stop' }
        }

        It 'does not hide or cache a manager authorization or transport failure: <Shape>' -ForEach @(
            @{ Shape = 'StatusCode' }, @{ Shape = 'ResponseStatusCode' }, @{ Shape = 'Response' }, @{ Shape = 'Transport' }
        ) {
            $script:managerFailureShape = $Shape
            Mock Get-MgUserManager {
                $exception = [Exception]::new('Manager lookup failed, preserve this error.')
                if ($script:managerFailureShape -eq 'Response') {
                    $exception | Add-Member NoteProperty Response ([pscustomobject]@{ StatusCode = [Net.HttpStatusCode]::Forbidden })
                } elseif ($script:managerFailureShape -ne 'Transport') {
                    $exception | Add-Member NoteProperty $script:managerFailureShape 403
                }
                throw $exception
            }
            $snapshot = [pscustomobject]@{ ManagerByUserId = @{} }
            { Get-UserManagerId -Snapshot $snapshot -UserId 'student-1' } | Should -Throw '*preserve this error*'
            $snapshot.ManagerByUserId.ContainsKey('student-1') | Should -BeFalse
        }

        It 'does not interpret an empty successful response as a documented missing manager' {
            Mock Get-MgUserManager { $null }
            $snapshot = [pscustomobject]@{ ManagerByUserId = @{} }
            { Get-UserManagerId -Snapshot $snapshot -UserId 'student-1' } | Should -Throw '*Manager*'
            $snapshot.ManagerByUserId.ContainsKey('student-1') | Should -BeFalse
        }

        It 'reuses a complete existing Graph context with scope casing differences' {
            [void]$global:EntraTestGraphContexts.Enqueue([pscustomobject]@{
                Scopes = @(
                    'user.readwrite.all'
                    'USER-MAIL.READWRITE.ALL'
                    'group.read.all'
                    'GROUPMEMBER.READWRITE.ALL'
                    'user.revokesessions.all'
                )
            })

            $context = Connect-SchuelerGraph

            $context.Scopes.Count | Should -Be 5
            Should -Invoke Connect-MgGraph -Times 0 -Exactly
            Should -Invoke Get-MgContext -Times 1 -Exactly
        }

        It 'connects with the complete required scope set when the existing context is incomplete' {
            [void]$global:EntraTestGraphContexts.Enqueue([pscustomobject]@{
                Scopes = @('User.ReadWrite.All')
            })
            [void]$global:EntraTestGraphContexts.Enqueue([pscustomobject]@{
                Scopes = @(
                    'User.ReadWrite.All'
                    'User-Mail.ReadWrite.All'
                    'Group.Read.All'
                    'GroupMember.ReadWrite.All'
                    'User.RevokeSessions.All'
                )
            })

            Connect-SchuelerGraph | Should -Not -BeNullOrEmpty

            Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
                $requiredScopes = @(
                    'User.ReadWrite.All'
                    'User-Mail.ReadWrite.All'
                    'Group.Read.All'
                    'GroupMember.ReadWrite.All'
                    'User.RevokeSessions.All'
                )
                $NoWelcome -and $Scopes.Count -eq $requiredScopes.Count -and
                @($requiredScopes | Where-Object { $Scopes -notcontains $_ }).Count -eq 0
            }
            Should -Invoke Get-MgContext -Times 2 -Exactly
        }

        It 'reports missing scopes when the post-connect context remains incomplete' {
            [void]$global:EntraTestGraphContexts.Enqueue([pscustomobject]@{
                Scopes = @('User.ReadWrite.All')
            })
            [void]$global:EntraTestGraphContexts.Enqueue([pscustomobject]@{
                Scopes = @(
                    'User.ReadWrite.All'
                    'Group.Read.All'
                    'User.RevokeSessions.All'
                )
            })

            { Connect-SchuelerGraph } |
                Should -Throw '*User-Mail.ReadWrite.All*GroupMember.ReadWrite.All*'
            Should -Invoke Connect-MgGraph -Times 1 -Exactly
        }

        It 'requests the complete user property set once' {
            Get-EntraSnapshot -Config $global:EntraTestConfig

            Should -Invoke Get-MgUser -Times 1 -Exactly -ParameterFilter {
                $requiredProperties = @(
                    'id', 'displayName', 'givenName', 'surname', 'userPrincipalName', 'mail', 'mailNickname', 'proxyAddresses',
                    'department', 'officeLocation', 'companyName', 'employeeType', 'usageLocation', 'ageGroup',
                    'consentProvidedForMinor', 'legalAgeGroupClassification', 'accountEnabled', 'userType'
                )
                $All -and @($requiredProperties | Where-Object { $Property -notcontains $_ }).Count -eq 0
            }
            Should -Invoke Get-MgGroup -Times 1 -Exactly -ParameterFilter {
                $requiredProperties = @(
                    'id', 'displayName', 'groupTypes', 'membershipRule', 'membershipRuleProcessingState'
                )
                $All -and @($requiredProperties | Where-Object { $Property -notcontains $_ }).Count -eq 0
            }
            Should -Invoke Get-MgGroupMember -Times 1 -Exactly -ParameterFilter {
                $All -and $GroupId -eq $global:EntraTestConfig.StudentRoleGroup.Id
            }
        }

        It 'reserves UPN, mail and SMTP proxy values case-insensitively' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig

            $snapshot.ReservedAddresses.Contains('mmuster@monteaufkirchen.com') | Should -BeTrue
            $snapshot.ReservedAddresses.Contains('MAX.MUSTER@SCHULE.EXAMPLE') | Should -BeTrue
            $snapshot.AddressOwners['mmuster@monteaufkirchen.com'] | Should -Be @('aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa')
            $snapshot.UsersByUpn['MMUSTER@MONTEAUFKIRCHEN.COM'].Id | Should -Be 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        }

        It 'keeps dynamic metadata and marks inherited transitive memberships non-removable' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig
            $direct = @(Get-UserDirectGroup -Snapshot $snapshot -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa')
            $transitive = @(Get-UserTransitiveGroup -Snapshot $snapshot -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa')
            $inherited = $transitive | Where-Object Id -eq '22222222-2222-2222-2222-222222222222'

            $direct.Count | Should -Be 1
            $inherited.IsDynamic | Should -BeTrue
            $inherited.IsInherited | Should -BeTrue
            $inherited.IsRemovable | Should -BeFalse
            Get-UserTransitiveGroup -Snapshot $snapshot -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            Should -Invoke Get-MgUserTransitiveMemberOfAsGroup -Times 1 -Exactly
        }

        It 'caches the manager identifier per user' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig

            Get-UserManagerId -Snapshot $snapshot -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' |
                Should -Be 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            Get-UserManagerId -Snapshot $snapshot -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' |
                Should -Be 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            Should -Invoke Get-MgUserManager -Times 1 -Exactly
        }

        It 'resolves exactly one teacher by UPN' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig

            (Resolve-StudentManager -Snapshot $snapshot -Teacher ' LEA.LEHRERIN@MONTEAUFKIRCHEN.COM ').Id |
                Should -Be 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
        }

        It 'resolves exactly one teacher by mail' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig

            (Resolve-StudentManager -Snapshot $snapshot -Teacher 'lea.lehrerin@schule.example').Id |
                Should -Be 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
        }

        It 'resolves exactly one teacher by exact case-insensitive display name' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig

            (Resolve-StudentManager -Snapshot $snapshot -Teacher 'lea lehrerin').Id |
                Should -Be 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
        }

        It 'rejects a missing teacher' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig

            { Resolve-StudentManager -Snapshot $snapshot -Teacher 'nicht vorhanden' } |
                Should -Throw '*Kein eindeutiger Manager*'
        }

        It 'rejects ambiguous teacher display names' {
            $snapshot = Get-EntraSnapshot -Config $global:EntraTestConfig
            $snapshot.UsersById['cccccccc-cccc-cccc-cccc-cccccccccccc'] = [pscustomobject]@{
                Id = 'cccccccc-cccc-cccc-cccc-cccccccccccc'
                DisplayName = 'Lea Lehrerin'
                UserPrincipalName = 'zweite.lea@monteaufkirchen.com'
                Mail = 'zweite.lea@schule.example'
            }

            { Resolve-StudentManager -Snapshot $snapshot -Teacher 'Lea Lehrerin' } |
                Should -Throw '*Kein eindeutiger Manager*'
        }
    }
}
