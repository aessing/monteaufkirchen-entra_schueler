BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $global:ExchangeTestStubCommands = [Collections.Generic.List[string]]::new()

    if ($null -eq (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        function global:Get-ConnectionInformation {
            [CmdletBinding()]
            param()
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Get-ConnectionInformation')
    }
    if ($null -eq (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        function global:Connect-ExchangeOnline {
            [CmdletBinding()]
            param([bool] $ShowBanner)
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Connect-ExchangeOnline')
    }
    if ($null -eq (Get-Command Get-Recipient -ErrorAction SilentlyContinue)) {
        function global:Get-Recipient {
            [CmdletBinding()]
            param([object] $ResultSize)
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Get-Recipient')
    }
    if ($null -eq (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        function global:Get-Mailbox {
            [CmdletBinding()]
            param([string] $Identity)
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Get-Mailbox')
    }
    if ($null -eq (Get-Command Get-CASMailbox -ErrorAction SilentlyContinue)) {
        function global:Get-CASMailbox {
            [CmdletBinding()]
            param([string] $Identity)
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Get-CASMailbox')
    }
    if ($null -eq (Get-Command Set-Mailbox -ErrorAction SilentlyContinue)) {
        function global:Set-Mailbox {
            [CmdletBinding(SupportsShouldProcess)]
            param(
                [string] $Identity,
                [object] $AddressBookPolicy,
                [string] $CustomAttribute1,
                [bool] $AuditEnabled,
                [object] $AuditLogAgeLimit,
                [object] $RetainDeletedItemsFor,
                [object] $RoleAssignmentPolicy,
                [object] $SharingPolicy,
                [object] $RetentionPolicy,
                [object] $AuditDelegate,
                [object] $AuditOwner,
                [object] $AuditAdmin
            )
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Set-Mailbox')
    }
    if ($null -eq (Get-Command Set-CASMailbox -ErrorAction SilentlyContinue)) {
        function global:Set-CASMailbox {
            [CmdletBinding(SupportsShouldProcess)]
            param(
                [string] $Identity,
                [bool] $ActiveSyncEnabled,
                [bool] $ImapEnabled,
                [bool] $MAPIEnabled,
                [bool] $OWAEnabled,
                [bool] $OWAforDevicesEnabled,
                [object] $OwaMailboxPolicy,
                [bool] $PopEnabled,
                [bool] $SmtpClientAuthenticationDisabled
            )
            throw 'Test stub must be mocked.'
        }
        $global:ExchangeTestStubCommands.Add('Set-CASMailbox')
    }

    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force

    $global:ExchangeTestConfig = @{
        AddressBookPolicy = 'MON-EXO-ABP-Schule_Schüler'
        CustomAttribute1 = 'Montessori Schule Aufkirchen - Schüler'
        AuditLogAgeLimitDays = 365
        RetainDeletedItemsForDays = 30
        RoleAssignmentPolicy = 'MON-EXO-UserRoles-Default'
        SharingPolicy = 'MON-EXO-Sharing-Default'
        RetentionPolicy = 'MON-EXO-Retention-Default'
        OwaMailboxPolicy = 'MON-EXO-OWA-Default'
    }
}

AfterAll {
    foreach ($commandName in @($global:ExchangeTestStubCommands)) {
        Remove-Item -LiteralPath "Function:\global:$commandName" -Force -ErrorAction SilentlyContinue
    }
    Remove-Variable -Name ExchangeTestStubCommands -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name ExchangeTestConfig -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name ExchangeTestState -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Exchange Online student mailbox adapter' {
    BeforeEach {
        $global:ExchangeTestState = @{
            Connections = [Collections.Generic.Queue[object]]::new()
            Mailboxes = [Collections.Generic.Queue[object]]::new()
            CasMailboxes = [Collections.Generic.Queue[object]]::new()
            MailboxAttempts = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
            AvailabilityAt = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
            Sleeps = [Collections.Generic.List[int]]::new()
        }

        Mock Get-ConnectionInformation -ModuleName SchuelerSync {
            if ($global:ExchangeTestState.Connections.Count -gt 0) {
                return $global:ExchangeTestState.Connections.Dequeue()
            }
            return @()
        }
        Mock Connect-ExchangeOnline -ModuleName SchuelerSync {}
        Mock Get-Recipient -ModuleName SchuelerSync { @() }
        Mock Get-Mailbox -ModuleName SchuelerSync {
            if ($global:ExchangeTestState.Mailboxes.Count -gt 0) {
                return $global:ExchangeTestState.Mailboxes.Dequeue()
            }
            return $global:ExchangeTestState.CompliantMailbox
        }
        Mock Get-CASMailbox -ModuleName SchuelerSync {
            if ($global:ExchangeTestState.CasMailboxes.Count -gt 0) {
                return $global:ExchangeTestState.CasMailboxes.Dequeue()
            }
            return $global:ExchangeTestState.CompliantCasMailbox
        }
        Mock -CommandName 'Set-Mailbox' -ModuleName SchuelerSync -MockWith { return }
        Mock -CommandName 'Set-CASMailbox' -ModuleName SchuelerSync -MockWith { return }

        $global:ExchangeTestState.CompliantMailbox = [pscustomobject]@{
            UserPrincipalName = 'mia.muster@monteaufkirchen.com'
            AddressBookPolicy = 'MON-EXO-ABP-Schule_Schüler'
            CustomAttribute1 = 'Montessori Schule Aufkirchen - Schüler'
            AuditEnabled = $true
            AuditLogAgeLimit = [TimeSpan]::FromDays(365)
            RetainDeletedItemsFor = [TimeSpan]::FromDays(30)
            RoleAssignmentPolicy = 'MON-EXO-UserRoles-Default'
            SharingPolicy = 'MON-EXO-Sharing-Default'
            RetentionPolicy = 'MON-EXO-Retention-Default'
            AuditDelegate = @(
                'Create', 'FolderBind', 'HardDelete', 'Move', 'MoveToDeletedItems',
                'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
                'UpdateFolderPermissions', 'UpdateInboxRules', 'ExtraDelegateAction'
            )
            AuditOwner = @(
                'Create', 'HardDelete', 'Move', 'MailboxLogin', 'MoveToDeletedItems',
                'SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateInboxRules',
                'UpdateCalendarDelegation', 'ExtraOwnerAction'
            )
            AuditAdmin = @(
                'Copy', 'Create', 'FolderBind', 'HardDelete', 'Move',
                'MoveToDeletedItems', 'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
                'UpdateFolderPermissions', 'UpdateInboxRules',
                'UpdateCalendarDelegation', 'ExtraAdminAction'
            )
            PersistedCapabilities = @('BPOS_S_Standard')
        }
        $global:ExchangeTestState.CompliantCasMailbox = [pscustomobject]@{
            ActiveSyncEnabled = $false
            ImapEnabled = $false
            MAPIEnabled = $true
            OWAEnabled = $true
            OWAforDevicesEnabled = $false
            OwaMailboxPolicy = 'MON-EXO-OWA-Default'
            PopEnabled = $false
            SmtpClientAuthenticationDisabled = $true
        }
    }

    InModuleScope SchuelerSync {
        Context 'connection and recipient inventory' {
            It 'reuses an active Exchange Online connection' {
                [void]$global:ExchangeTestState.Connections.Enqueue([pscustomobject]@{
                    Name = 'ExchangeOnline'
                    State = 'Connected'
                })

                $connection = Connect-SchuelerExchangeOnline

                $connection.State | Should -Be 'Connected'
                Should -Invoke Connect-ExchangeOnline -Times 0 -Exactly
            }

            It 'connects without a banner and verifies the resulting session' {
                [void]$global:ExchangeTestState.Connections.Enqueue(@())
                [void]$global:ExchangeTestState.Connections.Enqueue([pscustomobject]@{
                    Name = 'ExchangeOnline'
                    State = 'Connected'
                })

                Connect-SchuelerExchangeOnline | Should -Not -BeNullOrEmpty

                Should -Invoke Connect-ExchangeOnline -Times 1 -Exactly -ParameterFilter {
                    $ShowBanner -eq $false -and $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-ConnectionInformation -Times 2 -Exactly
            }

            It 'does not reuse an Exchange Online Protection compliance session' {
                [void]$global:ExchangeTestState.Connections.Enqueue([pscustomobject]@{
                    Name = 'ExchangeOnline'
                    State = 'Connected'
                    IsEopSession = $true
                    ConnectionUri = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                })
                [void]$global:ExchangeTestState.Connections.Enqueue([pscustomobject]@{
                    Name = 'ExchangeOnline'
                    State = 'Connected'
                    IsEopSession = $false
                    ConnectionUri = 'https://outlook.office365.com/powershell-liveid/'
                })

                $connection = Connect-SchuelerExchangeOnline

                $connection.IsEopSession | Should -BeFalse
                Should -Invoke Connect-ExchangeOnline -Times 1 -Exactly
                Should -Invoke Get-ConnectionInformation -Times 2 -Exactly
            }

            It 'fails when the connection does not become active' {
                { Connect-SchuelerExchangeOnline } | Should -Throw '*aktive Exchange-Online-Verbindung*'
                Should -Invoke Connect-ExchangeOnline -Times 1 -Exactly
            }

            It 'indexes every SMTP address and preserves known recipient owners' {
                Mock Get-Recipient -ModuleName SchuelerSync {
                    @(
                        [pscustomobject]@{
                            PrimarySmtpAddress = 'Mia.Muster@monteaufkirchen.com'
                            EmailAddresses = @(
                                'SMTP:Mia.Muster@monteaufkirchen.com'
                                'smtp:m.muster@monteaufkirchen.com'
                                'X500:/o=Example/ou=Exchange Administrative Group/cn=Recipients/cn=mia'
                            )
                            ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                        }
                        [pscustomobject]@{
                            PrimarySmtpAddress = 'Unowned@monteaufkirchen.com'
                            EmailAddresses = @('smtp:alias@monteaufkirchen.com')
                            ExternalDirectoryObjectId = $null
                        }
                    )
                }

                $inventory = Get-ExchangeRecipientAddresses

                @($inventory.ReservedAddresses).Count | Should -Be 4
                $inventory.ReservedAddresses.Contains('MIA.MUSTER@MONTEAUFKIRCHEN.COM') | Should -BeTrue
                $inventory.AddressOwners['m.muster@monteaufkirchen.com'] | Should -Be @('aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa')
                $inventory.AddressOwners.Contains('unowned@monteaufkirchen.com') | Should -BeTrue
                @($inventory.AddressOwners['unowned@monteaufkirchen.com']).Count | Should -Be 0
                $inventory.AddressOwners.Contains('x500:/o=example/ou=exchange administrative group/cn=recipients/cn=mia') | Should -BeFalse
                Should -Invoke Get-Recipient -Times 1 -Exactly -ParameterFilter {
                    $ResultSize -eq 'Unlimited' -and $ErrorAction -eq 'Stop'
                }
            }

            It 'accepts an initially empty reserved-address set for the first recipient address' {
                $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $owners = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)

                Add-ExchangeRecipientAddress `
                    -ReservedAddresses $reserved `
                    -AddressOwners $owners `
                    -Address 'first@monteaufkirchen.com' `
                    -OwnerId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'

                $reserved.Count | Should -Be 1
                $owners['first@monteaufkirchen.com'] | Should -Be @('aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa')
            }
        }

        Context 'mailbox snapshots and comparison' {
            It 'reads the mailbox and CAS state exactly once' {
                $state = Get-StudentMailboxState -UserPrincipalName 'mia.muster@monteaufkirchen.com'

                $state.Exists | Should -BeTrue
                $state.Status | Should -Be 'Ready'
                $state.Differences.Count | Should -Be 0
                Should -Invoke Get-Mailbox -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mia.muster@monteaufkirchen.com' -and $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-CASMailbox -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mia.muster@monteaufkirchen.com' -and $ErrorAction -eq 'Stop'
                }
            }

            It 'classifies only a recipient-not-found error as not ready' {
                Mock Get-Mailbox -ModuleName SchuelerSync {
                    throw [System.Management.Automation.ItemNotFoundException]::new(
                        "The operation couldn't be performed because object 'missing@monteaufkirchen.com' couldn't be found."
                    )
                }

                $state = Get-StudentMailboxState -UserPrincipalName 'missing@monteaufkirchen.com'

                $state.Exists | Should -BeFalse
                $state.Status | Should -Be 'MailboxNotReady'
                Should -Invoke Get-CASMailbox -Times 0 -Exactly
            }

            It 'propagates transport errors instead of treating them as provisioning delay' {
                Mock Get-Mailbox -ModuleName SchuelerSync {
                    throw [System.Net.Http.HttpRequestException]::new('The transport connection was interrupted.')
                }

                { Get-StudentMailboxState -UserPrincipalName 'mia.muster@monteaufkirchen.com' } |
                    Should -Throw '*transport connection*'
                Should -Invoke Get-CASMailbox -Times 0 -Exactly
            }

            It 'reports every drifted mailbox and CAS field separately' {
                $mailbox = [pscustomobject]@{
                    AddressBookPolicy = 'Wrong-ABP'
                    CustomAttribute1 = 'Wrong custom value'
                    AuditEnabled = $false
                    AuditLogAgeLimit = [TimeSpan]::FromDays(90)
                    RetainDeletedItemsFor = [TimeSpan]::FromDays(14)
                    RoleAssignmentPolicy = 'Wrong-Role'
                    SharingPolicy = 'Wrong-Sharing'
                    RetentionPolicy = 'Wrong-Retention'
                    AuditDelegate = @('Create')
                    AuditOwner = @('Create')
                    AuditAdmin = @('Copy')
                    PersistedCapabilities = @('BPOS_S_Standard')
                }
                $cas = [pscustomobject]@{
                    ActiveSyncEnabled = $true
                    ImapEnabled = $true
                    MAPIEnabled = $false
                    OWAEnabled = $false
                    OWAforDevicesEnabled = $true
                    OwaMailboxPolicy = 'Wrong-OWA'
                    PopEnabled = $true
                    SmtpClientAuthenticationDisabled = $false
                }

                $differences = @(Compare-StudentMailboxState -Mailbox $mailbox -CasMailbox $cas -Config $global:ExchangeTestConfig)

                @($differences.Field | Sort-Object) | Should -Be @(
                    'ActiveSyncEnabled', 'AddressBookPolicy', 'AuditAdmin', 'AuditDelegate',
                    'AuditEnabled', 'AuditLogAgeLimit', 'AuditOwner', 'CustomAttribute1',
                    'ImapEnabled', 'MAPIEnabled', 'OWAEnabled', 'OWAforDevicesEnabled',
                    'OwaMailboxPolicy', 'PopEnabled', 'RetainDeletedItemsFor',
                    'RetentionPolicy', 'RoleAssignmentPolicy', 'SharingPolicy',
                    'SmtpClientAuthenticationDisabled'
                )
                @($differences | Where-Object Area -ne 'Exchange').Count | Should -Be 0
            }

            It 'accepts required audit actions as subsets and preserves extras' {
                $differences = @(Compare-StudentMailboxState `
                        -Mailbox $global:ExchangeTestState.CompliantMailbox `
                        -CasMailbox $global:ExchangeTestState.CompliantCasMailbox `
                        -Config $global:ExchangeTestConfig)

                $differences | Should -BeNullOrEmpty
            }

            It 'requires MailItemsAccessed only for CommunicationsCompliance mailboxes' {
                $mailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $mailbox.PersistedCapabilities = @('BPOS_S_Standard', 'CommunicationsCompliance')

                $differences = @(Compare-StudentMailboxState `
                        -Mailbox $mailbox `
                        -CasMailbox $global:ExchangeTestState.CompliantCasMailbox `
                        -Config $global:ExchangeTestConfig)

                $differences.Count | Should -Be 1
                $differences[0].Field | Should -Be 'AuditAdmin'
                @($differences[0].Desired) | Should -Contain 'MailItemsAccessed'

                $mailbox.AuditAdmin += 'MailItemsAccessed'
                @(Compare-StudentMailboxState -Mailbox $mailbox -CasMailbox $global:ExchangeTestState.CompliantCasMailbox -Config $global:ExchangeTestConfig).Count |
                    Should -Be 0
            }
        }

        Context 'idempotent mailbox configuration' {
            It 'performs no writes for a compliant mailbox' {
                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Compliant'
                $result.Changed | Should -BeFalse
                Should -Invoke Set-Mailbox -Times 0 -Exactly
                Should -Invoke Set-CASMailbox -Times 0 -Exactly
            }

            It 'passes the complete exact student mailbox and CAS settings for full drift' {
                $initialMailbox = [pscustomobject]@{
                    AddressBookPolicy = 'Wrong-ABP'
                    CustomAttribute1 = 'Wrong custom value'
                    AuditEnabled = $false
                    AuditLogAgeLimit = [TimeSpan]::FromDays(90)
                    RetainDeletedItemsFor = [TimeSpan]::FromDays(14)
                    RoleAssignmentPolicy = 'Wrong-Role'
                    SharingPolicy = 'Wrong-Sharing'
                    RetentionPolicy = 'Wrong-Retention'
                    AuditDelegate = @()
                    AuditOwner = @()
                    AuditAdmin = @()
                    PersistedCapabilities = @('BPOS_S_Standard')
                }
                $initialCas = [pscustomobject]@{
                    ActiveSyncEnabled = $true
                    ImapEnabled = $true
                    MAPIEnabled = $false
                    OWAEnabled = $false
                    OWAforDevicesEnabled = $true
                    OwaMailboxPolicy = 'Wrong-OWA'
                    PopEnabled = $true
                    SmtpClientAuthenticationDisabled = $false
                }
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)
                [void]$global:ExchangeTestState.CasMailboxes.Enqueue($initialCas)
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($global:ExchangeTestState.CompliantMailbox)
                [void]$global:ExchangeTestState.CasMailboxes.Enqueue($global:ExchangeTestState.CompliantCasMailbox)

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Configured'
                Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mia.muster@monteaufkirchen.com' -and
                    $AddressBookPolicy -eq 'MON-EXO-ABP-Schule_Schüler' -and
                    $CustomAttribute1 -eq 'Montessori Schule Aufkirchen - Schüler' -and
                    $AuditEnabled -eq $true -and
                    $AuditLogAgeLimit.TotalDays -eq 365 -and
                    $RetainDeletedItemsFor.TotalDays -eq 30 -and
                    $RoleAssignmentPolicy -eq 'MON-EXO-UserRoles-Default' -and
                    $SharingPolicy -eq 'MON-EXO-Sharing-Default' -and
                    $RetentionPolicy -eq 'MON-EXO-Retention-Default' -and
                    @($AuditDelegate.Add).Count -eq 11 -and
                    @($AuditOwner.Add).Count -eq 10 -and
                    @($AuditAdmin.Add).Count -eq 13 -and
                    @($AuditAdmin.Add) -notcontains 'MailItemsAccessed' -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Set-CASMailbox -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mia.muster@monteaufkirchen.com' -and
                    $ActiveSyncEnabled -eq $false -and
                    $ImapEnabled -eq $false -and
                    $MAPIEnabled -eq $true -and
                    $OWAEnabled -eq $true -and
                    $OWAforDevicesEnabled -eq $false -and
                    $OwaMailboxPolicy -eq 'MON-EXO-OWA-Default' -and
                    $PopEnabled -eq $false -and
                    $SmtpClientAuthenticationDisabled -eq $true -and
                    $ErrorAction -eq 'Stop'
                }
            }

            It 'writes only drifted fields and missing audit actions, then verifies' {
                $initialMailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $initialMailbox.CustomAttribute1 = 'Wrong'
                $initialMailbox.AuditDelegate = @('Create', 'ExtraDelegateAction')
                $initialCas = $global:ExchangeTestState.CompliantCasMailbox | Select-Object *
                $initialCas.ImapEnabled = $true
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)
                [void]$global:ExchangeTestState.CasMailboxes.Enqueue($initialCas)
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($global:ExchangeTestState.CompliantMailbox)
                [void]$global:ExchangeTestState.CasMailboxes.Enqueue($global:ExchangeTestState.CompliantCasMailbox)

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Configured'
                $result.Changed | Should -BeTrue
                Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mia.muster@monteaufkirchen.com' -and
                    $CustomAttribute1 -eq 'Montessori Schule Aufkirchen - Schüler' -and
                    $null -eq $AddressBookPolicy -and
                    @($AuditDelegate.Add).Count -eq 10 -and
                    @($AuditDelegate.Add) -contains 'FolderBind' -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Set-CASMailbox -Times 1 -Exactly -ParameterFilter {
                    $Identity -eq 'mia.muster@monteaufkirchen.com' -and
                    $ImapEnabled -eq $false -and
                    $null -eq $ActiveSyncEnabled -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Get-Mailbox -Times 2 -Exactly
                Should -Invoke Get-CASMailbox -Times 2 -Exactly
            }

            It 'adds MailItemsAccessed for a CommunicationsCompliance mailbox' {
                $initialMailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $initialMailbox.PersistedCapabilities = @('CommunicationsCompliance')
                $verifiedMailbox = $initialMailbox | Select-Object *
                $verifiedMailbox.AuditAdmin = @($initialMailbox.AuditAdmin) + 'MailItemsAccessed'
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($verifiedMailbox)

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Configured'
                Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter {
                    @($AuditAdmin.Add).Count -eq 1 -and
                    @($AuditAdmin.Add) -contains 'MailItemsAccessed' -and
                    $ErrorAction -eq 'Stop'
                }
                Should -Invoke Set-CASMailbox -Times 0 -Exactly
            }

            It 'preserves the exact Exchange cmdlet error without retrying the write' {
                Mock -CommandName 'Set-Mailbox' -ModuleName SchuelerSync {
                    throw [InvalidOperationException]::new('A parameter cannot be found that matches parameter name AuditLogAgeLimit.')
                }
                $initialMailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $initialMailbox.AuditLogAgeLimit = [TimeSpan]::FromDays(90)
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Failed'
                $result.Error | Should -Be 'A parameter cannot be found that matches parameter name AuditLogAgeLimit.'
                Should -Invoke Set-Mailbox -Times 1 -Exactly
                Should -Invoke Set-CASMailbox -Times 0 -Exactly
                Should -Invoke Get-Mailbox -Times 1 -Exactly
            }

            It 'reports verification drift after writes as a failure' {
                $initialMailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $initialMailbox.CustomAttribute1 = 'Wrong'
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Failed'
                $result.Error | Should -Match 'CustomAttribute1'
                $result.RemainingDifferences.Field | Should -Contain 'CustomAttribute1'
                Should -Invoke Set-Mailbox -Times 1 -Exactly
            }

            It 'fails when the mailbox disappears during post-write verification' {
                $initialMailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $initialMailbox.CustomAttribute1 = 'Wrong'
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)
                Mock Get-Mailbox -ModuleName SchuelerSync {
                    if ($global:ExchangeTestState.Mailboxes.Count -gt 0) {
                        return $global:ExchangeTestState.Mailboxes.Dequeue()
                    }
                    throw [System.Management.Automation.ItemNotFoundException]::new(
                        "The operation couldn't be performed because object 'mia.muster@monteaufkirchen.com' couldn't be found."
                    )
                }

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -Confirm:$false

                $result.Status | Should -Be 'Failed'
                $result.Error | Should -Match 'Verifikation.*Postfach'
                Should -Invoke Set-Mailbox -Times 1 -Exactly
                Should -Invoke Set-CASMailbox -Times 0 -Exactly
                Should -Invoke Get-Mailbox -Times 2 -Exactly
                Should -Invoke Get-CASMailbox -Times 1 -Exactly
            }

            It 'plans drift under WhatIf without writing or rereading' {
                $initialMailbox = $global:ExchangeTestState.CompliantMailbox | Select-Object *
                $initialMailbox.CustomAttribute1 = 'Wrong'
                [void]$global:ExchangeTestState.Mailboxes.Enqueue($initialMailbox)

                $result = Set-StudentMailboxConfiguration `
                    -UserPrincipalName 'mia.muster@monteaufkirchen.com' `
                    -Config $global:ExchangeTestConfig `
                    -WhatIf

                $result.Status | Should -Be 'Planned'
                $result.Changed | Should -BeFalse
                $result.Differences.Field | Should -Contain 'CustomAttribute1'
                Should -Invoke Set-Mailbox -Times 0 -Exactly
                Should -Invoke Set-CASMailbox -Times 0 -Exactly
                Should -Invoke Get-Mailbox -Times 1 -Exactly
                Should -Invoke Get-CASMailbox -Times 1 -Exactly
            }
        }

        Context 'batched provisioning waits' {
            BeforeEach {
                Mock Get-StudentMailboxState -ModuleName SchuelerSync {
                    $current = [int]$global:ExchangeTestState.MailboxAttempts[$UserPrincipalName] + 1
                    $global:ExchangeTestState.MailboxAttempts[$UserPrincipalName] = $current
                    $availableAt = [int]$global:ExchangeTestState.AvailabilityAt[$UserPrincipalName]
                    if ($current -ge $availableAt) {
                        return [pscustomobject]@{
                            UserPrincipalName = $UserPrincipalName
                            Exists = $true
                            Status = 'Ready'
                            Mailbox = $global:ExchangeTestState.CompliantMailbox
                            CasMailbox = $global:ExchangeTestState.CompliantCasMailbox
                            Differences = @()
                        }
                    }
                    return [pscustomobject]@{
                        UserPrincipalName = $UserPrincipalName
                        Exists = $false
                        Status = 'MailboxNotReady'
                        Mailbox = $null
                        CasMailbox = $null
                        Differences = @()
                    }
                }
                Mock Set-StudentMailboxConfiguration -ModuleName SchuelerSync {
                    [pscustomobject]@{
                        UserPrincipalName = $UserPrincipalName
                        Status = 'Compliant'
                        Changed = $false
                        Differences = @()
                    }
                }
            }

            It 'sleeps once per pending batch and checks only remaining users' {
                $upns = @('ready@school.example', 'second@school.example', 'third@school.example')
                $global:ExchangeTestState.AvailabilityAt[$upns[0]] = 1
                $global:ExchangeTestState.AvailabilityAt[$upns[1]] = 2
                $global:ExchangeTestState.AvailabilityAt[$upns[2]] = 3

                $result = Wait-StudentMailboxes `
                    -UserPrincipalName $upns `
                    -MaxRetries 5 `
                    -RetryDelaySeconds 60 `
                    -SleepAction { param($seconds) $global:ExchangeTestState.Sleeps.Add($seconds) }

                @($global:ExchangeTestState.Sleeps) | Should -Be @(60, 60)
                @($result.Ready.UserPrincipalName | Sort-Object) | Should -Be @($upns | Sort-Object)
                $result.Missing.Count | Should -Be 0
                $result.AttemptsByUpn[$upns[0]] | Should -Be 1
                $result.AttemptsByUpn[$upns[1]] | Should -Be 2
                $result.AttemptsByUpn[$upns[2]] | Should -Be 3
                Should -Invoke Get-StudentMailboxState -Times 6 -Exactly
            }

            It 'returns immediately when every mailbox is ready' {
                $upns = @('one@school.example', 'two@school.example')
                foreach ($upn in $upns) { $global:ExchangeTestState.AvailabilityAt[$upn] = 1 }

                $result = Wait-StudentMailboxes -UserPrincipalName $upns -MaxRetries 5 -RetryDelaySeconds 60 `
                    -SleepAction { param($seconds) $global:ExchangeTestState.Sleeps.Add($seconds) }

                $result.Ready.Count | Should -Be 2
                $global:ExchangeTestState.Sleeps.Count | Should -Be 0
                Should -Invoke Get-StudentMailboxState -Times 2 -Exactly
            }

            It 'stops after one immediate attempt plus five retries' {
                $upn = 'missing@school.example'
                $global:ExchangeTestState.AvailabilityAt[$upn] = 99

                $result = Wait-StudentMailboxes -UserPrincipalName $upn -MaxRetries 5 -RetryDelaySeconds 60 `
                    -SleepAction { param($seconds) $global:ExchangeTestState.Sleeps.Add($seconds) }

                $result.Ready.Count | Should -Be 0
                $result.Missing | Should -Be @($upn)
                $result.AttemptsByUpn[$upn] | Should -Be 6
                @($global:ExchangeTestState.Sleeps) | Should -Be @(60, 60, 60, 60, 60)
                Should -Invoke Get-StudentMailboxState -Times 6 -Exactly
            }

            It 'performs only attempt zero under WhatIf and never sleeps or configures' {
                $upn = 'missing@school.example'
                $global:ExchangeTestState.AvailabilityAt[$upn] = 99

                $result = Wait-StudentMailboxes `
                    -UserPrincipalName $upn `
                    -Config $global:ExchangeTestConfig `
                    -Configure `
                    -MaxRetries 5 `
                    -RetryDelaySeconds 60 `
                    -SleepAction { param($seconds) $global:ExchangeTestState.Sleeps.Add($seconds) } `
                    -WhatIf

                $result.Missing | Should -Be @($upn)
                $result.AttemptsByUpn[$upn] | Should -Be 1
                $global:ExchangeTestState.Sleeps.Count | Should -Be 0
                Should -Invoke Get-StudentMailboxState -Times 1 -Exactly
                Should -Invoke Set-StudentMailboxConfiguration -Times 0 -Exactly
                Should -Invoke Set-Mailbox -Times 0 -Exactly
                Should -Invoke Set-CASMailbox -Times 0 -Exactly
            }

            It 'reports a configuration failure without retrying it as mailbox availability' {
                $upn = 'ready@school.example'
                $global:ExchangeTestState.AvailabilityAt[$upn] = 1
                Mock Set-StudentMailboxConfiguration -ModuleName SchuelerSync {
                    [pscustomobject]@{
                        UserPrincipalName = $UserPrincipalName
                        Status = 'Failed'
                        Changed = $false
                        Error = 'Set-Mailbox failed exactly once.'
                        Differences = @()
                    }
                }

                $result = Wait-StudentMailboxes `
                    -UserPrincipalName $upn `
                    -Config $global:ExchangeTestConfig `
                    -Configure `
                    -MaxRetries 5 `
                    -RetryDelaySeconds 60 `
                    -SleepAction { param($seconds) $global:ExchangeTestState.Sleeps.Add($seconds) } `
                    -Confirm:$false

                $result.Failed.Count | Should -Be 1
                $result.Failed[0].Error | Should -Be 'Set-Mailbox failed exactly once.'
                $result.Missing.Count | Should -Be 0
                $global:ExchangeTestState.Sleeps.Count | Should -Be 0
                Should -Invoke Get-StudentMailboxState -Times 1 -Exactly
                Should -Invoke Set-StudentMailboxConfiguration -Times 1 -Exactly
            }
        }
    }
}
