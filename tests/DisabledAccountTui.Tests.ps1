BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $global:DisabledTuiStubCommands = [Collections.Generic.List[string]]::new()
    foreach ($name in @('Get-MgContext', 'Connect-MgGraph', 'Get-MgUser', 'Get-MgUserOwnedObject', 'Get-MgUserTransitiveMemberOf', 'Get-EXORecipient', 'Remove-MgUser', 'Update-MgUser')) {
        if ($null -eq (Get-Command $name -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$name" -Value ([scriptblock]::Create('param($UserId, $ExternalDirectoryObjectId, $Property, $Properties, $ResultSize, [switch]$All, $Confirm, $ErrorAction)'))
            $global:DisabledTuiStubCommands.Add($name)
        }
    }
}

AfterAll {
    foreach ($name in $global:DisabledTuiStubCommands) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
    Remove-Variable DisabledTuiStubCommands -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Gesperrte Konten: Tastaturauswahl' {
    InModuleScope SchuelerSync {
        BeforeEach {
            Mock Assert-ExchangeTenant { }
            $script:tuiAccounts = @(
                [pscustomobject]@{ DisplayName = 'Erste Person'; UPN = 'first@example.org'; UserId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientType = 'UserMailbox'; ProtectionReasons = @(); OwnedObjects = @(); SecAGroups = @(); LastLogin = $null },
                [pscustomobject]@{ DisplayName = 'Zweite Person'; UPN = 'second@example.org'; UserId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; RecipientType = 'UserMailbox'; ProtectionReasons = @(); OwnedObjects = @(); SecAGroups = @(); LastLogin = $null }
            )
            Mock Show-DisabledAccountSelection {}
            Mock Show-DisabledAccountConfirmation {}
            Mock Show-DisabledAccountUnlockConfirmation {}
            Mock Invoke-SelectedDisabledAccountDeletion {
                @($Accounts | ForEach-Object { [pscustomobject]@{ UserId = $_.UserId; Status = 'WhatIf' } })
            }
            Mock Invoke-SelectedDisabledAccountUnlock {
                @($Accounts | ForEach-Object { [pscustomobject]@{ UserId = $_.UserId; Status = 'WhatIf' } })
            }
        }

        It 'returns to the command line on Escape without deleting' {
            Mock Read-DisabledAccountTuiKey { 'Escape' }
            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')
            $result.Count | Should -Be 0
            Should -Invoke Invoke-SelectedDisabledAccountDeletion -Exactly 0
        }

        It 'hides shared mailboxes and selects only the visible personal account' {
            $script:tuiAccounts[0].RecipientType = 'SharedMailbox'
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Spacebar', 'U', 'Y')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }

            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')

            $result.Count | Should -Be 1
            $result[0].UserId | Should -Be $script:tuiAccounts[1].UserId
            Should -Invoke Show-DisabledAccountSelection -ParameterFilter {
                $Accounts.Count -eq 1 -and $Accounts[0].UserId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            } -Exactly 2
            Should -Invoke Invoke-SelectedDisabledAccountUnlock -Exactly 1
        }

        It 'keeps selection after No and deletes only marked accounts after Yes' {
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Spacebar', 'DownArrow', 'Enter', 'N', 'Spacebar', 'Enter', 'Y')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }
            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')
            $result.Count | Should -Be 2
            $result[0].UserId | Should -Be $script:tuiAccounts[0].UserId
            $result[1].UserId | Should -Be $script:tuiAccounts[1].UserId
            Should -Invoke Show-DisabledAccountConfirmation -Exactly 2
            Should -Invoke Invoke-SelectedDisabledAccountDeletion -Exactly 1
        }

        It 'does not open confirmation when no account is selected' {
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Enter', 'Escape')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }
            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')
            $result.Count | Should -Be 0
            Should -Invoke Show-DisabledAccountConfirmation -Exactly 0
        }

        It 'unmarks an account when Space is pressed again' {
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Spacebar', 'Spacebar', 'Enter', 'Escape')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }
            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')
            $result.Count | Should -Be 0
            Should -Invoke Show-DisabledAccountConfirmation -Exactly 0
            Should -Invoke Invoke-SelectedDisabledAccountDeletion -Exactly 0
        }

        It 'allows a confirmed account without an Exchange recipient to be selected for deletion' {
            $script:tuiAccounts[0].RecipientType = ''
            $script:tuiAccounts[0] | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'None'
            $script:tuiAccounts[1].RecipientType = 'SharedMailbox'
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Spacebar', 'Enter', 'Y')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }

            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')
            $result.Count | Should -Be 1
            $result[0].UserId | Should -Be $script:tuiAccounts[0].UserId
            Should -Invoke Show-DisabledAccountSelection -ParameterFilter { $Accounts.Count -eq 1 } -Exactly 2
            Should -Invoke Show-DisabledAccountConfirmation -Exactly 1
            Should -Invoke Invoke-SelectedDisabledAccountDeletion -Exactly 1
            Should -Invoke Invoke-SelectedDisabledAccountUnlock -Exactly 0
        }

        It 'does not select ambiguous or unreadable Exchange recipients' {
            $script:tuiAccounts[0].RecipientType = ''
            $script:tuiAccounts[0] | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'Ambiguous'
            $script:tuiAccounts[1].RecipientType = ''
            $script:tuiAccounts[1] | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'Unreadable'
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Spacebar', 'DownArrow', 'Spacebar', 'Enter', 'Escape')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }

            @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test').Count | Should -Be 0
            Should -Invoke Show-DisabledAccountConfirmation -Exactly 0
            Should -Invoke Invoke-SelectedDisabledAccountDeletion -Exactly 0
        }

        It 'confirms unlocking separately and unlocks only marked accounts after Yes' {
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('Spacebar', 'DownArrow', 'U', 'N', 'Spacebar', 'U', 'Y')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }

            $result = @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test')

            $result.Count | Should -Be 2
            Should -Invoke Show-DisabledAccountUnlockConfirmation -Exactly 2
            Should -Invoke Invoke-SelectedDisabledAccountUnlock -Exactly 1
            Should -Invoke Invoke-SelectedDisabledAccountDeletion -Exactly 0
        }

        It 'does not open unlock confirmation without a marked account' {
            $script:keys = [Collections.Generic.Queue[string]]::new()
            foreach ($key in @('U', 'Escape')) { $script:keys.Enqueue($key) }
            Mock Read-DisabledAccountTuiKey { $script:keys.Dequeue() }

            @(Invoke-DisabledAccountTui -Accounts $script:tuiAccounts -TenantId 'tenant-test').Count | Should -Be 0
            Should -Invoke Show-DisabledAccountUnlockConfirmation -Exactly 0
            Should -Invoke Invoke-SelectedDisabledAccountUnlock -Exactly 0
        }
    }
}

Describe 'Gesperrte Konten: Löschung nach erneuter Prüfung' {
    InModuleScope SchuelerSync {
        BeforeEach {
            Mock Assert-ExchangeTenant { }
            $script:account = [pscustomobject]@{
                DisplayName = 'Test Person'; UPN = 'test@example.org'; UserId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                RecipientType = 'UserMailbox'; ProtectionReasons = @(); OwnedObjects = @()
            }
            Mock Connect-DisabledAccountDeleteGraph { [pscustomobject]@{ TenantId = 'tenant-test' } }
            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            Mock Get-EXORecipient { [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'UserMailbox' } }
            Mock Get-MgUserOwnedObject { @() }
            Mock Get-MgUserTransitiveMemberOf { @() }
            Mock Remove-MgUser {}
            Mock Write-Progress {}
            Mock Write-Warning {}
        }

        It 'shows owned object details in the deletion confirmation' {
            $script:account.OwnedObjects = @([pscustomobject]@{ Type = 'group'; DisplayName = 'Schulteam'; Id = 'group-1' })
            $script:confirmationLines = [Collections.Generic.List[string]]::new()
            Mock Clear-Host {}
            Mock Write-Host { $script:confirmationLines.Add([string]$Object) }

            Show-DisabledAccountConfirmation -Accounts @($script:account) -TenantId 'tenant-test'

            $text = $script:confirmationLines -join "`n"
            $text | Should -BeLike '*Test Person <test@example.org>*'
            $text | Should -BeLike '*group: Schulteam (group-1)*'
            $text | Should -BeLike '*ohne Eigentümer*'
        }

        It 'shows progress for each selected deletion and completes it' {
            $second = [pscustomobject]@{
                DisplayName = 'Second Person'; UPN = 'second@example.org'; UserId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
                RecipientType = 'UserMailbox'; ProtectionReasons = @(); OwnedObjects = @()
            }
            Mock Get-MgUser {
                $name = if ($UserId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb') { 'Second Person' } else { 'Test Person' }
                $upn = if ($UserId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb') { 'second@example.org' } else { 'test@example.org' }
                [pscustomobject]@{ Id = $UserId; DisplayName = $name; UserPrincipalName = $upn
                    AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            Mock Get-EXORecipient { [pscustomobject]@{ ExternalDirectoryObjectId = [string]$ExternalDirectoryObjectId; RecipientTypeDetails = 'UserMailbox' } }

            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account, $second) -TenantId 'tenant-test' -WhatIf -Confirm:$false)

            $result.Count | Should -Be 2
            Should -Invoke Write-Progress -ParameterFilter { $Status -like '*1 von 2*' } -Exactly 1
            Should -Invoke Write-Progress -ParameterFilter { $Status -like '*2 von 2*' } -Exactly 1
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'blocks deletion if Exchange switches tenant after selection' {
            Mock Assert-ExchangeTenant { throw 'Exchange-Tenant mismatch' }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be Skipped
            Should -Invoke Assert-ExchangeTenant -ParameterFilter { $TenantId -eq 'tenant-test' } -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Get-EXORecipient -Exactly 0
        }

        It 'deletes a rechecked disabled personal account by immutable ID' {
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Deleted'
            Should -Invoke Remove-MgUser -ParameterFilter { $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' } -Exactly 1
        }

        It 'skips an account reactivated after selection' {
            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $true; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
        }

        It 'skips a shared mailbox found during recheck' {
            Mock Get-EXORecipient { [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'SharedMailbox' } }
            $shared = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $shared[0].Status | Should -Be 'Skipped'
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'warns and deletes an account that owns Entra objects' {
            Mock Get-MgUserOwnedObject { [pscustomobject]@{ Id = 'owned-object' } }
            $owned = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $owned[0].Status | Should -Be 'Deleted'
            Should -Invoke Write-Warning -ParameterFilter { $Message -like '*1 Entra-Objekt*' } -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 1
        }

        It 'skips deletion when the ownership check fails' {
            Mock Get-MgUserOwnedObject { throw 'Graph ownership unavailable' }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            $result[0].Message | Should -BeLike '*Graph ownership unavailable*'
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'deletes an account whose missing Exchange recipient is reconfirmed by a fresh inventory' {
            $script:account.RecipientType = ''
            $script:account | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'None'
            Mock Get-EXORecipient { if ($ResultSize -eq 'Unlimited') { return @() }; throw 'Unexpected recipient lookup' }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Deleted'
            Should -Invoke Get-EXORecipient -ParameterFilter { $ResultSize -eq 'Unlimited' } -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 1
        }

        It 'skips a newly associated shared mailbox for an account previously without a recipient' {
            $script:account.RecipientType = ''
            $script:account | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'None'
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = ' AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA '; RecipientTypeDetails = 'SharedMailbox' }
            }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'does not delete an account with an ambiguous Exchange mapping even through a direct action call' {
            $script:account.RecipientType = ''
            $script:account | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'Ambiguous'
            Mock Get-EXORecipient { @() }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Get-EXORecipient -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'skips a no-recipient deletion if the fresh Exchange inventory fails' {
            $script:account.RecipientType = ''
            $script:account | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'None'
            Mock Get-EXORecipient { throw 'Exchange unavailable' }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            $result[0].Message | Should -BeLike '*Exchange unavailable*'
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'skips guests, synced users and current directory roles' {
            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $false; UserType = 'Guest'; OnPremisesSyncEnabled = $false }
            }
            @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)[0].Status | Should -Be 'Skipped'

            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $true }
            }
            @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)[0].Status | Should -Be 'Skipped'

            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            Mock Get-MgUserTransitiveMemberOf { [pscustomobject]@{ OdataType = '#microsoft.graph.directoryRole' } }
            @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)[0].Status | Should -Be 'Skipped'
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'skips an unreadable membership type' {
            Mock Get-MgUserTransitiveMemberOf { [pscustomobject]@{ Id = 'unknown-group' } }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            $result[0].Message | Should -BeLike '*nicht lesbar*'
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'reports Graph deletion errors separately from skipped validation' {
            Mock Remove-MgUser { throw 'Graph delete failed' }
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Failed'
            $result[0].Message | Should -BeLike '*Graph delete failed*'
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
        }

        It 'never deletes in WhatIf mode' {
            $result = @(Invoke-SelectedDisabledAccountDeletion -Accounts @($script:account) -TenantId 'tenant-test' -WhatIf -Confirm:$false)
            $result[0].Status | Should -Be 'WhatIf'
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Connect-DisabledAccountDeleteGraph -Exactly 0
        }
    }
}

Describe 'Gesperrte Konten: Mandantenbindung' {
    InModuleScope SchuelerSync {
        It 'rejects a different tenant before deleting' {
            Mock Get-MgContext {
                [pscustomobject]@{ TenantId = 'other-tenant'; Scopes = @('User.ReadWrite.All', 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            }
            { Connect-DisabledAccountDeleteGraph -TenantId 'tenant-test' } | Should -Throw '*Tenant hat sich geändert*'
        }

        It 'requests write scope only for a confirmed deletion and verifies it' {
            $script:deleteContext = [pscustomobject]@{ TenantId = 'tenant-test'; Scopes = @('User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            Mock Get-MgContext { $script:deleteContext }
            Mock Connect-MgGraph {
                $script:deleteContext = [pscustomobject]@{ TenantId = 'tenant-test'; Scopes = @('User.ReadWrite.All', 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            }
            $result = Connect-DisabledAccountDeleteGraph -TenantId 'tenant-test'
            $result.TenantId | Should -Be 'tenant-test'
            Should -Invoke Connect-MgGraph -Exactly 1
        }

        It 'stops when the refreshed Graph context still lacks write scope' {
            Mock Get-MgContext {
                [pscustomobject]@{ TenantId = 'tenant-test'; Scopes = @('User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            }
            Mock Connect-MgGraph {}
            { Connect-DisabledAccountDeleteGraph -TenantId 'tenant-test' } | Should -Throw '*User.ReadWrite.All*'
        }

        It 'requests the enable-disable scope only for a confirmed unlock' {
            $script:unlockContext = [pscustomobject]@{ TenantId = 'tenant-test'; Scopes = @('User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            Mock Get-MgContext { $script:unlockContext }
            Mock Connect-MgGraph {
                $script:unlockContext = [pscustomobject]@{ TenantId = 'tenant-test'; Scopes = @('User.EnableDisableAccount.All', 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            }
            $result = Connect-DisabledAccountWriteGraph -TenantId 'tenant-test' -Operation Unlock
            $result.TenantId | Should -Be 'tenant-test'
            Should -Invoke Connect-MgGraph -ParameterFilter {
                $Scopes -contains 'User.EnableDisableAccount.All' -and $Scopes -notcontains 'User.ReadWrite.All'
            } -Exactly 1
        }

        It 'rejects a changed tenant before unlocking' {
            Mock Get-MgContext {
                [pscustomobject]@{ TenantId = 'other-tenant'; Scopes = @('User.EnableDisableAccount.All', 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All') }
            }
            { Connect-DisabledAccountWriteGraph -TenantId 'tenant-test' -Operation Unlock } | Should -Throw '*Tenant hat sich geändert*'
        }
    }
}

Describe 'Gesperrte Konten: Entsperren nach erneuter Prüfung' {
    InModuleScope SchuelerSync {
        BeforeEach {
            Mock Assert-ExchangeTenant { }
            $script:account = [pscustomobject]@{
                DisplayName = 'Test Person'; UPN = 'test@example.org'; UserId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                RecipientType = 'UserMailbox'; ProtectionReasons = @(); OwnedObjects = @()
            }
            Mock Connect-DisabledAccountWriteGraph { [pscustomobject]@{ TenantId = 'tenant-test' } }
            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            Mock Get-EXORecipient { [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'UserMailbox' } }
            Mock Get-MgUserOwnedObject { @() }
            Mock Get-MgUserTransitiveMemberOf { @() }
            Mock Update-MgUser {}
            Mock Remove-MgUser { throw 'Löschen darf nicht aufgerufen werden.' }
            Mock Write-Progress {}
        }

        It 'shows progress for each selected unlock and completes it' {
            $second = [pscustomobject]@{
                DisplayName = 'Second Person'; UPN = 'second@example.org'; UserId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
                RecipientType = 'UserMailbox'; ProtectionReasons = @(); OwnedObjects = @()
            }
            Mock Get-MgUser {
                $name = if ($UserId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb') { 'Second Person' } else { 'Test Person' }
                $upn = if ($UserId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb') { 'second@example.org' } else { 'test@example.org' }
                [pscustomobject]@{ Id = $UserId; DisplayName = $name; UserPrincipalName = $upn
                    AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            Mock Get-EXORecipient { [pscustomobject]@{ ExternalDirectoryObjectId = [string]$ExternalDirectoryObjectId; RecipientTypeDetails = 'UserMailbox' } }

            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account, $second) -TenantId 'tenant-test' -WhatIf -Confirm:$false)

            $result.Count | Should -Be 2
            Should -Invoke Write-Progress -ParameterFilter { $Status -like '*1 von 2*' } -Exactly 1
            Should -Invoke Write-Progress -ParameterFilter { $Status -like '*2 von 2*' } -Exactly 1
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
            Should -Invoke Update-MgUser -Exactly 0
        }

        It 'enables a rechecked disabled user by immutable ID without deleting it' {
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Enabled'
            Should -Invoke Update-MgUser -ParameterFilter {
                $UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $AccountEnabled -eq $true
            } -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 0
        }

        It 'skips a user already enabled after selection' {
            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Test Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $true; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
        }

        It 'skips a mailbox converted to shared after selection' {
            Mock Get-EXORecipient { [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'SharedMailbox' } }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Update-MgUser -Exactly 0
        }

        It 'skips an account without an Exchange recipient even through a direct action call' {
            Mock Get-EXORecipient { @() }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Update-MgUser -Exactly 0
        }

        It 'enables an account whose missing Exchange recipient is reconfirmed by a fresh inventory' {
            $script:account.RecipientType = ''
            $script:account | Add-Member -NotePropertyName RecipientStatus -NotePropertyValue 'None'
            Mock Get-EXORecipient { if ($ResultSize -eq 'Unlimited') { return @() }; throw 'Unexpected recipient lookup' }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Enabled'
            Should -Invoke Get-EXORecipient -ParameterFilter { $ResultSize -eq 'Unlimited' } -Exactly 1
            Should -Invoke Update-MgUser -Exactly 1
        }

        It 'enables a disabled account that owns Entra objects without querying ownership again' {
            $script:account.OwnedObjects = @(1..7 | ForEach-Object { [pscustomobject]@{ Id = "owned-$_" } })
            Mock Get-MgUserOwnedObject { throw 'Besitzabfrage darf Entsperren nicht blockieren.' }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Enabled'
            Should -Invoke Get-MgUserOwnedObject -Exactly 0
            Should -Invoke Update-MgUser -Exactly 1
        }

        It 'still skips current directory roles when unlocking' {
            Mock Get-MgUserTransitiveMemberOf { [pscustomobject]@{ OdataType = '#microsoft.graph.directoryRole' } }
            @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)[0].Status | Should -Be 'Skipped'
            Should -Invoke Update-MgUser -Exactly 0
        }

        It 'skips a changed identity after selection' {
            Mock Get-MgUser {
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Other Person'
                    UserPrincipalName = 'test@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false }
            }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Skipped'
            Should -Invoke Update-MgUser -Exactly 0
        }

        It 'reports Graph enable errors without changing other resources' {
            Mock Update-MgUser { throw 'Graph enable failed' }
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -Confirm:$false)
            $result[0].Status | Should -Be 'Failed'
            $result[0].Message | Should -BeLike '*Graph enable failed*'
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
        }

        It 'never enables in WhatIf mode or requests write scope' {
            $result = @(Invoke-SelectedDisabledAccountUnlock -Accounts @($script:account) -TenantId 'tenant-test' -WhatIf -Confirm:$false)
            $result[0].Status | Should -Be 'WhatIf'
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Connect-DisabledAccountWriteGraph -Exactly 0
        }
    }
}
