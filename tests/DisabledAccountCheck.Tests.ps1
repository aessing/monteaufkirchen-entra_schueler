BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $global:DisabledCheckStubCommands = [Collections.Generic.List[string]]::new()
    foreach ($name in @('Get-MgUser', 'Get-MgUserTransitiveMemberOf', 'Get-MgUserOwnedObject', 'Get-EXORecipient')) {
        if ($null -eq (Get-Command $name -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$name" -Value ([scriptblock]::Create('param($UserId, $ExternalDirectoryObjectId, [switch]$All, $Filter, $Property, $Properties, $ResultSize, $ErrorAction)'))
            $global:DisabledCheckStubCommands.Add($name)
        }
    }
}

AfterAll {
    foreach ($name in $global:DisabledCheckStubCommands) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
    Remove-Variable DisabledCheckStubCommands -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Gesperrte Konten: Schutzprüfung' {
    InModuleScope SchuelerSync {
        BeforeAll {
            $script:DisabledCheckUser = [pscustomobject]@{
                Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                DisplayName = 'Test Person'
                UserPrincipalName = 'test@example.org'
                AccountEnabled = $false
                UserType = 'Member'
                OnPremisesSyncEnabled = $false
                SignInActivity = [pscustomobject]@{ LastSuccessfulSignInDateTime = '2026-01-02T10:30:00Z' }
            }
        }

        It 'rejects active accounts even if the server returns them' {
            $user = $script:DisabledCheckUser.PSObject.Copy()
            $user.AccountEnabled = $true
            { ConvertTo-DisabledAccountCheckRecord -User $user } | Should -Throw '*nicht gesperrt*'
        }

        It 'protects a blocked shared mailbox and keeps the disable date unknown' {
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser `
                -Memberships @([pscustomobject]@{ OdataType = '#microsoft.graph.group'; DisplayName = 'SEC-A-ROL-Schule_Schüler' }) `
                -Recipient ([pscustomobject]@{ RecipientTypeDetails = 'SharedMailbox' })
            $record.SecAGroups | Should -Be @('SEC-A-ROL-Schule_Schüler')
            $record.ProtectionReasons | Should -Contain 'Exchange-Empfänger: SharedMailbox'
            $record.LastLogin.ToString('yyyy-MM-dd HH:mm') | Should -Be '2026-01-02 10:30'
            $record.DisabledAt | Should -BeNullOrEmpty
            $record.Meets90Days | Should -BeNullOrEmpty
        }

        It 'flags guests, synced users, directory roles and missing SEC-A groups for review' {
            $user = $script:DisabledCheckUser.PSObject.Copy()
            $user.UserType = 'Guest'
            $user.OnPremisesSyncEnabled = $true
            $member = [pscustomobject]@{ AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.directoryRole'; displayName = 'Global Administrator' } }
            $record = ConvertTo-DisabledAccountCheckRecord -User $user -Memberships @($member)
            $record.ProtectionReasons | Should -Contain 'Gastkonto'
            $record.ProtectionReasons | Should -Contain 'Lokales AD synchronisiert'
            $record.ProtectionReasons | Should -Contain 'Verzeichnisrolle: Global Administrator'
            $record.ProtectionReasons | Should -Contain 'Keine SEC-A-Gruppe, Kontozweck prüfen'
        }

        It 'reads Graph AdditionalProperties from a generic dictionary' {
            $properties = [Collections.Generic.Dictionary[string,object]]::new()
            $properties['@odata.type'] = '#microsoft.graph.group'
            $properties['displayName'] = 'SEC-A-ROL-Schule_Schüler'
            $member = [pscustomobject]@{ AdditionalProperties = $properties }
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser -Memberships @($member)
            $record.SecAGroups | Should -Be @('SEC-A-ROL-Schule_Schüler')
        }

        It 'does not turn an old last login into a 90-day deletion decision' {
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser `
                -Memberships @([pscustomobject]@{ OdataType = '#microsoft.graph.group'; DisplayName = 'SEC-A-ROL-Schule_Schüler' }) `
                -Recipient ([pscustomobject]@{ RecipientTypeDetails = 'UserMailbox' })
            $record.LastLogin | Should -Not -BeNullOrEmpty
            $record.DisabledAt | Should -BeNullOrEmpty
            $record.Meets90Days | Should -BeNullOrEmpty
            $record.ReviewStatus | Should -Be 'Manuell prüfen'
        }

        It 'treats a remote user mailbox as a personal mailbox' {
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser `
                -Recipient ([pscustomobject]@{ RecipientTypeDetails = 'RemoteUserMailbox' })
            $record.ProtectionReasons | Should -Not -Contain 'Exchange-Empfänger: RemoteUserMailbox'
        }

        It 'separates missing sign-in data from a known empty sign-in history' {
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser -LastLoginAvailable:$false
            $record.LastLogin | Should -BeNullOrEmpty
            $record.LastLoginStatus | Should -Be 'Nicht abrufbar'
        }

        It 'marks unreadable group and recipient details instead of silently clearing them' {
            $member = [pscustomobject]@{ OdataType = '#microsoft.graph.group'; DisplayName = $null }
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser `
                -Memberships @($member) -Recipient ([pscustomobject]@{ RecipientTypeDetails = $null })
            $record.ProtectionReasons | Should -Contain 'Gruppenname nicht lesbar'
            $record.ProtectionReasons | Should -Contain 'Exchange-Empfängertyp nicht lesbar'
        }

        It 'marks accounts that own Entra objects' {
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser `
                -OwnedObjects @([pscustomobject]@{ Id = 'owned-group' })
            $record.ProtectionReasons | Should -Contain 'Besitzt 1 Entra-Objekt(e)'
        }

        It 'lists type, name and ID of every owned object from Graph properties' {
            $ownedObjects = @(
                [pscustomobject]@{ Id = 'group-1'; OdataType = '#microsoft.graph.group'; AdditionalProperties = @{ displayName = 'Schulteam' } },
                [pscustomobject]@{ AdditionalProperties = @{ id = 'app-1'; '@odata.type' = '#microsoft.graph.application'; displayName = 'Elternportal' } },
                [pscustomobject]@{ Id = 'unknown-1'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }
            )
            $record = ConvertTo-DisabledAccountCheckRecord -User $script:DisabledCheckUser -OwnedObjects $ownedObjects
            $record.ProtectionReasons | Should -Contain 'Besitzt 3 Entra-Objekt(e)'
            $record.OwnedObjects.Count | Should -Be 3
            $record.OwnedObjects[0].Type | Should -Be 'group'
            $record.OwnedObjects[0].DisplayName | Should -Be 'Schulteam'
            $record.OwnedObjects[0].Id | Should -Be 'group-1'
            $record.OwnedObjects[1].Type | Should -Be 'application'
            $record.OwnedObjects[1].DisplayName | Should -Be 'Elternportal'
            $record.OwnedObjects[1].Id | Should -Be 'app-1'
            $record.OwnedObjects[2].Type | Should -Be 'servicePrincipal'
            $record.OwnedObjects[2].DisplayName | Should -BeNullOrEmpty
            $record.OwnedObjects[2].Id | Should -Be 'unknown-1'
        }
    }
}

Describe 'Gesperrte Konten: lesender Gesamtlauf' {
    InModuleScope SchuelerSync {
        BeforeEach {
            Mock Connect-DisabledAccountGraph { [pscustomobject]@{ TenantId = 'tenant-test' } }
            Mock Connect-SchuelerExchangeOnline { [pscustomobject]@{ State = 'Connected' } }
            Mock Get-MgUser {
                @(
                    [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Gesperrt'
                        UserPrincipalName = 'disabled@example.org'; AccountEnabled = $false; UserType = 'Member'
                        OnPremisesSyncEnabled = $false; SignInActivity = $null },
                    [pscustomobject]@{ Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; DisplayName = 'Aktiv'
                        UserPrincipalName = 'active@example.org'; AccountEnabled = $true }
                )
            }
            Mock Get-MgUserTransitiveMemberOf {
                [pscustomobject]@{ OdataType = '#microsoft.graph.group'; DisplayName = 'SEC-A-ROL-Schule_Schüler' }
            }
            Mock Get-MgUserOwnedObject { @() }
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'UserMailbox' }
            }
            Mock Write-Progress {}
        }

        It 'rejects a missing Graph tenant before reading the Exchange inventory' {
            Mock Connect-DisabledAccountGraph { [pscustomobject]@{TenantId=''} }
            { Invoke-DisabledAccountCheck } | Should -Throw '*Tenant-ID*'
            Should -Invoke Connect-SchuelerExchangeOnline -Exactly 0
            Should -Invoke Get-MgUser -Exactly 0
            Should -Invoke Get-EXORecipient -Exactly 0
        }

        It 'lists disabled accounts and completes the progress bar' {
            $result = Invoke-DisabledAccountCheck
            $result.Mode | Should -Be 'Check'
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].UPN | Should -Be 'disabled@example.org'
            $result.Accounts[0].RecipientType | Should -Be 'UserMailbox'
            $result.Accounts[0].RecipientStatus | Should -Be 'Unique'
            $result.ExcludedCount | Should -Be 0
            Should -Invoke Get-MgUser -Exactly 1
            Should -Invoke Get-EXORecipient -ParameterFilter { $ResultSize -eq 'Unlimited' -and $Properties -contains 'ExternalDirectoryObjectId' } -Exactly 1
            Should -Invoke Get-MgUserOwnedObject -Exactly 1
            Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly 1
        }

        It 'lists a shared mailbox with a protection reason and loads its details' {
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'SharedMailbox' }
            }
            $result = Invoke-DisabledAccountCheck
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].RecipientType | Should -Be 'SharedMailbox'
            $result.Accounts[0].RecipientStatus | Should -Be 'Unique'
            $result.Accounts[0].ProtectionReasons | Should -Contain 'Exchange-Empfänger: SharedMailbox'
            $result.ExcludedCount | Should -Be 0
            Should -Invoke Get-MgUserTransitiveMemberOf -Exactly 1
            Should -Invoke Get-MgUserOwnedObject -Exactly 1
        }

        It 'lists room mailboxes and accounts without a reliable recipient type' {
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'RoomMailbox' }
            }
            $roomResult = Invoke-DisabledAccountCheck
            $roomResult.Accounts.Count | Should -Be 1
            $roomResult.Accounts[0].ProtectionReasons | Should -Contain 'Exchange-Empfänger: RoomMailbox'

            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = $null }
            }
            $unknownResult = Invoke-DisabledAccountCheck
            $unknownResult.Accounts.Count | Should -Be 1
            $unknownResult.Accounts[0].ProtectionReasons | Should -Contain 'Exchange-Empfängertyp nicht lesbar'
            $unknownResult.Accounts[0].RecipientStatus | Should -Be 'Unreadable'
        }

        It 'keeps a remote user mailbox in the list' {
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'RemoteUserMailbox' }
            }
            $result = Invoke-DisabledAccountCheck
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].RecipientType | Should -Be 'RemoteUserMailbox'
        }

        It 'lists accounts without an Exchange recipient without a warning' {
            Mock Get-EXORecipient { @() }
            $result = Invoke-DisabledAccountCheck
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].ProtectionReasons | Should -Contain 'Kein Exchange-Empfänger, Kontozweck prüfen'
            $result.Accounts[0].RecipientStatus | Should -Be 'None'
            $result.ExcludedCount | Should -Be 0
            $result.Warnings.Count | Should -Be 0
        }

        It 'keeps a matched personal mailbox when another disabled account has no recipient' {
            Mock Get-MgUser {
                @(
                    [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Ohne Postfach'
                        UserPrincipalName = 'without@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false; SignInActivity = $null },
                    [pscustomobject]@{ Id = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; DisplayName = 'Mit Postfach'
                        UserPrincipalName = 'with@example.org'; AccountEnabled = $false; UserType = 'Member'; OnPremisesSyncEnabled = $false; SignInActivity = $null }
                )
            }
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB'; RecipientTypeDetails = 'UserMailbox' }
            }
            $result = Invoke-DisabledAccountCheck
            $result.Accounts.Count | Should -Be 2
            $result.Accounts.UPN | Should -Contain 'with@example.org'
            $result.Accounts.UPN | Should -Contain 'without@example.org'
            $result.ExcludedCount | Should -Be 0
            $result.Warnings.Count | Should -Be 0
            Should -Invoke Get-EXORecipient -Exactly 1
        }

        It 'lists an ambiguous recipient mapping with a protection reason and reports it once' {
            Mock Get-EXORecipient {
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'UserMailbox' }
                [pscustomobject]@{ ExternalDirectoryObjectId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; RecipientTypeDetails = 'UserMailbox' }
            }
            $result = Invoke-DisabledAccountCheck
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].ProtectionReasons | Should -Contain 'Exchange-Empfänger nicht eindeutig zugeordnet'
            $result.Accounts[0].RecipientStatus | Should -Be 'Ambiguous'
            $result.ExcludedCount | Should -Be 0
            $result.Warnings.Count | Should -Be 1
            $result.Warnings[0] | Should -BeLike '*Mehrere Exchange-Empfänger*'
        }

        It 'stops the check when the Exchange inventory fails' {
            Mock Get-EXORecipient { throw 'Exchange nicht erreichbar' }
            { Invoke-DisabledAccountCheck } | Should -Throw '*Exchange nicht erreichbar*'
        }

        It 'still lists accounts when sign-in activity is unavailable and marks it unknown' {
            Mock Get-MgUser {
                if ($Property -contains 'signInActivity') { throw 'Sign-in not available' }
                [pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; DisplayName = 'Gesperrt'
                    UserPrincipalName = 'disabled@example.org'; AccountEnabled = $false; UserType = 'Member'
                    OnPremisesSyncEnabled = $false; SignInActivity = $null }
            }
            $result = Invoke-DisabledAccountCheck
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].LastLoginStatus | Should -Be 'Nicht abrufbar'
            $result.Warnings.Count | Should -Be 1
        }
    }
}
