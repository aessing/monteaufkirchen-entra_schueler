BeforeDiscovery {
    Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Manual update of an existing student' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:updateEntry = [pscustomobject]@{
                Student = [pscustomobject]@{ RowNumber = 0; NameMitRufname = 'Muster, Mia' }
                User = [pscustomobject]@{ Id = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; UserPrincipalName = 'mmuster@monteaufkirchen.com' }
                DesiredState = [pscustomobject]@{
                    UserPrincipalName = 'mmuster@monteaufkirchen.com'; ManagerId = 'teacher-id'
                    RequiredGroupNames = @('role', 'license', 'class')
                }
                Differences = @(
                    [pscustomobject]@{ Area = 'Entra'; Field = 'Department'; Action = 'Set' }
                    [pscustomobject]@{ Area = 'Manager'; Field = 'ManagerId'; Action = 'Set' }
                    [pscustomobject]@{ Area = 'Group'; Field = 'Membership'; Action = 'Add' }
                )
            }
            $script:updateSnapshot = [pscustomobject]@{ GroupsByDisplayName = @{ role = 'role'; license = 'license'; class = 'class' } }
            $script:updateConfig = @{ RoleGroupPrefix = 'SEC-A-ROL-'; ClassGroupPrefix = 'SEC-A-CLS-' }
            Mock Set-EntraStudentAttribute { [pscustomobject]@{ Verified = $true } }
            Mock Set-EntraStudentManager { [pscustomobject]@{ Verified = $true } }
            Mock Sync-EntraStudentGroup { [pscustomobject]@{ Verified = $true } }
            Mock Get-UserManagerId { 'old-teacher' }
            Mock Get-UserDirectGroup { @() }
            Mock New-StudentPassword { throw 'Existing users must not receive a password.' }
            Mock Write-StudentWorkbookUpdate { throw 'Manual update must not write Excel.' }
        }

        It 'applies Entra, manager and group differences without Excel or password changes' {
            $result = @(Invoke-ManualStudentUpdate -Entry $script:updateEntry -Snapshot $script:updateSnapshot -Config $script:updateConfig -Confirm:$false)

            $result[0].Status | Should -Be 'Succeeded'
            $result[0].Phase | Should -Be 'Add'
            Should -Invoke Set-EntraStudentAttribute -Times 1 -Exactly
            Should -Invoke Set-EntraStudentManager -Times 1 -Exactly
            Should -Invoke Sync-EntraStudentGroup -Times 1 -Exactly
            Should -Invoke New-StudentPassword -Times 0 -Exactly
            Should -Invoke Write-StudentWorkbookUpdate -Times 0 -Exactly
        }

        It 'enables a disabled existing student only after the existing required state was accepted' {
            $script:updateEntry.Differences = @()
            $script:updateEntry.User | Add-Member -NotePropertyName AccountEnabled -NotePropertyValue $false
            Mock Enable-EntraStudent { [pscustomobject]@{ Verified = $true } }

            $result = @(Invoke-ManualStudentUpdate -Entry $script:updateEntry -Snapshot $script:updateSnapshot -Config $script:updateConfig -Confirm:$false)

            $result[0].Status | Should -Be 'Succeeded'
            Should -Invoke Enable-EntraStudent -ParameterFilter { $WorkbookVerified -and $GroupsVerified -and $ManagerVerified -and $AttributesVerified } -Times 1 -Exactly
            Should -Invoke New-StudentPassword -Times 0 -Exactly
        }

        It 'retains the exact object ID when a resumed update fails again' {
            Mock Set-EntraStudentManager { throw 'manager temporarily unavailable' }

            $result = @(Invoke-ManualStudentUpdate -Entry $script:updateEntry -Snapshot $script:updateSnapshot -Config $script:updateConfig -Confirm:$false)

            $result[0].Status | Should -Be 'Failed'
            $result[0].RecoveryCommand | Should -Match ([regex]::Escape("-EntraObjectId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'"))
        }
    }
}

Describe 'Manual student operations' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:manualPassword = 'TigerWiese56'
            $script:manualDesired = [pscustomobject]@{
                UserPrincipalName = 'mmuster@monteaufkirchen.com'
                DisplayName = 'Mia Muster'
                ManagerId = 'teacher-id'
                RequiredGroupNames = @('role', 'license', 'class')
            }
            $script:manualEntry = [pscustomobject]@{
                Student = [pscustomobject]@{ RowNumber = 0; NameMitRufname = 'Muster, Mia' }
                User = $null
                DesiredState = $script:manualDesired
                Differences = @()
            }
            $script:manualSnapshot = [pscustomobject]@{
                GroupsByDisplayName = @{ role = 'role'; license = 'license'; class = 'class' }
            }
            $script:manualConfig = @{ RoleGroupPrefix = 'SEC-A-ROL-'; ClassGroupPrefix = 'SEC-A-CLS-' }

            Mock New-StudentPassword { $script:manualPassword }
            Mock New-StudentWithPasswordRetry {
                [pscustomobject]@{
                    User = [pscustomobject]@{ Id = '11111111-2222-3333-4444-555555555555'; UserPrincipalName = $Desired.UserPrincipalName }
                    Password = $InitialPassword
                }
            }
            Mock Assert-NewEntraStudentAttribute { $true }
            Mock Set-EntraStudentManager { [pscustomobject]@{ Verified = $true } }
            Mock Sync-EntraStudentGroup { [pscustomobject]@{ Verified = $true } }
            Mock Enable-EntraStudent { [pscustomobject]@{ Verified = $true } }
            Mock Write-StudentWorkbookUpdate { throw 'Manual add must not write Excel.' }
            Mock Write-Information { return }
        }

        It 'creates, verifies and enables one student without touching Excel and shows the password once' {
            $result = @(Invoke-ManualStudentAdd -Entry $script:manualEntry -Snapshot $script:manualSnapshot -Config $script:manualConfig -Confirm:$false)

            $result.Count | Should -Be 1
            $result[0].Phase | Should -Be 'Add'
            $result[0].Status | Should -Be 'Succeeded'
            Should -Invoke New-StudentWithPasswordRetry -Times 1 -Exactly
            Should -Invoke Assert-NewEntraStudentAttribute -Times 1 -Exactly
            Should -Invoke Set-EntraStudentManager -Times 1 -Exactly
            Should -Invoke Sync-EntraStudentGroup -Times 1 -Exactly
            Should -Invoke Enable-EntraStudent -Times 1 -Exactly
            Should -Invoke Write-StudentWorkbookUpdate -Times 0 -Exactly
            Should -Invoke Write-Information -ParameterFilter { [string]$MessageData -match [regex]::Escape($script:manualPassword) } -Times 1 -Exactly
            ($result | ConvertTo-Json -Depth 10) | Should -Not -Match [regex]::Escape($script:manualPassword)
        }

        It 'does not generate or display a password under WhatIf' {
            $result = @(Invoke-ManualStudentAdd -Entry $script:manualEntry -Snapshot $script:manualSnapshot -Config $script:manualConfig -WhatIf)

            $result[0].Status | Should -Be 'WhatIf'
            Should -Invoke New-StudentPassword -Times 0 -Exactly
            Should -Invoke New-StudentWithPasswordRetry -Times 0 -Exactly
            Should -Invoke Write-Information -Times 0 -Exactly
        }

        It 'shows the recoverable password once and reports an unknown state when activation verification fails' {
            Mock Enable-EntraStudent { throw 'activation failed' }

            $result = @(Invoke-ManualStudentAdd -Entry $script:manualEntry -Snapshot $script:manualSnapshot -Config $script:manualConfig -Confirm:$false)

            $result[0].Status | Should -Be 'Failed'
            $result[0].Phase | Should -Be 'Enable'
            Should -Invoke Write-Information -ParameterFilter { [string]$MessageData -match [regex]::Escape($script:manualPassword) } -Times 1 -Exactly
            ($result | ConvertTo-Json -Depth 10) | Should -Not -Match [regex]::Escape($script:manualPassword)
            $result[0].Message | Should -Match 'unbekannt'
            $result[0].Message | Should -Not -Match 'bleibt deaktiviert'
            $result[0].RecoveryCommand | Should -Match ([regex]::Escape("-EntraObjectId '11111111-2222-3333-4444-555555555555'"))
        }
    }
}
