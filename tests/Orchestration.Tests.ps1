BeforeDiscovery {
    Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Public orchestration' {
    InModuleScope SchuelerSync {
        BeforeAll {
            function New-TestEntry([string] $Id, [int] $Row = 2) {
                [pscustomobject]@{
                    Student = [pscustomobject]@{ RowNumber = $Row; NameMitRufname = "Test $Row"; ClassName = 'JK1-3g2_1'; Password = 'NeverReport12' }
                    User = if ($Id) { [pscustomobject]@{ Id = $Id; UserPrincipalName = "$Id@monteaufkirchen.com"; DisplayName = "Test $Row"; AccountEnabled = $true } } else { $null }
                    DesiredState = [pscustomobject]@{ UserPrincipalName = "test$Row@monteaufkirchen.com"; DisplayName = "Test $Row"; Department = 'JK1-3g2_1'; ManagerId = 'teacher'; RequiredGroupNames = @('role', 'license', 'class') }
                    Differences = @()
                }
            }
        }
        BeforeEach {
            $script:comparison = [pscustomobject]@{
                NewStudents = @((New-TestEntry '' 2)); ChangedStudents = @((New-TestEntry 'changed' 3))
                ExistingStudents = @((New-TestEntry 'existing' 4)); Departures = @((New-TestEntry 'departure' 5))
                Warnings = @(); Errors = @()
            }
            Mock Read-StudentWorkbook { [pscustomobject]@{ Path = $Path; SourceHash = 'original-version'; Students = @([pscustomobject]@{ Password = 'NeverReport12' }) } }
            Mock Assert-StudentWorkbookVersion { return }
            Mock Assert-WorkbookSafeForPasswordWrite { return }
            Mock Connect-SchuelerGraph { [pscustomobject]@{ TenantId = 'test-tenant'; Account = 'admin@test.invalid' } }
            Mock Get-EntraSnapshot { [pscustomobject]@{} }
            Mock Get-EntraStudentCurrentIdentity { [pscustomobject]@{ Id = $UserId; UserPrincipalName = "$UserId@monteaufkirchen.com" } }
            Mock Connect-SchuelerExchangeOnline { return }
            Mock Get-ExchangeRecipientAddresses { [pscustomobject]@{ AddressOwners = @{} } }
            Mock Compare-StudentDirectory { $script:comparison }
            Mock Get-StudentMailboxState { [pscustomobject]@{ Exists = $false; Status = 'MailboxNotReady'; Differences = @() } }
            Mock Invoke-StudentCreateBatch { New-StudentActionResult -UserId 'new' -UserPrincipalName 'new@monteaufkirchen.com' -Phase Create -Status Succeeded }
            Mock Invoke-StudentUpdates { New-StudentActionResult -UserId 'changed' -UserPrincipalName 'changed@monteaufkirchen.com' -Phase Update -Status Succeeded }
            Mock Invoke-StudentDepartures { return }
            Mock Wait-StudentMailboxes { [pscustomobject]@{ Ready = @(); Missing = @(); Failed = @() } }
            Mock Write-Information { return }
        }
        It 'compares each matched mailbox once without any mutation or wait' {
            $result = Invoke-SchuelerSync -File 'synthetic.xlsx'
            Should -Invoke Get-StudentMailboxState -Times 2 -Exactly
            Should -Invoke Invoke-StudentCreateBatch -Times 0
            Should -Invoke Invoke-StudentUpdates -Times 0
            Should -Invoke Invoke-StudentDepartures -Times 0
            Should -Invoke Wait-StudentMailboxes -Times 0
            $result.Comparison.NewStudents.Count | Should -Be 1
            $result.Comparison.Departures.Count | Should -Be 1
            ($result | ConvertTo-Json -Depth 30) | Should -Not -Match 'NeverReport12|"Password"'
        }
        It 'uses the repository workbook default and PowerShell-relative custom paths' {
            Invoke-SchuelerSync | Out-Null
            Should -Invoke Read-StudentWorkbook -ParameterFilter { $Path -eq (Join-Path $script:SchuelerSyncRepositoryRoot 'Schueler.xlsx') }
            Push-Location $TestDrive
            try {
                Invoke-SchuelerSync -File 'custom.xlsx' | Out-Null
                Should -Invoke Read-StudentWorkbook -ParameterFilter { $Path -eq (Join-Path $TestDrive 'custom.xlsx') }
            } finally { Pop-Location }
        }
        It 'blocks all mutation for preflight errors' {
            $script:comparison.Errors = @([pscustomobject]@{ Area = 'Identity'; Field = 'Duplicate'; Message = 'Duplicate identity'; Student = $null; User = $null })
            $result = Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -Confirm:$false
            $result.HasErrors | Should -BeTrue
            Should -Invoke Invoke-StudentCreateBatch -Times 0
            Should -Invoke Invoke-StudentUpdates -Times 0
            Should -Invoke Invoke-StudentDepartures -Times 0
            Should -Invoke Wait-StudentMailboxes -Times 0
        }
        It 'runs all Graph actions and Exchange for every active workbook student on plain Update' {
            Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -Confirm:$false | Out-Null
            Should -Invoke Invoke-StudentCreateBatch -Times 1
            Should -Invoke Invoke-StudentUpdates -Times 1
            Should -Invoke Invoke-StudentDepartures -ParameterFilter { $DisableUsers -and $RevokeSessions }
            Should -Invoke Wait-StudentMailboxes -ParameterFilter {
                $UserPrincipalName.Count -eq 3 -and $UserPrincipalName -contains 'new@monteaufkirchen.com' -and $UserPrincipalName -contains 'existing@monteaufkirchen.com'
            }
        }
        It 'limits selective Exchange to successfully created accounts' {
            Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -CreateNewUsers -Confirm:$false | Out-Null
            Should -Invoke Wait-StudentMailboxes -ParameterFilter { $UserPrincipalName.Count -eq 1 -and $UserPrincipalName[0] -eq 'new@monteaufkirchen.com' }
            Should -Invoke Invoke-StudentUpdates -Times 0
        }
        It 'limits selective Exchange to successfully updated accounts' {
            Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -UpdateUsers -Confirm:$false | Out-Null
            Should -Invoke Wait-StudentMailboxes -ParameterFilter { $UserPrincipalName.Count -eq 1 -and $UserPrincipalName[0] -eq 'changed@monteaufkirchen.com' }
            Should -Invoke Invoke-StudentCreateBatch -Times 0
        }
        It 'uses the reread UPN for full Exchange after a partial rename failure' {
            Mock Invoke-StudentUpdates { New-StudentActionResult -UserId 'changed' -UserPrincipalName 'changed@monteaufkirchen.com' -Phase Manager -Status Failed }
            Mock Get-EntraStudentCurrentIdentity { [pscustomobject]@{ Id = $UserId; UserPrincipalName = 'renamed@monteaufkirchen.com' } } -ParameterFilter { $UserId -eq 'changed' }
            Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -Confirm:$false | Out-Null
            Should -Invoke Wait-StudentMailboxes -ParameterFilter {
                $UserPrincipalName -contains 'renamed@monteaufkirchen.com' -and $UserPrincipalName -notcontains 'changed@monteaufkirchen.com'
            }
        }
        It 'does not configure partially failed renames in selective UpdateUsers mode' {
            Mock Invoke-StudentUpdates { New-StudentActionResult -UserId 'changed' -UserPrincipalName 'renamed@monteaufkirchen.com' -Phase Manager -Status Failed }
            Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -UpdateUsers -Confirm:$false | Out-Null
            Should -Invoke Wait-StudentMailboxes -Times 0
        }
        It 'checks rename workbook writeability before any Graph or Exchange write' {
            $script:comparison.ChangedStudents[0].Differences = @([pscustomobject]@{ Area = 'Entra'; Field = 'UserPrincipalName'; Action = 'Set' })
            Mock Assert-WorkbookSafeForPasswordWrite { throw 'Workbook is read-only' }
            $result = Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -Confirm:$false
            $result.HasErrors | Should -BeTrue
            Should -Invoke Invoke-StudentCreateBatch -Times 0
            Should -Invoke Invoke-StudentUpdates -Times 0
            Should -Invoke Invoke-StudentDepartures -Times 0
            Should -Invoke Wait-StudentMailboxes -Times 0
        }
        It 'rejects a workbook version change before any selected mutation' {
            Mock Assert-StudentWorkbookVersion { throw 'Workbook changed since preflight' }
            $result = Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -DisableUsers -Confirm:$false
            $result.HasErrors | Should -BeTrue
            Should -Invoke Invoke-StudentDepartures -Times 0
        }
        It 'does not configure active Exchange for departures-only actions' {
            Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -DisableUsers -RevokeSessions -Confirm:$false | Out-Null
            Should -Invoke Wait-StudentMailboxes -Times 0
        }
        It 'Exchange-only bypasses Graph and Excel and normalizes identities' {
            Invoke-SchuelerSync -ConfigureExchangeOnlineOnly -Mail ' TEST@monteaufkirchen.com ', 'test@monteaufkirchen.com' -Confirm:$false | Out-Null
            Should -Invoke Read-StudentWorkbook -Times 0
            Should -Invoke Connect-SchuelerGraph -Times 0
            Should -Invoke Wait-StudentMailboxes -ParameterFilter { $UserPrincipalName.Count -eq 1 -and $UserPrincipalName[0] -eq 'test@monteaufkirchen.com' }
        }
        It 'WhatIf retains reads and plans without invoking Graph orchestration mutations' {
            $result = Invoke-SchuelerSync -File 'synthetic.xlsx' -Update -WhatIf
            Should -Invoke Get-StudentMailboxState -Times 3 -Exactly
            Should -Invoke Invoke-StudentCreateBatch -Times 0
            Should -Invoke Invoke-StudentUpdates -Times 0
            Should -Invoke Invoke-StudentDepartures -Times 0
            Should -Invoke Wait-StudentMailboxes -Times 0
            $result.Actions.Status | Should -Contain 'WhatIf'
        }
    }
}

Describe 'Fail-safe new account batching' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:events = [Collections.Generic.List[string]]::new()
            $script:stored = @()
            $script:desired = [pscustomobject]@{ UserPrincipalName = 'new@monteaufkirchen.com'; DisplayName = 'New Test'; ManagerId = 'teacher'; RequiredGroupNames = @('role', 'license', 'class') }
            $script:entry = [pscustomobject]@{ Student = [pscustomobject]@{ RowNumber = 2 }; DesiredState = $script:desired }
            $script:snapshot = [pscustomobject]@{ GroupsByDisplayName = @{ role = 'role'; license = 'license'; class = 'class' } }
            $script:config = @{ RoleGroupPrefix = 'role'; ClassGroupPrefix = 'class' }
            $script:workbookState = [pscustomobject]@{ SourceHash = 'original-version' }
            Mock Assert-StudentWorkbookVersion { return }
            Mock New-StudentPassword { 'TigerWiese56' }
            Mock New-StudentWithPasswordRetry { $script:events.Add('create-disabled'); [pscustomobject]@{ User = [pscustomobject]@{ Id = 'new'; UserPrincipalName = 'new@monteaufkirchen.com' }; Password = $InitialPassword } }
            Mock Assert-NewEntraStudentAttributes { $script:events.Add('attributes'); $true }
            Mock Set-EntraStudentManager { $script:events.Add('manager'); [pscustomobject]@{ Verified = $true } }
            Mock Sync-EntraStudentGroups { $script:events.Add('groups'); [pscustomobject]@{ Verified = $true } }
            Mock Write-StudentWorkbookUpdates {
                $script:events.Add('excel-write')
                $script:stored = @($Updates | ForEach-Object { [pscustomobject]@{ RowNumber = $_.RowNumber; EntraObjectId = $_.EntraObjectId; StoredUpn = $_.UPN; Password = $_.Password } })
                [pscustomobject]@{ SourceHash = 'saved-version' }
            }
            Mock Read-StudentWorkbook { $script:events.Add('excel-read'); [pscustomobject]@{ Students = $script:stored; SourceHash = 'saved-version' } }
            Mock Enable-EntraStudent { $script:events.Add('enable'); [pscustomobject]@{ Verified = $true } }
        }
        It 'persists and rereads credentials before activation' {
            $result = @(Invoke-StudentCreateBatch -Entries @($script:entry) -Snapshot $script:snapshot -Config $script:config -File 'synthetic.xlsx' -WorkbookState $script:workbookState -Confirm:$false)
            ($script:events -join ',') | Should -Be 'create-disabled,attributes,manager,groups,excel-write,excel-read,enable'
            Should -Invoke Enable-EntraStudent -ParameterFilter { $WorkbookVerified -and $GroupsVerified -and $ManagerVerified -and $AttributesVerified }
            $result[0].Status | Should -Be Succeeded
            ($result | ConvertTo-Json -Depth 15) | Should -Not -Match 'TigerWiese56|"Password"'
        }
        It 'leaves every new account disabled when batch persistence fails' {
            Mock Write-StudentWorkbookUpdates { throw 'Workbook locked' }
            $result = @(Invoke-StudentCreateBatch -Entries @($script:entry) -Snapshot $script:snapshot -Config $script:config -File 'synthetic.xlsx' -WorkbookState $script:workbookState -Confirm:$false)
            Should -Invoke Enable-EntraStudent -Times 0
            $result[0].Status | Should -Be Failed
            $result[0].Phase | Should -Be Workbook
        }
        It 'does not activate if reread credentials differ' {
            Mock Read-StudentWorkbook { [pscustomobject]@{ Students = @(); SourceHash = 'saved-version' } }
            $result = @(Invoke-StudentCreateBatch -Entries @($script:entry) -Snapshot $script:snapshot -Config $script:config -File 'synthetic.xlsx' -WorkbookState $script:workbookState -Confirm:$false)
            Should -Invoke Enable-EntraStudent -Times 0
            $result[0].Status | Should -Be Failed
        }
        It 'does not generate passwords or touch Excel for WhatIf' {
            Invoke-StudentCreateBatch -Entries @($script:entry) -Snapshot $script:snapshot -Config $script:config -File 'synthetic.xlsx' -WorkbookState $script:workbookState -WhatIf | Out-Null
            Should -Invoke New-StudentPassword -Times 0
            Should -Invoke Write-StudentWorkbookUpdates -Times 0
            Should -Invoke Enable-EntraStudent -Times 0
        }
        It 'writes all prepared accounts in one batch before enabling either account' {
            $second = [pscustomobject]@{ Student = [pscustomobject]@{ RowNumber = 3 }; DesiredState = $script:desired }
            $result = @(Invoke-StudentCreateBatch -Entries @($script:entry, $second) -Snapshot $script:snapshot -Config $script:config -File 'synthetic.xlsx' -WorkbookState $script:workbookState -Confirm:$false)
            Should -Invoke Write-StudentWorkbookUpdates -Times 1 -Exactly -ParameterFilter { $Updates.Count -eq 2 }
            ($script:events -join ',') | Should -Be 'create-disabled,attributes,manager,groups,create-disabled,attributes,manager,groups,excel-write,excel-read,enable,enable'
            $result.Count | Should -Be 2
        }
        It 'continues independent students and redacts password-bearing API exceptions' {
            $second = [pscustomobject]@{ Student = [pscustomobject]@{ RowNumber = 3 }; DesiredState = $script:desired }
            Mock Set-EntraStudentManager { throw 'Rejected TigerWiese56' }
            $result = @(Invoke-StudentCreateBatch -Entries @($script:entry, $second) -Snapshot $script:snapshot -Config $script:config -File 'synthetic.xlsx' -WorkbookState $script:workbookState -Confirm:$false)
            $result.Count | Should -Be 2
            Should -Invoke New-StudentWithPasswordRetry -Times 2
            ($result | ConvertTo-Json -Depth 15) | Should -Not -Match 'TigerWiese56'
        }
    }
}

Describe 'Existing students and departures' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:entry = [pscustomobject]@{
                User = [pscustomobject]@{ Id = 'existing'; UserPrincipalName = 'existing@monteaufkirchen.com'; AccountEnabled = $true }
                DesiredState = [pscustomobject]@{ UserPrincipalName = 'existing@monteaufkirchen.com' }
                Differences = @([pscustomobject]@{ Area = 'Entra'; Field = 'Department'; Action = 'Set' })
            }
            Mock Set-EntraStudentAttributes { [pscustomobject]@{ Verified = $true } }
            Mock Assert-StudentWorkbookVersion { return }
            Mock Set-EntraStudentManager { throw 'Unexpected manager write' }
            Mock Sync-EntraStudentGroups { throw 'Unexpected group write' }
            Mock New-StudentPassword { throw 'Unexpected password generation' }
            Mock Write-StudentWorkbookUpdates { throw 'Unexpected workbook mutation' }
            Mock Disable-EntraStudent { throw 'Disable failed' }
            Mock Revoke-EntraStudentSessions { $true }
        }
        It 'changes only listed existing fields without generating or storing credentials' {
            $result = @(Invoke-StudentUpdates -Entries @($script:entry) -Snapshot ([pscustomobject]@{}) -Config @{} -File 'synthetic.xlsx' -WorkbookState ([pscustomobject]@{ SourceHash = 'original' }) -Confirm:$false)
            $result[0].Status | Should -Be Succeeded
            Should -Invoke Set-EntraStudentAttributes -ParameterFilter { $Differences.Count -eq 1 -and $Differences[0].Field -eq 'Department' }
            Should -Invoke Set-EntraStudentManager -Times 0
            Should -Invoke Sync-EntraStudentGroups -Times 0
            Should -Invoke New-StudentPassword -Times 0
            Should -Invoke Write-StudentWorkbookUpdates -Times 0
        }
        It 'still revokes after a failed disable and records each outcome separately' {
            $result = @(Invoke-StudentDepartures -Entries @($script:entry) -DisableUsers -RevokeSessions -File 'synthetic.xlsx' -Confirm:$false)
            ($result | Where-Object Phase -eq Disable).Status | Should -Be Failed
            ($result | Where-Object Phase -eq RevokeSessions).Status | Should -Be Succeeded
        }
        It 'does not report an unconfirmed session revoke as successful' {
            Mock Revoke-EntraStudentSessions { $false }
            $result = @(Invoke-StudentDepartures -Entries @($script:entry) -RevokeSessions -File 'synthetic.xlsx' -Confirm:$false)
            $result[0].Status | Should -Be Failed
        }
        It 'avoids re-disabling accounts already disabled' {
            $script:entry.User.AccountEnabled = $false
            $result = @(Invoke-StudentDepartures -Entries @($script:entry) -DisableUsers -File 'synthetic.xlsx' -Confirm:$false)
            $result[0].Status | Should -Be Compliant
            Should -Invoke Disable-EntraStudent -Times 0
        }
    }
}

Describe 'Exchange-only WhatIf through real batch orchestration' {
    InModuleScope SchuelerSync {
        It 'reads once and never sleeps or writes for a missing mailbox' {
            Mock Connect-SchuelerExchangeOnline { return }
            Mock Get-StudentMailboxState { [pscustomobject]@{ Exists = $false; Status = 'MailboxNotReady' } }
            Mock Start-Sleep { throw 'Unexpected sleep' }
            Mock Set-StudentMailboxConfiguration { throw 'Unexpected write' }
            Mock Write-Information { return }
            $result = Invoke-SchuelerSync -ConfigureExchangeOnlineOnly -Mail 'test@monteaufkirchen.com' -WhatIf
            Should -Invoke Get-StudentMailboxState -Times 1 -Exactly
            Should -Invoke Start-Sleep -Times 0
            Should -Invoke Set-StudentMailboxConfiguration -Times 0
            $result.Actions[0].Status | Should -Be WhatIf
        }
    }
}

Describe 'Direct workbook mutation boundary' {
    InModuleScope SchuelerSync {
        It 'honors WhatIf before reading or copying a workbook' {
            Mock Read-StudentWorkbook { throw 'Unexpected workbook read' }
            Mock Copy-Item { throw 'Unexpected backup' }
            Write-StudentWorkbookUpdates -Path 'does-not-exist.xlsx' -Updates @([pscustomobject]@{ RowNumber = 2 }) -WhatIf
            Should -Invoke Read-StudentWorkbook -Times 0
            Should -Invoke Copy-Item -Times 0
        }
    }
}
