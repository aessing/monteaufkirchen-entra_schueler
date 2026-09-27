BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
}

Describe 'Update selection' {
    InModuleScope SchuelerSync {
        BeforeEach { Mock Assert-CurrentStudentMembership { } }
        It 'selects every action for plain Update' {
            $selection = Resolve-UpdateSelection -Update
            foreach ($name in 'CreateNewUsers', 'UpdateUsers', 'DisableUsers', 'RevokeSessions', 'ConfigureAllActiveExchange') {
                $selection.$name | Should -BeTrue
            }
        }
        It 'selects no mutations for default comparison' {
            (Resolve-UpdateSelection).PSObject.Properties.Value | Should -Not -Contain $true
        }
        It 'limits selective creation to creation and affected-user Exchange' {
            $selection = Resolve-UpdateSelection -Update -CreateNewUsers
            $selection.CreateNewUsers | Should -BeTrue
            $selection.UpdateUsers | Should -BeFalse
            $selection.DisableUsers | Should -BeFalse
            $selection.RevokeSessions | Should -BeFalse
            $selection.ConfigureAllActiveExchange | Should -BeFalse
        }
        It 'rejects each action without Update' -ForEach @(
            @{ Action = 'CreateNewUsers' }, @{ Action = 'UpdateUsers' },
            @{ Action = 'DisableUsers' }, @{ Action = 'RevokeSessions' }
        ) {
            $parameters = @{ $Action = $true }
            { Resolve-UpdateSelection @parameters } | Should -Throw '*Update*'
        }

        It 'activates a listed disabled student only after the update prerequisites are verified' {
            $script:UpdateEvents = [Collections.Generic.List[string]]::new()
            $entry = [pscustomobject]@{
                User = [pscustomobject]@{ Id = 'disabled-id'; UserPrincipalName = 'student@example.invalid'; AccountEnabled = $false }
                DesiredState = [pscustomobject]@{ UserPrincipalName = 'student@example.invalid' }
                Differences = @([pscustomobject]@{ Area = 'Entra'; Field = 'AccountEnabled'; Current = $false; Desired = $true; Action = 'Set' })
            }
            Mock Assert-StudentWorkbookVersion { $script:UpdateEvents.Add('workbook') }
            Mock Set-EntraStudentAttribute { $script:UpdateEvents.Add('attributes'); [pscustomobject]@{ Verified = $true } }
            Mock Enable-EntraStudent { $script:UpdateEvents.Add('enable'); [pscustomobject]@{ Verified = $true } }
            $result = @(Invoke-StudentUpdate -Entries @($entry) -Snapshot ([pscustomobject]@{}) -Config @{} `
                -File 'test.xlsx' -WorkbookState ([pscustomobject]@{ SourceHash = 'test' }) -Confirm:$false)
            $result[0].Status | Should -Be 'Succeeded'
            @($script:UpdateEvents) | Should -Be @('workbook', 'attributes', 'enable')
            Should -Invoke Enable-EntraStudent -Exactly 1 -ParameterFilter {
                $UserId -eq 'disabled-id' -and $WorkbookVerified -and $GroupsVerified -and $ManagerVerified -and $AttributesVerified
            }
        }

        It 'keeps a listed disabled student locked when the update fails' {
            $entry = [pscustomobject]@{
                User = [pscustomobject]@{ Id = 'disabled-id'; UserPrincipalName = 'student@example.invalid'; AccountEnabled = $false }
                DesiredState = [pscustomobject]@{ UserPrincipalName = 'student@example.invalid' }
                Differences = @([pscustomobject]@{ Area = 'Entra'; Field = 'AccountEnabled'; Current = $false; Desired = $true; Action = 'Set' })
            }
            Mock Assert-StudentWorkbookVersion { }
            Mock Set-EntraStudentAttribute { [pscustomobject]@{ Verified = $false } }
            Mock Enable-EntraStudent { throw 'Unexpected activation' }
            $result = @(Invoke-StudentUpdate -Entries @($entry) -Snapshot ([pscustomobject]@{}) -Config @{} `
                -File 'test.xlsx' -WorkbookState ([pscustomobject]@{ SourceHash = 'test' }) -Confirm:$false)
            $result[0].Status | Should -Be 'Failed'
            Should -Invoke Enable-EntraStudent -Exactly 0
        }
    }
    It 'defines exclusive parameter sets and mandatory Mail aliases' {
        $command = Get-Command (Join-Path $repoRoot 'Sync-SchuelerEntra.ps1')
        $command.Parameters['Mail'].ParameterType | Should -Be ([string[]])
        $command.Parameters['Mail'].Aliases | Should -Contain 'UPN'
        $command.Parameters['Mail'].Aliases | Should -Contain 'UserPrincipalName'
        $exchange = $command.ParameterSets | Where-Object Name -eq ExchangeOnly
        ($exchange.Parameters | Where-Object Name -eq Mail).IsMandatory | Should -BeTrue
        foreach ($name in 'File', 'Update', 'CreateNewUsers', 'UpdateUsers', 'DisableUsers', 'RevokeSessions') {
            $exchange.Parameters.Name | Should -Not -Contain $name
        }
    }
}
