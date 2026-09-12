BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
}

Describe 'Update selection' {
    InModuleScope SchuelerSync {
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
