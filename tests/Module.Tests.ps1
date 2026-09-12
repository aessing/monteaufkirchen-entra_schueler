BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $manifest = Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1'
    $configPath = Join-Path $repoRoot 'config/SchuelerSync.psd1'
}

Describe 'SchuelerSync module contract' {
    It 'imports and exports only Invoke-SchuelerSync' {
        Import-Module $manifest -ErrorAction Stop
        (Get-Command Invoke-SchuelerSync -Module SchuelerSync).Name |
            Should -Be 'Invoke-SchuelerSync'
        @(Get-Command -Module SchuelerSync).Name |
            Should -Be @('Invoke-SchuelerSync')
    }

    It 'contains the immutable student configuration' {
        $config = Import-PowerShellDataFile $configPath
        $config.Domain | Should -Be 'monteaufkirchen.com'
        $config.StudentRoleGroup.Id |
            Should -Be 'cebc1326-1174-4126-ba84-7a8960850e0a'
        $config.PasswordLength | Should -Be 12
        $config.Exchange.MaxMailboxRetries | Should -Be 5
        $config.Exchange.RetryDelaySeconds | Should -Be 60
    }
}
