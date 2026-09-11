BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force
}

Describe 'Student password generation' {
    InModuleScope SchuelerSync {
        It 'creates exactly twelve friendly characters' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $password = New-StudentPassword -UsedPasswords $used
            $password.Length | Should -Be 12
            $password | Should -Match '^[A-Z][a-z]+[A-Z][a-z]+[0-9]{2}$'
        }

        It 'does not return a password already reserved in the run' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $first = New-StudentPassword -UsedPasswords $used
            $second = New-StudentPassword -UsedPasswords $used
            $first | Should -Not -Be $second
            $used.Count | Should -Be 2
        }
    }
}
