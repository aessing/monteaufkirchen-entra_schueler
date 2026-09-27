BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Student password generation' {
    InModuleScope SchuelerSync {
        It 'offers more than 64 bits of valid twelve-character passwords' {
            $chars = $script:StudentPasswordCharacters.ToCharArray()
            @($chars | Select-Object -Unique).Count | Should -Be $chars.Count
            $upper = @($chars | Where-Object { [char]::IsUpper($_) }).Count
            $lower = @($chars | Where-Object { [char]::IsLower($_) }).Count
            $digits = @($chars | Where-Object { [char]::IsDigit($_) }).Count
            $count = [math]::Pow($chars.Count,12) - [math]::Pow($lower+$digits,12) - [math]::Pow($upper+$digits,12) -
                [math]::Pow($upper+$lower,12) + [math]::Pow($upper,12) + [math]::Pow($lower,12) + [math]::Pow($digits,12)
            [math]::Log2($count) | Should -BeGreaterThan 64
        }
        It 'creates twelve readable characters with all required classes and no duplicates in one run' {
            $used = [Collections.Generic.HashSet[string]]::new()
            foreach ($iteration in 1..100) {
                $password = New-StudentPassword -UsedPasswords $used
                $password.Length | Should -Be 12
                $password | Should -Match '^[A-HJ-NP-Za-hj-km-np-z2-9]{12}$'
                $password | Should -MatchExactly '[A-Z]'
                $password | Should -MatchExactly '[a-z]'
                $password | Should -Match '[2-9]'
            }
            $used.Count | Should -Be 100
        }
        It 'stops after bounded attempts if a valid password is already reserved' {
            $used = [Collections.Generic.HashSet[string]]::new()
            [void]$used.Add('Aa2Aa2Aa2Aa2')
            $script:indexCalls = 0
            { New-StudentPassword -UsedPasswords $used -RandomIndexScriptBlock {
                param($count)
                $character = 'Aa2'[$script:indexCalls % 3]
                $script:indexCalls++
                $script:StudentPasswordCharacters.IndexOf($character)
            } } | Should -Throw '*100*'
            $script:indexCalls | Should -Be 1200
            $used.Count | Should -Be 1
        }
        It 'rejects candidates that lack a required character class' {
            $used = [Collections.Generic.HashSet[string]]::new()
            { New-StudentPassword -UsedPasswords $used -RandomIndexScriptBlock { param($count) 0 } } | Should -Throw '*100*'
            $used.Count | Should -Be 0
        }
        It 'rejects an invalid injected selection index' -ForEach @(
            @{ Index = -1 }, @{ Index = 'not-an-index' }, @{ Index = [int]::MaxValue }
        ) {
            $script:invalidIndex = $Index
            $used = [Collections.Generic.HashSet[string]]::new()
            { New-StudentPassword -UsedPasswords $used -RandomIndexScriptBlock { param($count) $script:invalidIndex } } | Should -Throw '*Zufallsindex*'
        }
    }
}
