BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Student password generation' {
    InModuleScope SchuelerSync {
        It 'offers at least 24 bits of reachable friendly passwords without duplicate prefix weighting' {
            $prefixes = Get-StudentPasswordPrefixes
            $unique = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            $unique.UnionWith([string[]]$prefixes)
            $unique.Count | Should -Be $prefixes.Count
            @($prefixes | Where-Object { $_.Length -ne 10 -or $_ -cnotmatch '^[A-Z][a-z]+[A-Z][a-z]+$' }).Count | Should -Be 0
            # Each distinct prefix has exactly 90 possible cryptographic numeric suffixes.
            ([long]$unique.Count * 90) | Should -BeGreaterOrEqual 16777216
            foreach ($index in @(0, ($prefixes.Count - 1))) {
                $script:chosenPrefix = $index
                $used = [Collections.Generic.HashSet[string]]::new()
                $password = New-StudentPassword -UsedPasswords $used -RandomIndexScriptBlock { param($count) $script:chosenPrefix }
                $password.Substring(0, 10) | Should -Be $prefixes[$index]
                [int]$password.Substring(10) | Should -BeGreaterOrEqual 10
                [int]$password.Substring(10) | Should -BeLessOrEqual 99
            }
        }

        It 'accepts an initially empty collection and creates exactly twelve friendly characters' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $password = New-StudentPassword -UsedPasswords $used
            $password.Length | Should -Be 12
            $password | Should -Match '^[A-Z][a-z]+[A-Z][a-z]+[0-9]{2}$'
            $used.Count | Should -Be 1
        }

        It 'does not return a password already reserved in the run' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $first = New-StudentPassword -UsedPasswords $used
            $second = New-StudentPassword -UsedPasswords $used
            $first | Should -Not -Be $second
            $used.Count | Should -Be 2
        }

        It 'stops after the bounded uniqueness attempts when every suffix for the selected prefix is reserved' {
            $prefixes = Get-StudentPasswordPrefixes
            $used = [Collections.Generic.HashSet[string]]::new()
            foreach ($number in 10..99) { [void]$used.Add("$($prefixes[0])$number") }
            $script:indexCalls = 0
            { New-StudentPassword -UsedPasswords $used -RandomIndexScriptBlock { param($count) $script:indexCalls++; 0 } } | Should -Throw '*100*'
            $script:indexCalls | Should -Be 100
            $used.Count | Should -Be 90
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
