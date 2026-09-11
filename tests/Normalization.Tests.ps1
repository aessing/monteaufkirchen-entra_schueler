BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Student identity normalization' {
    InModuleScope SchuelerSync {
        It 'selects and reserves the first UPN with no occupied addresses or owners' {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            Test-UsedAddress -UsedAddresses $used -Candidate 'mmueller@monteaufkirchen.com' | Should -BeFalse
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners @{}
            $result.Upn | Should -Be 'mmueller@monteaufkirchen.com'
            $result.CollisionCount | Should -Be 0
            $used.Count | Should -Be 1
            $used.Contains($result.Upn) | Should -BeTrue
        }

        It 'adds the first comparison address to empty accumulators without a null owner' {
            $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $owners = @{}
            Add-ComparisonAddressOwner -AddressOwners $owners -ReservedAddresses $reserved -Address 'FIRST@school.example' -Owner 'student-1'
            $reserved.Count | Should -Be 1
            $owners.Count | Should -Be 1
            @($owners['first@school.example']).Count | Should -Be 1
            $owners['first@school.example'][0] | Should -Be 'student-1'
        }

        It 'transliterates German characters' {
            ConvertTo-UpnToken 'Änne Weiß' | Should -Be 'aenneweiss'
            ConvertTo-UpnToken 'JÖRG-MÜLLER' | Should -Be 'joergmueller'
        }

        It 'derives allowed office locations' -ForEach @(
            @{ Class = 'JK1-3g2_1'; Expected = 'G2' }
            @{ Class = 'JK4-6m2_4'; Expected = 'M2' }
            @{ Class = 'JK7-9O1_2'; Expected = 'O1' }
            @{ Class = 'JK7-9a2_3'; Expected = 'A2' }
        ) {
            Get-OfficeLocation $Class | Should -Be $Expected
        }

        It 'rejects an unsupported office token' {
            { Get-OfficeLocation 'JK1-3x9_1' } | Should -Throw '*Office Location*'
        }

        It 'generates prefix candidates in order' {
            $actual = @(Get-UpnCandidate -GivenName 'Maria' -Surname 'Müller' -Domain 'monteaufkirchen.com')
            $actual | Should -Be @(
                'mmueller@monteaufkirchen.com'
                'mamueller@monteaufkirchen.com'
                'marmueller@monteaufkirchen.com'
                'marimueller@monteaufkirchen.com'
                'mariamueller@monteaufkirchen.com'
            )
        }

        It 'uses the first free address and records UPN collisions' {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            [void] $used.Add('mmueller@monteaufkirchen.com')
            $owners = @{ 'mmueller@monteaufkirchen.com' = @('other-id') }
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners $owners
            $result.Upn | Should -Be 'mamueller@monteaufkirchen.com'
            $result.WasFallback | Should -BeFalse
            $result.CollisionCount | Should -Be 1
            $result.Collisions | Should -Contain 'mmueller@monteaufkirchen.com'
            $used | Should -Contain 'mamueller@monteaufkirchen.com'
        }

        It 'allows an address only when all owners are the current object' {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $owners = @{ 'mmueller@monteaufkirchen.com' = @('same-id') }
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners $owners -CurrentObjectId same-id
            $result.Upn | Should -Be 'mmueller@monteaufkirchen.com'
        }

        It 'treats used addresses as case-insensitive with a case-sensitive set' {
            $used = [Collections.Generic.HashSet[string]]::new()
            [void] $used.Add('MMueller@Monteaufkirchen.com')
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners @{}
            $result.Upn | Should -Be 'mamueller@monteaufkirchen.com'
            $result.Collisions | Should -Contain 'mmueller@monteaufkirchen.com'
        }

        It 'blocks UPN, mail, proxy and Exchange recipient address sources' -ForEach @(
            @{ Source = 'UPN'; Value = @('other-upn') }
            @{ Source = 'mail'; Value = [pscustomobject]@{ GraphObjectId = 'other-mail' } }
            @{ Source = 'proxy address'; Value = [pscustomobject]@{ OwnerIds = @('other-proxy') } }
            @{ Source = 'Exchange recipient'; Value = [pscustomobject]@{ ExchangeObjectId = 'other-exchange' } }
        ) {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $owners = @{ 'MMUELLER@MONTEAUFKIRCHEN.COM' = $Value }
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners $owners
            $result.Upn | Should -Be 'mamueller@monteaufkirchen.com'
            $result.Collisions | Should -Contain 'mmueller@monteaufkirchen.com'
        }

        It 'rejects reuse when any known owner differs from the current object' {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $owners = @{ 'MmUeLlEr@Monteaufkirchen.com' = @('same-id', 'other-id') }
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners $owners -CurrentObjectId same-id
            $result.Upn | Should -Be 'mamueller@monteaufkirchen.com'
        }

        It 'falls back to a suffixed full name after all collisions' {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $owners = @{}
            foreach ($candidate in (Get-UpnCandidate Maria Müller monteaufkirchen.com)) {
                [void] $used.Add($candidate)
                $owners[$candidate] = @('other-id')
            }
            $result = Select-AvailableUpn -GivenName Maria -Surname Müller -Domain monteaufkirchen.com -UsedAddresses $used -AddressOwners $owners
            $result.Upn | Should -Be 'mariamueller2@monteaufkirchen.com'
            $result.WasFallback | Should -BeTrue
        }
    }
}
