BeforeDiscovery {
    Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Creation address policy and existing account protection' {
    InModuleScope SchuelerSync {
        BeforeAll {
            function New-CreationSnapshot {
                param($Config)
                $groups = @{}; $licenses = @{}
                foreach ($role in $Config.Roles.Values) {
                    $groups[$role.Id] = [pscustomobject]@{ Id=$role.Id; DisplayName=$role.Name; AssignedLicenses=@(); GroupTypes=@() }
                }
                foreach ($profile in $Config.Profiles.Values) {
                    $licenses[$profile.LicenseGroupName] = [pscustomobject]@{ Id=$profile.LicenseGroupName; DisplayName=$profile.LicenseGroupName; AssignedLicenses=@('sku'); GroupTypes=@() }
                }
                [pscustomobject]@{
                    UsersById=@{}; UsersByUpn=@{}; GroupsById=$groups; LicenseGroups=$licenses
                    PersonnelMemberIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    PersonnelRolesByUserId=@{}; StudentMemberIds=[Collections.Generic.HashSet[string]]::new()
                    DirectGroupsByUserId=@{}; ReservedAddresses=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    AddressOwners=@{}
                }
            }
            function New-CreationPerson {
                param([string] $Job='L')
                [pscustomobject]@{ RowNumber=2; NameMitRufname='Testwald, Lea'; GivenName='Lea'; Surname='Testwald'; Job=$Job; EntraObjectId=''; StoredUpn=''; Password='' }
            }
        }
        BeforeEach {
            $script:creationConfig = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $script:creationSnapshot = New-CreationSnapshot $script:creationConfig
        }

        It 'creates school staff with the new primary domain and school alias for <Job>' -ForEach @(@{Job='L'},@{Job='CO-L'},@{Job='PA'},@{Job='L, JAS'}) {
            $result = Compare-LehrerDirectory -People @((New-CreationPerson $Job)) -Snapshot $script:creationSnapshot -Config $script:creationConfig
            $result.Errors | Should -BeNullOrEmpty
            $desired = $result.NewPeople[0].DesiredState
            $desired.UserPrincipalName | Should -Be 'ltestwald@montessori-erding.de'
            $desired.Mail | Should -Be $desired.UserPrincipalName
            $desired.Exchange.PrimarySmtpAddress | Should -Be $desired.Mail
            $desired.Exchange.RequiredAliases | Should -Contain 'ltestwald@monteaufkirchen.com'
        }
        It 'creates OGTS with the personnel domain and without a school alias' {
            $result = Compare-LehrerDirectory -People @((New-CreationPerson OGTS)) -Snapshot $script:creationSnapshot -Config $script:creationConfig
            $result.Errors | Should -BeNullOrEmpty
            $desired = $result.NewPeople[0].DesiredState
            $desired.Mail | Should -Be 'ltestwald@montessori-erding.de'
            $desired.Exchange.ContainsKey('RequiredAliases') | Should -BeFalse
        }
        It 'advances the name prefix when either primary address or alias is reserved by Graph or Exchange' -ForEach @(
            @{Address='ltestwald@montessori-erding.de'; Exchange=$false}
            @{Address='ltestwald@monteaufkirchen.com'; Exchange=$false}
            @{Address='ltestwald@monteaufkirchen.com'; Exchange=$true}
        ) {
            $exchangeOwners = @{}
            if ($Exchange) { $exchangeOwners[$Address]=@('other') }
            else { $script:creationSnapshot.AddressOwners[$Address]=@('other') }
            $result = Compare-LehrerDirectory -People @((New-CreationPerson)) -Snapshot $script:creationSnapshot -Config $script:creationConfig -ExchangeAddressOwners $exchangeOwners
            $result.Errors | Should -BeNullOrEmpty
            $result.NewPeople[0].DesiredState.Mail | Should -Be 'letestwald@montessori-erding.de'
            $result.NewPeople[0].DesiredState.Exchange.RequiredAliases | Should -Contain 'letestwald@monteaufkirchen.com'
        }
        It 'uses normalized first.last addresses for external personnel with numeric collisions' {
            $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/PersonalSync.psd1')
            $snapshot = New-CreationSnapshot $config
            $snapshot.AddressOwners['lea.testwald@montessori-erding.de']=@('other')
            $result = Compare-LehrerDirectory -People @((New-CreationPerson ExterneBenutzer)) -Snapshot $snapshot -Config $config
            $result.Errors | Should -BeNullOrEmpty
            $result.NewPeople[0].DesiredState.Mail | Should -Be 'lea.testwald2@montessori-erding.de'
        }
        It 'creates employees with first initial and expands it on collision for <Job>' -ForEach @(
            @{Job='Schule_Sekretariat'},@{Job='Kinderhaus_PädagogischesTeam'},@{Job='Unterstützung'}
        ) {
            $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/PersonalSync.psd1')
            $snapshot = New-CreationSnapshot $config
            $person = New-CreationPerson $Job
            $person.GivenName='Oliver'; $person.Surname='Testwald'
            $result = Compare-LehrerDirectory -People @($person) -Snapshot $snapshot -Config $config
            $result.NewPeople[0].DesiredState.Mail | Should -Be 'otestwald@montessori-erding.de'
            $snapshot.AddressOwners['otestwald@montessori-erding.de']=@('other')
            $result = Compare-LehrerDirectory -People @($person) -Snapshot $snapshot -Config $config
            $result.NewPeople[0].DesiredState.Mail | Should -Be 'oltestwald@montessori-erding.de'
        }
        It 'preserves all address fields on an existing school account even when mail differs from UPN' {
            $person = New-CreationPerson
            $person.EntraObjectId='existing'
            $user = [pscustomobject]@{ Id='existing'; GivenName='Lea'; Surname='Testwald'; DisplayName='Lea Testwald'
                UserPrincipalName='custom@monteaufkirchen.com'; Mail='different@monteaufkirchen.com'; MailNickname='custom-alias'
                AccountEnabled=$true; UserType='Member'; AssignedLicenses=@('sku'); LicenseAssignmentStates=@()
                LegalAgeGroupClassification='Adult'; ProxyAddresses=@('smtp:old@example.invalid') }
            $script:creationSnapshot.UsersById[$user.Id]=$user
            [void]$script:creationSnapshot.PersonnelMemberIds.Add($user.Id)
            $script:creationSnapshot.DirectGroupsByUserId[$user.Id]=@($script:creationSnapshot.GroupsById[$script:creationConfig.Roles.School.Id])
            $result = Compare-LehrerDirectory -People @($person) -Snapshot $script:creationSnapshot -Config $script:creationConfig
            $result.Errors | Should -BeNullOrEmpty
            $desired = $result.ChangedPeople[0].DesiredState
            $desired.UserPrincipalName | Should -Be $user.UserPrincipalName
            $desired.Mail | Should -Be $user.Mail
            $desired.MailNickname | Should -Be $user.MailNickname
            $desired.Exchange.ContainsKey('PrimarySmtpAddress') | Should -BeFalse
            $desired.Exchange.ContainsKey('RequiredAliases') | Should -BeFalse
        }
        It 'reserves primary and alias addresses together and falls back to numeric suffixes' {
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($upn in @(Get-UpnCandidate Lea Testwald montessori-erding.de)) { [void]$used.Add($upn) }
            $selection = Select-AvailableUpn -GivenName Lea -Surname Testwald -Domain montessori-erding.de -AliasDomains @('monteaufkirchen.com') -UsedAddresses $used -AddressOwners @{}
            $selection.Upn | Should -Be 'leatestwald2@montessori-erding.de'
            $used.Contains('leatestwald2@monteaufkirchen.com') | Should -BeTrue
        }
        It 'normalizes external names without losing the separating dot' {
            $used = [Collections.Generic.HashSet[string]]::new()
            $selection = Select-AvailableUpn -GivenName 'Änne' -Surname 'Weiß' -Domain montessori-erding.de -NamingStyle FirstLast -UsedAddresses $used -AddressOwners @{}
            $selection.Upn | Should -Be 'aenne.weiss@montessori-erding.de'
        }
        It 'merges new mailbox addresses while retaining SMTP aliases and non SMTP entries' {
            $config = @{} + $script:creationConfig.Exchange
            $config.AddressBookPolicy='test-policy'; $config.CustomAttribute1='test-staff'
            $config.PrimarySmtpAddress='ltestwald@montessori-erding.de'; $config.RequiredAliases=@('ltestwald@monteaufkirchen.com')
            $mailbox = [pscustomobject]@{ PrimarySmtpAddress='old@example.invalid'; EmailAddresses=@('SMTP:old@example.invalid','smtp:extra@example.invalid','X500:/o=synthetic') }
            $differences = @([pscustomobject]@{Field='PrimarySmtpAddress'},[pscustomobject]@{Field='EmailAddresses'})
            $parameters = Get-ExchangeMailboxWriteParameter -Mailbox $mailbox -Differences $differences -Config $config
            $parameters.EmailAddresses | Should -Contain 'SMTP:ltestwald@montessori-erding.de'
            $parameters.EmailAddresses | Should -Contain 'smtp:ltestwald@monteaufkirchen.com'
            $parameters.EmailAddresses | Should -Contain 'smtp:old@example.invalid'
            $parameters.EmailAddresses | Should -Contain 'smtp:extra@example.invalid'
            $parameters.EmailAddresses | Should -Contain 'X500:/o=synthetic'
            @($parameters.EmailAddresses | Where-Object { $_ -cmatch '^SMTP:' }).Count | Should -Be 1
            $parameters.ContainsKey('PrimarySmtpAddress') | Should -BeFalse
        }
    }
}
