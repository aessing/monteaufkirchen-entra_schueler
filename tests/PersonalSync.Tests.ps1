BeforeDiscovery {
    Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

BeforeAll {
    $global:PersonalStubCommands = [Collections.Generic.List[string]]::new()
    foreach ($name in @('Get-MgUser','Get-MgContext','Get-MgGroup','Get-MgGroupMember','Get-MgUserMemberOfAsGroup',
        'Get-MgUserTransitiveMemberOf','Get-MgUserOwnedObject','Get-EXORecipient','Update-MgUser','New-MgUser',
        'New-MgGroupMemberByRef','Remove-MgGroupMemberDirectoryObjectByRef','Remove-MgUser','Get-AddressBookPolicy','Set-Mailbox','Set-CASMailbox','Get-MgUserManager','Set-MgUserManagerByRef')) {
        if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
            Set-Item "function:global:$name" ([scriptblock]::Create('[CmdletBinding(SupportsShouldProcess)]param($UserId,$GroupId,$DirectoryObjectId,$BodyParameter,$AccountEnabled,$Identity,$CustomAttribute1,$MAPIEnabled,$ExternalDirectoryObjectId,[switch]$All,$Property,$Properties,$ResultSize)'))
            $global:PersonalStubCommands.Add($name)
        }
    }
}
AfterAll {
    foreach ($name in $global:PersonalStubCommands) { Remove-Item "function:global:$name" -ErrorAction SilentlyContinue }
    Remove-Variable PersonalStubCommands -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Personal 0.3.0: inventory, guards, sync and actions' {
    InModuleScope SchuelerSync {
        BeforeAll {
            function New-PersonalSyntheticUser {
                param([string]$Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa', [bool]$Enabled = $true)
                [pscustomobject]@{
                    Id=$Id; DisplayName='Lea Testwald'; GivenName='Lea'; Surname='Testwald'
                    UserPrincipalName='ltestwald@example.invalid'; Mail='ltestwald@example.invalid'; MailNickname='ltestwald'
                    ProxyAddresses=@(); AccountEnabled=$Enabled; UserType='Member'; OnPremisesSyncEnabled=$false
                    CompanyName='Montessori Schule Aufkirchen'; EmployeeType='Sekretariat'; JobTitle='Schule Sekretariat'
                    AgeGroup='Adult'; LegalAgeGroupClassification='Adult'; ConsentProvidedForMinor='NotRequired'; UsageLocation='DE'
                    AssignedLicenses=@(); LicenseAssignmentStates=@(); SignInActivity=[pscustomobject]@{LastSuccessfulSignInDateTime=$null}
                }
            }
            function New-PersonalSyntheticPerson {
                param([string]$Job='Schule_Sekretariat', [string]$Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa')
                [pscustomobject]@{RowNumber=2;NameMitRufname='Testwald, Lea';GivenName='Lea';Surname='Testwald';Job=$Job
                    EntraObjectId=$Id;StoredUpn='';StoredMail='';Password='';ResumeCreation=$false}
            }
            function Get-PersonalSyntheticSnapshot {
                $snapshot = Get-LehrerSnapshot -Config $script:PersonalConfig
                $snapshot | Add-Member -NotePropertyMembers @{PersonalContext=$script:PersonalConfig.SafetyContext
                    ManagedRoleIds=$script:PersonalConfig.SafetyContext.ManagedRoleIds
                    RecipientIndex=(Get-AccountRecipientIndex -Recipients $script:PersonalRecipients)}
                return $snapshot
            }
            function Set-StudentTuiFixture {
                $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
                $script:PersonalMemberships[$id]=@((Get-ConfiguredStudentRole).Id,'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee')
                $script:PersonalUsers[$id] | Add-Member -NotePropertyMembers @{Department='JK1-3g2_1';OfficeLocation='g2'}
                foreach ($entry in @(@{Id='eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';Name='SEC-A-CLS-JK1-3g2_1'},@{Id='ffffffff-ffff-ffff-ffff-ffffffffffff';Name='SEC-A-CLS-JK4-6m1_1'})) {
                    $script:PersonalGroups[$entry.Id]=[pscustomobject]@{Id=$entry.Id;DisplayName=$entry.Name;OdataType='#microsoft.graph.group';GroupTypes=@();MembershipRule=$null;AssignedLicenses=@()}
                }
                $teacher=New-PersonalSyntheticUser -Id '11111111-1111-1111-1111-111111111111'
                $teacher.DisplayName='Nora Lehrwald';$teacher.UserPrincipalName='nlehrwald@example.invalid'
                $script:PersonalUsers[$teacher.Id]=$teacher
                $script:PersonalMemberships[$teacher.Id]=@('103a1c4c-036b-40bc-bbe8-e3f7a866cb01')
                $second=New-PersonalSyntheticUser -Id '22222222-2222-2222-2222-222222222222'
                $second.DisplayName='Zoe Lehrwald';$second.UserPrincipalName='zlehrwald@example.invalid'
                $script:PersonalUsers[$second.Id]=$second;$script:PersonalMemberships[$second.Id]=@('103a1c4c-036b-40bc-bbe8-e3f7a866cb01')
                $script:StudentTuiManagers=@{$id=$teacher.Id}
                Mock Get-MgUserManager { if ($script:StudentTuiManagers.ContainsKey([string]$UserId)) { [pscustomobject]@{Id=$script:StudentTuiManagers[[string]$UserId]} } else {throw 'synthetic manager unavailable'} }
                Mock Set-MgUserManagerByRef { $script:StudentTuiManagers[[string]$UserId]=($BodyParameter['@odata.id'] -split '/')[-1] }
            }
        }
        BeforeEach {
            $script:PersonalConfig=Get-PersonalConfig
            $script:PersonalGroups=@{}
            $teacherConfig=Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            foreach ($definition in @($script:PersonalConfig.Roles.Values)+@($teacherConfig.Roles.Values)+@(Get-ConfiguredStudentRole)+@(
                @{Id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';Name='SEC-A-ADM-GlobalAdministrators'},
                @{Id='cccccccc-cccc-cccc-cccc-cccccccccccc';Name='SEC-A-LIC-O365A1Faculty'},
                @{Id='dddddddd-dddd-dddd-dddd-dddddddddddd';Name='SEC-A-ROL-SyntheticOther'})) {
                $script:PersonalGroups[[string]$definition.Id]=[pscustomobject]@{Id=$definition.Id;DisplayName=$definition.Name
                    OdataType='#microsoft.graph.group';GroupTypes=@();MembershipRule=$null;AssignedLicenses=@()}
            }
            $script:PersonalGroups['cccccccc-cccc-cccc-cccc-cccccccccccc'].AssignedLicenses=@([pscustomobject]@{SkuId='synthetic-sku'})
            $script:PersonalConfig.SessionTenantId='synthetic-tenant'
            $script:PersonalConfig.SafetyContext=Get-PersonalProtectionContext -GroupsById $script:PersonalGroups -Config $script:PersonalConfig
            $script:PersonalUsers=@{'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'=(New-PersonalSyntheticUser)}
            $script:PersonalMemberships=@{'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'=@($script:PersonalConfig.Roles.Schule_Sekretariat.Id)}
            $script:PersonalTransitive=@{}
            $script:PersonalRecipients=@()
            $script:PersonalPeople=@(New-PersonalSyntheticPerson)
            Mock Connect-SchuelerGraph { [pscustomobject]@{TenantId='synthetic-tenant';Account='operator@example.invalid'} }
            Mock Connect-DisabledAccountGraph { [pscustomobject]@{TenantId='synthetic-tenant'} }
            Mock Get-MgContext { [pscustomobject]@{TenantId='synthetic-tenant'} }
            Mock Connect-SchuelerExchangeOnline { }
            Mock Assert-ExchangeTenant { }
            Mock Get-MgGroup { if ($GroupId) { $script:PersonalGroups[[string]$GroupId] } else { $script:PersonalGroups.Values } }
            Mock Get-MgGroupMember { foreach($key in $script:PersonalMemberships.Keys) { if ($GroupId -in $script:PersonalMemberships[$key]) { $script:PersonalUsers[$key] } } }
            Mock Get-MgUser { if ($UserId) { if ($script:PersonalUsers.ContainsKey([string]$UserId)) { $script:PersonalUsers[[string]$UserId] } else { $script:PersonalUsers.Values | Where-Object UserPrincipalName -eq $UserId } } else { $script:PersonalUsers.Values } }
            Mock Get-MgUserMemberOfAsGroup { foreach($id in $script:PersonalMemberships[[string]$UserId]) { $script:PersonalGroups[$id] } }
            Mock Get-MgUserTransitiveMemberOf { foreach($id in $script:PersonalMemberships[[string]$UserId]) { $script:PersonalGroups[$id] }; if ($script:PersonalTransitive.ContainsKey([string]$UserId)) { $script:PersonalTransitive[[string]$UserId] } }
            Mock Get-MgUserOwnedObject { }
            Mock Get-EXORecipient { if ($ExternalDirectoryObjectId) { $script:PersonalRecipients | Where-Object ExternalDirectoryObjectId -eq $ExternalDirectoryObjectId } else { $script:PersonalRecipients } }
            Mock Get-ExchangeRecipientAddress { [pscustomobject]@{AddressOwners=@{}} }
            Mock Assert-LehrerExchangePolicy { }
            Mock Get-StudentMailboxState { [pscustomobject]@{Exists=$true;Differences=@()} }
            Mock Write-LehrerReport { }
            Mock Read-StudentWorkbook { [pscustomobject]@{Students=$script:PersonalPeople;SourceHash='synthetic'} }
            Mock Assert-WorkbookSafeForPasswordWrite { }
            Mock Assert-StudentWorkbookVersion { }
            Mock Write-StudentWorkbookUpdate { [pscustomobject]@{SourceHash='synthetic-updated'} }
            Mock Update-MgUser {
                if ($null -ne $BodyParameter) { foreach($key in @('DisplayName','GivenName','Surname','JobTitle','CompanyName','EmployeeType','UsageLocation','AgeGroup','ConsentProvidedForMinor','Department','OfficeLocation')) { if ($null -ne $BodyParameter.$key) { $script:PersonalUsers[[string]$UserId].$key=$BodyParameter.$key } } }
                else { $script:PersonalUsers[[string]$UserId].AccountEnabled=[bool]$AccountEnabled }
            }
            Mock New-MgGroupMemberByRef { $id=($BodyParameter['@odata.id'] -split '/')[-1]; $script:PersonalMemberships[$id]=@($script:PersonalMemberships[$id])+[string]$GroupId }
            Mock Remove-MgGroupMemberDirectoryObjectByRef { $script:PersonalMemberships[[string]$DirectoryObjectId]=@($script:PersonalMemberships[[string]$DirectoryObjectId] | Where-Object {$_ -ne $GroupId}) }
            Mock Connect-DisabledAccountWriteGraph { }
            Mock Connect-DisabledAccountDeleteGraph { }
            Mock Revoke-EntraStudentSession { $true }
            Mock Remove-MgUser { $script:PersonalUsers.Remove([string]$UserId) }
            Mock Write-Progress { }
            Mock Clear-Host { }
        }

        It 'exports the actual personnel login and warns when sign-in data is unavailable' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].SignInActivity.LastSuccessfulSignInDateTime='2026-01-02T12:30:00+02:00'
            $path=Join-Path $TestDrive 'Personal-Login.xlsx'
            $inventory=Get-PersonalInventory
            $null=Export-PersonalWorkbook -Inventory $inventory -Path $path -Confirm:$false
            $package=Open-StudentWorkbookPackage $path
            try { $package.Workbook.Worksheets['Tabelle1'].Cells[2,9].Value | Should -BeExactly '2026-01-02T10:30:00Z' } finally {$package.Dispose()}
            Mock Get-MgUser { if ($Property -contains 'signInActivity') { throw 'synthetic sign-in failure' }; $script:PersonalUsers.Values }
            Mock Write-Warning { }
            $null=Invoke-PersonalSync -ListToExcel -File (Join-Path $TestDrive 'Personal-NoLogin.xlsx') -Confirm:$false
            Should -Invoke Write-Warning -ParameterFilter { $Message -like '*Letzte Anmeldung nicht abrufbar*synthetic sign-in failure*' } -Exactly 1
        }
        It 'maps each released speaking job to its own profile' -ForEach @(
            @{Job='Schule_Sekretariat';Company='Montessori Schule Aufkirchen';Type='Sekretariat'},
            @{Job='Schule_Vertretungskräfte';Company='Montessori Schule Aufkirchen';Type='Vertretungskräfte'},
            @{Job='Kinderhaus_PädagogischesTeam';Company='Montessori Kinderhaus Erding';Type='Pädagogisches Team'},
            @{Job='Kinderhaus_Sekretariat';Company='Montessori Kinderhaus Erding';Type='Sekretariat'},
            @{Job='Kinderhaus_Vertretungskräfte';Company='Montessori Kinderhaus Erding';Type='Vertretungskräfte'},
            @{Job='Unterstützung';Company='Montessori Verein Landkreis Erding e.V.';Type='Unterstützung'},
            @{Job='ExterneBenutzer';Company='Montessori Verein Landkreis Erding e.V.';Type='Externe Benutzer'}
        ) {
            $p=Get-LehrerProfile -Job " $($Job.ToLowerInvariant()) " -Config $script:PersonalConfig
            $p.Job | Should -BeExactly $Job
            $p.CompanyName | Should -BeExactly $Company
            $p.EmployeeType | Should -BeExactly $Type
            $p.Role.Id | Should -Be $script:PersonalConfig.Roles[$Job].Id
            $p.LicenseGroupName | Should -Be 'SEC-A-LIC-O365A1Faculty'
            $p.Exchange.AddressBookPolicy | Should -Be "MON-EXO-ABP-$Job"
        }
        It 'resolves readable <Label> to the existing personnel profile' -ForEach @(
            @{Job='Schule_Sekretariat';Label='Schule Sekretariat'},
            @{Job='Schule_Vertretungskräfte';Label='Schule Vertretungskräfte'},
            @{Job='Kinderhaus_PädagogischesTeam';Label='Kinderhaus Pädagogisches Team'},
            @{Job='Kinderhaus_Sekretariat';Label='Kinderhaus Sekretariat'},
            @{Job='Kinderhaus_Vertretungskräfte';Label='Kinderhaus Vertretungskräfte'},
            @{Job='Unterstützung';Label='Unterstützung'},
            @{Job='ExterneBenutzer';Label='Externe Benutzer'}
        ) {
            foreach ($inputJob in @($Job, $Label, ('  ' + $Label.ToLowerInvariant().Replace(' ', '   ') + '  '))) {
                $profile=Get-LehrerProfile -Job $inputJob -Config $script:PersonalConfig
                $profile.Job | Should -BeExactly $Job
                $profile.Role.Id | Should -BeExactly $script:PersonalConfig.Roles[$Job].Id
                $profile.Exchange.AddressBookPolicy | Should -BeExactly $script:PersonalConfig.Profiles[$Job].AddressBookPolicy
            }
        }
        It 'uses <Title> for personnel JobTitle on creation and update' -ForEach @(
            @{Job='Schule_Sekretariat';Label='Schule Sekretariat';Title='Schule Sekretariat'},
            @{Job='Schule_Vertretungskräfte';Label='Schule Vertretungskräfte';Title='Schule Vertretungskräfte'},
            @{Job='Kinderhaus_PädagogischesTeam';Label='Kinderhaus Pädagogisches Team';Title='Pädagogisches Team'},
            @{Job='Kinderhaus_Sekretariat';Label='Kinderhaus Sekretariat';Title='Sekretariat'},
            @{Job='Kinderhaus_Vertretungskräfte';Label='Kinderhaus Vertretungskräfte';Title='Vertretungskräfte'},
            @{Job='Unterstützung';Label='Unterstützung';Title='Unterstützung'},
            @{Job='ExterneBenutzer';Label='Externe Benutzer';Title='Externe Benutzer'}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $profile=Get-LehrerProfile -Job $Job -Config $script:PersonalConfig
            $user=$script:PersonalUsers[$id]
            $user.CompanyName=$profile.CompanyName; $user.EmployeeType=$profile.EmployeeType
            $script:PersonalMemberships[$id]=@($profile.Role.Id,'cccccccc-cccc-cccc-cccc-cccccccccccc')
            $person=New-PersonalSyntheticPerson -Job $Job
            foreach ($previousTitle in @($Job,$Label)) {
                $user.JobTitle=$previousTitle
                $comparison=Compare-LehrerDirectory -People @($person) -Snapshot (Get-PersonalSyntheticSnapshot) -Config $script:PersonalConfig
                $comparison.Errors | Should -BeNullOrEmpty
                $entry=@($comparison.ChangedPeople)+@($comparison.ExistingPeople)
                $entry.Count | Should -Be 1
                $entry[0].DesiredState.JobTitle | Should -BeExactly $Title
                $entry[0].DesiredState.CompanyName | Should -BeExactly $profile.CompanyName
                $entry[0].DesiredState.Role.Id | Should -BeExactly $profile.Role.Id
                @($entry[0].Differences | Where-Object Field -in @('UserPrincipalName','Mail','MailNickname')).Count | Should -Be 0
                (Invoke-LehrerUpdate -Entry $entry[0] -Config $script:PersonalConfig -Confirm:$false).Status | Should -Be Succeeded
                $user.JobTitle | Should -BeExactly $Title
                $user.CompanyName | Should -BeExactly $profile.CompanyName
                $unchanged=Compare-LehrerDirectory -People @($person) -Snapshot (Get-PersonalSyntheticSnapshot) -Config $script:PersonalConfig
                $unchanged.Errors | Should -BeNullOrEmpty
                $unchanged.ChangedPeople.Count | Should -Be 0
                $unchanged.ExistingPeople.Count | Should -Be 1
            }

            $script:PersonalUsers=@{}; $script:PersonalMemberships=@{}
            $newPerson=New-PersonalSyntheticPerson -Job $Label -Id ''
            $newComparison=Compare-LehrerDirectory -People @($newPerson) -Snapshot (Get-PersonalSyntheticSnapshot) -Config $script:PersonalConfig
            $newComparison.Errors | Should -BeNullOrEmpty
            $newComparison.NewPeople.Count | Should -Be 1
            $newComparison.NewPeople[0].DesiredState.JobTitle | Should -BeExactly $Title
            Mock New-MgUser { [pscustomobject]@{Id='new-synthetic-person'} }
            $null=New-DisabledEntraLehrer -Desired $newComparison.NewPeople[0].DesiredState -Password 'SyntheticPassw0rd' -Confirm:$false
            Should -Invoke New-MgUser -Exactly 1 -ParameterFilter {$BodyParameter.JobTitle -ceq $Title -and $BodyParameter.CompanyName -ceq $profile.CompanyName -and $BodyParameter.AccountEnabled -eq $false}
        }
        It 'shows readable jobs in the personnel comparison without changing profile keys' {
            $script:PersonalUsers=@{}; $script:PersonalMemberships=@{}
            $person=New-PersonalSyntheticPerson -Job Kinderhaus_PädagogischesTeam -Id ''
            $comparison=Compare-LehrerDirectory -People @($person) -Snapshot (Get-PersonalSyntheticSnapshot) -Config $script:PersonalConfig
            $safe=ConvertTo-SafeLehrerComparison -Comparison $comparison
            $safe.NewPeople[0].Job | Should -BeExactly 'Kinderhaus Pädagogisches Team'
            $person.Job | Should -BeExactly 'Kinderhaus_PädagogischesTeam'
        }
        It 'shows readable jobs in the personnel inventory report without changing profile keys' {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalMemberships[$id]=@($script:PersonalConfig.Roles.Kinderhaus_PädagogischesTeam.Id)
            $inventory=Get-PersonalInventory
            Mock Write-Information { }
            Write-PersonalInventoryReport -Inventory $inventory
            Should -Invoke Write-Information -Exactly 1 -ParameterFilter {$MessageData -match 'Job\s*:\s*Kinderhaus Pädagogisches Team'}
            $inventory.Accounts[0].Job | Should -BeExactly 'Kinderhaus_PädagogischesTeam'
        }
        It 'rejects protected, unknown and competing readable jobs' -ForEach @(
            @{Job='Externe Admins'},@{Job='  externe   admins  '},@{Job='Schule Verwaltung'},
            @{Job='Kinderhaus Pädagogische Team'},@{Job='Schule Sekretariat, Unterstützung'}
        ) {
            {Get-LehrerProfile -Job $Job -Config $script:PersonalConfig} | Should -Throw
        }
        It 'rejects an ambiguous display name instead of choosing a personnel role' -ForEach @(
            @{Job='Externe Benutzer'},@{Job='ExterneBenutzer'},@{Job='externebenutzer'},@{Job='EXTERNEBENUTZER'}
        ) {
            $script:PersonalConfig.Jobs.Externe_Benutzer='Unterstützung'
            {Get-LehrerProfile -Job $Job -Config $script:PersonalConfig} | Should -Throw '*mehrdeutig*'
        }
        It 'rejects teacher codes, protected jobs and competing profiles' -ForEach @(@{Job='L'},@{Job='OGTS'},@{Job='ExterneAdmins'},@{Job='GlobalAdministrators'},@{Job='Schule_Sekretariat, Unterstützung'}) {
            {Get-LehrerProfile -Job $Job -Config $script:PersonalConfig} | Should -Throw
        }
        It 'excludes transitive teachers, Ganztag and pupils from the complete inventory' -ForEach @(@{Index=0},@{Index=1},@{Index=2}) {
            $id=$script:PersonalConfig.SafetyContext.ExcludedIds[$Index]
            $script:PersonalTransitive['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@($script:PersonalGroups[$id])
            $result=Get-PersonalInventory
            $result.Accounts.Count | Should -Be 0
            $result.ExcludedCount | Should -Be 1
        }
        It 'excludes shared mailboxes even behind an ambiguous recipient' {
            $script:PersonalRecipients=@(
                [pscustomobject]@{ExternalDirectoryObjectId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';RecipientTypeDetails='UserMailbox'},
                [pscustomobject]@{ExternalDirectoryObjectId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';RecipientTypeDetails='SharedMailbox'})
            (Get-PersonalInventory).Accounts.Count | Should -Be 0
        }
        It 'lists active and disabled roleless accounts with an empty job' {
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@()
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee']=New-PersonalSyntheticUser -Id 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee' -Enabled:$false
            $result=Invoke-PersonalSync -List -PassThru
            $result.Accounts.Count | Should -Be 2
            @($result.Accounts | Where-Object Job).Count | Should -Be 0
            @($result.Accounts | Where-Object MutationAllowed).Count | Should -Be 2
            $result.PSObject.Properties.Name | Should -Not -Contain 'Config'
        }
        It 'shows administrator groups but protects every action' -ForEach @(@{Name='SEC-A-ROL-ExterneAdmins'},@{Name='SEC-A-ADM-GlobalAdministrators'}) {
            $group=$script:PersonalGroups.Values | Where-Object DisplayName -eq $Name
            $script:PersonalTransitive['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@($group)
            $result=Get-PersonalInventory
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].MutationAllowed | Should -BeFalse
            (Test-DisabledAccountSelectable $result.Accounts[0]) | Should -BeFalse
            {Assert-PersonalMutationTarget -UserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ExpectedUpn 'ltestwald@example.invalid' -Config $result.Config} | Should -Throw '*geschützt*'
        }
        It 'protects the ADM administrator group in <Population> management' -ForEach @(@{Population='Personnel'},@{Population='Teacher'},@{Population='Student'}) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            if ($Population -eq 'Teacher') {$script:PersonalMemberships[$id]=@('103a1c4c-036b-40bc-bbe8-e3f7a866cb01')}
            elseif ($Population -eq 'Student') {$script:PersonalMemberships[$id]=@((Get-ConfiguredStudentRole).Id)}
            $script:PersonalMemberships[$id]+='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $result=Get-PersonalInventory -Population $Population
            $result.Config.SafetyContext.Complete | Should -BeTrue
            $result.Accounts.Count | Should -Be 1
            $result.Accounts[0].SecAGroups | Should -Contain 'SEC-A-ADM-GlobalAdministrators'
            $result.Accounts[0].MutationAllowed | Should -BeFalse
            $result.Accounts[0].ProtectionReasons -join ' ' | Should -Match 'SEC-A-ADM-GlobalAdministrators'
            {Assert-PersonalMutationTarget -UserId $id -ExpectedUpn 'ltestwald@example.invalid' -Config $result.Config} | Should -Throw '*geschützt*'
        }
        It 'does not accept the obsolete role name as the current administrator protection group' {
            $obsoleteName='SEC-A-ROL-'+'GlobalAdministrators'
            $script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'].DisplayName=$obsoleteName
            $script:PersonalConfig.ProtectedGroupNames | Should -Contain 'SEC-A-ADM-GlobalAdministrators'
            $script:PersonalConfig.ProtectedGroupNames | Should -Not -Contain $obsoleteName
            $result=Get-PersonalInventory
            $result.Config.SafetyContext.Complete | Should -BeFalse
            $result.Accounts[0].MutationAllowed | Should -BeFalse
            $result.Warnings -join ' ' | Should -Match 'SEC-A-ADM-GlobalAdministrators'
        }
        It 'blocks writes if the ADM administrator group name is ambiguous' {
            $duplicate=$script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'] | Select-Object *
            $duplicate.Id='dddddddd-dddd-dddd-dddd-dddddddddddd'
            $script:PersonalGroups[$duplicate.Id]=$duplicate
            $result=Get-PersonalInventory
            $result.Config.SafetyContext.Complete | Should -BeFalse
            $result.Accounts[0].MutationAllowed | Should -BeFalse
            (Invoke-PersonalSync -Remove -UPN 'ltestwald@example.invalid' -WhatIf).HasErrors | Should -BeTrue
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'fails closed for writes while still listing when an admin group is missing' {
            $script:PersonalGroups.Remove('bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb')
            $result=Get-PersonalInventory
            $result.Accounts[0].MutationAllowed | Should -BeFalse
            $result.Warnings.Count | Should -Be 1
            (Invoke-PersonalSync -Remove -UPN 'ltestwald@example.invalid' -WhatIf).HasErrors | Should -BeTrue
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'protects locally synchronized accounts while listing them' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].OnPremisesSyncEnabled=$true
            (Get-PersonalInventory).Accounts[0].MutationAllowed | Should -BeFalse
        }
        It 'blocks and then deletes a roleless guest through the refreshed TUI with recipient <Type>' -ForEach @(
            @{Type='None'},@{Type='MailUser'},@{Type='GuestMailUser'},@{Type='UserMailbox'}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalUsers[$id].UserType='Guest'
            $script:PersonalMemberships[$id]=@()
            if ($Type -ne 'None') { $script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails=$Type}) }
            $inventory=Get-PersonalInventory
            $inventory.Accounts[0].UserType | Should -Be Guest
            $inventory.Accounts[0].MutationAllowed | Should -BeTrue
            (Test-DisabledAccountSelectable $inventory.Accounts[0]) | Should -BeTrue
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar','B','Y','Spacebar','Backspace','Y','Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            $result=@(Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config)
            $result.Status | Should -Be @('Disabled','Deleted')
            Should -Invoke Update-MgUser -ParameterFilter {$UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $AccountEnabled -eq $false} -Exactly 1
            Should -Invoke Revoke-EntraStudentSession -Exactly 1
            Should -Invoke Remove-MgUser -Exactly 1
            Should -Invoke Show-DisabledAccountSelection -ParameterFilter {$Accounts.Count -eq 0} -Exactly 1
        }
        It 'keeps guest unlocking and ordinary profile mutations blocked' {
            Mock Read-DisabledAccountTuiKey { throw 'Unexpected guest unlock confirmation' }
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].UserType='Guest'
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=$false
            $inventory=Get-PersonalInventory
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation Unlock -Confirm:$false).Status | Should -Be Skipped
            {Assert-PersonalMutationTarget -UserId $inventory.Accounts[0].UserId -ExpectedUpn $inventory.Accounts[0].UPN -Config $inventory.Config -Account $inventory.Accounts[0]} | Should -Throw '*Kontotyp*'
            {Assert-PersonalSnapshotTarget -User $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'] -Snapshot (Get-PersonalSyntheticSnapshot)} | Should -Throw '*geschützt*'
            (Invoke-SelectedDisabledAccountDeletion -Accounts $inventory.Accounts -TenantId $inventory.TenantId -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Read-DisabledAccountTuiKey -Exactly 0
        }
        It 'omits the writing-import review warning for a guest with a valid export job' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].UserType='Guest'
            Mock Write-Warning { }
            $null=Invoke-PersonalSync -ListToExcel -File (Join-Path $TestDrive 'Personal-Guest.xlsx') -Confirm:$false
            Should -Invoke Write-Warning -Exactly 0 -ParameterFilter {$Message -like '*Exportzeile benötigt vor einem schreibenden Import eine Prüfung*'}
        }
        It 'retains export data warnings for roleless <Type> accounts as appropriate' -ForEach @(
            @{Type='Guest';ExpectedWarnings=0},@{Type='Member';ExpectedWarnings=1}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalUsers[$id].UserType=$Type
            $script:PersonalMemberships[$id]=@()
            Mock Write-Warning { }
            $null=Invoke-PersonalSync -ListToExcel -File (Join-Path $TestDrive "Roleless-$Type.xlsx") -Confirm:$false
            Should -Invoke Write-Warning -Exactly $ExpectedWarnings -ParameterFilter {$Message -like '*Exportzeile benötigt vor einem schreibenden Import eine Prüfung*'}
        }
        It 'preserves guest protections before <Operation> after <Change>' -ForEach @(
            @{Operation='Block';Change='Admin'},@{Operation='Delete';Change='Admin'},
            @{Operation='Block';Change='Excluded'},@{Operation='Delete';Change='Excluded'},
            @{Operation='Block';Change='DirectoryRole'},@{Operation='Delete';Change='DirectoryRole'},
            @{Operation='Block';Change='Shared'},@{Operation='Delete';Change='Shared'},
            @{Operation='Block';Change='Identity'},@{Operation='Delete';Change='Identity'},
            @{Operation='Block';Change='Type'},@{Operation='Delete';Change='Type'},
            @{Operation='Block';Change='LocalAD'},@{Operation='Delete';Change='LocalAD'},
            @{Operation='Block';Change='Ambiguous'},@{Operation='Delete';Change='Ambiguous'}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalUsers[$id].UserType='Guest'
            $script:PersonalUsers[$id].AccountEnabled=($Operation -eq 'Block')
            $inventory=Get-PersonalInventory
            switch($Change) {
                Admin {$script:PersonalTransitive[$id]=@($script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'])}
                Excluded {$script:PersonalTransitive[$id]=@($script:PersonalGroups[$script:PersonalConfig.SafetyContext.ExcludedIds[0]])}
                DirectoryRole {$script:PersonalTransitive[$id]=@([pscustomobject]@{Id='ffffffff-ffff-ffff-ffff-ffffffffffff';OdataType='#microsoft.graph.directoryRole';DisplayName='Synthetic admin'})}
                Shared {$script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails='SharedMailbox'})}
                Identity {$script:PersonalUsers[$id].UserPrincipalName='different@example.invalid'}
                Type {$script:PersonalUsers[$id].UserType='Member'}
                LocalAD {$script:PersonalUsers[$id].OnPremisesSyncEnabled=$true}
                Ambiguous {$script:PersonalRecipients=@(1..2 | ForEach-Object {[pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails='MailUser'}})}
            }
            Mock Read-DisabledAccountTuiKey {'Y'}
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'does not select a protected guest or authorize a member converted to a guest after selection' {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $inventory=Get-PersonalInventory
            $script:PersonalUsers[$id].UserType='Guest'
            (Invoke-PersonalAccountBlock -Accounts $inventory.Accounts -Config $inventory.Config -Confirm:$false).Status | Should -Be Skipped
            $script:PersonalTransitive[$id]=@($script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'])
            $protected=Get-PersonalInventory
            (Test-DisabledAccountSelectable $protected.Accounts[0]) | Should -BeFalse
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'protects resource mailboxes and ambiguous recipients' -ForEach @(@{Type='RoomMailbox';Count=1},@{Type='UserMailbox';Count=2}) {
            $script:PersonalRecipients=@(1..$Count | ForEach-Object {[pscustomobject]@{ExternalDirectoryObjectId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';RecipientTypeDetails=$Type}})
            (Get-PersonalInventory).Accounts[0].MutationAllowed | Should -BeFalse
        }
        It 'does not continue when exclusions cannot be read completely' {
            Mock Get-MgUserTransitiveMemberOf { throw 'synthetic membership failure' }
            {Get-PersonalInventory} | Should -Throw '*synthetic membership failure*'
            Should -Invoke Write-Progress -ParameterFilter {$Completed} -Exactly 1
        }
        It 'checks new excluded membership immediately before a block' {
            $inventory=Get-PersonalInventory
            $script:PersonalTransitive['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@($script:PersonalGroups[$script:PersonalConfig.SafetyContext.ExcludedIds[0]])
            $result=Invoke-PersonalAccountBlock -Accounts $inventory.Accounts -Config $inventory.Config -Confirm:$false
            $result.Status | Should -Be Skipped
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'checks protection again between mailbox and CAS writes' {
            $inventory=Get-PersonalInventory
            $config=(Get-LehrerProfile -Job Schule_Sekretariat -Config $inventory.Config).Exchange
            Mock Compare-StudentMailboxState {
                @([pscustomobject]@{Field='CustomAttribute1'},[pscustomobject]@{Field='MAPIEnabled'})
            }
            Mock Set-Mailbox {
                $script:PersonalTransitive['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@($script:PersonalGroups[$script:PersonalConfig.SafetyContext.ExcludedIds[0]])
            }
            Mock Set-CASMailbox {throw 'CAS write for excluded target'}
            $state=[pscustomobject]@{Exists=$true;Mailbox=[pscustomobject]@{PersistedCapabilities=@()};CasMailbox=[pscustomobject]@{}}
            $result=Set-StudentMailboxConfiguration -UserPrincipalName 'ltestwald@example.invalid' -Config $config -MailboxState $state `
                -SafetyConfig $inventory.Config -SafetyUserId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Confirm:$false
            $result.Status | Should -Be Failed
            $result.Changed | Should -BeTrue
            Should -Invoke Set-Mailbox -Exactly 1
            Should -Invoke Set-CASMailbox -Exactly 0
        }
        It 'marks non-writable personnel as review rows only in read-only comparison for <Reason>' -ForEach @(
            @{Reason='Admin'},@{Reason='OnPremises'},@{Reason='Shared'}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            switch ($Reason) {
                Admin {$script:PersonalMemberships[$id]+=$script:PersonalConfig.Roles.ExterneAdmins.Id}
                OnPremises {$script:PersonalUsers[$id].OnPremisesSyncEnabled=$true}
                Shared {$script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails='SharedMailbox'})}
            }
            $snapshot=Get-PersonalSyntheticSnapshot
            $people=@(New-PersonalSyntheticPerson)
            $result=Compare-LehrerDirectory -People $people -Snapshot $snapshot -Config $script:PersonalConfig -ReviewOnlyPersonnel
            $result.Errors.Count | Should -Be 0
            $result.Warnings.Count | Should -Be 1
            $result.Warnings[0].Field | Should -Be Review
            $result.NewPeople.Count | Should -Be 0
            $result.Departures.Count | Should -Be 0
            $result.ChangedPeople.Count | Should -Be 0
            $strict=Compare-LehrerDirectory -People $people -Snapshot $snapshot -Config $script:PersonalConfig
            $strict.Errors.Count | Should -Be 1
        }
        It 'ignores incomplete guest rows identified by <Identity> with OtherMember <OtherMember> without sync issues' -ForEach @(
            @{Identity='Id';OtherMember=$false},@{Identity='Upn';OtherMember=$false},@{Identity='Both';OtherMember=$false},@{Identity='Both';OtherMember=$true}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $user=$script:PersonalUsers[$id]
            $user.UserType='Guest'; $user.GivenName=''; $user.Surname=''
            $person=New-PersonalSyntheticPerson
            $person.GivenName=''; $person.Surname=''; $person.Job=''
            $person.EntraObjectId=if ($Identity -eq 'Upn') { '' } else { $id }
            $person.StoredUpn=if ($Identity -eq 'Id') { '' } else { $user.UserPrincipalName.ToUpperInvariant() }
            $person | Add-Member -NotePropertyName ReviewReason -NotePropertyValue 'Unvollständige Personaldaten'
            if ($OtherMember) {
                $member=New-PersonalSyntheticUser -Id 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
                $member.UserPrincipalName='member@example.invalid'
                $script:PersonalUsers[$member.Id]=$member
                $script:PersonalMemberships[$member.Id]=@($script:PersonalConfig.Roles.Schule_Sekretariat.Id)
            }
            $result=Compare-LehrerDirectory -People @($person) -Snapshot (Get-PersonalSyntheticSnapshot) -Config $script:PersonalConfig -ReviewOnlyPersonnel
            $result.Warnings | Should -BeNullOrEmpty
            $result.Errors | Should -BeNullOrEmpty
            $result.NewPeople | Should -BeNullOrEmpty
            $result.ChangedPeople | Should -BeNullOrEmpty
            $result.ExistingPeople | Should -BeNullOrEmpty
            $result.Departures | Should -BeNullOrEmpty
        }
        It 'updates members while skipping incomplete guests without guest mutations' {
            $guest=New-PersonalSyntheticUser -Id 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $guest.UserPrincipalName='guest@example.invalid'; $guest.UserType='Guest'; $guest.GivenName=''; $guest.Surname=''
            $script:PersonalUsers[$guest.Id]=$guest
            $script:PersonalMemberships[$guest.Id]=@($script:PersonalConfig.Roles.ExterneBenutzer.Id)
            $guestPerson=New-PersonalSyntheticPerson -Id $guest.Id
            $guestPerson.GivenName=''; $guestPerson.Surname=''; $guestPerson.Job=''
            $script:PersonalPeople+= $guestPerson
            Mock New-MgUser { throw 'Unexpected guest creation' }
            $result=Invoke-PersonalSync -File (Join-Path $TestDrive 'Synthetic-Mixed.xlsx') -Update -WhatIf
            $result.HasErrors | Should -BeFalse
            $result.Comparison.Warnings | Should -BeNullOrEmpty
            @($result.Comparison.ChangedPeople).Count | Should -Be 1
            $result.Comparison.ChangedPeople[0].UserId | Should -BeExactly 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            @($result.Actions | Where-Object UserId -eq $guest.Id).Count | Should -Be 0
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke New-MgUser -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
        }
        It 'updates valid personnel while skipping protected admins identified by <Identity>' -ForEach @(
            @{Identity='Id'},@{Identity='Upn'},@{Identity='Both'}
        ) {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalUsers[$id].JobTitle='Schule_Sekretariat'
            $admin=New-PersonalSyntheticUser -Id 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $admin.GivenName='Eva'; $admin.Surname='Adminwald'; $admin.DisplayName='Eva Adminwald'
            $admin.UserPrincipalName='admin@example.invalid'; $admin.JobTitle='ExterneAdmins'
            $script:PersonalUsers[$admin.Id]=$admin
            $script:PersonalMemberships[$admin.Id]=@($script:PersonalConfig.Roles.ExterneAdmins.Id)
            $adminPerson=New-PersonalSyntheticPerson -Job 'Externe Admins' -Id $admin.Id
            $adminPerson.GivenName='Eva'; $adminPerson.Surname='Adminwald'; $adminPerson.NameMitRufname='Adminwald, Eva'
            $adminPerson.EntraObjectId=if ($Identity -eq 'Upn') { '' } else { $admin.Id }
            $adminPerson.StoredUpn=if ($Identity -eq 'Id') { '' } else { $admin.UserPrincipalName.ToUpperInvariant() }
            $script:PersonalPeople+= $adminPerson
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            @($result.Comparison.ChangedPeople).Count | Should -Be 1
            $result.Comparison.ChangedPeople[0].UserId | Should -BeExactly $id
            @($result.Comparison.Warnings).Count | Should -Be 1
            $result.Comparison.Warnings[0].Code | Should -BeExactly 'Workbook.Review'
            $result.Actions | Where-Object Phase -eq Update | Select-Object -ExpandProperty Status | Should -BeExactly Succeeded
            Should -Invoke Update-MgUser -Exactly 1 -ParameterFilter {$UserId -eq $id -and $BodyParameter.JobTitle -ceq 'Schule Sekretariat'}
            Should -Invoke Update-MgUser -Exactly 0 -ParameterFilter {$UserId -eq $admin.Id}
            $admin.JobTitle | Should -BeExactly ExterneAdmins
            $script:PersonalUsers[$id].JobTitle | Should -BeExactly 'Schule Sekretariat'
            $comparison=Compare-LehrerDirectory -People $script:PersonalPeople -Snapshot (Get-PersonalSyntheticSnapshot) -Config $script:PersonalConfig -ReviewOnlyPersonnel
            $comparison.Errors | Should -BeNullOrEmpty
            $comparison.ChangedPeople | Should -BeNullOrEmpty
            $comparison.ExistingPeople.Count | Should -Be 1
        }
        It 'does not derive member departures from a workbook containing only protected admins' {
            $admin=New-PersonalSyntheticUser -Id 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $admin.UserPrincipalName='admin@example.invalid'
            $script:PersonalUsers[$admin.Id]=$admin
            $script:PersonalMemberships[$admin.Id]=@($script:PersonalConfig.Roles.ExterneAdmins.Id)
            $script:PersonalPeople=@(New-PersonalSyntheticPerson -Job 'Externe Admins' -Id $admin.Id)
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -DisableUsers -RevokeSessions -Confirm:$false
            $result.HasErrors | Should -BeFalse
            $result.Comparison.Departures | Should -BeNullOrEmpty
            $result.Actions | Should -BeNullOrEmpty
            $result.Comparison.Warnings.Count | Should -Be 1
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'does not report an explicitly identified review row as a departure' {
            $snapshot=Get-PersonalSyntheticSnapshot
            $person=New-PersonalSyntheticPerson
            $person.Job=''
            $person | Add-Member -NotePropertyName ReviewReason -NotePropertyValue 'Job fehlt, nur zur Prüfung.'
            $result=Compare-LehrerDirectory -People @($person) -Snapshot $snapshot -Config $script:PersonalConfig -ReviewOnlyPersonnel
            $result.Warnings.Count | Should -Be 1
            $result.Departures.Count | Should -Be 0
            $result.Errors.Count | Should -Be 0
        }
        It 'requires an explicit ID to take over a roleless account' {
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@()
            $person=New-PersonalSyntheticPerson -Id ''
            $snapshot=Get-PersonalSyntheticSnapshot
            {Resolve-LehrerIdentity -Person $person -Snapshot $snapshot -Config $script:PersonalConfig} | Should -Throw '*explizite EntraObjectId*'
            $person.EntraObjectId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            (Resolve-LehrerIdentity -Person $person -Snapshot $snapshot -Config $script:PersonalConfig).Id | Should -Be $person.EntraObjectId
        }
        It 'updates valid personnel while skipping an incomplete roleless row identified by <Identity>' -ForEach @(
            @{Identity='Id'},@{Identity='Upn'},@{Identity='Both'}
        ) {
            $roleless=New-PersonalSyntheticUser -Id 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $roleless.UserPrincipalName='device@example.invalid'; $roleless.GivenName=''; $roleless.Surname=''; $roleless.JobTitle=''
            $script:PersonalUsers[$roleless.Id]=$roleless
            $script:PersonalMemberships[$roleless.Id]=@()
            $person=New-PersonalSyntheticPerson -Job '' -Id $roleless.Id
            $person.RowNumber=18; $person.GivenName=''; $person.Surname=''
            $person.StoredUpn=if ($Identity -eq 'Id') {''} else {$roleless.UserPrincipalName}
            if ($Identity -eq 'Upn') {$person.EntraObjectId=''}
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].JobTitle='Schule_Sekretariat'
            $script:PersonalPeople=@((New-PersonalSyntheticPerson),$person)
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -Confirm:$false
            $result.HasErrors | Should -BeFalse
            $result.Comparison.Warnings.Count | Should -Be 1
            $result.Comparison.Warnings[0].Message | Should -Match 'Zeile 18.*Keine Personalrolle und kein Job'
            Should -Invoke Update-MgUser -ParameterFilter {$UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $BodyParameter.JobTitle -eq 'Schule Sekretariat'} -Exactly 1
            Should -Invoke Update-MgUser -ParameterFilter {$UserId -eq $roleless.Id} -Exactly 0
            $roleless.JobTitle | Should -BeNullOrEmpty
        }
        It 'does not derive member departures from a workbook containing only a roleless row without a job' {
            $roleless=New-PersonalSyntheticUser -Id 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $roleless.UserPrincipalName='device@example.invalid'
            $script:PersonalUsers[$roleless.Id]=$roleless; $script:PersonalMemberships[$roleless.Id]=@()
            $script:PersonalPeople=@(New-PersonalSyntheticPerson -Job '' -Id $roleless.Id)
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -DisableUsers -RevokeSessions -Confirm:$false
            $result.HasErrors | Should -BeFalse
            $result.Comparison.Departures | Should -BeNullOrEmpty
            $result.Actions | Should -BeNullOrEmpty
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'updates valid personnel while leaving known incomplete named accounts unchanged for <Identity>' -ForEach @(
            @{Identity='Id'},@{Identity='Upn'},@{Identity='Both'}
        ) {
            $id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
            $user=New-PersonalSyntheticUser -Id $id
            $user.UserPrincipalName='team@example.invalid'; $user.GivenName=''; $user.Surname=''
            $user.DisplayName='Synthetic team'; $user.JobTitle='Kinderhaus_PädagogischesTeam'
            $script:PersonalUsers[$id]=$user
            $script:PersonalMemberships[$id]=@($script:PersonalConfig.Roles.Kinderhaus_PädagogischesTeam.Id)
            $person=New-PersonalSyntheticPerson -Job 'Kinderhaus Pädagogisches Team' -Id $id
            $person.RowNumber=41; $person.NameMitRufname=$user.DisplayName; $person.GivenName=''; $person.Surname=''
            $person.StoredUpn=if ($Identity -eq 'Id') {''} else {$user.UserPrincipalName}
            if ($Identity -eq 'Upn') {$person.EntraObjectId=''}
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].JobTitle='Schule_Sekretariat'
            $script:PersonalPeople=@((New-PersonalSyntheticPerson),$person)
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -Confirm:$false
            $result.HasErrors | Should -BeFalse
            $result.Comparison.Warnings.Count | Should -Be 1
            $result.Comparison.Warnings[0].Message | Should -Match 'Zeile 41.*Unvollständige Namensdaten, nicht abgeglichen'
            $result.Comparison.Departures | Should -BeNullOrEmpty
            Should -Invoke Update-MgUser -ParameterFilter {$UserId -eq 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -and $BodyParameter.JobTitle -eq 'Schule Sekretariat'} -Exactly 1
            Should -Invoke Update-MgUser -ParameterFilter {$UserId -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'} -Exactly 0
            $user.JobTitle | Should -BeExactly 'Kinderhaus_PädagogischesTeam'
        }
        It 'does not derive member departures when all workbook rows have known incomplete names' {
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalUsers[$id].GivenName=''; $script:PersonalUsers[$id].Surname=''
            $person=New-PersonalSyntheticPerson
            $person.GivenName=''; $person.Surname=''
            $script:PersonalPeople=@($person)
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -DisableUsers -RevokeSessions -Confirm:$false
            $result.HasErrors | Should -BeFalse
            $result.Comparison.Departures | Should -BeNullOrEmpty
            $result.Actions | Should -BeNullOrEmpty
            $result.Comparison.Warnings.Count | Should -Be 1
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'preserves a disabled account during comparison and single maintenance' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=$false
            $result=Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -EntraObjectId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled | Should -BeFalse
            Should -Invoke Update-MgUser -ParameterFilter {$AccountEnabled -eq $true} -Exactly 0
            $result.Comparison.ChangedPeople[0].Differences.Field | Should -Not -Contain 'AccountEnabled'
        }
        It 'permits activation only for explicit matching creation recovery' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=$false
            $result=Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -EntraObjectId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ResumeCreation -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled | Should -BeTrue
            Should -Invoke Update-MgUser -ParameterFilter {$AccountEnabled -eq $true} -Exactly 1
        }
        It 'requires an ID for explicit recovery and rejects malformed object IDs at binding' {
            {Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -ResumeCreation} | Should -Throw '*EntraObjectId*'
            {Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -EntraObjectId ('-'*36)} | Should -Throw
        }
        It 'does not disable departures on an ordinary partial Update' {
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee']=New-PersonalSyntheticUser -Id 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'].GivenName='Nora'
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'].DisplayName='Nora Testwald'
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'].UserPrincipalName='ntestwald@example.invalid'
            $script:PersonalMemberships['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee']=@($script:PersonalConfig.Roles.Unterstützung.Id)
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $result.Comparison.Departures.Count | Should -Be 1
            Should -Invoke Update-MgUser -ParameterFilter {$AccountEnabled -eq $false} -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'allows explicit single Remove for a roleless account without removing licenses or groups' {
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@('dddddddd-dddd-dddd-dddd-dddddddddddd')
            $result=Invoke-PersonalSync -Remove -UserPrincipalName 'ltestwald@example.invalid' -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled | Should -BeFalse
            Should -Invoke Revoke-EntraStudentSession -Exactly 1
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
        }
        It 'keeps unrelated roles during profile changes' {
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']+= 'dddddddd-dddd-dddd-dddd-dddddddddddd'
            $result=Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Unterstützung -EntraObjectId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'] | Should -Contain 'dddddddd-dddd-dddd-dddd-dddddddddddd'
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'] | Should -Not -Contain $script:PersonalConfig.Roles.Schule_Sekretariat.Id
        }
        It 'preserves an already assigned license without adding standard licenses' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AssignedLicenses=@([pscustomobject]@{SkuId='other-synthetic'})
            $result=Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }
        It 'preserves a group license that has not propagated yet' {
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']+='cccccccc-cccc-cccc-cccc-cccccccccccc'
            $result=Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }
        It 'never creates an account or password under WhatIf' {
            $script:PersonalUsers=@{}; $script:PersonalMemberships=@{}
            Mock New-LehrerPassword { throw 'password generated under WhatIf' }
            Mock New-MgUser { throw 'write under WhatIf' }
            $result=Invoke-PersonalSync -Add -Vorname Lea -Nachname Testwald -Job Schule_Sekretariat -WhatIf
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $result.Actions.Status | Should -Contain 'WhatIf'
            Should -Invoke New-LehrerPassword -Exactly 0
            Should -Invoke New-MgUser -Exactly 0
        }
        It 'creates disabled, saves the early checkpoint and only then activates' {
            $script:PersonalUsers=@{}; $script:PersonalMemberships=@{}
            $person=New-PersonalSyntheticPerson -Id ''
            $script:PersonalPeople=@($person)
            $script:PersonalWriteOrder=[Collections.Generic.List[string]]::new()
            Mock New-MgUser {
                $script:PersonalWriteOrder.Add('CreateDisabled')
                $BodyParameter.AccountEnabled | Should -BeFalse
                $BodyParameter.JobTitle | Should -BeExactly 'Schule Sekretariat'
                $BodyParameter.PasswordProfile.ForceChangePasswordNextSignIn | Should -BeTrue
                $BodyParameter.PasswordProfile.Password.Length | Should -Be 12
                $user=New-PersonalSyntheticUser -Enabled:$false
                foreach($key in @('DisplayName','GivenName','Surname','UserPrincipalName','Mail','MailNickname','JobTitle','CompanyName','EmployeeType','UsageLocation','AgeGroup','ConsentProvidedForMinor')) { if ($null -ne $BodyParameter.$key) {$user.$key=$BodyParameter.$key} }
                $script:PersonalUsers[$user.Id]=$user; $script:PersonalMemberships[$user.Id]=@()
                $user
            }
            Mock Write-StudentWorkbookUpdate { $script:PersonalWriteOrder.Add('Checkpoint'); [pscustomobject]@{SourceHash='checkpoint'} }
            Mock Update-MgUser { $script:PersonalWriteOrder.Add('Enable'); $script:PersonalUsers[[string]$UserId].AccountEnabled=[bool]$AccountEnabled }
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -CreateNewUsers -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            @($script:PersonalWriteOrder) | Should -Be @('CreateDisabled','Checkpoint','Enable')
            Should -Invoke New-MgGroupMemberByRef -Exactly 2
            Should -Invoke Write-StudentWorkbookUpdate -ParameterFilter {$Kind -eq 'Personnel' -and $Updates[0].EntraObjectId -and $Updates[0].Password.Length -eq 12} -Exactly 1
        }
        It 'keeps a partial creation disabled and reports explicit recovery when checkpoint saving fails' {
            $script:PersonalUsers=@{}; $script:PersonalMemberships=@{}
            $script:PersonalPeople=@(New-PersonalSyntheticPerson -Id '')
            Mock New-MgUser { $user=New-PersonalSyntheticUser -Enabled:$false; $script:PersonalUsers[$user.Id]=$user; $user }
            Mock Write-StudentWorkbookUpdate {throw 'synthetic checkpoint failure'}
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -CreateNewUsers -Confirm:$false
            $result.HasErrors | Should -BeTrue
            $result.Actions[0].Phase | Should -Be Workbook
            $result.Actions[0].RecoveryCommand | Should -Match 'Sync-PersonalEntra.ps1.+ResumeCreation'
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled | Should -BeFalse
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }
        It 'disables only explicitly selected managed departures' {
            $script:PersonalPeople=@(New-PersonalSyntheticPerson -Id '')
            $script:PersonalPeople[0].GivenName='Nora'
            $script:PersonalPeople[0].NameMitRufname='Testwald, Nora'
            $result=Invoke-PersonalSync -File 'synthetic.xlsx' -Update -DisableUsers -Confirm:$false
            $result.HasErrors | Should -BeFalse -Because ($result | ConvertTo-Json -Depth 5 -Compress)
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled | Should -BeFalse
            Should -Invoke Update-MgUser -ParameterFilter {$AccountEnabled -eq $false} -Exactly 1
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'rechecks an assignment acquired after comparison instead of adding a second license' {
            $snapshot=Get-PersonalSyntheticSnapshot
            $comparison=Compare-LehrerDirectory -People $script:PersonalPeople -Snapshot $snapshot -Config $script:PersonalConfig
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AssignedLicenses=@([pscustomobject]@{SkuId='late-synthetic'})
            $entry=$comparison.ChangedPeople[0]
            $entry.DesiredState.AddLicenseGroup | Should -BeTrue
            (Invoke-LehrerUpdate -Entry $entry -Config $script:PersonalConfig -Confirm:$false).Status | Should -Be Succeeded
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }
        It 'rejects a standard group that lost its license after comparison' {
            $snapshot=Get-PersonalSyntheticSnapshot
            $comparison=Compare-LehrerDirectory -People $script:PersonalPeople -Snapshot $snapshot -Config $script:PersonalConfig
            $script:PersonalGroups['cccccccc-cccc-cccc-cccc-cccccccccccc'].AssignedLicenses=$null
            (Invoke-LehrerUpdate -Entry $comparison.ChangedPeople[0] -Config $script:PersonalConfig -Confirm:$false).Status | Should -Be Failed
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }
        It 'blocks a newly shared recipient, identity change or tenant switch before mutation' -ForEach @(
            @{Change='Shared'},@{Change='Identity'},@{Change='Tenant'}
        ) {
            $inventory=Get-PersonalInventory
            switch($Change) {
                Shared {$script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';RecipientTypeDetails='SharedMailbox'})}
                Identity {$script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].UserPrincipalName='different@example.invalid'}
                Tenant {Mock Get-MgContext {[pscustomobject]@{TenantId='different-tenant'}}}
            }
            (Invoke-PersonalAccountBlock -Accounts $inventory.Accounts -Config $inventory.Config -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'skips active selections during deletion and cancels on Escape' {
            $inventory=Get-PersonalInventory
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation Delete -Confirm:$false).Status | Should -Be Skipped
            Mock Read-DisabledAccountTuiKey {'Escape'}
            @(Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation Block -Confirm:$false).Count | Should -Be 0
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'keeps an active guest selected after <Key> and explains how to enable deletion' -ForEach @(@{Key='Backspace'},@{Key='Delete'}) {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].UserType='Guest'
            $inventory=Get-PersonalInventory
            Mock Get-PersonalInventory {throw 'Unexpected inventory refresh for skipped deletion'}
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar',$Key,'Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            $result=@(Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config)
            $result.Status | Should -Be Skipped
            $result.Message | Should -Match 'aktiv.*B.*sperren.*Backspace/Entf'
            Should -Invoke Show-DisabledAccountSelection -Exactly 1 -ParameterFilter {$Message -match 'aktiv.*B.*sperren.*Backspace/Entf' -and $SelectedIds.Count -eq 1}
            Should -Invoke Get-PersonalInventory -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
        }
        It 'can block and delete a guest after the rejected active deletion without remarking' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].UserType='Guest'
            $inventory=Get-PersonalInventory
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar','Backspace','B','Y','Spacebar','Backspace','Y','Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            $result=@(Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config)
            $result.Status | Should -Be @('Skipped','Disabled','Deleted')
            Should -Invoke Remove-MgUser -Exactly 1
        }
        It 'reports a partial block when revocation fails and always closes progress' {
            Mock Revoke-EntraStudentSession { $false }
            $inventory=Get-PersonalInventory
            $result=Invoke-PersonalAccountBlock -Accounts $inventory.Accounts -Config $inventory.Config -Confirm:$false
            $result.Status | Should -Be Failed
            $result.AccountDisabled | Should -BeTrue
            $result.Message | Should -Match 'bereits gesperrt'
            Should -Invoke Write-Progress -ParameterFilter {$Id -eq 2 -and $Completed} -Exactly 1
        }
        It 'accepts block, unlock and deletion only with their matching status' -ForEach @(
            @{Operation='Block';Enabled=$true;Expected='Disabled'},@{Operation='Unlock';Enabled=$false;Expected='Enabled'},@{Operation='Delete';Enabled=$false;Expected='Deleted'}
        ) {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=$Enabled
            $inventory=Get-PersonalInventory
            Mock Read-DisabledAccountTuiKey { 'Y' }
            $result=Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Confirm:$false
            $result.Status | Should -Be $Expected
        }
        It 'keeps the TUI open and refreshes the population after a confirmed action' {
            $inventory=Get-PersonalInventory
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar','B','Y','Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            $result=Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config
            $result.Status | Should -Be Disabled
            Should -Invoke Show-DisabledAccountSelection -ParameterFilter {$PersonalMode -and $Accounts[0].AccountEnabled -eq $false} -Exactly 1
        }
        It 'retains the selection after No without refreshing or mutating' {
            $inventory=Get-PersonalInventory
            Mock Get-PersonalInventory {throw 'Unexpected refresh after No'}
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar','B','N','Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            @(Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config).Count | Should -Be 0
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Get-PersonalInventory -Exactly 0
            Should -Invoke Show-DisabledAccountSelection -Exactly 1 -ParameterFilter {$Message -match 'abgebrochen' -and $SelectedIds.Count -eq 1}
        }
        It 'processes only active accounts in a mixed block selection' {
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee']=New-PersonalSyntheticUser -Id 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee' -Enabled:$false
            $inventory=Get-PersonalInventory
            Mock Read-DisabledAccountTuiKey {'J'}
            $result=@(Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation Block -Confirm:$false)
            @($result | Where-Object Status -eq Disabled).Count | Should -Be 1
            @($result | Where-Object Status -eq Skipped).Count | Should -Be 1
            Should -Invoke Update-MgUser -ParameterFilter {$UserId -eq 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'} -Exactly 0
        }
        It 'protects deletion when the fresh owned-object lookup fails' {
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=$false
            $inventory=Get-PersonalInventory
            Mock Get-MgUserOwnedObject {throw 'synthetic owned-object failure'}
            $result=Invoke-SelectedDisabledAccountDeletion -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config -Confirm:$false
            $result.Status | Should -Be Skipped
            $result.Message | Should -Match 'synthetic owned-object failure'
            Should -Invoke Remove-MgUser -Exactly 0
        }
        It 'limits teacher management to active and disabled direct <Role> members' -ForEach @(@{Role='School'},@{Role='Ganztag'}) {
            $teacherConfig=Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalMemberships[$id]=@($teacherConfig.Roles[$Role].Id)
            $script:PersonalUsers['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee']=New-PersonalSyntheticUser -Id 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee' -Enabled:$false
            $script:PersonalMemberships['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee']=@($teacherConfig.Roles[$Role].Id)
            $script:PersonalUsers['ffffffff-ffff-ffff-ffff-ffffffffffff']=New-PersonalSyntheticUser -Id 'ffffffff-ffff-ffff-ffff-ffffffffffff'
            $script:PersonalMemberships['ffffffff-ffff-ffff-ffff-ffffffffffff']=@($script:PersonalConfig.Roles.Schule_Sekretariat.Id)
            $inventory=Get-PersonalInventory -Population Teacher
            $inventory.Accounts.Count | Should -Be 2
            @($inventory.Accounts | Where-Object MutationAllowed).Count | Should -Be 2
            $inventory.Config.ManagementPopulation | Should -Be Teacher
            $inventory.Accounts.UserId | Should -Not -Contain 'ffffffff-ffff-ffff-ffff-ffffffffffff'
            (Get-PersonalInventory).Accounts.UserId | Should -Not -Contain $id
        }
        It 'excludes teacher accounts with pupil or shared-mailbox assignments and indirect-only membership' -ForEach @(
            @{Reason='Pupil'},@{Reason='Shared'},@{Reason='Indirect'}
        ) {
            $teacherConfig=Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalMemberships[$id]=@($teacherConfig.Roles.School.Id)
            switch($Reason) {
                Pupil {$script:PersonalTransitive[$id]=@($script:PersonalGroups[(Get-ConfiguredStudentRole).Id])}
                Shared {$script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails='SharedMailbox'})}
                Indirect {$script:PersonalMemberships[$id]=@();$script:PersonalTransitive[$id]=@($script:PersonalGroups[$teacherConfig.Roles.School.Id])}
            }
            (Get-PersonalInventory -Population Teacher).Accounts.Count | Should -Be 0
        }
        It 'runs teacher <Operation> with the existing confirmation and action checks' -ForEach @(
            @{Operation='Block';Enabled=$true;Expected='Disabled'},@{Operation='Unlock';Enabled=$false;Expected='Enabled'},@{Operation='Delete';Enabled=$false;Expected='Deleted'}
        ) {
            $teacherConfig=Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalMemberships[$id]=@($teacherConfig.Roles.School.Id)
            $script:PersonalUsers[$id].AccountEnabled=$Enabled
            $inventory=Get-PersonalInventory -Population Teacher
            Mock Read-DisabledAccountTuiKey {'Y'}
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Confirm:$false).Status | Should -Be $Expected
        }
        It 'blocks teacher <Operation> after <Change>' -ForEach @(
            @{Operation='Block';Change='RoleLost'},@{Operation='Unlock';Change='RoleLost'},@{Operation='Delete';Change='RoleLost'},
            @{Operation='Block';Change='Admin'},@{Operation='Unlock';Change='Admin'},@{Operation='Delete';Change='Admin'},
            @{Operation='Block';Change='Pupil'},@{Operation='Unlock';Change='Pupil'},@{Operation='Delete';Change='Pupil'},
            @{Operation='Block';Change='Shared'},@{Operation='Unlock';Change='Shared'},@{Operation='Delete';Change='Shared'}
        ) {
            $teacherConfig=Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalMemberships[$id]=@($teacherConfig.Roles.School.Id)
            $script:PersonalUsers[$id].AccountEnabled=($Operation -eq 'Block')
            $inventory=Get-PersonalInventory -Population Teacher
            switch($Change) {
                RoleLost {$script:PersonalMemberships[$id]=@($script:PersonalConfig.Roles.Schule_Sekretariat.Id)}
                Admin {$script:PersonalTransitive[$id]=@($script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'])}
                Pupil {$script:PersonalTransitive[$id]=@($script:PersonalGroups[(Get-ConfiguredStudentRole).Id])}
                Shared {$script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails='SharedMailbox'})}
            }
            Mock Read-DisabledAccountTuiKey {'Y'}
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'keeps teacher scope when refreshing the TUI after a block' {
            $teacherConfig=Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@($teacherConfig.Roles.Ganztag.Id)
            $inventory=Get-PersonalInventory -Population Teacher
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar','B','Y','Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            $result=Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config
            $result.Status | Should -Be Disabled
            Should -Invoke Show-DisabledAccountSelection -ParameterFilter {$Title -eq 'Lehrer- und Ganztagskonten' -and $Accounts.Count -eq 1 -and $Accounts[0].AccountEnabled -eq $false} -Exactly 1
        }
        It 'loads student scope with actual CLS groups, attributes and manager' {
            Set-StudentTuiFixture
            $inventory=Get-PersonalInventory -Population Student
            $inventory.Accounts.Count | Should -Be 1
            $inventory.Accounts[0].ClassGroups.DisplayName | Should -Be 'SEC-A-CLS-JK1-3g2_1'
            $inventory.Accounts[0].Teacher | Should -Be 'Nora Lehrwald <nlehrwald@example.invalid>'
            $inventory.Accounts[0].OfficeLocation | Should -Be g2
            $inventory.Accounts[0].MutationAllowed | Should -BeTrue
        }
        It 'runs student <Operation> with the shared lifecycle checks' -ForEach @(
            @{Operation='Block';Enabled=$true;Expected='Disabled'},@{Operation='Unlock';Enabled=$false;Expected='Enabled'},@{Operation='Delete';Enabled=$false;Expected='Deleted'}
        ) {
            Set-StudentTuiFixture
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=$Enabled
            $inventory=Get-PersonalInventory -Population Student
            Mock Read-DisabledAccountTuiKey {'Y'}
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Confirm:$false).Status | Should -Be $Expected
        }
        It 'blocks student <Operation> after the direct student role is lost' -ForEach @(@{Operation='Block'},@{Operation='Unlock'},@{Operation='Delete'}) {
            Set-StudentTuiFixture
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].AccountEnabled=($Operation -eq 'Block')
            $inventory=Get-PersonalInventory -Population Student
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@('eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee')
            Mock Read-DisabledAccountTuiKey {'Y'}
            (Invoke-PersonalTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke Remove-MgUser -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'changes the class for multiple selected students independently' {
            Set-StudentTuiFixture
            $second=New-PersonalSyntheticUser -Id '33333333-3333-3333-3333-333333333333'
            $second.UserPrincipalName='second-student@example.invalid'
            $second | Add-Member -NotePropertyMembers @{Department='JK1-3g2_1';OfficeLocation='G2'}
            $script:PersonalUsers[$second.Id]=$second
            $script:PersonalMemberships[$second.Id]=@((Get-ConfiguredStudentRole).Id,'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee')
            $script:StudentTuiManagers[$second.Id]='11111111-1111-1111-1111-111111111111'
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation Class | Select-Object -Last 1
            $results=@(Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation Class -Target $target -Confirm:$false)
            $results.Count | Should -Be 2
            @($results | Where-Object Status -eq Succeeded).Count | Should -Be 2
            Should -Invoke New-MgGroupMemberByRef -Exactly 2
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 2
            Should -Invoke Update-MgUser -Exactly 2
            Should -Invoke Write-Progress -Exactly 1 -ParameterFilter {$Completed -and $Activity -eq 'Klasse/Jahrgang wechseln'}
        }
        It 'changes only CLS groups and class attributes and keeps a disabled student disabled' {
            Set-StudentTuiFixture
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $script:PersonalUsers[$id].AccountEnabled=$false
            $script:PersonalMemberships[$id]+='cccccccc-cccc-cccc-cccc-cccccccccccc'
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation Class | Where-Object Id -eq 'ffffffff-ffff-ffff-ffff-ffffffffffff'
            $result=Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation Class -Target $target -Confirm:$false
            $result.Status | Should -Be Succeeded -Because $result.Message
            $script:PersonalMemberships[$id] | Should -Contain 'ffffffff-ffff-ffff-ffff-ffffffffffff'
            $script:PersonalMemberships[$id] | Should -Not -Contain 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
            $script:PersonalMemberships[$id] | Should -Contain 'cccccccc-cccc-cccc-cccc-cccccccccccc'
            $script:PersonalMemberships[$id] | Should -Contain (Get-ConfiguredStudentRole).Id
            $script:PersonalUsers[$id].Department | Should -Be JK4-6m1_1
            $script:PersonalUsers[$id].OfficeLocation | Should -Be m1
            $script:PersonalUsers[$id].AccountEnabled | Should -BeFalse
            Should -Invoke Set-MgUserManagerByRef -Exactly 0
            Should -Invoke Update-MgUser -Exactly 1 -ParameterFilter {
                $json=if ($BodyParameter -is [System.Collections.IDictionary]) {$BodyParameter | ConvertTo-Json} else {$BodyParameter.ToJsonString()}
                $submitted=$json | ConvertFrom-Json
                @($submitted.PSObject.Properties).Count -eq 2 -and $submitted.Department -eq 'JK4-6m1_1' -and $submitted.OfficeLocation -eq 'm1'
            }
        }
        It 'updates the manager using only a freshly confirmed school-team member' {
            Set-StudentTuiFixture
            $teacher=New-PersonalSyntheticUser -Id '22222222-2222-2222-2222-222222222222'
            $teacher.DisplayName='Lea Lehrwald';$teacher.UserPrincipalName='llehrwald@example.invalid'
            $script:PersonalUsers[$teacher.Id]=$teacher;$script:PersonalMemberships[$teacher.Id]=@('103a1c4c-036b-40bc-bbe8-e3f7a866cb01')
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation Teacher | Where-Object Id -eq $teacher.Id
            (Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation Teacher -Target $target -Confirm:$false).Status | Should -Be Succeeded
            $script:StudentTuiManagers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'] | Should -Be $teacher.Id
            Should -Invoke Set-MgUserManagerByRef -Exactly 1
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
        }
        It 'prevents assignment writes for <Operation> after <Change>' -ForEach @(
            @{Operation='Class';Change='RoleLost'},@{Operation='Teacher';Change='RoleLost'},
            @{Operation='Class';Change='Admin'},@{Operation='Teacher';Change='Admin'},
            @{Operation='Class';Change='Shared'},@{Operation='Teacher';Change='Shared'},
            @{Operation='Class';Change='SourceChanged'},@{Operation='Teacher';Change='SourceChanged'},
            @{Operation='Class';Change='TargetChanged'},@{Operation='Teacher';Change='TargetChanged'}
        ) {
            Set-StudentTuiFixture
            $id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation $Operation | Select-Object -Last 1
            switch($Change) {
                RoleLost {$script:PersonalMemberships[$id]=@()}
                Admin {$script:PersonalTransitive[$id]=@($script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'])}
                Shared {$script:PersonalRecipients=@([pscustomobject]@{ExternalDirectoryObjectId=$id;RecipientTypeDetails='SharedMailbox'})}
                SourceChanged {if ($Operation -eq 'Class') {$script:PersonalUsers[$id].Department='changed'} else {$script:StudentTuiManagers[$id]='33333333-3333-3333-3333-333333333333'}}
                TargetChanged {if ($Operation -eq 'Class') {$script:PersonalGroups[$target.Id].GroupTypes=@('DynamicMembership')} else {$script:PersonalMemberships[$target.Id]=@()}}
            }
            (Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Target $target -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
            Should -Invoke Set-MgUserManagerByRef -Exactly 0
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'previews <Operation> without write scopes or mutations' -ForEach @(@{Operation='Class'},@{Operation='Teacher'}) {
            Set-StudentTuiFixture
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation $Operation | Select-Object -Last 1
            (Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Target $target -WhatIf).Status | Should -Be WhatIf
            Should -Invoke Connect-DisabledAccountWriteGraph -Exactly 0
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
            Should -Invoke Set-MgUserManagerByRef -Exactly 0
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'retains old classes and reports a partial failure when <Failure> fails' -ForEach @(@{Failure='Addition'},@{Failure='Attributes'},@{Failure='Protection'}) {
            Set-StudentTuiFixture
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation Class | Select-Object -Last 1
            switch ($Failure) {
                Addition {Mock New-MgGroupMemberByRef { }}
                Attributes {Mock Update-MgUser {throw 'synthetic attribute failure'}}
                Protection {Mock New-MgGroupMemberByRef {$script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']+='ffffffff-ffff-ffff-ffff-ffffffffffff';$script:PersonalTransitive['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@($script:PersonalGroups['bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'])}}
            }
            $result=Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation Class -Target $target -Confirm:$false
            $result.Status | Should -Be Failed
            $result.Message | Should -Match 'Teiländerung'
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'] | Should -Contain 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
        }
        It 'protects licensing and dynamic classes before any mutation' -ForEach @(@{Field='AssignedLicenses';Value=@([pscustomobject]@{SkuId='synthetic-sku'})},@{Field='GroupTypes';Value=@('DynamicMembership')}) {
            Set-StudentTuiFixture
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation Class | Select-Object -Last 1
            $script:PersonalGroups['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'].$Field=$Value
            (Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation Class -Target $target -Confirm:$false).Status | Should -Be Skipped
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
            Should -Invoke Update-MgUser -Exactly 0
        }
        It 'assigns a first CLS group to a student with no current class' {
            Set-StudentTuiFixture
            $script:PersonalMemberships['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa']=@((Get-ConfiguredStudentRole).Id)
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].Department=''
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].OfficeLocation=''
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation Class | Select-Object -Last 1
            $result=Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation Class -Target $target -Confirm:$false
            $result.Status | Should -Be Succeeded -Because $result.Message
            Should -Invoke New-MgGroupMemberByRef -Exactly 1
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
        }
        It 'does not write an already matching <Operation> assignment' -ForEach @(@{Operation='Class'},@{Operation='Teacher'}) {
            Set-StudentTuiFixture
            $script:PersonalUsers['aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'].OfficeLocation='G2'
            $inventory=Get-PersonalInventory -Population Student
            $target=Get-StudentTuiOption -Config $inventory.Config -Operation $Operation | Select-Object -First 1
            (Invoke-StudentAssignment -Accounts $inventory.Accounts -Config $inventory.Config -Operation $Operation -Target $target -Confirm:$false).Status | Should -Be Compliant
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
            Should -Invoke Set-MgUserManagerByRef -Exactly 0
        }
        It 'confirms <Operation> through target selection and refreshes student scope' -ForEach @(@{Operation='Class';Key='K'},@{Operation='Teacher';Key='L'}) {
            Set-StudentTuiFixture
            $inventory=Get-PersonalInventory -Population Student
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in @('Spacebar',$Key,'DownArrow','Enter','Y','Escape')) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            Mock Show-DisabledAccountSelection { }
            $result=Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config
            $result.Status | Should -Be Succeeded
            Should -Invoke Show-DisabledAccountSelection -Exactly 3 -ParameterFilter {$StudentMode -and $Title -eq 'Schülerkonten' -and $Accounts.Count -eq 1}
        }
        It 'cancels student target selection and confirmation without writes' -ForEach @(@{Keys=@('Escape')},@{Keys=@('Enter','N')}) {
            Set-StudentTuiFixture
            $inventory=Get-PersonalInventory -Population Student
            $script:PersonalKeys=[Collections.Generic.Queue[string]]::new()
            foreach($key in $Keys) {$script:PersonalKeys.Enqueue($key)}
            Mock Read-DisabledAccountTuiKey {$script:PersonalKeys.Dequeue()}
            @(Invoke-StudentTuiAction -Accounts $inventory.Accounts -Config $inventory.Config -Operation Class).Count | Should -Be 0
            Should -Invoke Update-MgUser -Exactly 0
            Should -Invoke New-MgGroupMemberByRef -Exactly 0
        }
    }
}

Describe 'Personal 0.3.0: actual isolated CLI processes' {
    BeforeAll {
        $script:PersonalCliHarness=Join-Path $PSScriptRoot 'Invoke-SyntheticPersonalCli.ps1'
        $script:PersonalCliWorkbook=Join-Path $TestDrive 'Personal.xlsx'
    }
    It 'executes List and PassThru through the real script' {
        foreach($mode in @('List','Object')) {
            $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode $mode 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match 'ltestwald@example.invalid'
            $output -join "`n" | Should -Not -Match 'Remote operation forbidden'
        }
    }
    It 'exports a new workbook and actually compares and previews its Update' {
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Export -File $script:PersonalCliWorkbook 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        Test-Path $script:PersonalCliWorkbook | Should -BeTrue
        $exportCells = & (Get-Module SchuelerSync) {
            param($WorkbookPath)
            $package = Open-StudentWorkbookPackage $WorkbookPath
            try { @{Timestamp=$package.Workbook.Worksheets['Tabelle1'].Cells[2,9].Value; Job=$package.Workbook.Worksheets['Tabelle1'].Cells[2,4].Value} } finally { $package.Dispose() }
        } $script:PersonalCliWorkbook
        $exportCells.Timestamp | Should -BeExactly '2026-01-02T10:30:00Z'
        $exportCells.Job | Should -BeExactly 'Schule Sekretariat'
        $hash=(Get-FileHash $script:PersonalCliWorkbook).Hash
        foreach($mode in @('Compare','Update')) {
            $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode $mode -File $script:PersonalCliWorkbook 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match '"HasErrors": false'
        }
        (Get-FileHash $script:PersonalCliWorkbook).Hash | Should -Be $hash
    }
    It 'ignores incomplete guests through the actual <Mode> CLI without changing the workbook' -ForEach @(
        @{Mode='Compare'},@{Mode='Update'}
    ) {
        $path=Join-Path $TestDrive "Guest-$Mode.xlsx"
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Export -File $path -GuestFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $exportOutput=$output -join "`n"
        $hash=(Get-FileHash $path).Hash
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode $Mode -File $path -GuestFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $output -join "`n" | Should -Match '"HasErrors": false'
        $output -join "`n" | Should -Not -Match 'Unvollständige Personaldaten|Workbook.Review|Remote operation forbidden'
        $output -join "`n" | Should -Match '"Warnings": \[\s*\]'
        $output -join "`n" | Should -Match '"Errors": \[\s*\]'
        $output -join "`n" | Should -Match '"Actions": \[\s*\]'
        $exportOutput | Should -Not -Match 'Exportzeile benötigt'
        (Get-FileHash $path).Hash | Should -Be $hash
    }
    It 'previews real personnel updates despite a protected admin in the actual exported workbook' {
        $path=Join-Path $TestDrive 'ProtectedAdmin.xlsx'
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Export -File $path -ProtectedFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $hash=(Get-FileHash $path).Hash
        foreach ($mode in @('Compare','Update')) {
            $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode $mode -File $path -ProtectedFixture 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match '"HasErrors": false'
            $output -join "`n" | Should -Match 'Geschütztes Administratorkonto, nicht abgeglichen'
            $output -join "`n" | Should -Match 'JobTitle'
            $output -join "`n" | Should -Not -Match 'Ungültiger Job|Remote operation forbidden'
            if ($mode -eq 'Update') {
                $output -join "`n" | Should -Match '"Phase": "Update"'
                $output -join "`n" | Should -Match '"Status": "WhatIf"'
                [regex]::Matches(($output -join "`n"),'"Phase": "Update"').Count | Should -Be 1
            }
        }
        (Get-FileHash $path).Hash | Should -Be $hash
    }
    It 'updates a complete personnel export containing protected, roleless and guest rows without aborting' {
        $path=Join-Path $TestDrive 'ReviewExport.xlsx'
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Export -File $path -ReviewFixture -ProtectedFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $hash=(Get-FileHash $path).Hash
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Compare -File $path -ReviewFixture -ProtectedFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $output -join "`n" | Should -Match '"HasErrors": false'
        $output -join "`n" | Should -Match 'Workbook.Review|Workbook.*Review'
        $output -join "`n" | Should -Match 'Änderungen \(1\)'
        [regex]::Matches(($output -join "`n"),'"Code": "Workbook.Review"').Count | Should -Be 2
        $output -join "`n" | Should -Match 'Geschütztes Administratorkonto, nicht abgeglichen'
        $output -join "`n" | Should -Not -Match 'Remote operation forbidden'
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Update -File $path -ReviewFixture -ProtectedFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $output -join "`n" | Should -Match '"HasErrors": false'
        $output -join "`n" | Should -Match 'Keine Personalrolle und kein Job, nicht abgeglichen'
        $output -join "`n" | Should -Not -Match 'Unvollständige Personaldaten'
        [regex]::Matches(($output -join "`n"),'"Phase": "Update"').Count | Should -Be 1
        $output -join "`n" | Should -Not -Match 'Remote operation forbidden'
        (Get-FileHash $path).Hash | Should -Be $hash
    }
    It 'updates a full export despite six known personnel accounts with missing names through the actual CLI' {
        $path=Join-Path $TestDrive 'IncompleteNames.xlsx'
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Export -File $path -ReviewFixture -ProtectedFixture -IncompleteNameFixture 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $hash=(Get-FileHash $path).Hash
        foreach ($mode in @('Compare','Update')) {
            $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode $mode -File $path -ReviewFixture -ProtectedFixture -IncompleteNameFixture 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match '"HasErrors": false'
            [regex]::Matches(($output -join "`n"),'"Code": "Workbook.Review"').Count | Should -Be 8
            $output -join "`n" | Should -Match 'Änderungen \(1\)'
            $output -join "`n" | Should -Not -Match 'Preflight.Inventory|Remote operation forbidden'
            if ($mode -eq 'Update') {
                [regex]::Matches(($output -join "`n"),'"Phase": "Update"').Count | Should -Be 1
                $output -join "`n" | Should -Match 'Unvollständige Namensdaten, nicht abgeglichen'
            }
        }
        (Get-FileHash $path).Hash | Should -Be $hash
    }
    It 'executes Add, both UPN aliases and Exchange without remote mutations' {
        foreach($mode in @('Add','Remove','Exchange')) {
            $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode $mode 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match '"HasErrors": false'
            $output -join "`n" | Should -Not -Match 'Remote operation forbidden'
        }
    }
    It 'previews recovery of an existing teacher through the actual ExchangeOnly CLI' {
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Population Teacher -Mode Exchange 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
        $output -join "`n" | Should -Match '"Mode": "ExchangeOnly"'
        $output -join "`n" | Should -Match '"Phase": "Exchange"'
        $output -join "`n" | Should -Match '"Status": "WhatIf"'
        $output -join "`n" | Should -Match '"HasErrors": false'
        $output -join "`n" | Should -Not -Match 'Remote operation forbidden'
    }
    It 'returns exit code 1 for a protected profile and rejects a redirected TUI' {
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Invalid 2>&1
        $LASTEXITCODE | Should -Be 1
        $output -join "`n" | Should -Match 'geschützter Personaljob'
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Mode Tui 2>&1
        $LASTEXITCODE | Should -Be 1
        $output -join "`n" | Should -Match 'interaktives Terminal'
    }
    It 'binds teacher ManageAccounts through the actual entry and rejects a redirected terminal' {
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Population Teacher -Mode Tui 2>&1
        $LASTEXITCODE | Should -Be 1
        $output -join "`n" | Should -Match 'interaktives Terminal'
        $output -join "`n" | Should -Not -Match 'Remote operation forbidden|parameter.*not.*found|Parametersatz'
    }
    It 'binds student ManageAccounts and rejects redirected input' {
        $output=& pwsh -NoLogo -NoProfile -NonInteractive -File $script:PersonalCliHarness -Population Student -Mode Tui 2>&1
        $LASTEXITCODE | Should -Be 1
        $output -join "`n" | Should -Match 'interaktives Terminal'
    }
    It 'rejects <Population> management combined with <Argument> before executing any action' -ForEach @(
        @{Population='Lehrer';Argument='-Update'},@{Population='Lehrer';Argument='-ListToExcel'},@{Population='Lehrer';Argument='-File'},
        @{Population='Schueler';Argument='-Update'},@{Population='Schueler';Argument='-ListToExcel'},@{Population='Schueler';Argument='-File'}
    ) {
        $entry=Join-Path (Split-Path $PSScriptRoot -Parent) "Sync-$($Population)Entra.ps1"
        $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-File',$entry,'-ManageAccounts',$Argument)
        if ($Argument -eq '-File') {$arguments+='synthetic.xlsx'}
        $output=& pwsh @arguments 2>&1
        $LASTEXITCODE | Should -Be 1
        $output -join "`n" | Should -Match 'Parametersatz|parameter set'
    }
}

Describe 'Personal 0.3.0: real Excel export and ABP omission' {
    InModuleScope SchuelerSync {
        It 'exports the precise teacher structure as safe text including roleless accounts' {
            $path=Join-Path $TestDrive 'Personal.xlsx'
            $inventory=[pscustomobject]@{Accounts=@(
                [pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';Job='Schule_Sekretariat';UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail='ltestwald@example.invalid'},
                [pscustomobject]@{DisplayName='=SYNTHETIC()';GivenName='';Surname='';Job='';UserId='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';UPN='roleless@example.invalid';Mail=''}
            )}
            $result=Export-PersonalWorkbook -Inventory $inventory -Path $path -Confirm:$false
            $result.Count | Should -Be 2
            $package=Open-StudentWorkbookPackage -Path $path
            try {
                $sheet=$package.Workbook.Worksheets['Tabelle1']
                @((1..9) | ForEach-Object {$sheet.Cells[1,$_].Value}) | Should -Be @('Name mit Rufname','Vorname','Nachname','Job','Passwort','EntraObjectId','UPN','Mail','Last Login Timestamp')
                $sheet.Tables[0].Name | Should -BeExactly 'Tabelle1'
                $sheet.Tables[0].StyleName | Should -BeExactly 'TableStyleMedium14'
                $sheet.Tables[0].ShowRowStripes | Should -BeTrue
                $sheet.Column(4).Width | Should -Be 32
                $sheet.Cells[2,5].Value | Should -BeNullOrEmpty
                $sheet.Cells[3,1].Formula | Should -BeNullOrEmpty
                $sheet.Cells[3,1].Value | Should -BeExactly '=SYNTHETIC()'
                $sheet.Cells[3,4].Value | Should -BeNullOrEmpty
                $sheet.Cells[3,6].Value | Should -BeExactly $inventory.Accounts[1].UserId
            } finally { $package.Dispose() }
            {Export-PersonalWorkbook -Inventory $inventory -Path $path -Confirm:$false} | Should -Throw '*existiert bereits*'
        }
        It 'exports readable <Label> and imports the canonical personnel job' -ForEach @(
            @{Job='Schule_Sekretariat';Label='Schule Sekretariat';Protected=$false},
            @{Job='Schule_Vertretungskräfte';Label='Schule Vertretungskräfte';Protected=$false},
            @{Job='Kinderhaus_PädagogischesTeam';Label='Kinderhaus Pädagogisches Team';Protected=$false},
            @{Job='Kinderhaus_Sekretariat';Label='Kinderhaus Sekretariat';Protected=$false},
            @{Job='Kinderhaus_Vertretungskräfte';Label='Kinderhaus Vertretungskräfte';Protected=$false},
            @{Job='Unterstützung';Label='Unterstützung';Protected=$false},
            @{Job='ExterneBenutzer';Label='Externe Benutzer';Protected=$false},
            @{Job='ExterneAdmins';Label='Externe Admins';Protected=$true}
        ) {
            $path=Join-Path $TestDrive "Readable-$Job.xlsx"
            $account=[pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';Job=$Job;UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail='ltestwald@example.invalid'}
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Confirm:$false
            $package=Open-StudentWorkbookPackage -Path $path
            try {
                $cell=$package.Workbook.Worksheets['Tabelle1'].Cells[2,4]
                $cell.Value | Should -BeExactly $Label
                $cell.Formula | Should -BeNullOrEmpty
            } finally {$package.Dispose()}
            $account.Job | Should -BeExactly $Job
            $inputJobs=@($Label, ('  ' + $Label.ToLowerInvariant().Replace(' ', '   ') + '  '), $Job.ToLowerInvariant())
            if (-not $Protected) { $inputJobs += "$Job, $Label" }
            foreach ($inputJob in $inputJobs) {
                $package=Open-StudentWorkbookPackage -Path $path
                try { $package.Workbook.Worksheets['Tabelle1'].Cells[2,4].Value=$inputJob; $package.Save() } finally {$package.Dispose()}
                $data=Read-StudentWorkbook -Path $path -Kind Personnel -AllowPersonnelReviewRows
                $data.Students[0].Job | Should -BeExactly $Job
                if ($Protected) {
                    $data.Students[0].ReviewReason | Should -Match 'geschützter Personaljob'
                    {Read-StudentWorkbook -Path $path -Kind Personnel} | Should -Throw '*geschützter Personaljob*'
                } else {
                    $data.Students[0].ReviewReason | Should -BeNullOrEmpty
                    (Read-StudentWorkbook -Path $path -Kind Personnel).Students[0].Job | Should -BeExactly $Job
                }
            }
        }
        It 'supports a header-only export of an empty inventory' {
            $path=Join-Path $TestDrive 'EmptyPersonal.xlsx'
            $result=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@()}) -Path $path -Confirm:$false
            $result.Count | Should -Be 0
            $package=Open-StudentWorkbookPackage $path
            try { $package.Workbook.Worksheets['Tabelle1'].Dimension.End.Row | Should -Be 1 } finally {$package.Dispose()}
        }
        It 'does not create an export directory or workbook with WhatIf' {
            $path=Join-Path $TestDrive 'NewDir/WhatIf.xlsx'
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@()}) -Path $path -WhatIf
            Test-Path (Split-Path $path -Parent) | Should -BeFalse
        }
        It 'retains review rows only in explicitly enabled read-only personnel imports' {
            $path=Join-Path $TestDrive 'ReadOnlyReview.xlsx'
            $accounts=@(
                [pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';Job='Schule_Sekretariat';UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail=''},
                [pscustomobject]@{DisplayName='Eva Adminwald';GivenName='Eva';Surname='Adminwald';Job='ExterneAdmins';UserId='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';UPN='admin@example.invalid';Mail=''},
                [pscustomobject]@{DisplayName='Synthetic device';GivenName='';Surname='';Job='';UserId='cccccccc-cccc-cccc-cccc-cccccccccccc';UPN='device@example.invalid';Mail=''}
            )
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=$accounts}) -Path $path -Confirm:$false
            {Read-StudentWorkbook -Path $path -Kind Personnel} | Should -Throw '*geschützter Personaljob*'
            $data=Read-StudentWorkbook -Path $path -Kind Personnel -AllowPersonnelReviewRows
            $data.Students.Count | Should -Be 3
            $data.Students[1].ReviewReason | Should -Match 'geschützter Personaljob'
            $data.Students[2].ReviewReason | Should -Match 'Unvollständige'
            $data.Students[1].Job | Should -Be ExterneAdmins
            $data.Students[2].StoredUpn | Should -Be device@example.invalid
            {Read-StudentWorkbook -Path $path -Kind Personnel} | Should -Throw '*geschützter Personaljob*'
            {Read-StudentWorkbook -Path $path -Kind Teacher -AllowPersonnelReviewRows} | Should -Throw '*lesenden Personalvergleich*'
        }
        It 'exempts only confirmed guests from workbook data validation for <Identity>' -ForEach @(
            @{Identity='Id';RecognizedGuest=$true},@{Identity='Upn';RecognizedGuest=$true},
            @{Identity='Both';RecognizedGuest=$true},@{Identity='AllBlank';RecognizedGuest=$true},
            @{Identity='Member';RecognizedGuest=$false},@{Identity='UnknownId';RecognizedGuest=$false},
            @{Identity='ConflictingUpn';RecognizedGuest=$false},@{Identity='UnknownUpn';RecognizedGuest=$false},
            @{Identity='Unidentified';RecognizedGuest=$false}
        ) {
            $guest=[pscustomobject]@{Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UserPrincipalName='guest@example.invalid';UserType='Guest'}
            $member=[pscustomobject]@{Id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';UserPrincipalName='member@example.invalid';UserType='Member'}
            $snapshot=[pscustomobject]@{UsersById=@{$guest.Id=$guest;$member.Id=$member};UsersByUpn=@{$guest.UserPrincipalName=$guest;$member.UserPrincipalName=$member}}
            $account=[pscustomobject]@{DisplayName='Synthetic guest';GivenName='';Surname='';Job='';UserId=$guest.Id;UPN=$guest.UserPrincipalName.ToUpperInvariant();Mail=''}
            switch ($Identity) {
                Id {$account.UPN=''}
                Upn {$account.UserId=''}
                AllBlank {$account.DisplayName=''}
                Member {$account.UserId=$member.Id;$account.UPN=$member.UserPrincipalName}
                UnknownId {$account.UserId='cccccccc-cccc-cccc-cccc-cccccccccccc'}
                ConflictingUpn {$account.UPN=$member.UserPrincipalName}
                UnknownUpn {$account.UPN='missing@example.invalid'}
                Unidentified {$account.UserId='';$account.UPN=''}
            }
            $path=Join-Path $TestDrive "GuestValidation-$Identity.xlsx"
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Confirm:$false
            $hash=(Get-FileHash $path).Hash
            foreach ($review in @($false,$true)) {
                if ($RecognizedGuest -or $review) {
                    $data=Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot -AllowPersonnelReviewRows:$review
                    $data.Students.Count | Should -Be 1
                    if ($review) {
                        if ($RecognizedGuest) { $data.Students[0].ReviewReason | Should -BeNullOrEmpty }
                        else { $data.Students[0].ReviewReason | Should -Match 'Unvollständige Personaldaten' }
                    }
                } else {
                    {Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot} | Should -Throw '*Unvollständige Personaldaten*'
                }
            }
            {Read-StudentWorkbook -Path $path -Kind Teacher -PersonnelSnapshot $snapshot} | Should -Throw '*Personal*'
            (Get-FileHash $path).Hash | Should -Be $hash
        }
        It 'skips only confirmed protected accounts in workbook validation for <Case>' -ForEach @(
            @{Case='ExterneRole';Protected=$true},@{Case='GlobalAdmin';Protected=$true},@{Case='DirectoryRole';Protected=$true},
            @{Case='Unprotected';Protected=$false},@{Case='UnknownId';Protected=$false},@{Case='Conflict';Protected=$false}
        ) {
            $admin=[pscustomobject]@{Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UserPrincipalName='admin@example.invalid';UserType='Member'}
            $member=[pscustomobject]@{Id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';UserPrincipalName='member@example.invalid';UserType='Member'}
            $config=Get-PersonalConfig
            $groupId=if ($Case -eq 'ExterneRole') { $config.Roles.ExterneAdmins.Id } else { 'cccccccc-cccc-cccc-cccc-cccccccccccc' }
            $context=[pscustomobject]@{Complete=$true;ProtectedIds=@($config.Roles.ExterneAdmins.Id,'cccccccc-cccc-cccc-cccc-cccccccccccc')}
            $snapshot=[pscustomobject]@{UsersById=@{$admin.Id=$admin;$member.Id=$member};UsersByUpn=@{$admin.UserPrincipalName=$admin;$member.UserPrincipalName=$member};PersonalContext=$context}
            $script:WorkbookAdminMembership=if ($Case -eq 'Unprotected') { @() } else {
                @([pscustomobject]@{Id=$groupId;OdataType=$(if ($Case -eq 'DirectoryRole') { '#microsoft.graph.directoryRole' } else { '#microsoft.graph.group' })})
            }
            Mock Get-MgUserTransitiveMemberOf { $script:WorkbookAdminMembership }
            $account=[pscustomobject]@{DisplayName='Synthetic admin';GivenName='Eva';Surname='Adminwald';Job='ExterneAdmins';UserId=$admin.Id;UPN=$admin.UserPrincipalName;Mail=''}
            if ($Case -eq 'UnknownId') { $account.UserId='dddddddd-dddd-dddd-dddd-dddddddddddd' }
            if ($Case -eq 'Conflict') { $account.UPN=$member.UserPrincipalName }
            $path=Join-Path $TestDrive "Protected-$Case.xlsx"
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Confirm:$false
            $hash=(Get-FileHash $path).Hash
            if ($Protected) {
                $data=Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot
                $data.Students.Count | Should -Be 1
                $data.Students[0].Job | Should -BeExactly ExterneAdmins
            } else {
                {Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot} | Should -Throw '*geschützter Personaljob*'
            }
            (Get-FileHash $path).Hash | Should -Be $hash
        }
        It 'skips only confirmed roleless blank-job workbook rows for <Case>' -ForEach @(
            @{Case='Id';Skip=$true},@{Case='Upn';Skip=$true},@{Case='Both';Skip=$true},@{Case='AllBlank';Skip=$true},
            @{Case='Managed';Skip=$false},@{Case='IncompleteContext';Skip=$false},@{Case='InvalidMembership';Skip=$false},
            @{Case='UnknownId';Skip=$false},@{Case='Conflict';Skip=$false},@{Case='New';Skip=$false},
            @{Case='AssignedJob';Skip=$false},@{Case='UnknownType';Skip=$false}
        ) {
            $user=[pscustomobject]@{Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UserPrincipalName='device@example.invalid';UserType='Member'}
            $other=[pscustomobject]@{Id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';UserPrincipalName='other@example.invalid';UserType='Member'}
            $config=Get-PersonalConfig
            $context=[pscustomobject]@{Complete=$true;ManagedRoleIds=@($config.Roles.Schule_Sekretariat.Id);ProtectedIds=@();ProtectedNames=@();ExcludedIds=@()}
            $snapshot=[pscustomobject]@{UsersById=@{$user.Id=$user;$other.Id=$other};UsersByUpn=@{$user.UserPrincipalName=$user;$other.UserPrincipalName=$other};PersonalContext=$context}
            $account=[pscustomobject]@{DisplayName='Synthetic device';GivenName='';Surname='';Job='';UserId=$user.Id;UPN=$user.UserPrincipalName;Mail=''}
            $script:WorkbookRolelessMembership=@()
            switch ($Case) {
                Id {$account.UPN=''}
                Upn {$account.UserId=''}
                AllBlank {$account.DisplayName=''}
                Managed {$script:WorkbookRolelessMembership=@([pscustomobject]@{Id=$config.Roles.Schule_Sekretariat.Id;OdataType='#microsoft.graph.group';DisplayName='Schule_Sekretariat'})}
                IncompleteContext {$context.Complete=$false}
                InvalidMembership {$script:WorkbookRolelessMembership=@([pscustomobject]@{Id='cccccccc-cccc-cccc-cccc-cccccccccccc'})}
                UnknownId {$account.UserId='dddddddd-dddd-dddd-dddd-dddddddddddd'}
                Conflict {$account.UPN=$other.UserPrincipalName}
                New {$account.UserId='';$account.UPN=''}
                AssignedJob {$account.Job='Schule Sekretariat'}
                UnknownType {$user.UserType=''}
            }
            Mock Get-MgUserTransitiveMemberOf { $script:WorkbookRolelessMembership }
            $path=Join-Path $TestDrive "Roleless-$Case.xlsx"
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Confirm:$false
            $hash=(Get-FileHash $path).Hash
            if ($Skip) {
                $data=Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot
                $data.Students.Count | Should -Be 1
                $data.Students[0].Job | Should -BeNullOrEmpty
            } else {
                {Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot} | Should -Throw '*Unvollständige Personaldaten*'
            }
            (Get-FileHash $path).Hash | Should -Be $hash
        }
        It 'skips missing-name workbook rows only for confirmed personnel with a valid job for <Case>' -ForEach @(
            @{Case='Id';Skip=$true},@{Case='Upn';Skip=$true},@{Case='Both';Skip=$true},
            @{Case='MissingGiven';Skip=$true},@{Case='MissingSurname';Skip=$true},@{Case='MissingDisplay';Skip=$true},
            @{Case='UnknownId';Skip=$false},@{Case='Conflict';Skip=$false},@{Case='Unidentified';Skip=$false},
            @{Case='UnknownJob';Skip=$false},@{Case='BlankJob';Skip=$false},@{Case='IncompleteContext';Skip=$false},
            @{Case='InvalidMembership';Skip=$false},@{Case='Roleless';Skip=$false},@{Case='MultipleRoles';Skip=$false},
            @{Case='UnknownType';Skip=$false}
        ) {
            $user=[pscustomobject]@{Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UserPrincipalName='team@example.invalid';UserType='Member'}
            $other=[pscustomobject]@{Id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';UserPrincipalName='other@example.invalid';UserType='Member'}
            $config=Get-PersonalConfig
            $roleId=$config.Roles.Kinderhaus_PädagogischesTeam.Id
            $context=[pscustomobject]@{Complete=$true;ManagedRoleIds=@($roleId,$config.Roles.Schule_Sekretariat.Id);ProtectedIds=@();ProtectedNames=@();ExcludedIds=@()}
            $snapshot=[pscustomobject]@{UsersById=@{$user.Id=$user;$other.Id=$other};UsersByUpn=@{$user.UserPrincipalName=$user;$other.UserPrincipalName=$other};PersonalContext=$context}
            $account=[pscustomobject]@{DisplayName='Synthetic team';GivenName='';Surname='';Job='Kinderhaus Pädagogisches Team';UserId=$user.Id;UPN=$user.UserPrincipalName;Mail=''}
            $script:WorkbookIncompleteMembership=@([pscustomobject]@{Id=$roleId;OdataType='#microsoft.graph.group';DisplayName='Kinderhaus_PädagogischesTeam'})
            switch ($Case) {
                Id {$account.UPN=''}
                Upn {$account.UserId=''}
                MissingGiven {$account.Surname='Team'}
                MissingSurname {$account.GivenName='Synthetic'}
                MissingDisplay {$account.GivenName='Synthetic';$account.Surname='Team';$account.DisplayName=''}
                UnknownId {$account.UserId='cccccccc-cccc-cccc-cccc-cccccccccccc'}
                Conflict {$account.UPN=$other.UserPrincipalName}
                Unidentified {$account.UserId='';$account.UPN=''}
                UnknownJob {$account.Job='Unknown job'}
                BlankJob {$account.Job=''}
                IncompleteContext {$context.Complete=$false}
                InvalidMembership {$script:WorkbookIncompleteMembership=@([pscustomobject]@{Id=$roleId})}
                Roleless {$script:WorkbookIncompleteMembership=@()}
                MultipleRoles {$script:WorkbookIncompleteMembership+=@([pscustomobject]@{Id=$config.Roles.Schule_Sekretariat.Id;OdataType='#microsoft.graph.group';DisplayName='Schule_Sekretariat'})}
                UnknownType {$user.UserType=''}
            }
            Mock Get-MgUserTransitiveMemberOf { $script:WorkbookIncompleteMembership }
            $path=Join-Path $TestDrive "IncompleteNames-$Case.xlsx"
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Confirm:$false
            $hash=(Get-FileHash $path).Hash
            if ($Skip) {
                $data=Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot
                $data.Students.Count | Should -Be 1
                $data.Students[0].Job | Should -BeExactly 'Kinderhaus_PädagogischesTeam'
            } else {
                {Read-StudentWorkbook -Path $path -Kind Personnel -PersonnelSnapshot $snapshot} | Should -Throw '*Unvollständige Personaldaten*'
            }
            (Get-FileHash $path).Hash | Should -Be $hash
        }
        It 'reads canonical personnel jobs with the shared adapter' {
            $path=Join-Path $TestDrive 'ImportPersonal.xlsx'
            $account=[pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';Job='Schule_Sekretariat';UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail='ltestwald@example.invalid'}
            $null=Export-PersonalWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Confirm:$false
            $data=Read-StudentWorkbook -Path $path -Kind Personnel
            $data.Students[0].Job | Should -BeExactly 'Schule_Sekretariat'
            $data.Students[0].EntraObjectId | Should -BeExactly $account.UserId
            $data.Students[0].Password | Should -BeNullOrEmpty
        }
        It 'never passes an unconfigured ABP to Set-Mailbox even if supplied as a difference' {
            $config=@{AddressBookPolicy=$null;CustomAttribute1='synthetic';AuditLogAgeLimitDays=365;RetainDeletedItemsForDays=30;RoleAssignmentPolicy='roles';SharingPolicy='sharing';RetentionPolicy='retention'}
            $parameters=Get-ExchangeMailboxWriteParameter -Mailbox ([pscustomobject]@{PersistedCapabilities=@()}) -Config $config -Differences @([pscustomobject]@{Field='AddressBookPolicy';Action='Set';Desired=$null})
            $parameters.ContainsKey('AddressBookPolicy') | Should -BeFalse
        }
        It 'omits a blank ABP from the real mailbox comparison' {
            $config=@{AddressBookPolicy=$null;CustomAttribute1='synthetic';AuditLogAgeLimitDays=365;RetainDeletedItemsForDays=30
                RoleAssignmentPolicy='roles';SharingPolicy='sharing';RetentionPolicy='retention';OwaMailboxPolicy='owa'}
            $mailbox=[pscustomobject]@{AddressBookPolicy='existing-policy';CustomAttribute1='synthetic';AuditEnabled=$true
                AuditLogAgeLimit=[timespan]::FromDays(365);RetainDeletedItemsFor=[timespan]::FromDays(30)
                RoleAssignmentPolicy='roles';SharingPolicy='sharing';RetentionPolicy='retention';PersistedCapabilities=@()
                AuditDelegate=@();AuditOwner=@();AuditAdmin=@()}
            $required=Get-RequiredExchangeAuditAction -Mailbox $mailbox
            foreach($field in @('AuditDelegate','AuditOwner','AuditAdmin')) {$mailbox.$field=$required.$field}
            $cas=[pscustomobject]@{ActiveSyncEnabled=$false;ImapEnabled=$false;MAPIEnabled=$true;OWAEnabled=$true
                OWAforDevicesEnabled=$false;OwaMailboxPolicy='owa';PopEnabled=$false;SmtpClientAuthenticationDisabled=$true}
            @(Compare-StudentMailboxState -Mailbox $mailbox -CasMailbox $cas -Config $config).Count | Should -Be 0
        }
    }
}
