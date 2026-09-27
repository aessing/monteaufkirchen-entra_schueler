BeforeDiscovery {
    Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Audit regressions: destructive input and tenant boundaries' {
    InModuleScope SchuelerSync {
        It 'does not classify an empty student import as departures' {
            $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/SchuelerSync.psd1')
            $ids = [Collections.Generic.HashSet[string]]::new(); [void]$ids.Add('student')
            $snapshot = [pscustomobject]@{ UsersById = @{ student = [pscustomobject]@{ Id = 'student'; AccountEnabled = $true } }
                StudentRoleMemberIds = $ids; ReservedAddresses = @(); AddressOwners = @{} }
            $result = Compare-StudentDirectory -Students @() -Snapshot $snapshot -Config $config
            $result.Departures.Count | Should -Be 0
            $result.Errors.Count | Should -BeGreaterThan 0
        }

        It 'rejects a workbook as report destination without changing its bytes' {
            $path = Join-Path $TestDrive 'source.xlsx'
            Copy-Item (Join-Path $script:SchuelerSyncRepositoryRoot 'tests/fixtures/Lehrer-Testdaten.xlsx') $path
            $hash = (Get-FileHash $path).Hash
            $empty = [pscustomobject]@{NewPeople=@();Departures=@();ChangedPeople=@();ExistingPeople=@();Warnings=@();Errors=@()}
            { Write-LehrerReport -Comparison $empty -Actions @() -HeaderLines @() -OutputFile $path } | Should -Throw
            (Get-FileHash $path).Hash | Should -Be $hash
        }

        It 'never overwrites an existing text report' {
            $path = Join-Path $TestDrive 'existing.txt'
            Set-Content $path 'original'
            $empty = [pscustomobject]@{NewPeople=@();Departures=@();ChangedPeople=@();ExistingPeople=@();Warnings=@();Errors=@()}
            { Write-LehrerReport -Comparison $empty -Actions @() -HeaderLines @() -OutputFile $path } | Should -Throw
            (Get-Content $path) | Should -Be 'original'
        }
    }
}

Describe 'Audit regressions: personnel role protection' {
    InModuleScope SchuelerSync {
        It 'does not remove a role that acquired a license after the comparison' {
            $role = [pscustomobject]@{Id='target';DisplayName='SEC-A-ROL-Target';AssignedLicenses=@();GroupTypes=@();MembershipRule=$null}
            $old = [pscustomobject]@{Id='old';DisplayName='SEC-A-ROL-Old';AssignedLicenses=@([pscustomobject]@{SkuId='license'});GroupTypes=@();MembershipRule=$null}
            Mock Get-FreshEntraUserDirectGroup { @($role,$old) }
            Mock Get-MgGroup { if ($GroupId -eq 'old') { $old } else { $role } }
            Mock Remove-MgGroupMemberDirectoryObjectByRef { }
            $snapshot = [pscustomobject]@{GroupsById=@{old=[pscustomobject]@{AssignedLicenses=@()}}}
            { Sync-LehrerGroup -UserId person -Desired ([pscustomobject]@{Role=$role}) -Confirm:$false } | Should -Throw '*Lizenz*'
            Should -Invoke Remove-MgGroupMemberDirectoryObjectByRef -Exactly 0
        }

        It 'rejects a personnel identity whose current membership and profile no longer match' {
            $members = [Collections.Generic.HashSet[string]]::new(); [void]$members.Add('person')
            $snapshot = [pscustomobject]@{PersonnelMemberIds=$members;StudentMemberIds=[Collections.Generic.HashSet[string]]::new()}
            $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            Mock Get-MgUser { [pscustomobject]@{Id='person';GivenName='Synthetic';Surname='Person';CompanyName='Other';EmployeeType='Other';UserType='Member';AccountEnabled=$true} }
            Mock Get-FreshEntraUserDirectGroup { @() }
            { Assert-LehrerRecoveryIdentity -UserId person -Person ([pscustomobject]@{GivenName='Synthetic';Surname='Person';Job='L'}) -Config $config } | Should -Throw
        }
    }
}

Describe 'Audit regressions: shared safety boundaries' {
    InModuleScope SchuelerSync {
        It 'rejects an unknown, foreign or ambiguous Exchange tenant' -ForEach @(
            @{ Tenants = @('other') }, @{ Tenants = @('') }, @{ Tenants = @('expected','expected') }
        ) {
            $connections = @($Tenants | ForEach-Object { [pscustomobject]@{Name='ExchangeOnline';State='Connected';TenantID=$_} })
            { Assert-ExchangeTenant -TenantId expected -Connections $connections } | Should -Throw
        }
        It 'accepts a single Exchange session belonging to the Graph tenant' {
            $connection = [pscustomobject]@{Name='ExchangeOnline';State='Connected';TenantID='expected'}
            (Assert-ExchangeTenant -TenantId expected -Connections @($connection)).TenantID | Should -Be expected
        }
        It 'blocks both student departure actions after removal from the student role' {
            Mock Get-FreshEntraUserDirectGroup { @() }
            Mock Disable-EntraStudent { throw 'Must not disable' }
            Mock Revoke-EntraStudentSession { throw 'Must not revoke' }
            $entry = [pscustomobject]@{User=[pscustomobject]@{Id='former-student';UserPrincipalName='synthetic@example.org';AccountEnabled=$true}}
            $actions = @(Invoke-StudentDeparture -Entries @($entry) -DisableUsers -RevokeSessions -Confirm:$false)
            $actions.Count | Should -Be 2
            @($actions | Where-Object Status -ne Failed).Count | Should -Be 0
            Should -Invoke Disable-EntraStudent -Exactly 0
            Should -Invoke Revoke-EntraStudentSession -Exactly 0
        }
        It 'blocks student updates after removal from the student role' {
            Mock Get-FreshEntraUserDirectGroup { @() }
            Mock Set-EntraStudentAttribute { throw 'Must not update' }
            $entry = [pscustomobject]@{User=[pscustomobject]@{Id='former-student';UserPrincipalName='synthetic@example.org'}}
            $actions = @(Invoke-StudentUpdate -Entries @($entry) -Snapshot ([pscustomobject]@{}) -Config @{} -File 'synthetic.xlsx' -WorkbookState ([pscustomobject]@{}) -Confirm:$false)
            $actions[0].Status | Should -Be Failed
            Should -Invoke Set-EntraStudentAttribute -Exactly 0
        }
        It 'blocks reports in unignored directories including directories not yet created' {
            $repo = Join-Path $TestDrive 'repo'
            $null = New-Item $repo -ItemType Directory
            & git -C $repo init --quiet
            $LASTEXITCODE | Should -Be 0
            $path = Join-Path $repo 'not-created-yet/report.txt'
            { Write-ComparisonReportFile -Path $path -Text 'Synthetic Person' -Confirm:$false } | Should -Throw '*ignoriert*'
            Test-Path $path | Should -BeFalse
            Test-Path (Split-Path $path) | Should -BeFalse
            Set-Content (Join-Path $repo '.gitignore') '/not-created-yet/'
            Write-ComparisonReportFile -Path $path -Text 'Synthetic Person' -Confirm:$false
            Get-Content $path | Should -Be 'Synthetic Person'
        }
        It 'blocks tracked report destinations even when subsequently ignored' {
            $repo = Join-Path $TestDrive 'tracked'
            $null = New-Item $repo -ItemType Directory
            & git -C $repo init --quiet
            $path = Join-Path $repo 'report.txt'
            Set-Content $path 'original'
            & git -C $repo add report.txt
            $LASTEXITCODE | Should -Be 0
            Remove-Item $path
            Set-Content (Join-Path $repo '.gitignore') '*.txt'
            { Write-ComparisonReportFile -Path $path -Text 'Synthetic Person' -Confirm:$false } | Should -Throw '*versioniert*'
            Test-Path $path | Should -BeFalse
        }
        It 'does not write a report under WhatIf' {
            $path = Join-Path $TestDrive 'whatif.txt'
            Write-ComparisonReportFile -Path $path -Text 'Synthetic Person' -WhatIf
            Test-Path $path | Should -BeFalse
        }
    }
}

Describe 'Audit regressions: personnel recovery checkpoints' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:auditConfig = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $script:auditPerson = [pscustomobject]@{GivenName='Synthetic';Surname='Person';Job='L';RowNumber=2;EntraObjectId='';StoredUpn='';Password='';NameMitRufname='Person, Synthetic'}
            $script:auditEntry = [pscustomobject]@{Person=$script:auditPerson;DesiredState=[pscustomobject]@{UserPrincipalName='person@example.org'}}
            $script:auditUser = [pscustomobject]@{Id='created-person';GivenName='Synthetic';Surname='Person';AccountEnabled=$false;UserType='Member';CompanyName=$script:auditConfig.Profiles.School.CompanyName;EmployeeType=$script:auditConfig.Profiles.School.EmployeeType}
            $script:auditSnapshot = [pscustomobject]@{UsersById=@{'created-person'=$script:auditUser};UsersByUpn=@{};StudentMemberIds=[Collections.Generic.HashSet[string]]::new();PersonnelMemberIds=[Collections.Generic.HashSet[string]]::new()}
            Mock Assert-StudentWorkbookVersion { }
            Mock New-DisabledEntraLehrer { [pscustomobject]@{Id='created-person'} }
            Mock Write-StudentWorkbookUpdate {
                $script:auditPerson.EntraObjectId = $Updates[0].EntraObjectId
                $script:auditPerson.Password = $Updates[0].Password
                [pscustomobject]@{SourceHash='saved'}
            }
            Mock Assert-LehrerAttribute { throw 'Synthetic verification failure' }
            Mock Enable-EntraStudent { throw 'Must remain disabled' }
        }
        It 'persists credentials before group configuration and can resolve the saved identity on retry' {
            $state = [pscustomobject]@{SourceHash='original'}
            $result = Invoke-LehrerCreate -Entry $script:auditEntry -Config $script:auditConfig -File 'synthetic.xlsx' -WorkbookState $state -UsedPasswords ([Collections.Generic.HashSet[string]]::new()) -Confirm:$false
            $result.Status | Should -Be Failed
            $state.SourceHash | Should -Be saved
            $script:auditPerson.Password.Length | Should -Be 12
            $result.RecoveryCommand | Should -Match '-File'
            (Resolve-LehrerIdentity -Person $script:auditPerson -Snapshot $script:auditSnapshot -Config $script:auditConfig).Id | Should -Be created-person
            Should -Invoke Enable-EntraStudent -Exactly 0
        }
        It 'offers a valid retry when creation fails before returning an object ID' -ForEach @(
            @{ File = 'synthetic.xlsx' }, @{ File = '' }
        ) {
            Mock New-DisabledEntraLehrer { throw 'Synthetic create failure' }
            $result = Invoke-LehrerCreate -Entry $script:auditEntry -Config $script:auditConfig -File $File -WorkbookState ([pscustomobject]@{SourceHash='original'}) -UsedPasswords ([Collections.Generic.HashSet[string]]::new()) -Confirm:$false
            $result.Status | Should -Be Failed
            $result.RecoveryCommand | Should -Not -Match '-EntraObjectId'
            if ($File) { $result.RecoveryCommand | Should -Match '-File' }
            else { $result.RecoveryCommand | Should -Match '-Add' }
            Should -Invoke Write-StudentWorkbookUpdate -Exactly 0
            Should -Invoke Enable-EntraStudent -Exactly 0
        }
        It 'requires explicit recovery and password reset when the checkpoint cannot be saved' {
            Mock Write-StudentWorkbookUpdate { throw 'Workbook locked' }
            $result = Invoke-LehrerCreate -Entry $script:auditEntry -Config $script:auditConfig -File 'synthetic.xlsx' -WorkbookState ([pscustomobject]@{SourceHash='original'}) -UsedPasswords ([Collections.Generic.HashSet[string]]::new()) -Confirm:$false
            $result.Status | Should -Be Failed
            $result.RecoveryCommand | Should -Match "-EntraObjectId 'created-person'"
            $result.Message | Should -Match 'Startpasswort administrativ zurücksetzen'
            Should -Invoke Enable-EntraStudent -Exactly 0
        }
        It 'does not treat an arbitrary name match as an unfinished account creation' {
            { Resolve-LehrerIdentity -Person $script:auditPerson -Snapshot $script:auditSnapshot -Config $script:auditConfig } | Should -Throw '*Personalrolle*'
        }
        It 'rejects a checkpoint belonging to another profile' {
            $script:auditPerson.EntraObjectId = 'created-person'; $script:auditPerson.Password = 'SyntheticOnly'
            $script:auditUser.CompanyName = 'Different organisation'
            { Resolve-LehrerIdentity -Person $script:auditPerson -Snapshot $script:auditSnapshot -Config $script:auditConfig } | Should -Throw '*Personalrolle*'
        }
    }
}

Describe 'Audit regressions: real CLI report validation' {
    It 'rejects Excel output before connecting in <Script>' -ForEach @(
        @{Script='Sync-SchuelerEntra.ps1'}, @{Script='Sync-LehrerEntra.ps1'}
    ) {
        $repo = Split-Path $PSScriptRoot -Parent
        $copy = Join-Path $TestDrive "$Script.xlsx"
        Copy-Item (Join-Path $repo 'tests/fixtures/Schueler-Testdaten.xlsx') $copy
        $hash = (Get-FileHash $copy).Hash
        $output = & (Get-Command pwsh).Source -NoProfile -NonInteractive -File (Join-Path $repo $Script) -File $copy -OutputFile $copy 2>&1
        $LASTEXITCODE | Should -Be 1
        ($output -join "`n") | Should -Match 'neue .txt-Datei'
        (Get-FileHash $copy).Hash | Should -Be $hash
    }
}

Describe 'Audit regressions: empty workbook CLI' {
    InModuleScope SchuelerSync {
        It 'rejects a header-only student workbook before any cloud operation' {
            $copy = Join-Path $TestDrive 'empty.xlsx'
            Copy-Item (Join-Path $script:SchuelerSyncRepositoryRoot 'tests/fixtures/Schueler-Testdaten.xlsx') $copy
            $package = Open-StudentWorkbookPackage -Path $copy
            try {
                $sheet = $package.Workbook.Worksheets['Tabelle1']
                $sheet.DeleteRow(2, $sheet.Dimension.End.Row - 1)
                $package.Save()
            } finally { $package.Dispose() }
            $hash = (Get-FileHash $copy).Hash
            $output = & (Get-Command pwsh).Source -NoProfile -NonInteractive -File (Join-Path $script:SchuelerSyncRepositoryRoot 'Sync-SchuelerEntra.ps1') -File $copy -Update 2>&1
            $LASTEXITCODE | Should -Be 1
            ($output -join "`n") | Should -Match 'Leere Schülerliste'
            (Get-FileHash $copy).Hash | Should -Be $hash
        }
    }
}
