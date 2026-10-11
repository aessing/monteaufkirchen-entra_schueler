BeforeDiscovery {
    Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}
BeforeAll {
    $global:WorkbookExportStubs = @()
    foreach ($name in @('Get-MgGroup','Get-MgUser','Get-MgGroupMember','Get-EXORecipient')) {
        if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
            Set-Item "function:global:$name" {[CmdletBinding()]param($UserId,$GroupId,$ExternalDirectoryObjectId,[switch]$All,$Property,$Properties,$ResultSize)}
            $global:WorkbookExportStubs += $name
        }
    }
}
AfterAll { foreach ($name in $global:WorkbookExportStubs) { Remove-Item "function:global:$name" }; Remove-Variable WorkbookExportStubs -Scope Global }

Describe 'Student and teacher workbook inventory' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:ExportStudent = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/SchuelerSync.psd1')
            $script:ExportTeacher = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $script:ExportGroups = @(@($script:ExportStudent.StudentRoleGroup) + @($script:ExportTeacher.Roles.Values) | ForEach-Object { [pscustomobject]@{Id=$_.Id;DisplayName=$_.Name} })
            $script:ExportUser = [pscustomobject]@{Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald'
                UserPrincipalName='ltestwald@example.invalid';Mail='ltestwald@example.invalid';JobTitle='PA, JAS';AccountEnabled=$true;UserType='Member';OnPremisesSyncEnabled=$false;SignInActivity=[pscustomobject]@{LastSuccessfulSignInDateTime='2026-01-02T12:30:00+02:00'}}
            $script:ExportUsers = @($script:ExportUser,[pscustomobject]@{Id='ffffffff-ffff-ffff-ffff-ffffffffffff';UserPrincipalName='teacher@example.invalid'})
            $script:ExportMembers = @($script:ExportUser)
            $script:ExportRoleIds = @($script:ExportStudent.StudentRoleGroup.Id,$script:ExportTeacher.Roles.School.Id)
            $script:ExportRecipients = @()
            Mock Connect-DisabledAccountGraph { [pscustomobject]@{TenantId='synthetic-tenant'} }
            Mock Connect-SchuelerExchangeOnline { }
            Mock Get-MgGroup { $script:ExportGroups }
            Mock Get-MgUser { $script:ExportUsers }
            Mock Get-MgGroupMember { if ($GroupId -in $script:ExportRoleIds) { $script:ExportMembers } }
            Mock Get-EXORecipient { $script:ExportRecipients }
            Mock Get-FreshEntraUserDirectGroup { [pscustomobject]@{Id='eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';DisplayName='SEC-A-CLS-JK1-3g2_1'} }
            Mock Get-UserManagerId { 'ffffffff-ffff-ffff-ffff-ffffffffffff' }
        }
        It 'exports the direct class and manager UPN with read scopes for sign-in activity' {
            $inventory = Get-RoleWorkbookInventory -Kind Student
            $inventory.Accounts.Count | Should -Be 1
            $inventory.Accounts[0].ClassName | Should -BeExactly 'JK1-3g2_1'
            $inventory.Accounts[0].Teacher | Should -BeExactly 'teacher@example.invalid'
            $inventory.Warnings.Count | Should -Be 0
            $inventory.Accounts[0].LastLogin.ToUniversalTime().ToString('yyyy-MM-dd HH:mm') | Should -Be '2026-01-02 10:30'
            Should -Invoke Connect-DisabledAccountGraph -Exactly 1 -ParameterFilter { $RequiredScopes.Count -eq 3 -and 'User.Read.All' -in $RequiredScopes -and 'Group.Read.All' -in $RequiredScopes -and 'AuditLog.Read.All' -in $RequiredScopes }
            Should -Invoke Connect-SchuelerExchangeOnline -Exactly 1 -ParameterFilter { $TenantId -eq 'synthetic-tenant' }
        }
        It 'keeps disabled role members and reports the existing reactivation semantics' {
            $script:ExportUser.AccountEnabled = $false
            foreach ($kind in @('Student','Teacher')) {
                $inventory = Get-RoleWorkbookInventory -Kind $kind
                $inventory.Accounts.Count | Should -Be 1
                $inventory.Warnings -join ' ' | Should -Match 'gesperrt.*reaktivieren'
            }
        }
        It 'preserves the last successful login for both role exports and never uses attempted logins' {
            $script:ExportUser.SignInActivity | Add-Member -NotePropertyMembers @{LastSignInDateTime='2026-09-28T12:00:00Z';LastNonInteractiveSignInDateTime='2026-09-29T12:00:00Z'}
            foreach ($kind in @('Student','Teacher')) {
                $account = (Get-RoleWorkbookInventory -Kind $kind).Accounts[0]
                $account.LastLogin.ToUniversalTime().ToString('yyyy-MM-dd HH:mm') | Should -Be '2026-01-02 10:30'
                $account.LastLoginStatus | Should -BeExactly 'Bekannt'
            }
            $script:ExportUser.SignInActivity.LastSuccessfulSignInDateTime = $null
            (Get-RoleWorkbookInventory -Kind Teacher).Accounts[0].LastLogin | Should -BeNullOrEmpty
        }
        It 'retries without sign-in activity and reports unavailable values for both role exports' {
            Mock Get-MgUser { if ($Property -contains 'signInActivity') { throw 'synthetic sign-in access denied' }; $script:ExportUsers }
            foreach ($kind in @('Student','Teacher')) {
                $inventory = Get-RoleWorkbookInventory -Kind $kind
                $inventory.Accounts[0].LastLogin | Should -BeNullOrEmpty
                $inventory.Accounts[0].LastLoginStatus | Should -BeExactly 'Nicht abrufbar'
                $inventory.Warnings -join ' ' | Should -Match 'Letzte Anmeldung nicht abrufbar.*synthetic sign-in access denied'
            }
            Should -Invoke Get-MgUser -Exactly 2 -ParameterFilter { $Property -contains 'signInActivity' }
            Should -Invoke Get-MgUser -Exactly 2 -ParameterFilter { $Property -notcontains 'signInActivity' }
        }
        It 'normalizes the actual teacher job instead of guessing from the role' {
            $inventory = Get-RoleWorkbookInventory -Kind Teacher
            $inventory.Accounts[0].Job | Should -BeExactly 'PA, JAS'
            $script:ExportUser.JobTitle = ''
            $inventory = Get-RoleWorkbookInventory -Kind Teacher
            $inventory.Accounts[0].Job | Should -BeNullOrEmpty
            $inventory.Warnings -join ' ' | Should -Match 'Jobzelle bleibt leer'
        }
        It 'exports OGTS with its actual Ganztag role' {
            $script:ExportUser.JobTitle = 'OGTS'
            $script:ExportRoleIds = @($script:ExportTeacher.Roles.Ganztag.Id)
            (Get-RoleWorkbookInventory -Kind Teacher).Accounts[0].Job | Should -BeExactly 'OGTS'
        }
        It 'deduplicates cross-role members and leaves a conflicting job empty' {
            $script:ExportRoleIds += $script:ExportTeacher.Roles.Ganztag.Id
            $inventory = Get-RoleWorkbookInventory -Kind Teacher
            $inventory.Accounts.Count | Should -Be 1
            $inventory.Accounts[0].Job | Should -BeNullOrEmpty
            $inventory.Warnings.Count | Should -Be 1
        }
        It 'keeps missing class and unreadable manager cells empty with reasons' {
            Mock Get-FreshEntraUserDirectGroup { }
            Mock Get-UserManagerId { throw 'synthetic manager read failure' }
            $inventory = Get-RoleWorkbookInventory -Kind Student
            $inventory.Accounts[0].ClassName | Should -BeNullOrEmpty
            $inventory.Accounts[0].Teacher | Should -BeNullOrEmpty
            $inventory.Warnings -join ' ' | Should -Match 'Klasse.*Klassenlehrer'
        }
        It 'does not choose the first of multiple class groups' {
            Mock Get-FreshEntraUserDirectGroup { [pscustomobject]@{DisplayName='SEC-A-CLS-A'}; [pscustomobject]@{DisplayName='SEC-A-CLS-B'} }
            (Get-RoleWorkbookInventory -Kind Student).Accounts[0].ClassName | Should -BeNullOrEmpty
        }
        It 'excludes shared mailboxes even behind duplicate recipient IDs' {
            $script:ExportRecipients = @([pscustomobject]@{ExternalDirectoryObjectId=$script:ExportUser.Id;RecipientTypeDetails='UserMailbox'},
                [pscustomobject]@{ExternalDirectoryObjectId=$script:ExportUser.Id;RecipientTypeDetails='SharedMailbox'})
            $inventory = Get-RoleWorkbookInventory -Kind Teacher
            $inventory.Accounts.Count | Should -Be 0
            $inventory.ExcludedCount | Should -Be 1
        }
        It 'aborts on missing roles, unreadable members and failed Exchange reads' {
            $script:ExportGroups = @()
            { Get-RoleWorkbookInventory -Kind Student } | Should -Throw '*Exportrolle*'
            $script:ExportGroups = @([pscustomobject]@{Id=$script:ExportStudent.StudentRoleGroup.Id;DisplayName=$script:ExportStudent.StudentRoleGroup.Name})
            $script:ExportMembers = @([pscustomobject]@{Id='missing-user'})
            { Get-RoleWorkbookInventory -Kind Student } | Should -Throw '*Benutzerbestand*'
            $script:ExportMembers = @($script:ExportUser)
            Mock Get-EXORecipient { throw 'synthetic recipient failure' }
            { Get-RoleWorkbookInventory -Kind Student } | Should -Throw '*synthetic recipient failure*'
        }
        It 'refuses an existing target before reading any tenant and preserves it' {
            $path = Join-Path $TestDrive 'Existing.xlsx'
            Set-Content $path 'synthetic unchanged'
            { Invoke-RoleWorkbookExport -Kind Student -File $path -Confirm:$false } | Should -Throw '*existiert bereits*'
            Get-Content $path | Should -BeExactly 'synthetic unchanged'
            Should -Invoke Connect-DisabledAccountGraph -Exactly 0
        }
        It 'rejects missing or unexpected tenant identities before reading Exchange' {
            Mock Connect-DisabledAccountGraph { [pscustomobject]@{TenantId=''} }
            { Get-RoleWorkbookInventory -Kind Student } | Should -Throw '*Tenant-ID*'
            Mock Connect-DisabledAccountGraph { [pscustomobject]@{TenantId='synthetic-tenant'} }
            Mock Import-PowerShellDataFile { $config = @{} + $script:ExportStudent; $config.ExpectedTenantId='other-synthetic-tenant'; $config }
            { Get-RoleWorkbookInventory -Kind Student } | Should -Throw '*Graph-Tenant*'
            Should -Invoke Connect-SchuelerExchangeOnline -Exactly 0
        }
        It 'ignores nested group objects without inventing user rows' {
            $script:ExportMembers += [pscustomobject]@{Id='nested-group';AdditionalProperties=@{'@odata.type'='#microsoft.graph.group'}}
            (Get-RoleWorkbookInventory -Kind Student).Accounts.Count | Should -Be 1
        }
        It 'completes export progress after a failed inventory read without writing a file' {
            Mock Get-RoleWorkbookInventory { throw 'synthetic inventory failure' }
            Mock Write-Progress { }
            $path = Join-Path $TestDrive 'Failed.xlsx'
            { Invoke-RoleWorkbookExport -Kind Student -File $path } | Should -Throw '*synthetic inventory failure*'
            Test-Path $path | Should -BeFalse
            Should -Invoke Write-Progress -Exactly 1 -ParameterFilter { $Completed }
        }
        It 'keeps the public module defaults and Confirm forwarding for the export mode' {
            Mock Invoke-RoleWorkbookExport { [pscustomobject]@{HasErrors=$false} }
            $null = Invoke-SchuelerSync -ListToExcel -WhatIf -Confirm:$false
            $null = Invoke-LehrerSync -ListToExcel -WhatIf -Confirm:$false
            Should -Invoke Invoke-RoleWorkbookExport -Exactly 1 -ParameterFilter { $Kind -eq 'Student' -and $File -eq (Join-Path $script:SchuelerSyncRepositoryRoot 'Schueler.xlsx') -and $WhatIf -and -not $Confirm }
            Should -Invoke Invoke-RoleWorkbookExport -Exactly 1 -ParameterFilter { $Kind -eq 'Teacher' -and $File -eq (Join-Path $script:SchuelerSyncRepositoryRoot 'Lehrer.xlsx') -and $WhatIf -and -not $Confirm }
        }
    }
}

Describe 'Native student and teacher export schemas' {
    InModuleScope SchuelerSync {
        It 'appends a UTC login timestamp without changing identity or password columns for <Kind>' -ForEach @(
            @{Kind='Student';Column=10;PasswordColumn=6}, @{Kind='Teacher';Column=9;PasswordColumn=5}, @{Kind='Personnel';Column=9;PasswordColumn=5}
        ) {
            $account = [pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';ClassName='JK1-3g2_1';Teacher='teacher@example.invalid'
                Job=$(if ($Kind -eq 'Personnel') { 'Schule_Sekretariat' } else { 'PA, JAS' });UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail='ltestwald@example.invalid'
                LastLogin=[datetimeoffset]'2026-01-02T12:30:00+02:00'}
            foreach ($hasLogin in @($true,$false)) {
                if (-not $hasLogin) { $account.LastLogin = $null }
                $path = Join-Path $TestDrive "$Kind-Login-$hasLogin.xlsx"
                $null = Export-AccountWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Kind $Kind -Confirm:$false
                $package = Open-StudentWorkbookPackage $path
                try {
                    $sheet = $package.Workbook.Worksheets['Tabelle1']
                    $sheet.Dimension.End.Column | Should -Be $Column
                    $sheet.Cells[1,$Column].Value | Should -BeExactly 'Last Login Timestamp'
                    $sheet.Tables[0].Address.End.Column | Should -Be $Column
                    $sheet.Cells[2,$PasswordColumn].Value | Should -BeNullOrEmpty
                    $sheet.Cells[2,$Column].Formula | Should -BeNullOrEmpty
                    if ($hasLogin) { $sheet.Cells[2,$Column].Value | Should -BeExactly '2026-01-02T10:30:00Z' }
                    else { $sheet.Cells[2,$Column].Value | Should -BeNullOrEmpty }
                } finally { $package.Dispose() }
                $data = Read-StudentWorkbook -Path $path -Kind $Kind
                $data.Students[0].EntraObjectId | Should -BeExactly $account.UserId
                $data.Students[0].Password | Should -BeNullOrEmpty
            }
        }
        It 'round-trips a student with ten columns, class, manager and empty password' {
            $account = [pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';ClassName='JK1-3g2_1';Teacher='teacher@example.invalid'
                UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail='ltestwald@example.invalid'}
            $path = Join-Path $TestDrive 'Student.xlsx'
            $null = Export-AccountWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Kind Student -Confirm:$false
            $package = Open-StudentWorkbookPackage $path
            try {
                $sheet = $package.Workbook.Worksheets['Tabelle1']
                @((1..10) | ForEach-Object { $sheet.Cells[1,$_].Value }) | Should -Be @('Name mit Rufname','Vorname','Nachname','Klassen','Klassenlehrer','Passwort','EntraObjectId','UPN','Mail','Last Login Timestamp')
                $sheet.Tables[0].StyleName | Should -BeExactly 'TableStyleMedium14'
                $sheet.Cells[2,6].Value | Should -BeNullOrEmpty
                $sheet.Cells[2,7].Value | Should -BeExactly $account.UserId
            } finally { $package.Dispose() }
            $data = Read-StudentWorkbook -Path $path
            $data.Students[0].ClassName | Should -BeExactly $account.ClassName
            $data.Students[0].Teacher | Should -BeExactly $account.Teacher
            $data.Students[0].Password | Should -BeNullOrEmpty
        }
        It 'round-trips a teacher with nine columns and canonical jobs' {
            $account = [pscustomobject]@{DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald';Job='L, PA';UserId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';UPN='ltestwald@example.invalid';Mail='ltestwald@example.invalid'}
            $path = Join-Path $TestDrive 'Teacher.xlsx'
            $null = Export-AccountWorkbook -Inventory ([pscustomobject]@{Accounts=@($account)}) -Path $path -Kind Teacher -Confirm:$false
            $data = Read-StudentWorkbook -Path $path -Kind Teacher
            $data.Students[0].Job | Should -BeExactly 'L, PA'
            $data.Students[0].EntraObjectId | Should -BeExactly $account.UserId
            $data.Students[0].Password | Should -BeNullOrEmpty
        }
        It 'keeps formula-looking student values literal and supports an empty schema' {
            $account = [pscustomobject]@{DisplayName='=SYNTHETIC()';GivenName='';Surname='';ClassName='=CLASS()';Teacher='=MANAGER()';UserId='synthetic-id';UPN='student@example.invalid';Mail=''}
            foreach ($empty in @($false,$true)) {
                $path = Join-Path $TestDrive "Student-$empty.xlsx"
                $accounts = if ($empty) { @() } else { @($account) }
                $null = Export-AccountWorkbook -Inventory ([pscustomobject]@{Accounts=$accounts}) -Path $path -Kind Student -Confirm:$false
                $package = Open-StudentWorkbookPackage $path
                try {
                    $sheet = $package.Workbook.Worksheets['Tabelle1']
                    $sheet.Dimension.End.Column | Should -Be 10
                    if ($empty) { $sheet.Dimension.End.Row | Should -Be 1 } else {
                        $sheet.Cells[2,4].Value | Should -BeExactly '=CLASS()'
                        $sheet.Cells[2,5].Formula | Should -BeNullOrEmpty
                    }
                } finally { $package.Dispose() }
            }
        }
    }
}

Describe 'Actual isolated student and teacher export CLI' {
    It 'executes both entrypoints and WhatIf without touching remote systems' {
        foreach ($population in @('Student','Teacher')) {
            $path = Join-Path $TestDrive "$population-CLI.xlsx"
            $harness = Join-Path $PSScriptRoot 'Invoke-SyntheticPersonalCli.ps1'
            $output = & pwsh -NoLogo -NoProfile -NonInteractive -File $harness -Mode Export -Population $population -File $path 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match '"HasErrors": false'
            Test-Path $path | Should -BeTrue
            & (Get-Module SchuelerSync) {
                param($WorkbookPath)
                $package = Open-StudentWorkbookPackage $WorkbookPath
                try {
                    $sheet = $package.Workbook.Worksheets['Tabelle1']
                    $sheet.Cells[2,$sheet.Dimension.End.Column].Value | Should -BeExactly '2026-01-02T10:30:00Z'
                } finally { $package.Dispose() }
            } $path
            $hash = (Get-FileHash $path).Hash
            $output = & pwsh -NoLogo -NoProfile -NonInteractive -File $harness -Mode Export -Population $population -File $path 2>&1
            $LASTEXITCODE | Should -Be 1
            $output -join "`n" | Should -Match 'existiert bereits'
            (Get-FileHash $path).Hash | Should -Be $hash
            $preview = Join-Path $TestDrive "$population-Preview/New.xlsx"
            $output = & pwsh -NoLogo -NoProfile -NonInteractive -File $harness -Mode Export -Population $population -File $preview -Preview 2>&1
            $LASTEXITCODE | Should -Be 0 -Because ($output -join "`n")
            $output -join "`n" | Should -Match '"Status": "WhatIf"'
            Test-Path (Split-Path $preview -Parent) | Should -BeFalse
        }
    }
    It 'rejects conflicting selectors before connecting' {
        foreach ($entry in @('Sync-SchuelerEntra.ps1','Sync-LehrerEntra.ps1')) {
            $path = Join-Path (Split-Path $PSScriptRoot -Parent) $entry
            foreach ($selector in @('-Update','-Add','-OutputFile')) {
                $arguments = @('-NoLogo','-NoProfile','-NonInteractive','-File',$path,'-ListToExcel',$selector)
                if ($selector -eq '-OutputFile') { $arguments += 'synthetic.txt' }
                $output = & pwsh @arguments 2>&1
                $LASTEXITCODE | Should -Be 1
                $output -join "`n" | Should -Match 'Parameter set cannot be resolved|Parametersatz.*nicht aufgelöst'
            }
        }
    }
}
