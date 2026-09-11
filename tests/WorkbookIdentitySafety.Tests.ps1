BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
}

Describe 'Version-bound workbook identity writes' {
    InModuleScope SchuelerSync {
        BeforeEach {
            $script:testDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
            $null = New-Item -Path $script:testDirectory -ItemType Directory
            $script:workbookFixture = Join-Path $script:SchuelerSyncRepositoryRoot 'tests/fixtures/Schueler-Testdaten.xlsx'
            $script:copy = Join-Path $script:testDirectory 'Studenten.xlsx'
            Copy-Item $script:workbookFixture $script:copy -Force
        }
        It 'refuses an unignored <Unprotected> artifact before the first confidential copy' -ForEach @(
            @{ Unprotected = 'backup'; Patterns = @('/Studenten.xlsx', '/.*.tmp.xlsx') }
            @{ Unprotected = 'temporary'; Patterns = @('/Studenten.xlsx', '/*.backup-*.xlsx') }
            @{ Unprotected = 'both'; Patterns = @('/Studenten.xlsx') }
        ) {
            & git -C $script:testDirectory init --quiet
            $LASTEXITCODE | Should -Be 0
            Set-Content -LiteralPath (Join-Path $script:testDirectory '.gitignore') -Value $Patterns
            $original = Read-StudentWorkbook -Path $script:copy
            Mock Copy-Item { throw 'Confidential copy must not begin.' }
            {
                Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                    [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
                ) -Confirm:$false
            } | Should -Throw '*ignoriert*'
            Should -Invoke Copy-Item -Times 0 -Exactly
            (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash | Should -Be $original.SourceHash
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.xlsx' -Force).Count | Should -Be 1
        }

        It 'protects source backup and temporary workbook inside a custom repository subdirectory' {
            & git -C $script:testDirectory init --quiet
            $LASTEXITCODE | Should -Be 0
            Set-Content -LiteralPath (Join-Path $script:testDirectory '.gitignore') -Value '*.[xX][lL][sS][xX]'
            $nested = Join-Path $script:testDirectory 'import [2026]'
            $null = New-Item -Path $nested -ItemType Directory
            $source = Join-Path $nested 'Eigene Liste.XLSX'
            Copy-Item -LiteralPath $script:workbookFixture -Destination $source
            $original = Read-StudentWorkbook -Path $source
            $result = Write-StudentWorkbookUpdates -Path $source -ExpectedSourceHash $original.SourceHash -Updates @(
                [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
            ) -Confirm:$false
            (Read-StudentWorkbook -Path $source).Students[0].Password | Should -Be 'TigerWiese56'
            Test-Path -LiteralPath $result.BackupPath | Should -BeTrue
            @(& git -C $script:testDirectory ls-files --others --exclude-standard '*.XLSX' '*.xlsx').Count | Should -Be 0
            $LASTEXITCODE | Should -Be 0
        }

        It 'rejects a generated path already tracked in Git even when its file was removed' {
            & git -C $script:testDirectory init --quiet
            $LASTEXITCODE | Should -Be 0
            Set-Content -LiteralPath (Join-Path $script:testDirectory '.gitignore') -Value '*.xlsx'
            $artifact = Join-Path $script:testDirectory '.Studenten.known.tmp.xlsx'
            Set-Content -LiteralPath $artifact -Value 'synthetic test placeholder'
            & git -C $script:testDirectory add --force -- '.Studenten.known.tmp.xlsx'
            $LASTEXITCODE | Should -Be 0
            Remove-Item -LiteralPath $artifact
            { Assert-WorkbookArtifactGitSafety -Path $artifact } | Should -Throw '*versioniert*'
            Test-Path -LiteralPath $artifact | Should -BeFalse
        }

        It 'does not mistake a dot-prefixed artifact name for a path outside Git' {
            & git -C $script:testDirectory init --quiet
            $LASTEXITCODE | Should -Be 0
            $artifact = Join-Path $script:testDirectory '..Studenten.known.tmp.xlsx'
            { Assert-WorkbookArtifactGitSafety -Path $artifact } | Should -Throw '*ignoriert*'
        }

        It 'rejects swapped rows before writing credentials to either student' {
            $original = Read-StudentWorkbook -Path $script:copy
            $package = Open-ExcelPackage -Path $script:copy
            try {
                $sheet = $package.Workbook.Worksheets['Tabelle1']
                for ($column = 1; $column -le 5; $column++) {
                    $first = $sheet.Cells[2, $column].Value
                    $sheet.Cells[2, $column].Value = $sheet.Cells[3, $column].Value
                    $sheet.Cells[3, $column].Value = $first
                }
                $package.Save()
            } finally { $package.Dispose() }
            $editedHash = (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash
            {
                Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                    [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
                ) -Confirm:$false
            } | Should -Throw '*Vorprüfung*'
            (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash | Should -Be $editedHash
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.backup-*.xlsx').Count | Should -Be 0
        }
        It 'restores late concurrent edits made after the final source check without installing credentials' {
            $original = Read-StudentWorkbook -Path $script:copy
            Mock Move-StudentWorkbookFile {
                if ($Source -eq $script:copy) {
                    $package = Open-ExcelPackage -Path $Source
                    try {
                        $sheet = $package.Workbook.Worksheets['Tabelle1']
                        for ($column = 1; $column -le 5; $column++) {
                            $first = $sheet.Cells[2, $column].Value
                            $sheet.Cells[2, $column].Value = $sheet.Cells[3, $column].Value
                            $sheet.Cells[3, $column].Value = $first
                        }
                        $package.Save()
                    } finally { $package.Dispose() }
                    $script:lateEditHash = (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash
                }
                [IO.File]::Move($Source, $Destination)
            }
            {
                Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                    [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
                ) -Confirm:$false
            } | Should -Throw '*Vorprüfung*'
            (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash | Should -Be $script:lateEditHash
            (Read-StudentWorkbook -Path $script:copy).Students.Password | Should -Not -Contain 'TigerWiese56'
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.tmp.xlsx' -Force).Count | Should -Be 0
        }
        It 'preserves a concurrent replacement plus displaced source and staged workbook when the original path is recreated' {
            $original = Read-StudentWorkbook -Path $script:copy
            Mock Move-StudentWorkbookFile {
                if ($Destination -eq $script:copy -and $Source -like '*.tmp.xlsx') {
                    Copy-Item -LiteralPath $script:workbookFixture -Destination $Destination
                    $package = Open-ExcelPackage -Path $Destination
                    try {
                        $package.Workbook.Worksheets['Tabelle1'].Cells[2, 1].Value = 'Parallel bearbeitete Datei'
                        $package.Save()
                    } finally { $package.Dispose() }
                    $script:replacementHash = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
                }
                [IO.File]::Move($Source, $Destination)
            }
            $failure = $null
            try {
                Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                    [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
                ) -Confirm:$false
            } catch { $failure = $_.Exception.Message }
            $failure | Should -Not -BeNullOrEmpty
            $failure | Should -Not -Match 'TigerWiese56'
            (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash | Should -Be $script:replacementHash
            (Read-StudentWorkbook -Path $script:copy).Students.Password | Should -Not -Contain 'TigerWiese56'
            $backup = @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.backup-*.xlsx')
            $staged = @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.tmp.xlsx' -Force)
            $backup.Count | Should -Be 1
            $staged.Count | Should -Be 1
            (Get-FileHash -LiteralPath $backup[0].FullName -Algorithm SHA256).Hash | Should -Be $original.SourceHash
            $failure | Should -Match ([regex]::Escape($backup[0].FullName))
            $failure | Should -Match ([regex]::Escape($staged[0].FullName))
        }
        It 'restores the displaced source if installing the prepared workbook fails while the original path is vacant' {
            $original = Read-StudentWorkbook -Path $script:copy
            Mock Move-StudentWorkbookFile {
                if ($Source -like '*.tmp.xlsx') { throw 'Simulated install failure' }
                [IO.File]::Move($Source, $Destination)
            }
            {
                Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                    [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
                ) -Confirm:$false
            } | Should -Throw '*Simulated install failure*'
            (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash | Should -Be $original.SourceHash
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.tmp.xlsx' -Force).Count | Should -Be 0
        }
        It 'leaves exactly the previous bytes in the backup after a successful commit' {
            $original = Read-StudentWorkbook -Path $script:copy
            $result = Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
            ) -Confirm:$false
            (Get-FileHash -LiteralPath $result.BackupPath -Algorithm SHA256).Hash | Should -Be $original.SourceHash
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.backup-*.xlsx').Count | Should -Be 1
            (Read-StudentWorkbook -Path $script:copy).Students[0].Password | Should -Be 'TigerWiese56'
        }
        It 'leaves the source intact when a Windows handle denies renaming at commit time' -Skip:(-not $IsWindows) {
            $original = Read-StudentWorkbook -Path $script:copy
            Mock Move-StudentWorkbookFile {
                if ($Source -eq $script:copy) {
                    $handle = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
                    try { [IO.File]::Move($Source, $Destination) } finally { $handle.Dispose() }
                } else { [IO.File]::Move($Source, $Destination) }
            }
            {
                Write-StudentWorkbookUpdates -Path $script:copy -ExpectedSourceHash $original.SourceHash -Updates @(
                    [pscustomobject]@{ RowNumber = 2; Password = 'TigerWiese56'; EntraObjectId = 'new-user'; UPN = 'mmuster@monteaufkirchen.com' }
                ) -Confirm:$false
            } | Should -Throw
            (Get-FileHash -LiteralPath $script:copy -Algorithm SHA256).Hash | Should -Be $original.SourceHash
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.backup-*.xlsx').Count | Should -Be 0
            @(Get-ChildItem -LiteralPath $script:testDirectory -Filter '*.tmp.xlsx' -Force).Count | Should -Be 0
        }
        It 'keeps a new account disabled when rows are swapped after Graph creation' {
            $original = Read-StudentWorkbook -Path $script:copy
            $entry = [pscustomobject]@{
                Student = $original.Students[0]
                DesiredState = [pscustomobject]@{ UserPrincipalName = 'mmuster@monteaufkirchen.com'; ManagerId = 'teacher'; RequiredGroupNames = @('role', 'license', 'class') }
            }
            Mock New-StudentPassword { 'TigerWiese56' }
            Mock New-StudentWithPasswordRetry {
                $package = Open-ExcelPackage -Path $script:copy
                try {
                    $sheet = $package.Workbook.Worksheets['Tabelle1']
                    for ($column = 1; $column -le 5; $column++) {
                        $first = $sheet.Cells[2, $column].Value
                        $sheet.Cells[2, $column].Value = $sheet.Cells[3, $column].Value
                        $sheet.Cells[3, $column].Value = $first
                    }
                    $package.Save()
                } finally { $package.Dispose() }
                [pscustomobject]@{ User = [pscustomobject]@{ Id = 'new-user' }; Password = $InitialPassword }
            }
            Mock Assert-NewEntraStudentAttributes { $true }
            Mock Set-EntraStudentManager { [pscustomobject]@{ Verified = $true } }
            Mock Sync-EntraStudentGroups { [pscustomobject]@{ Verified = $true } }
            Mock Enable-EntraStudent { throw 'Must remain disabled' }
            $snapshot = [pscustomobject]@{ GroupsByDisplayName = @{ role = 'role'; license = 'license'; class = 'class' } }
            $result = @(Invoke-StudentCreateBatch -Entries @($entry) -Snapshot $snapshot -Config @{ RoleGroupPrefix = 'role'; ClassGroupPrefix = 'class' } -File $script:copy -WorkbookState $original -Confirm:$false)
            $result[0].Status | Should -Be Failed
            $result[0].Phase | Should -Be Workbook
            Should -Invoke Enable-EntraStudent -Times 0
            (Read-StudentWorkbook -Path $script:copy).Students.Password | Should -Not -Contain 'TigerWiese56'
        }
        It 'preserves stable matching and password when attribute verification fails: <AttributesFail>' -ForEach @(
            @{ AttributesFail = $false }, @{ AttributesFail = $true }
        ) {
            $script:attributeFailure = $AttributesFail
            $package = Open-ExcelPackage -Path $script:copy
            try {
                $sheet = $package.Workbook.Worksheets['Tabelle1']
                $sheet.Cells[2, 3].Value = 'Neu'
                $sheet.Cells[1, 6].Value = 'Passwort'
                $sheet.Cells[1, 7].Value = 'EntraObjectId'
                $sheet.Cells[1, 8].Value = 'UPN'
                $sheet.Cells[2, 6].Value = 'TigerWiese56'
                $sheet.Cells[2, 7].Value = ''
                $sheet.Cells[2, 8].Value = 'mmuster@monteaufkirchen.com'
                $package.Save()
            } finally { $package.Dispose() }
            $original = Read-StudentWorkbook -Path $script:copy
            $script:currentUser = [pscustomobject]@{ Id = 'stable-user-id'; UserPrincipalName = 'mmuster@monteaufkirchen.com' }
            $snapshot = [pscustomobject]@{ UsersById = @{ 'stable-user-id' = $script:currentUser } }
            (Resolve-StudentIdentity -Student $original.Students[0] -Snapshot $snapshot).Method | Should -Be StoredUpn
            $entry = [pscustomobject]@{
                Student = $original.Students[0]; User = $script:currentUser
                DesiredState = [pscustomobject]@{ UserPrincipalName = 'mneu@monteaufkirchen.com' }
                Differences = @([pscustomobject]@{ Area = 'Entra'; Field = 'UserPrincipalName'; Action = 'Set' })
            }
            Mock Set-EntraStudentAttributes {
                $checkpoint = Read-StudentWorkbook -Path $script:copy
                $checkpoint.Students[0].EntraObjectId | Should -Be 'stable-user-id'
                $checkpoint.Students[0].Password | Should -Be 'TigerWiese56'
                $script:currentUser.UserPrincipalName = 'mneu@monteaufkirchen.com'
                if ($script:attributeFailure) { throw 'Attribute verification failed after accepted rename' }
                [pscustomobject]@{ Verified = $true }
            }
            Mock Get-EntraStudentCurrentIdentity { $script:currentUser }
            $result = @(Invoke-StudentUpdates -Entries @($entry) -Snapshot $snapshot -Config @{} -File $script:copy -WorkbookState $original -Confirm:$false)
            $result[0].Status | Should -Be $(if ($AttributesFail) { 'Failed' } else { 'Succeeded' })
            $next = Read-StudentWorkbook -Path $script:copy
            $next.Students[0].StoredUpn | Should -Be 'mneu@monteaufkirchen.com'
            $next.Students[0].Password | Should -Be 'TigerWiese56'
            $identity = Resolve-StudentIdentity -Student $next.Students[0] -Snapshot $snapshot
            $identity.Method | Should -Be EntraObjectId
            $identity.User.Id | Should -Be 'stable-user-id'
        }
    }
}
