BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $fixture = Join-Path $PSScriptRoot 'fixtures/Schueler-Testdaten.xlsx'
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force
}

Describe 'Version-bound workbook identity writes' {
    InModuleScope SchuelerSync -Parameters @{ Fixture = $fixture } {
        BeforeEach {
            $script:copy = Join-Path $TestDrive 'Studenten.xlsx'
            Copy-Item $Fixture $script:copy -Force
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
            @(Get-ChildItem -LiteralPath $TestDrive -Filter '*.backup-*.xlsx').Count | Should -Be 0
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
