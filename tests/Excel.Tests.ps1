BeforeDiscovery {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $manifest = Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1'
    Import-Module $manifest -ErrorAction Stop
}

Describe 'Student workbook adapter' {
    InModuleScope SchuelerSync {
        BeforeAll {
            $script:ExcelTestFixture = Join-Path $script:SchuelerSyncRepositoryRoot 'tests/fixtures/Schueler-Testdaten.xlsx'
        }
        It 'reads the required sheet and maps the student rows' {
            $copy = Join-Path $TestDrive 'Eigener-Dateiname.xlsx'
            Copy-Item $script:ExcelTestFixture $copy

            $context = Read-StudentWorkbook -Path $copy

            $context.WorksheetName | Should -Be 'Tabelle1'
            $context.Students.Count | Should -Be 2
            $context.Students[0].RowNumber | Should -Be 2
            $context.Students[0].ClassName | Should -Be 'JK1-3g2_1'
            $context.Students[0].GivenName | Should -Be 'Mia'
            $context.Students[0].Surname | Should -Be 'Muster'
        }

        It 'reads a literal bracket filename without treating it as a wildcard pattern' {
            $literalPath = Join-Path $TestDrive 'Students[2026].xlsx'
            Copy-Item -LiteralPath $script:ExcelTestFixture -Destination $literalPath

            $context = Read-StudentWorkbook -Path $literalPath

            $context.Path | Should -Be $literalPath
            $context.Students.Count | Should -Be 2
            $context.Students[0].GivenName | Should -Be 'Mia'
        }

        It 'rejects a workbook path that is not an xlsx leaf' {
            { Resolve-StudentWorkbookPath -Path (Join-Path $TestDrive 'alt.xls') } |
                Should -Throw '*.xlsx*'
        }

        It 'rejects a partially filled mandatory student row before mutation' {
            $copy = Join-Path $TestDrive 'Teilweise.xlsx'
            Copy-Item $script:ExcelTestFixture $copy
            $package = Open-ExcelPackage -Path $copy
            try {
                $package.Workbook.Worksheets['Tabelle1'].Cells[4, 1].Value = 'Unvollständig, Uma'
                Close-ExcelPackage -ExcelPackage $package
                $package = $null
            } finally {
                if ($null -ne $package) { Close-ExcelPackage -ExcelPackage $package }
            }

            { Read-StudentWorkbook -Path $copy } | Should -Throw '*Tabelle1*4*'
        }

        It 'writes verified identity data and retains a backup' {
            $copy = Join-Path $TestDrive 'Eigener-Dateiname.xlsx'
            Copy-Item $script:ExcelTestFixture $copy
            $objectId = [guid]::NewGuid().Guid

            $result = Write-StudentWorkbookUpdate -Path $copy -Updates @(
                [pscustomobject]@{
                    RowNumber = 2
                    Password = 'TigerWiese56'
                    EntraObjectId = $objectId
                    UPN = 'mmuster@monteaufkirchen.com'
                    Mail = 'mmuster@monteaufkirchen.com'
                }
            ) -SkipGitSafetyCheck

            $result.Path | Should -Be $copy
            (Test-Path -LiteralPath $result.BackupPath) | Should -BeTrue
            $saved = @(Import-Excel -Path $copy -WorksheetName Tabelle1)
            $saved[0].Passwort | Should -Be 'TigerWiese56'
            $saved[0].EntraObjectId | Should -Be $objectId
            $saved[0].UPN | Should -Be 'mmuster@monteaufkirchen.com'
            $saved[0].Mail | Should -Be 'mmuster@monteaufkirchen.com'
            (Read-StudentWorkbook -Path $copy).Students[0].StoredMail | Should -Be 'mmuster@monteaufkirchen.com'
        }

        It 'does not replace an existing password with an empty update' {
            $copy = Join-Path $TestDrive 'Bestehendes-Passwort.xlsx'
            Copy-Item $script:ExcelTestFixture $copy
            $initialObjectId = [guid]::NewGuid().Guid
            Write-StudentWorkbookUpdate -Path $copy -Updates @(
                [pscustomobject]@{
                    RowNumber = 2
                    Password = 'TigerWiese56'
                    EntraObjectId = $initialObjectId
                    UPN = 'mmuster@monteaufkirchen.com'
                    Mail = 'mmuster@monteaufkirchen.com'
                }
            ) -SkipGitSafetyCheck | Out-Null

            Write-StudentWorkbookUpdate -Path $copy -Updates @(
                [pscustomobject]@{
                    RowNumber = 2
                    Password = ''
                    EntraObjectId = $initialObjectId
                    UPN = 'mmuster@monteaufkirchen.com'
                    Mail = 'mmuster@monteaufkirchen.com'
                }
            ) -SkipGitSafetyCheck | Out-Null

            $saved = @(Import-Excel -Path $copy -WorksheetName Tabelle1)
            $saved[0].Passwort | Should -Be 'TigerWiese56'
        }
    }
}
