BeforeAll {
    $testFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1')
}

Describe 'Pester discovery isolation' {
    It 'discovers the suite in a fresh process without a pre-imported application module' {
        $executable = Join-Path $PSHOME $(if ($IsWindows) { 'pwsh.exe' } else { 'pwsh' })
        $pesterManifest = Join-Path (Get-Module Pester).ModuleBase 'Pester.psd1'
        $output = & $executable -NoLogo -NoProfile -NonInteractive -File (Join-Path $PSScriptRoot 'Test-PesterDiscovery.ps1') -TestPath $PSScriptRoot -PesterManifest $pesterManifest 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
    }

    It 'imports before discovery-time module scope and never replaces the bound module during the run' {
        foreach ($file in $testFiles) {
            $tokens = $null
            $errors = $null
            $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            $errors | Should -BeNullOrEmpty -Because $file.Name
            $commands = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true))
            $imports = @($commands | Where-Object { $_.GetCommandName() -eq 'Import-Module' })
            foreach ($import in $imports) {
                @($import.CommandElements | Where-Object { $_ -is [Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'Force' }).Count | Should -Be 0 -Because $file.Name
            }
            foreach ($scope in @($commands | Where-Object { $_.GetCommandName() -eq 'InModuleScope' })) {
                $discovery = @($commands | Where-Object {
                    $_.GetCommandName() -eq 'BeforeDiscovery' -and $_.Extent.EndOffset -lt $scope.Extent.StartOffset -and
                    $null -ne $_.Find({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Import-Module' }, $true)
                })
                $discovery.Count | Should -Be 1 -Because $file.Name
                # Runtime fixtures belong in BeforeAll/BeforeEach inside module scope.
                # Discovery-time parameter injection does not capture future test state.
                @($scope.CommandElements | Where-Object { $_ -is [Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'Parameters' }).Count | Should -Be 0 -Because $file.Name
            }
        }
    }
}
