[CmdletBinding()]
param(
    [string] $PesterManifest = 'Pester',
    [string] $AnalyzerManifest = 'PSScriptAnalyzer'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Import-Module $PesterManifest -ErrorAction Stop
Import-Module $AnalyzerManifest -ErrorAction Stop
$files = @(Get-ChildItem (Join-Path $repo 'src') -Recurse -File | Where-Object Extension -in @('.ps1','.psm1','.psd1')) +
    @(Get-ChildItem $repo -Filter '*.ps1' -File) + @(Get-ChildItem (Join-Path $repo 'config') -Filter '*.psd1' -File)
foreach ($file in $files) {
    $tokens = $null; $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw "Parserfehler in $($file.Name): $($parseErrors.Message -join ', ')" }
}
$diagnostics = @($files | ForEach-Object {
    Invoke-ScriptAnalyzer -Path $_.FullName -Settings (Join-Path $repo 'PSScriptAnalyzerSettings.psd1')
})
# Colored section headings deliberately use the host information stream.
$unexpected = @($diagnostics | Where-Object RuleName -ne PSAvoidUsingWriteHost)
$diagnostics | Format-Table Severity, RuleName, ScriptName, Line -AutoSize
if ($unexpected.Count -gt 0) { throw "$($unexpected.Count) ungeklärte Analyzer-Befunde." }
$config = New-PesterConfiguration
$config.Run.Path = $PSScriptRoot
$config.Run.PassThru = $true
$config.Output.Verbosity = 'Normal'
$result = Invoke-Pester -Configuration $config
if ($result.FailedCount -gt 0 -or $result.FailedContainersCount -gt 0 -or $result.TotalCount -eq 0) { throw 'Lokale Tests fehlgeschlagen.' }
Write-Output "Prüfung erfolgreich: $($result.TotalCount) Tests, $($result.PassedCount) bestanden, $($result.SkippedCount) übersprungen."
