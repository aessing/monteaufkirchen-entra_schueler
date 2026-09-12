[CmdletBinding()]
param(
    [string] $TestPath = $PSScriptRoot,
    [string] $PesterManifest = 'Pester'
)

$ErrorActionPreference = 'Stop'
if (Get-Module SchuelerSync) {
    throw 'Die Discovery-Prüfung muss in einer frischen PowerShell-Sitzung ohne SchuelerSync starten.'
}
Import-Module $PesterManifest -ErrorAction Stop
$configuration = New-PesterConfiguration
$configuration.Run.Path = $TestPath
$configuration.Run.SkipRun = $true
$configuration.Run.PassThru = $true
$configuration.Output.Verbosity = 'Minimal'
$result = Invoke-Pester -Configuration $configuration
if ($result.Result -eq 'Failed' -or $result.FailedCount -gt 0 -or $result.FailedContainersCount -gt 0 -or $result.TotalCount -eq 0) {
    throw 'Pester konnte die Tests in der frischen Sitzung nicht vollständig entdecken.'
}
Write-Output "Discovery erfolgreich: $($result.TotalCount) Tests, keine Testausführung."
