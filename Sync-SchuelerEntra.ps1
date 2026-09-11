[CmdletBinding(DefaultParameterSetName = 'Sync', SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(ParameterSetName = 'Sync')]
    [ValidateScript({ [IO.Path]::GetExtension($_) -ieq '.xlsx' })]
    [string] $File = (Join-Path $PSScriptRoot 'Schueler.xlsx'),
    [Parameter(ParameterSetName = 'Sync')][switch] $Update,
    [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,
    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [switch] $ConfigureExchangeOnlineOnly,
    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [Alias('UPN', 'UserPrincipalName')][ValidateNotNullOrEmpty()][string[]] $Mail,
    [Parameter(ParameterSetName = 'Sync')]
    [Parameter(ParameterSetName = 'ExchangeOnly')]
    [ValidateNotNullOrEmpty()][string] $OutputFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$manifest = Join-Path $PSScriptRoot 'src/SchuelerSync/SchuelerSync.psd1'
Import-Module $manifest -Force

$forwardParameters = @{} + $PSBoundParameters
if ($PSCmdlet.ParameterSetName -eq 'Sync' -and -not $PSBoundParameters.ContainsKey('File')) {
    $forwardParameters['File'] = $File
}

$result = Invoke-SchuelerSync @forwardParameters
$result
if ($result.HasErrors) { exit 1 }
