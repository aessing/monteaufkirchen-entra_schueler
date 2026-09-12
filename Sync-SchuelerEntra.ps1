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
    [Parameter(ParameterSetName = 'Add', Mandatory)][switch] $Add,
    [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Vorname,
    [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Nachname,
    [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Klasse,
    [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Klassenlehrer,
    [Parameter(ParameterSetName = 'Add')][ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')][string] $EntraObjectId,
    [Parameter(ParameterSetName = 'Remove', Mandatory)][switch] $Remove,
    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [switch] $ConfigureExchangeOnlineOnly,
    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [Parameter(ParameterSetName = 'Remove', Mandatory)]
    [Alias('UPN', 'UserPrincipalName')][ValidateNotNullOrEmpty()][string[]] $Mail,
    [Parameter(ParameterSetName = 'Sync')]
    [Parameter(ParameterSetName = 'ExchangeOnly')]
    [Parameter(ParameterSetName = 'Add')]
    [Parameter(ParameterSetName = 'Remove')]
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
