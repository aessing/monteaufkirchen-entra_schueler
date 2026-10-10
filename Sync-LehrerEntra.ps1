[CmdletBinding(DefaultParameterSetName = 'Sync', SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(ParameterSetName = 'Sync')][Parameter(ParameterSetName = 'Export')][ValidateScript({ [IO.Path]::GetExtension($_) -ieq '.xlsx' })]
    [string] $File = (Join-Path $PSScriptRoot 'Lehrer.xlsx'),
    [Parameter(ParameterSetName = 'Export', Mandatory)][switch] $ListToExcel,
    [Parameter(ParameterSetName = 'Manage', Mandatory)][switch] $ManageAccounts,
    [Parameter(ParameterSetName = 'Sync')][switch] $Update,
    [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
    [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,
    [Parameter(ParameterSetName = 'Add', Mandatory)][switch] $Add,
    [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Vorname,
    [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Nachname,
    [Parameter(ParameterSetName = 'Add', Mandatory)]
    [ValidatePattern('(?i)^\s*(?:L|CO-L|PA|OGTS|JAS)(?:\s*,\s*(?:L|CO-L|PA|OGTS|JAS))*\s*$')][string] $Job,
    [Parameter(ParameterSetName = 'Add')][ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')][string] $EntraObjectId,
    [Parameter(ParameterSetName = 'Remove', Mandatory)][switch] $Remove,
    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)][switch] $ConfigureExchangeOnlineOnly,
    [Parameter(ParameterSetName = 'Remove', Mandatory)]
    [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
    [Alias('UPN', 'UserPrincipalName')][ValidateNotNullOrEmpty()][string[]] $Mail,
    [Parameter(ParameterSetName = 'Sync')]
    [Parameter(ParameterSetName = 'Add')]
    [Parameter(ParameterSetName = 'Remove')]
    [Parameter(ParameterSetName = 'ExchangeOnly')]
    [ValidateNotNullOrEmpty()][string] $OutputFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force
$forwardParameters = @{} + $PSBoundParameters
if ($PSCmdlet.ParameterSetName -in @('Sync','Export') -and -not $PSBoundParameters.ContainsKey('File')) {
    $forwardParameters['File'] = $File
}
$result = Invoke-LehrerSync @forwardParameters
$result
if ($result.HasErrors) { exit 1 }
