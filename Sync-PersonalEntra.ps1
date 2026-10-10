[CmdletBinding(DefaultParameterSetName = 'Sync', SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Sync')]
        [Parameter(ParameterSetName = 'Export')]
        [ValidateScript({ [IO.Path]::GetExtension($_) -ieq '.xlsx' })]
        [string] $File = (Join-Path $PSScriptRoot 'Personal.xlsx'),
        [Parameter(ParameterSetName = 'Sync')][switch] $Update,
        [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,
        [Parameter(ParameterSetName = 'Add', Mandatory)][switch] $Add,
        [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Vorname,
        [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Nachname,
        [Parameter(ParameterSetName = 'Add', Mandatory)][ValidateNotNullOrEmpty()][string] $Job,
        [Parameter(ParameterSetName = 'Add')][ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')][string] $EntraObjectId,
        [Parameter(ParameterSetName = 'Add')][switch] $ResumeCreation,
        [Parameter(ParameterSetName = 'Remove', Mandatory)][switch] $Remove,
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)][switch] $ConfigureExchangeOnlineOnly,
        [Parameter(ParameterSetName = 'Remove', Mandatory)]
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
        [Alias('UPN','UserPrincipalName')][ValidateNotNullOrEmpty()][string[]] $Mail,
        [Parameter(ParameterSetName = 'List', Mandatory)][switch] $List,
        [Parameter(ParameterSetName = 'Export', Mandatory)][switch] $ListToExcel,
        [Parameter(ParameterSetName = 'Manage', Mandatory)][switch] $ManageAccounts,
        [Parameter(ParameterSetName = 'List')]
        [Parameter(ParameterSetName = 'Export')]
        [Parameter(ParameterSetName = 'Object', Mandatory)][switch] $PassThru,
        [Parameter(ParameterSetName = 'List')]
        [Parameter(ParameterSetName = 'Sync')]
        [Parameter(ParameterSetName = 'Add')]
        [Parameter(ParameterSetName = 'Remove')]
        [Parameter(ParameterSetName = 'ExchangeOnly')][ValidateNotNullOrEmpty()][string] $OutputFile
    )

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force
$parameters = @{} + $PSBoundParameters
if ($PSCmdlet.ParameterSetName -in @('Sync','Export') -and -not $parameters.ContainsKey('File')) { $parameters.File = $File }
$result = Invoke-PersonalSync @parameters
$result
if ($null -ne $result -and $null -ne $result.PSObject.Properties['HasErrors'] -and $result.HasErrors) { exit 1 }
