Set-StrictMode -Version Latest

Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    Sort-Object Name |
    ForEach-Object { . $_.FullName }

function Invoke-SchuelerSync {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Sync', Mandatory)][string] $File,
        [Parameter(ParameterSetName = 'Sync')][switch] $Update,
        [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
        [switch] $ConfigureExchangeOnlineOnly,
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
        [Alias('UPN', 'UserPrincipalName')][string[]] $Mail
    )
    throw 'OrchestrationNotImplemented'
}

Export-ModuleMember -Function Invoke-SchuelerSync
