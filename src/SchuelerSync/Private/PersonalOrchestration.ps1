function Export-PersonalWorkbook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '',
        Justification = 'Delegates the file write to Export-AccountWorkbook, forwarding WhatIf and Confirm.')]
    [CmdletBinding(SupportsShouldProcess)]
    param([object] $Inventory, [string] $Path)
    $common = @{ WhatIf = [bool]$WhatIfPreference }
    if ($PSBoundParameters.ContainsKey('Confirm')) { $common.Confirm = $PSBoundParameters.Confirm }
    Export-AccountWorkbook -Inventory $Inventory -Path $Path -Kind Personnel @common
}

function Invoke-PersonalSync {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '',
        Justification = 'Delegates mutations to guarded action functions, forwarding WhatIf and Confirm.')]
    [CmdletBinding(DefaultParameterSetName = 'Sync', SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Sync')]
        [Parameter(ParameterSetName = 'Export')]
        [ValidateScript({ [IO.Path]::GetExtension($_) -ieq '.xlsx' })]
        [string] $File = (Join-Path $script:SchuelerSyncRepositoryRoot 'Personal.xlsx'),
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
    $ErrorActionPreference = 'Stop'
    $common = @{ WhatIf = [bool]$WhatIfPreference }
    if ($PSBoundParameters.ContainsKey('Confirm')) { $common.Confirm = $PSBoundParameters.Confirm }
    if ($OutputFile) { $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile); Assert-ComparisonReportPath -Path $OutputFile }
    if ($PSCmdlet.ParameterSetName -in @('List','Object','Export','Manage')) {
        if ($ListToExcel) {
            $File = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($File)
            Assert-WorkbookExportPath -Path $File
        }
        $inventory = Get-PersonalInventory
        if ($ManageAccounts) {
            if ([Console]::IsInputRedirected) { throw 'Die Tastaturauswahl benötigt ein interaktives Terminal. Für die lesende Ausgabe -List verwenden.' }
            $actions = @(Invoke-DisabledAccountTui -Accounts $inventory.Accounts -TenantId $inventory.TenantId -PersonalConfig $inventory.Config @common)
            return [pscustomobject]@{ Mode = 'Manage'; Actions = $actions; HasErrors = @($actions | Where-Object Status -eq Failed).Count -gt 0 }
        }
        if ($List) { Write-PersonalInventoryReport -Inventory $inventory -OutputFile $OutputFile }
        if ($ListToExcel) {
            $export = Export-PersonalWorkbook -Inventory $inventory -Path $File @common
            foreach ($warning in $inventory.Warnings) { Write-Warning $warning }
            foreach ($account in $inventory.Accounts | Where-Object { $_.UserType -ne 'Guest' -and (-not $_.Job -or -not $_.MutationAllowed -or $_.UserType -ne 'Member') }) { Write-Warning "$($account.UPN): Exportzeile benötigt vor einem schreibenden Import eine Prüfung." }
            if ($null -ne $export) { Write-Information "Personalexport: $($export.Count) Konten, $($export.Path)" -InformationAction Continue }
        }
        if ($PassThru) { $inventory.PSObject.Properties.Remove('Config'); return $inventory }
        return
    }
    if ($ResumeCreation -and -not $EntraObjectId) { throw '-ResumeCreation benötigt eine explizite -EntraObjectId.' }
    $parameters = @{} + $PSBoundParameters
    if ($PSCmdlet.ParameterSetName -eq 'Sync' -and -not $parameters.ContainsKey('File')) { $parameters.File = $File }
    $parameters.Personnel = $true
    Invoke-PersonnelSyncCore @parameters
}
