Set-StrictMode -Version Latest
$script:SchuelerSyncRepositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    Sort-Object Name |
    ForEach-Object { . $_.FullName }

function Invoke-SchuelerSync {
    [CmdletBinding(DefaultParameterSetName = 'Sync', SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Sync')]
        [ValidateScript({ [IO.Path]::GetExtension($_) -ieq '.xlsx' })]
        [string] $File = (Join-Path $script:SchuelerSyncRepositoryRoot 'Schueler.xlsx'),
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
    $ErrorActionPreference = 'Stop'
    $progressActivity = 'Schülerabgleich'
    $reportHeader = [Collections.Generic.List[string]]::new()
    if ($PSBoundParameters.ContainsKey('OutputFile')) {
        $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    }
    try {
    $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/SchuelerSync.psd1')
    $actions = [Collections.Generic.List[object]]::new()
    $comparisonPrinted = $false
    $secrets = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $comparison = [pscustomobject]@{ NewStudents = @(); Departures = @(); ChangedStudents = @(); ExistingStudents = @(); Warnings = @(); Errors = @() }
    $common = @{ WhatIf = [bool]$WhatIfPreference }
    if ($PSBoundParameters.ContainsKey('Confirm')) { $common.Confirm = $PSBoundParameters.Confirm }
    if ($PSCmdlet.ParameterSetName -eq 'ExchangeOnly') {
        Write-Progress -Id 1 -Activity $progressActivity -Status 'Verbinde mit Exchange Online ...' -PercentComplete 10
        $targets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($address in $Mail) {
            $normalized = ([string]$address).Trim().ToLowerInvariant()
            if ($normalized -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { throw "Ungültige Postfachidentität '$address'." }
            [void]$targets.Add($normalized)
        }
        try {
            $null = Connect-SchuelerExchangeOnline
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Prüfe und konfiguriere Exchange-Postfächer ...' -PercentComplete 35
            $batch = Wait-StudentMailbox -UserPrincipalName @($targets) -Config $config.Exchange -Configure `
                -MaxRetries $config.Exchange.MaxMailboxRetries -RetryDelaySeconds $config.Exchange.RetryDelaySeconds @common
            foreach ($action in @(ConvertTo-ExchangeActionResult -Batch $batch -WhatIfMode:$WhatIfPreference)) { $actions.Add($action) }
        } catch {
            $actions.Add((New-StudentActionResult -Phase Exchange -Status Failed -Message $_.Exception.Message))
        }
    } else {
        $selection = Resolve-UpdateSelection -Update:$Update -CreateNewUsers:$CreateNewUsers -UpdateUsers:$UpdateUsers -DisableUsers:$DisableUsers -RevokeSessions:$RevokeSessions
        try {
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Lese Excel-Datei ...' -PercentComplete 5
            $File = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($File)
            $workbook = Read-StudentWorkbook -Path $File
            foreach ($student in $workbook.Students) {
                $password = [string](Get-ComparisonPropertyValue $student Password)
                if ($password) { [void]$secrets.Add($password) }
            }
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Verbinde mit Microsoft Entra ID ...' -PercentComplete 10
            $context = Connect-SchuelerGraph
            $tenantId = [string](Get-ComparisonPropertyValue $context TenantId)
            if ([string]::IsNullOrWhiteSpace($tenantId)) { throw 'Graph-Kontext enthält keine Tenant-ID.' }
            $tenantLine = "Entra-Tenant: $tenantId | Konto: $([string](Get-ComparisonPropertyValue $context Account))"
            $reportHeader.Add($tenantLine)
            Write-Information $tenantLine -InformationAction Continue
            if ($config.ExpectedTenantId -and $tenantId -ine [string]$config.ExpectedTenantId) { throw 'Der verbundene Graph-Tenant stimmt nicht mit ExpectedTenantId überein.' }
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Lade Entra-Benutzer, Gruppen und Schülerrolle ...' -PercentComplete 15
            $snapshot = Get-EntraSnapshot -Config $config
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Verbinde mit Exchange Online ...' -PercentComplete 30
            $null = Connect-SchuelerExchangeOnline
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Lade Exchange-Empfängeradressen ...' -PercentComplete 35
            $recipients = Get-ExchangeRecipientAddress
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Vergleiche Schülerdaten ...' -PercentComplete 40
            $comparison = Compare-StudentDirectory -Students @($workbook.Students) -Snapshot $snapshot -Config $config -ExchangeAddressOwners $recipients.AddressOwners
            # Read current identities, even when a rename is planned, exactly once during preflight.
            $mailboxEntries = @($comparison.ChangedStudents) + @($comparison.ExistingStudents)
            $mailboxIndex = 0
            foreach ($entry in $mailboxEntries) {
                $mailboxIndex++
                $mailboxPercent = if ($mailboxEntries.Count -eq 0) { 70 } else { 55 + [int](20 * $mailboxIndex / $mailboxEntries.Count) }
                Write-Progress -Id 1 -Activity $progressActivity -Status "Prüfe Exchange-Postfach $mailboxIndex von $($mailboxEntries.Count): $($entry.User.UserPrincipalName)" -PercentComplete $mailboxPercent
                try {
                    $state = Get-StudentMailboxState -UserPrincipalName $entry.User.UserPrincipalName -Config $config.Exchange
                    $comparison = Add-ExchangeComparison -Comparison $comparison -UserId $entry.User.Id -MailboxState $state
                } catch {
                    $comparison.Errors += New-ComparisonIssue -Severity Error -Area Exchange -Field Preflight -Message $_.Exception.Message -Student $entry.Student -User $entry.User
                }
            }
            if ($WhatIfPreference -and $selection.CreateNewUsers) {
                foreach ($entry in $comparison.NewStudents) {
                    try {
                        $null = Get-StudentMailboxState -UserPrincipalName $entry.DesiredState.UserPrincipalName -Config $config.Exchange
                    } catch {
                        $comparison.Errors += New-ComparisonIssue -Severity Error -Area Exchange -Field Preflight -Message $_.Exception.Message -Student $entry.Student
                    }
                }
            }
            if ($Update -and -not $WhatIfPreference) {
                $hasRenames = $selection.UpdateUsers -and @($comparison.ChangedStudents | Where-Object { Test-StudentUpnRename -Entry $_ }).Count -gt 0
                if (($selection.CreateNewUsers -and @($comparison.NewStudents).Count -gt 0) -or $hasRenames) {
                    Assert-WorkbookSafeForPasswordWrite -Path $File
                }
                Assert-StudentWorkbookVersion -Path $File -ExpectedSourceHash $workbook.SourceHash
            }
        } catch {
            $comparison.Errors += New-ComparisonIssue -Severity Error -Area Preflight -Field Inventory -Message $_.Exception.Message
        }
        $safeComparison = ConvertTo-SafeStudentComparison -Comparison $comparison -File $File -Secrets @($secrets)
        if ($Update -and @($comparison.Errors).Count -eq 0) {
            Write-StudentComparisonReport -Comparison $safeComparison
            $comparisonPrinted = $true
        }
        if ($Update -and @($comparison.Errors).Count -eq 0) {
            if ($WhatIfPreference) {
                foreach ($definition in @(
                    @{ Enabled = $selection.CreateNewUsers; Entries = $comparison.NewStudents; Phase = 'Create' }
                    @{ Enabled = $selection.UpdateUsers; Entries = $comparison.ChangedStudents; Phase = 'Update' }
                    @{ Enabled = $selection.DisableUsers; Entries = $comparison.Departures; Phase = 'Disable' }
                    @{ Enabled = $selection.RevokeSessions; Entries = $comparison.Departures; Phase = 'RevokeSessions' }
                )) {
                    if (-not $definition.Enabled) { continue }
                    foreach ($entry in $definition.Entries) {
                        $upn = if ($null -ne $entry.DesiredState) { $entry.DesiredState.UserPrincipalName } else { $entry.User.UserPrincipalName }
                        [void]$PSCmdlet.ShouldProcess($upn, $definition.Phase)
                        $actions.Add((New-StudentActionResult -UserId ([string](Get-ComparisonPropertyValue $entry.User Id)) -UserPrincipalName $upn -Phase $definition.Phase -Status WhatIf))
                    }
                }
                $exchangeEntries = @()
                if ($selection.ConfigureAllActiveExchange) { $exchangeEntries = @($comparison.NewStudents) + @($comparison.ChangedStudents) + @($comparison.ExistingStudents) }
                else {
                    if ($selection.CreateNewUsers) { $exchangeEntries += @($comparison.NewStudents) }
                    if ($selection.UpdateUsers) { $exchangeEntries += @($comparison.ChangedStudents) }
                }
                foreach ($entry in $exchangeEntries) {
                    $actions.Add((New-StudentActionResult -UserId ([string](Get-ComparisonPropertyValue $entry.User Id)) -UserPrincipalName $entry.DesiredState.UserPrincipalName -Phase Exchange -Status WhatIf))
                }
            } else {
                if ($selection.CreateNewUsers) {
                    foreach ($action in @(Invoke-StudentCreateBatch -Entries @($comparison.NewStudents) -Snapshot $snapshot -Config $config -File $File -WorkbookState $workbook -UsedPasswords $secrets @common)) { $actions.Add($action) }
                }
                if ($selection.UpdateUsers) {
                    foreach ($action in @(Invoke-StudentUpdate -Entries @($comparison.ChangedStudents) -Snapshot $snapshot -Config $config -File $File -WorkbookState $workbook -Secrets @($secrets) @common)) { $actions.Add($action) }
                }
                if ($selection.DisableUsers -or $selection.RevokeSessions) {
                    foreach ($action in @(Invoke-StudentDeparture -Entries @($comparison.Departures) -DisableUsers:$selection.DisableUsers -RevokeSessions:$selection.RevokeSessions -File $File -Secrets @($secrets) @common)) { $actions.Add($action) }
                }
                $targetsById = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($action in @($actions | Where-Object { $_.Status -eq 'Succeeded' -and $_.Phase -in @('Create', 'Update') })) {
                    $targetsById[$action.UserId] = $action.UserPrincipalName
                }
                if ($selection.ConfigureAllActiveExchange) {
                    foreach ($entry in @($comparison.ChangedStudents) + @($comparison.ExistingStudents)) {
                        $userId = [string]$entry.User.Id
                        try {
                            $currentIdentity = Get-EntraStudentCurrentIdentity -UserId $userId
                            $targetsById[$userId] = [string]$currentIdentity.UserPrincipalName
                        } catch {
                            $targetsById.Remove($userId)
                            $actions.Add((New-StudentActionResult -UserId $userId -Phase ExchangeIdentity -Status Failed -Message $_.Exception.Message -Secrets @($secrets) -RecoveryCommand (Get-StudentRecoveryCommand -File $File)))
                        }
                    }
                }
                if ($targetsById.Count -gt 0) {
                    try {
                        $batch = Wait-StudentMailbox -UserPrincipalName @($targetsById.Values) -Config $config.Exchange -Configure `
                            -MaxRetries $config.Exchange.MaxMailboxRetries -RetryDelaySeconds $config.Exchange.RetryDelaySeconds @common
                        foreach ($action in @(ConvertTo-ExchangeActionResult -Batch $batch)) { $actions.Add($action) }
                    } catch {
                        $actions.Add((New-StudentActionResult -Phase Exchange -Status Failed -Message $_.Exception.Message -Secrets @($secrets)))
                    }
                }
            }
        }
    }
    $safeComparison = ConvertTo-SafeStudentComparison -Comparison $comparison -File $File -Secrets @($secrets)
    # Only allowlisted projections cross the public output boundary, never workbook rows or password profiles.
    foreach ($action in $actions) { $action.Message = Protect-StudentMessage -Message $action.Message -Secrets @($secrets) }
    Write-Progress -Id 1 -Activity $progressActivity -Status 'Erzeuge Ergebnisbericht ...' -PercentComplete 95
    Write-StudentComparisonReport -Comparison $safeComparison -Actions @($actions) -ActionsOnly:$comparisonPrinted -HeaderLines @($reportHeader) -OutputFile $OutputFile
    [pscustomobject]@{
        Mode = if ($ConfigureExchangeOnlineOnly) { 'ExchangeOnly' } elseif ($Update) { 'Update' } else { 'Compare' }
        Comparison = $safeComparison
        Actions = [object[]]@($actions)
        HasErrors = (@($comparison.Errors).Count -gt 0 -or @($actions | Where-Object Status -eq Failed).Count -gt 0)
    }
    } finally {
        Write-Progress -Id 1 -Activity $progressActivity -Completed
    }
}

Export-ModuleMember -Function Invoke-SchuelerSync
