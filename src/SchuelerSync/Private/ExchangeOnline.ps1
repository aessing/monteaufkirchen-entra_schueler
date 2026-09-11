function Test-ExchangeOnlineConnectionActive {
    param([AllowNull()][object] $Connection)

    if ($null -eq $Connection) { return $false }
    $name = [string](Get-ComparisonPropertyValue -InputObject $Connection -Name Name)
    $moduleName = [string](Get-ComparisonPropertyValue -InputObject $Connection -Name ModuleName)
    $connectionUri = [string](Get-ComparisonPropertyValue -InputObject $Connection -Name ConnectionUri)
    $state = [string](Get-ComparisonPropertyValue -InputObject $Connection -Name State)
    if ([string]::IsNullOrWhiteSpace($state)) {
        $state = [string](Get-ComparisonPropertyValue -InputObject $Connection -Name ConnectionState)
    }

    $isEopValue = Get-ComparisonPropertyValue -InputObject $Connection -Name IsEopSession
    $isEopSession = $false
    if ($isEopValue -is [bool]) {
        $isEopSession = [bool]$isEopValue
    } elseif (-not [string]::IsNullOrWhiteSpace([string]$isEopValue)) {
        [void][bool]::TryParse([string]$isEopValue, [ref]$isEopSession)
    }
    $isComplianceIdentity = $name -match '(?i)SecurityCompliance|Compliance|Protection' -or
        $connectionUri -match '(?i)compliance\.protection\.outlook\.com|ps\.protection\.outlook\.com'
    if ($isEopSession -or $isComplianceIdentity) { return $false }

    $isExchangeOnline = $name -match '(?i)ExchangeOnline' -or
        $moduleName -match '(?i)ExchangeOnline' -or
        $connectionUri -match '(?i)outlook\.office365\.com|exchange\.microsoft\.com'
    return $isExchangeOnline -and [string]::Equals($state, 'Connected', [StringComparison]::OrdinalIgnoreCase)
}

function Connect-SchuelerExchangeOnline {
    [CmdletBinding()]
    param()

    $connections = @(Get-ConnectionInformation -ErrorAction Stop)
    $active = @($connections | Where-Object { Test-ExchangeOnlineConnectionActive -Connection $_ } | Select-Object -First 1)
    if ($active.Count -eq 1) { return $active[0] }

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    $connections = @(Get-ConnectionInformation -ErrorAction Stop)
    $active = @($connections | Where-Object { Test-ExchangeOnlineConnectionActive -Connection $_ } | Select-Object -First 1)
    if ($active.Count -ne 1) {
        throw 'Nach Connect-ExchangeOnline wurde keine aktive Exchange-Online-Verbindung gefunden.'
    }
    return $active[0]
}

function Add-ExchangeRecipientAddress {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.HashSet[string]] $ReservedAddresses,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.IDictionary] $AddressOwners,
        [AllowNull()][object] $Address,
        [AllowNull()][object] $OwnerId,
        [switch] $HasSmtpPrefix
    )

    $normalized = ([string]$Address).Trim()
    if ($HasSmtpPrefix) {
        if ($normalized -notmatch '(?i)^smtp:') { return }
        $normalized = ($normalized -replace '(?i)^smtp:', '').Trim()
    }
    if ([string]::IsNullOrWhiteSpace($normalized)) { return }
    $normalized = $normalized.ToLowerInvariant()

    [void]$ReservedAddresses.Add($normalized)
    if (-not $AddressOwners.Contains($normalized)) { $AddressOwners[$normalized] = @() }
    $owner = ([string]$OwnerId).Trim()
    if ([string]::IsNullOrWhiteSpace($owner)) { return }
    if (@($AddressOwners[$normalized] | Where-Object {
                [string]::Equals([string]$_, $owner, [StringComparison]::OrdinalIgnoreCase)
            }).Count -eq 0) {
        $AddressOwners[$normalized] = @($AddressOwners[$normalized]) + $owner
    }
}

function Get-ExchangeRecipientAddress {
    [CmdletBinding()]
    param()

    $reservedAddresses = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $addressOwners = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($recipient in @(Get-Recipient -ResultSize Unlimited -ErrorAction Stop)) {
        $ownerId = Get-ComparisonPropertyValue -InputObject $recipient -Name ExternalDirectoryObjectId
        Add-ExchangeRecipientAddress -ReservedAddresses $reservedAddresses -AddressOwners $addressOwners `
            -Address (Get-ComparisonPropertyValue -InputObject $recipient -Name PrimarySmtpAddress) -OwnerId $ownerId
        foreach ($emailAddress in @((Get-ComparisonPropertyValue -InputObject $recipient -Name EmailAddresses))) {
            Add-ExchangeRecipientAddress -ReservedAddresses $reservedAddresses -AddressOwners $addressOwners `
                -Address $emailAddress -OwnerId $ownerId -HasSmtpPrefix
        }
    }

    return [pscustomobject]@{
        ReservedAddresses = $reservedAddresses
        AddressOwners = $addressOwners
    }
}

function Test-ExchangeMailboxNotFoundError {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord] $ErrorRecord)

    $exceptionName = $ErrorRecord.Exception.GetType().Name
    $fullyQualifiedErrorId = [string]$ErrorRecord.FullyQualifiedErrorId
    $category = [string]$ErrorRecord.CategoryInfo.Category
    $reason = [string]$ErrorRecord.CategoryInfo.Reason
    $message = [string]$ErrorRecord.Exception.Message
    $isNotFoundType = $exceptionName -in @('ManagementObjectNotFoundException', 'ItemNotFoundException') -or
        $reason -eq 'ManagementObjectNotFoundException' -or
        $fullyQualifiedErrorId -match '(?i)ManagementObjectNotFoundException|Ex6F9304'
    $isNotFoundCategory = [string]::Equals($category, 'ObjectNotFound', [StringComparison]::OrdinalIgnoreCase)
    $hasDocumentedMessage = $message -match "(?i)(couldn't|could not|cannot|wasn't|was not) be found|does not exist"
    return ($isNotFoundType -or $isNotFoundCategory) -and $hasDocumentedMessage
}

function Get-StudentMailboxState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $UserPrincipalName,
        [System.Collections.IDictionary] $Config
    )

    $upn = $UserPrincipalName.Trim()
    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
    } catch {
        if (Test-ExchangeMailboxNotFoundError -ErrorRecord $_) {
            return [pscustomobject]@{
                UserPrincipalName = $upn
                Exists = $false
                Status = 'MailboxNotReady'
                Mailbox = $null
                CasMailbox = $null
                Differences = @()
            }
        }
        throw
    }

    $casMailbox = Get-CASMailbox -Identity $upn -ErrorAction Stop
    $differences = if ($null -eq $Config) {
        @()
    } else {
        @(Compare-StudentMailboxState -Mailbox $mailbox -CasMailbox $casMailbox -Config $Config)
    }
    return [pscustomobject]@{
        UserPrincipalName = $upn
        Exists = $true
        Status = 'Ready'
        Mailbox = $mailbox
        CasMailbox = $casMailbox
        Differences = [object[]]@($differences)
    }
}

function Get-RequiredExchangeAuditAction {
    param([Parameter(Mandatory)][object] $Mailbox)

    $auditAdmin = @(
        'Copy', 'Create', 'FolderBind', 'HardDelete', 'Move',
        'MoveToDeletedItems', 'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
        'UpdateFolderPermissions', 'UpdateInboxRules', 'UpdateCalendarDelegation'
    )
    $hasCommunicationsCompliance = @((Get-ComparisonPropertyValue -InputObject $Mailbox -Name PersistedCapabilities) | Where-Object {
            [string]::Equals([string]$_, 'CommunicationsCompliance', [StringComparison]::OrdinalIgnoreCase)
        }).Count -gt 0
    if ($hasCommunicationsCompliance) { $auditAdmin += 'MailItemsAccessed' }

    return [pscustomobject]@{
        AuditDelegate = @(
            'Create', 'FolderBind', 'HardDelete', 'Move', 'MoveToDeletedItems',
            'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
            'UpdateFolderPermissions', 'UpdateInboxRules'
        )
        AuditOwner = @(
            'Create', 'HardDelete', 'Move', 'MailboxLogin', 'MoveToDeletedItems',
            'SoftDelete', 'Update', 'UpdateFolderPermissions', 'UpdateInboxRules',
            'UpdateCalendarDelegation'
        )
        AuditAdmin = $auditAdmin
    }
}

function Get-MissingExchangeValue {
    param(
        [AllowNull()][object[]] $Current,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Required
    )

    $currentSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($value in @($Current)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { [void]$currentSet.Add([string]$value) }
    }
    return [object[]]@($Required | Where-Object { -not $currentSet.Contains([string]$_) })
}

function ConvertTo-ExchangePolicyName {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value) { return $null }
    $name = Get-ComparisonPropertyValue -InputObject $Value -Name Name
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { return [string]$name }
    return [string]$Value
}

function Get-ExchangeDurationDay {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value) { return $null }
    $totalDays = Get-ComparisonPropertyValue -InputObject $Value -Name TotalDays
    if ($null -ne $totalDays) { return [double]$totalDays }
    $days = Get-ComparisonPropertyValue -InputObject $Value -Name Days
    if ($null -ne $days) { return [double]$days }
    try { return [TimeSpan]::Parse([string]$Value).TotalDays } catch { return $null }
}

function Compare-StudentMailboxState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Mailbox,
        [Parameter(Mandatory)][object] $CasMailbox,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config
    )

    $differences = [Collections.Generic.List[object]]::new()
    $scalarFields = @(
        [pscustomobject]@{ Field = 'AddressBookPolicy'; Current = ConvertTo-ExchangePolicyName (Get-ComparisonPropertyValue -InputObject $Mailbox -Name AddressBookPolicy); Desired = [string]$Config.AddressBookPolicy }
        [pscustomobject]@{ Field = 'CustomAttribute1'; Current = Get-ComparisonPropertyValue -InputObject $Mailbox -Name CustomAttribute1; Desired = [string]$Config.CustomAttribute1 }
        [pscustomobject]@{ Field = 'AuditEnabled'; Current = Get-ComparisonPropertyValue -InputObject $Mailbox -Name AuditEnabled; Desired = $true }
        [pscustomobject]@{ Field = 'AuditLogAgeLimit'; Current = Get-ExchangeDurationDay (Get-ComparisonPropertyValue -InputObject $Mailbox -Name AuditLogAgeLimit); Desired = [double]$Config.AuditLogAgeLimitDays }
        [pscustomobject]@{ Field = 'RetainDeletedItemsFor'; Current = Get-ExchangeDurationDay (Get-ComparisonPropertyValue -InputObject $Mailbox -Name RetainDeletedItemsFor); Desired = [double]$Config.RetainDeletedItemsForDays }
        [pscustomobject]@{ Field = 'RoleAssignmentPolicy'; Current = ConvertTo-ExchangePolicyName (Get-ComparisonPropertyValue -InputObject $Mailbox -Name RoleAssignmentPolicy); Desired = [string]$Config.RoleAssignmentPolicy }
        [pscustomobject]@{ Field = 'SharingPolicy'; Current = ConvertTo-ExchangePolicyName (Get-ComparisonPropertyValue -InputObject $Mailbox -Name SharingPolicy); Desired = [string]$Config.SharingPolicy }
        [pscustomobject]@{ Field = 'RetentionPolicy'; Current = ConvertTo-ExchangePolicyName (Get-ComparisonPropertyValue -InputObject $Mailbox -Name RetentionPolicy); Desired = [string]$Config.RetentionPolicy }
    )
    foreach ($field in $scalarFields) {
        $difference = New-StateDifference -Area Exchange -Field $field.Field -Current $field.Current -Desired $field.Desired
        if ($null -ne $difference) { $differences.Add($difference) }
    }

    $requiredAudit = Get-RequiredExchangeAuditAction -Mailbox $Mailbox
    foreach ($field in @('AuditDelegate', 'AuditOwner', 'AuditAdmin')) {
        $current = @((Get-ComparisonPropertyValue -InputObject $Mailbox -Name $field))
        $required = @((Get-ComparisonPropertyValue -InputObject $requiredAudit -Name $field))
        $missing = @(Get-MissingExchangeValue -Current $current -Required $required)
        if ($missing.Count -gt 0) {
            $differences.Add([pscustomobject]@{
                    Area = 'Exchange'
                    Field = $field
                    Current = [object[]]@($current)
                    Desired = [object[]]@($required)
                    Action = 'Add'
                })
        }
    }

    $casFields = @(
        [pscustomobject]@{ Field = 'ActiveSyncEnabled'; Desired = $false }
        [pscustomobject]@{ Field = 'ImapEnabled'; Desired = $false }
        [pscustomobject]@{ Field = 'MAPIEnabled'; Desired = $true }
        [pscustomobject]@{ Field = 'OWAEnabled'; Desired = $true }
        [pscustomobject]@{ Field = 'OWAforDevicesEnabled'; Desired = $false }
        [pscustomobject]@{ Field = 'OwaMailboxPolicy'; Desired = [string]$Config.OwaMailboxPolicy }
        [pscustomobject]@{ Field = 'PopEnabled'; Desired = $false }
        [pscustomobject]@{ Field = 'SmtpClientAuthenticationDisabled'; Desired = $true }
    )
    foreach ($field in $casFields) {
        $current = Get-ComparisonPropertyValue -InputObject $CasMailbox -Name $field.Field
        if ($field.Field -eq 'OwaMailboxPolicy') { $current = ConvertTo-ExchangePolicyName $current }
        $difference = New-StateDifference -Area Exchange -Field $field.Field -Current $current -Desired $field.Desired
        if ($null -ne $difference) { $differences.Add($difference) }
    }

    return [object[]]@($differences)
}

function Get-ExchangeMailboxWriteParameter {
    param(
        [Parameter(Mandatory)][object] $Mailbox,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Differences,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config
    )

    $parameters = @{}
    $desiredValues = @{
        AddressBookPolicy = [string]$Config.AddressBookPolicy
        CustomAttribute1 = [string]$Config.CustomAttribute1
        AuditEnabled = $true
        AuditLogAgeLimit = [TimeSpan]::FromDays([double]$Config.AuditLogAgeLimitDays)
        RetainDeletedItemsFor = [TimeSpan]::FromDays([double]$Config.RetainDeletedItemsForDays)
        RoleAssignmentPolicy = [string]$Config.RoleAssignmentPolicy
        SharingPolicy = [string]$Config.SharingPolicy
        RetentionPolicy = [string]$Config.RetentionPolicy
    }
    foreach ($difference in @($Differences | Where-Object { $desiredValues.ContainsKey([string]$_.Field) })) {
        $parameters[[string]$difference.Field] = $desiredValues[[string]$difference.Field]
    }

    $requiredAudit = Get-RequiredExchangeAuditAction -Mailbox $Mailbox
    foreach ($field in @('AuditDelegate', 'AuditOwner', 'AuditAdmin')) {
        if (@($Differences | Where-Object Field -eq $field).Count -eq 0) { continue }
        $missing = @(Get-MissingExchangeValue `
                -Current @((Get-ComparisonPropertyValue -InputObject $Mailbox -Name $field)) `
                -Required @((Get-ComparisonPropertyValue -InputObject $requiredAudit -Name $field)))
        if ($missing.Count -gt 0) { $parameters[$field] = @{ Add = [object[]]@($missing) } }
    }
    return $parameters
}

function Get-ExchangeCasWriteParameter {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Differences,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config
    )

    $desiredValues = @{
        ActiveSyncEnabled = $false
        ImapEnabled = $false
        MAPIEnabled = $true
        OWAEnabled = $true
        OWAforDevicesEnabled = $false
        OwaMailboxPolicy = [string]$Config.OwaMailboxPolicy
        PopEnabled = $false
        SmtpClientAuthenticationDisabled = $true
    }
    $parameters = @{}
    foreach ($difference in @($Differences | Where-Object { $desiredValues.ContainsKey([string]$_.Field) })) {
        $parameters[[string]$difference.Field] = $desiredValues[[string]$difference.Field]
    }
    return $parameters
}

function New-ExchangeConfigurationResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions',
        '',
        Justification = 'Creates and returns an in-memory Exchange result record without changing external state.'
    )]
    param(
        [Parameter(Mandatory)][string] $UserPrincipalName,
        [Parameter(Mandatory)][string] $Status,
        [bool] $Changed,
        [AllowNull()][object[]] $Differences,
        [AllowNull()][object[]] $RemainingDifferences,
        [AllowNull()][string] $ErrorMessage
    )

    return [pscustomobject]@{
        UserPrincipalName = $UserPrincipalName
        Status = $Status
        Changed = $Changed
        Differences = [object[]]@($Differences)
        RemainingDifferences = [object[]]@($RemainingDifferences)
        Error = $ErrorMessage
    }
}

function Set-StudentMailboxConfiguration {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $UserPrincipalName,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [AllowNull()][object] $MailboxState
    )

    $upn = $UserPrincipalName.Trim()
    try {
        $state = if ($null -eq $MailboxState) {
            Get-StudentMailboxState -UserPrincipalName $upn -Config $Config
        } else {
            $MailboxState
        }
    } catch {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Failed -Changed:$false `
            -Differences @() -RemainingDifferences @() -ErrorMessage $_.Exception.Message
    }
    if (-not [bool]$state.Exists) {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status MailboxNotReady -Changed:$false `
            -Differences @() -RemainingDifferences @() -ErrorMessage $null
    }

    $differences = @(Compare-StudentMailboxState -Mailbox $state.Mailbox -CasMailbox $state.CasMailbox -Config $Config)
    if ($differences.Count -eq 0) {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Compliant -Changed:$false `
            -Differences @() -RemainingDifferences @() -ErrorMessage $null
    }

    $mailboxParameters = Get-ExchangeMailboxWriteParameter -Mailbox $state.Mailbox -Differences $differences -Config $Config
    $casParameters = Get-ExchangeCasWriteParameter -Differences $differences -Config $Config
    if ($WhatIfPreference) {
        if ($mailboxParameters.Count -gt 0) { [void]$PSCmdlet.ShouldProcess($upn, 'Exchange mailbox settings') }
        if ($casParameters.Count -gt 0) { [void]$PSCmdlet.ShouldProcess($upn, 'Exchange CAS mailbox settings') }
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Planned -Changed:$false `
            -Differences $differences -RemainingDifferences $differences -ErrorMessage $null
    }

    $changed = $false
    try {
        if ($mailboxParameters.Count -gt 0 -and $PSCmdlet.ShouldProcess($upn, 'Exchange mailbox settings')) {
            Set-Mailbox -Identity $upn @mailboxParameters -Confirm:$false -ErrorAction Stop
            $changed = $true
        }
        if ($casParameters.Count -gt 0 -and $PSCmdlet.ShouldProcess($upn, 'Exchange CAS mailbox settings')) {
            Set-CASMailbox -Identity $upn @casParameters -Confirm:$false -ErrorAction Stop
            $changed = $true
        }
    } catch {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Failed -Changed:$changed `
            -Differences $differences -RemainingDifferences $differences -ErrorMessage $_.Exception.Message
    }
    if (-not $changed) {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Skipped -Changed:$false `
            -Differences $differences -RemainingDifferences $differences -ErrorMessage $null
    }

    try {
        $verified = Get-StudentMailboxState -UserPrincipalName $upn -Config $Config
    } catch {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Failed -Changed:$true `
            -Differences $differences -RemainingDifferences $differences -ErrorMessage $_.Exception.Message
    }
    $verifiedReady = [bool](Get-ComparisonPropertyValue -InputObject $verified -Name Exists) -and
        [string]::Equals(
            [string](Get-ComparisonPropertyValue -InputObject $verified -Name Status),
            'Ready',
            [StringComparison]::OrdinalIgnoreCase
        ) -and
        $null -ne (Get-ComparisonPropertyValue -InputObject $verified -Name Mailbox) -and
        $null -ne (Get-ComparisonPropertyValue -InputObject $verified -Name CasMailbox)
    if (-not $verifiedReady) {
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Failed -Changed:$true `
            -Differences $differences -RemainingDifferences $differences `
            -ErrorMessage 'Exchange-Verifikation fehlgeschlagen: Das Postfach ist nach der Schreiboperation nicht bereit.'
    }
    $remaining = @($verified.Differences)
    if ($remaining.Count -gt 0) {
        $fields = @($remaining.Field | Sort-Object -Unique)
        return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Failed -Changed:$true `
            -Differences $differences -RemainingDifferences $remaining `
            -ErrorMessage "Exchange-Verifikation fehlgeschlagen. Abweichend: $($fields -join ', ')."
    }
    return New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Configured -Changed:$true `
        -Differences $differences -RemainingDifferences @() -ErrorMessage $null
}

function Wait-StudentMailbox {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]] $UserPrincipalName,
        [ValidateRange(0, 100)][int] $MaxRetries = 5,
        [ValidateRange(0, 86400)][int] $RetryDelaySeconds = 60,
        [scriptblock] $SleepAction = { param($seconds) Start-Sleep -Seconds $seconds },
        [System.Collections.IDictionary] $Config,
        [switch] $Configure
    )

    if ($Configure -and $null -eq $Config) {
        throw 'Config ist erforderlich, wenn Exchange-Postfächer konfiguriert werden sollen.'
    }

    $pending = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in $UserPrincipalName) {
        $upn = ([string]$candidate).Trim()
        if ([string]::IsNullOrWhiteSpace($upn)) { throw 'UserPrincipalName darf keinen leeren Wert enthalten.' }
        [void]$pending.Add($upn)
    }
    $attemptsByUpn = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($upn in $pending) { $attemptsByUpn[$upn] = 0 }
    $ready = [Collections.Generic.List[object]]::new()
    $failed = [Collections.Generic.List[object]]::new()

    for ($attempt = 0; $attempt -le $MaxRetries -and $pending.Count -gt 0; $attempt++) {
        if ($attempt -gt 0) { & $SleepAction $RetryDelaySeconds }
        foreach ($upn in @($pending)) {
            $attemptsByUpn[$upn] = [int]$attemptsByUpn[$upn] + 1
            try {
                $state = Get-StudentMailboxState -UserPrincipalName $upn -Config $Config
            } catch {
                $failed.Add([pscustomobject]@{
                        UserPrincipalName = $upn
                        Status = 'Failed'
                        Error = $_.Exception.Message
                    })
                [void]$pending.Remove($upn)
                continue
            }
            if (-not [bool]$state.Exists) { continue }

            $configuration = $null
            if ($Configure) {
                if ($WhatIfPreference) {
                    $configuration = Set-StudentMailboxConfiguration `
                        -UserPrincipalName $upn -Config $Config -MailboxState $state -WhatIf
                } elseif ($PSCmdlet.ShouldProcess($upn, 'Configure Exchange Online student mailbox')) {
                    $configuration = Set-StudentMailboxConfiguration `
                        -UserPrincipalName $upn -Config $Config -MailboxState $state -Confirm:$false
                } else {
                    $configuration = New-ExchangeConfigurationResult -UserPrincipalName $upn `
                        -Status Skipped -Changed:$false -Differences $state.Differences `
                        -RemainingDifferences $state.Differences -ErrorMessage $null
                }
                if ([string]::Equals([string]$configuration.Status, 'Failed', [StringComparison]::OrdinalIgnoreCase)) {
                    $failed.Add($configuration)
                    [void]$pending.Remove($upn)
                    continue
                }
            }
            $ready.Add([pscustomobject]@{
                    UserPrincipalName = $upn
                    MailboxState = $state
                    Configuration = $configuration
                })
            [void]$pending.Remove($upn)
        }
        if ($WhatIfPreference) { break }
    }

    return [pscustomobject]@{
        Ready = [object[]]@($ready)
        Missing = [string[]]@($pending)
        Failed = [object[]]@($failed)
        AttemptsByUpn = $attemptsByUpn
    }
}
