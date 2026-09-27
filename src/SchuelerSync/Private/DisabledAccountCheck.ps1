function Connect-DisabledAccountGraph {
    [CmdletBinding()]
    param()

    $requiredScopes = @('User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All')
    $context = Get-MgContext
    $grantedScopes = if ($null -eq $context) { @() } else { @($context.Scopes) }
    $missingScopes = @(Test-GraphScopeSet -GrantedScopes $grantedScopes -RequiredScopes $requiredScopes)
    if ($missingScopes.Count -eq 0) { return $context }

    Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
    $context = Get-MgContext
    $grantedScopes = if ($null -eq $context) { @() } else { @($context.Scopes) }
    $missingScopes = @(Test-GraphScopeSet -GrantedScopes $grantedScopes -RequiredScopes $requiredScopes)
    if ($missingScopes.Count -gt 0) {
        throw "Microsoft Graph context lacks required scopes: $($missingScopes -join ', ')."
    }
    return $context
}

function Get-DisabledAccountMemberValue {
    param(
        [Parameter(Mandatory)][object] $Member,
        [Parameter(Mandatory)][string] $Name
    )

    $value = Get-ComparisonPropertyValue -InputObject $Member -Name $Name
    if ($null -ne $value) { return $value }
    if ($Name -eq '@odata.type') {
        $value = Get-ComparisonPropertyValue -InputObject $Member -Name 'OdataType'
        if ($null -ne $value) { return $value }
    }
    $additional = Get-ComparisonPropertyValue -InputObject $Member -Name AdditionalProperties
    if ($additional -is [System.Collections.IDictionary]) {
        foreach ($key in $additional.Keys) {
            if ([string]::Equals([string]$key, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                return $additional[$key]
            }
        }
    }
    return $null
}

function ConvertTo-DisabledAccountCheckRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $User,
        [AllowNull()][object[]] $Memberships,
        [AllowNull()][object[]] $OwnedObjects,
        [AllowNull()][object] $Recipient,
        [string] $RecipientStatus,
        [string[]] $CheckErrors = @(),
        [bool] $LastLoginAvailable = $true
    )

    if ($User.AccountEnabled -ne $false) {
        throw "Konto '$($User.Id)' ist nicht gesperrt."
    }

    $secAGroups = [Collections.Generic.List[string]]::new()
    $reasons = [Collections.Generic.List[string]]::new()
    $ownedObjectDetails = [Collections.Generic.List[object]]::new()
    foreach ($member in @($Memberships)) {
        if ($null -eq $member) { continue }
        $type = [string](Get-DisabledAccountMemberValue -Member $member -Name '@odata.type')
        $name = [string](Get-DisabledAccountMemberValue -Member $member -Name 'displayName')
        if ($type -eq '#microsoft.graph.directoryRole') {
            if ($name) { $reasons.Add("Verzeichnisrolle: $name") }
            else { $reasons.Add('Verzeichnisrolle ohne lesbaren Namen') }
        } elseif ($type -eq '#microsoft.graph.group') {
            if (-not $name) { $reasons.Add('Gruppenname nicht lesbar') }
            elseif ($name.StartsWith('SEC-A-', [StringComparison]::OrdinalIgnoreCase)) { $secAGroups.Add($name) }
        }
    }

    if ([string]$User.UserType -eq 'Guest') { $reasons.Add('Gastkonto') }
    if ($User.OnPremisesSyncEnabled -eq $true) { $reasons.Add('Lokales AD synchronisiert') }
    foreach ($ownedObject in @($OwnedObjects)) {
        if ($null -eq $ownedObject) { continue }
        $type = [string](Get-DisabledAccountMemberValue -Member $ownedObject -Name '@odata.type')
        if ($type.StartsWith('#microsoft.graph.', [StringComparison]::OrdinalIgnoreCase)) {
            $type = $type.Substring('#microsoft.graph.'.Length)
        }
        if ([string]::IsNullOrWhiteSpace($type)) { $type = 'Unbekannt' }
        $ownedObjectDetails.Add([pscustomobject]@{
            Type = $type
            DisplayName = [string](Get-DisabledAccountMemberValue -Member $ownedObject -Name 'displayName')
            Id = [string](Get-DisabledAccountMemberValue -Member $ownedObject -Name 'id')
        })
    }
    if ($ownedObjectDetails.Count -gt 0) { $reasons.Add("Besitzt $($ownedObjectDetails.Count) Entra-Objekt(e)") }

    $recipientType = if ($null -eq $Recipient) { '' } else { [string](Get-ComparisonPropertyValue -InputObject $Recipient -Name 'RecipientTypeDetails') }
    if ([string]::IsNullOrWhiteSpace($RecipientStatus)) {
        $RecipientStatus = if ($null -eq $Recipient) { 'None' } elseif ([string]::IsNullOrWhiteSpace($recipientType)) { 'Unreadable' } else { 'Unique' }
    }
    if (-not [string]::IsNullOrWhiteSpace($recipientType) -and $recipientType -notin @('UserMailbox', 'RemoteUserMailbox')) {
        $reasons.Add("Exchange-Empfänger: $recipientType")
    } elseif ($null -ne $Recipient -and [string]::IsNullOrWhiteSpace($recipientType)) {
        $reasons.Add('Exchange-Empfängertyp nicht lesbar')
    } elseif ($RecipientStatus -eq 'None') {
        $reasons.Add('Kein Exchange-Empfänger, Kontozweck prüfen')
    }
    if (@($CheckErrors).Count -gt 0) {
        foreach ($errorText in @($CheckErrors)) { $reasons.Add($errorText) }
    }
    if ($secAGroups.Count -eq 0 -and @($CheckErrors | Where-Object { $_ -match 'Gruppen' }).Count -eq 0) {
        $reasons.Add('Keine SEC-A-Gruppe, Kontozweck prüfen')
    }

    $lastLogin = $null
    if ($LastLoginAvailable -and $null -ne $User.SignInActivity) {
        $rawLastLogin = $User.SignInActivity.LastSuccessfulSignInDateTime
        if ($null -ne $rawLastLogin -and -not [string]::IsNullOrWhiteSpace([string]$rawLastLogin)) {
            $lastLogin = [datetimeoffset]$rawLastLogin
        }
    }

    [pscustomobject]@{
        DisplayName = [string]$User.DisplayName
        UPN = [string]$User.UserPrincipalName
        UserId = [string]$User.Id
        SecAGroups = @(($secAGroups | Sort-Object -Unique))
        LastLogin = $lastLogin
        LastLoginStatus = if (-not $LastLoginAvailable) { 'Nicht abrufbar' } elseif ($null -eq $lastLogin) { 'Kein Wert' } else { 'Bekannt' }
        RecipientType = $recipientType
        RecipientStatus = $RecipientStatus
        OwnedObjects = @($ownedObjectDetails)
        ProtectionReasons = @($reasons)
        ReviewStatus = if ($reasons.Count -gt 0) { 'Schutz/Prüfung nötig' } else { 'Manuell prüfen' }
        DisabledAt = $null
        Meets90Days = $null
    }
}

function Invoke-DisabledAccountCheck {
    [CmdletBinding()]
    param()

    $ErrorActionPreference = 'Stop'
    $activity = 'Gesperrte Entra-Konten prüfen'
    $warnings = [Collections.Generic.List[string]]::new()
    $records = [Collections.Generic.List[object]]::new()
    try {
        Write-Progress -Id 1 -Activity $activity -Status 'Verbinde mit Microsoft Graph ...' -PercentComplete 5
        $context = Connect-DisabledAccountGraph
        $tenantId = [string](Get-ComparisonPropertyValue $context TenantId)
        if ([string]::IsNullOrWhiteSpace($tenantId)) { throw 'Graph-Kontext enthält keine Tenant-ID.' }
        Write-Progress -Id 1 -Activity $activity -Status 'Verbinde mit Exchange Online ...' -PercentComplete 15
        $null = Connect-SchuelerExchangeOnline -TenantId $tenantId

        Write-Progress -Id 1 -Activity $activity -Status 'Lade gesperrte Konten ...' -PercentComplete 25
        $userProperties = @('id', 'displayName', 'userPrincipalName', 'accountEnabled', 'userType', 'onPremisesSyncEnabled', 'signInActivity')
        $lastLoginAvailable = $true
        try {
            $users = @(Get-MgUser -All -Filter 'accountEnabled eq false' -Property $userProperties -ErrorAction Stop)
        } catch {
            $lastLoginAvailable = $false
            $warnings.Add("Letzte Anmeldung konnte nicht mitgeladen werden: $($_.Exception.Message)")
            $users = @(Get-MgUser -All -Filter 'accountEnabled eq false' -Property @('id', 'displayName', 'userPrincipalName', 'accountEnabled', 'userType', 'onPremisesSyncEnabled') -ErrorAction Stop)
        }

        $disabledUsers = @($users | Where-Object { $_.AccountEnabled -eq $false } | Sort-Object DisplayName, UserPrincipalName)
        Write-Progress -Id 1 -Activity $activity -Status 'Lade Exchange-Empfänger ...' -PercentComplete 28
        $exchangeRecipients = @(Get-EXORecipient -ResultSize Unlimited -Properties @('ExternalDirectoryObjectId', 'RecipientTypeDetails') -ErrorAction Stop)
        $recipientsById = @{}
        $duplicateRecipientIds = @{}
        foreach ($exchangeRecipient in $exchangeRecipients) {
            $recipientId = [string](Get-ComparisonPropertyValue -InputObject $exchangeRecipient -Name 'ExternalDirectoryObjectId')
            if ([string]::IsNullOrWhiteSpace($recipientId)) { continue }
            $recipientId = $recipientId.Trim()
            if ($recipientsById.ContainsKey($recipientId)) {
                $duplicateRecipientIds[$recipientId] = $true
            } else {
                $recipientsById[$recipientId] = $exchangeRecipient
            }
        }

        $index = 0
        foreach ($user in $disabledUsers) {
            $index++
            $percent = 30 + [int](60 * $index / [math]::Max(1, $disabledUsers.Count))
            Write-Progress -Id 1 -Activity $activity -Status "Prüfe gesperrtes Konto $index von $($disabledUsers.Count) ..." -PercentComplete $percent

            $recipient = $null
            $recipientStatus = 'None'
            $userId = [string]$user.Id
            $checkErrors = [Collections.Generic.List[string]]::new()
            if ($duplicateRecipientIds.ContainsKey($userId)) {
                $recipientStatus = 'Ambiguous'
                $checkErrors.Add('Exchange-Empfänger nicht eindeutig zugeordnet')
                $warnings.Add("$($user.UserPrincipalName): Mehrere Exchange-Empfänger, Zuordnung muss geprüft werden.")
            } elseif ($recipientsById.ContainsKey($userId)) {
                $recipient = $recipientsById[$userId]
                $recipientStatus = 'Unique'
            }
            $recipientType = if ($null -eq $recipient) { '' } else {
                [string](Get-ComparisonPropertyValue -InputObject $recipient -Name 'RecipientTypeDetails')
            }
            if ($null -ne $recipient -and [string]::IsNullOrWhiteSpace($recipientType)) {
                $recipientStatus = 'Unreadable'
                $warnings.Add("$($user.UserPrincipalName): Exchange-Empfängertyp nicht lesbar, Kontozweck muss geprüft werden.")
            }

            $memberships = @()
            try {
                $memberships = @(Get-MgUserTransitiveMemberOf -UserId $user.Id -All -ErrorAction Stop)
            } catch {
                $checkErrors.Add('Gruppen und Rollen nicht vollständig abrufbar')
                $warnings.Add("$($user.UserPrincipalName): Gruppen/Rollen: $($_.Exception.Message)")
            }

            $ownedObjects = @()
            try {
                $ownedObjects = @(Get-MgUserOwnedObject -UserId $user.Id -All -ErrorAction Stop)
            } catch {
                $checkErrors.Add('Besitz von Entra-Objekten nicht prüfbar')
                $warnings.Add("$($user.UserPrincipalName): Objektbesitz: $($_.Exception.Message)")
            }

            $record = ConvertTo-DisabledAccountCheckRecord -User $user -Memberships $memberships -OwnedObjects $ownedObjects `
                -Recipient $recipient -RecipientStatus $recipientStatus -CheckErrors @($checkErrors) -LastLoginAvailable $lastLoginAvailable
            $records.Add($record)
        }

        Write-Progress -Id 1 -Activity $activity -Status 'Erzeuge Prüfliste ...' -PercentComplete 95
        return [pscustomobject]@{
            Mode = 'Check'
            TenantId = [string]$context.TenantId
            CheckedAt = [datetimeoffset]::UtcNow
            Accounts = @($records)
            ExcludedCount = 0
            Warnings = @($warnings)
            DisabledDateSource = 'Unbekannt'
        }
    } finally {
        Write-Progress -Id 1 -Activity $activity -Completed
    }
}
