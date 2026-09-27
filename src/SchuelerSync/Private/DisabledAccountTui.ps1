function Read-DisabledAccountTuiKey {
    try {
        return [string][Console]::ReadKey($true).Key
    } catch {
        throw 'Die Tastaturauswahl benötigt ein interaktives Terminal.'
    }
}

function Test-DisabledAccountSelectable {
    param([Parameter(Mandatory)][object] $Account)

    $recipientType = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientType')
    $recipientStatus = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientStatus')
    if ($recipientStatus -eq 'None' -and [string]::IsNullOrWhiteSpace($recipientType)) { return $true }
    return ($recipientStatus -in @('', 'Unique') -and $recipientType -in @('UserMailbox', 'RemoteUserMailbox'))
}

function Show-DisabledAccountSelection {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.HashSet[string]] $SelectedIds,
        [Parameter(Mandatory)][int] $Cursor,
        [Parameter(Mandatory)][string] $TenantId,
        [string] $Message
    )

    Clear-Host
    Write-Host "Gesperrte Entra-Konten ($($Accounts.Count)) | Tenant $TenantId" -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Pfeile: bewegen  Leertaste: markieren  Enter: löschen  U: entsperren  Esc: zurück zur CLI'
    Write-Host 'Sperrzeitpunkt und echte Person können nicht automatisch bestätigt werden.'
    Write-Host 'Persönliche Benutzerpostfächer und Konten ohne Exchange-Empfänger sind auswählbar.'
    Write-Host ''
    Write-Host "Markiert: $($SelectedIds.Count)"
    if ($Message) { Write-Host $Message -ForegroundColor Yellow }
    Write-Host ''

    $height = try { $Host.UI.RawUI.WindowSize.Height } catch { 24 }
    $pageSize = [math]::Max(5, $height - 10)
    $pageStart = [int]([math]::Floor($Cursor / $pageSize) * $pageSize)
    $pageEnd = [math]::Min($Accounts.Count, $pageStart + $pageSize)
    for ($index = $pageStart; $index -lt $pageEnd; $index++) {
        $account = $Accounts[$index]
        $pointer = if ($index -eq $Cursor) { '>' } else { ' ' }
        $mark = if ($SelectedIds.Contains([string]$account.UserId)) { 'x' } else { ' ' }
        $recipientType = [string]$account.RecipientType
        $recipientStatus = [string](Get-ComparisonPropertyValue -InputObject $account -Name 'RecipientStatus')
        $warning = if ($recipientStatus -eq 'None' -and [string]::IsNullOrWhiteSpace($recipientType)) {
            '  [Kein Exchange-Empfänger; Prüfgründe]'
        } elseif (-not (Test-DisabledAccountSelectable -Account $account)) {
            $label = if ([string]::IsNullOrWhiteSpace($recipientType)) { 'ohne eindeutigen Exchange-Empfänger' } else { $recipientType }
            "  [Nur prüfen: $label]"
        } elseif (@($account.ProtectionReasons).Count -gt 0) { '  [Prüfgründe]' } else { '' }
        $line = "$pointer [$mark] $($account.DisplayName) <$($account.UPN)>$warning"
        if ($index -eq $Cursor) { Write-Host $line -ForegroundColor Cyan }
        else { Write-Host $line }
    }
    if ($Accounts.Count -eq 0) { Write-Host 'Keine Konten in der Prüfliste.' }
    if ($Accounts.Count -gt $pageSize) { Write-Host "`nEinträge $($pageStart + 1)-$pageEnd von $($Accounts.Count)" }
    if ($Accounts.Count -gt 0) {
        $current = $Accounts[$Cursor]
        Write-Host "`nAktuell: $($current.DisplayName) | $($current.UserId)"
        $recipientType = if ([string]$current.RecipientStatus -eq 'None') { 'Keiner' } elseif ([string]::IsNullOrWhiteSpace([string]$current.RecipientType)) { 'Nicht eindeutig oder nicht lesbar' } else { [string]$current.RecipientType }
        Write-Host "Exchange-Empfänger: $recipientType"
        if (@($current.ProtectionReasons).Count -gt 0) {
            Write-Host "Prüfgründe: $($current.ProtectionReasons -join ', ')" -ForegroundColor Yellow
        }
        foreach ($ownedObject in @($current.OwnedObjects)) {
            Write-Host "Besitz: $($ownedObject.Type): $($ownedObject.DisplayName) ($($ownedObject.Id))" -ForegroundColor Yellow
        }
    }
}

function Show-DisabledAccountConfirmation {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId
    )

    Clear-Host
    Write-Host "Konten löschen | Tenant $TenantId" -ForegroundColor Red
    Write-Host ''
    Write-Host 'Diese Benutzerkonten werden in Entra gelöscht (zunächst 30 Tage wiederherstellbar):'
    foreach ($account in $Accounts) {
        Write-Host "  $($account.DisplayName) <$($account.UPN)>"
        $ownedObjects = @($account.OwnedObjects | Where-Object { $null -ne $_ })
        if ($ownedObjects.Count -gt 0) {
            Write-Host "    Besitzt $($ownedObjects.Count) Entra-Objekt(e). Bei Löschung können diese ohne Eigentümer bleiben:" -ForegroundColor Yellow
            foreach ($ownedObject in $ownedObjects) {
                $type = if ($ownedObject.Type) { [string]$ownedObject.Type } else { 'Unbekannt' }
                $name = if ($ownedObject.DisplayName) { [string]$ownedObject.DisplayName } else { 'Name nicht lesbar' }
                Write-Host "    $($type): $name ($($ownedObject.Id))" -ForegroundColor Yellow
            }
        }
    }
    Write-Host ''
    Write-Host 'Prüfe selbst, dass jede Auswahl eine echte Person ist und gelöscht werden darf.' -ForegroundColor Yellow
    Write-Host 'Ein Sperrdatum und eine 90-Tage-Frist sind nicht belegt.' -ForegroundColor Yellow
    Write-Host 'Ja (Y/J): löschen    Nein (N/Esc): zurück zur Auswahl' -ForegroundColor Red
}

function Show-DisabledAccountUnlockConfirmation {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId
    )

    Clear-Host
    Write-Host "Konten entsperren | Tenant $TenantId" -ForegroundColor Green
    Write-Host ''
    Write-Host 'Diese Benutzerkonten werden in Entra wieder für die Anmeldung aktiviert:'
    foreach ($account in $Accounts) { Write-Host "  $($account.DisplayName) <$($account.UPN)>" }
    Write-Host ''
    Write-Host 'Die Anmeldung ist danach mit dem bestehenden Passwort möglich. Es wird kein Passwort zurückgesetzt.' -ForegroundColor Yellow
    Write-Host 'Ja (Y/J): entsperren    Nein (N/Esc): zurück zur Auswahl' -ForegroundColor Green
}

function Connect-DisabledAccountWriteGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $TenantId,
        [Parameter(Mandatory)][ValidateSet('Delete', 'Unlock')][string] $Operation
    )

    $writeScope = if ($Operation -eq 'Delete') { 'User.ReadWrite.All' } else { 'User.EnableDisableAccount.All' }
    $requiredScopes = @($writeScope, 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All')
    $context = Get-MgContext
    $grantedScopes = if ($null -eq $context) { @() } else { @($context.Scopes) }
    $missingScopes = @(Test-GraphScopeSet -GrantedScopes $grantedScopes -RequiredScopes $requiredScopes)
    if ($missingScopes.Count -gt 0) {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        $context = Get-MgContext
        $grantedScopes = if ($null -eq $context) { @() } else { @($context.Scopes) }
        $missingScopes = @(Test-GraphScopeSet -GrantedScopes $grantedScopes -RequiredScopes $requiredScopes)
    }
    if ($missingScopes.Count -gt 0) {
        throw "Graph-Kontext hat keine Berechtigung $writeScope für die Aktion $Operation."
    }
    if ([string]$context.TenantId -ine $TenantId) {
        throw "Graph-Tenant hat sich geändert. Erwartet: $TenantId, verbunden: $($context.TenantId)."
    }
    return $context
}

function Connect-DisabledAccountDeleteGraph {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $TenantId)

    return Connect-DisabledAccountWriteGraph -TenantId $TenantId -Operation Delete
}

function Assert-DisabledAccountActionTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Account,
        [Parameter(Mandatory)][string] $TenantId
    )

    $null = Assert-ExchangeTenant -TenantId $TenantId
    $userId = [string]$Account.UserId
    if ([string]::IsNullOrWhiteSpace($userId)) { throw 'Auswahl ohne Entra-Objekt-ID.' }
    $user = Get-MgUser -UserId $userId -Property @('id', 'displayName', 'userPrincipalName', 'accountEnabled', 'userType', 'onPremisesSyncEnabled') -ErrorAction Stop
    if ([string]$user.Id -ine $userId -or [string]$user.UserPrincipalName -ine [string]$Account.UPN -or
        [string]$user.DisplayName -cne [string]$Account.DisplayName) {
        throw 'Identität hat sich seit der Auswahl geändert.'
    }
    if ($user.AccountEnabled -ne $false) { throw 'Konto ist nicht mehr gesperrt.' }
    if ([string]$user.UserType -ne 'Member') { throw 'Konto ist kein internes Mitgliedskonto.' }
    if ($user.OnPremisesSyncEnabled -eq $true) { throw 'Konto wird aus lokalem AD synchronisiert.' }

    $selectedRecipientType = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientType')
    $selectedRecipientStatus = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientStatus')
    if ($selectedRecipientStatus -eq 'None' -and [string]::IsNullOrWhiteSpace($selectedRecipientType)) {
        $recipients = @(Get-EXORecipient -ResultSize Unlimited -Properties @('ExternalDirectoryObjectId', 'RecipientTypeDetails') -ErrorAction Stop |
            Where-Object { ([string](Get-ComparisonPropertyValue -InputObject $_ -Name 'ExternalDirectoryObjectId')).Trim() -ieq $userId })
        if ($recipients.Count -gt 0) { throw 'Seit der Auswahl ist ein Exchange-Empfänger vorhanden. Konto erneut prüfen.' }
    } elseif ($selectedRecipientStatus -in @('', 'Unique') -and $selectedRecipientType -in @('UserMailbox', 'RemoteUserMailbox')) {
        $recipients = @(Get-EXORecipient -ExternalDirectoryObjectId ([guid]$userId) `
            -Properties @('ExternalDirectoryObjectId', 'RecipientTypeDetails') -ErrorAction Stop)
        if ($recipients.Count -ne 1) { throw 'Exchange-Empfänger ist nicht eindeutig vorhanden.' }
        $recipientId = [string](Get-ComparisonPropertyValue -InputObject $recipients[0] -Name 'ExternalDirectoryObjectId')
        if ($recipientId -ine $userId) { throw 'Exchange-Empfänger passt nicht zur Entra-Objekt-ID.' }
        $recipientType = [string](Get-ComparisonPropertyValue -InputObject $recipients[0] -Name 'RecipientTypeDetails')
        if ($recipientType -notin @('UserMailbox', 'RemoteUserMailbox')) { throw "Exchange-Empfängertyp ist '$recipientType'." }
    } else {
        throw 'Exchange-Empfänger aus der Auswahl ist nicht eindeutig als Benutzerpostfach oder fehlend bestätigt.'
    }

    $memberships = @(Get-MgUserTransitiveMemberOf -UserId $userId -All -ErrorAction Stop)
    foreach ($member in $memberships) {
        $type = [string](Get-DisabledAccountMemberValue -Member $member -Name '@odata.type')
        if ([string]::IsNullOrWhiteSpace($type)) { throw 'Mitgliedschaftstyp nicht lesbar.' }
        if ($type -eq '#microsoft.graph.directoryRole') { throw 'Konto hat eine Verzeichnisrolle.' }
    }
    return $user
}

function Invoke-SelectedDisabledAccountDeletion {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId
    )

    $ErrorActionPreference = 'Stop'
    if (-not $WhatIfPreference) { $null = Connect-DisabledAccountDeleteGraph -TenantId $TenantId }
    $seenIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $activity = 'Gesperrte Entra-Konten löschen'
    $index = 0
    try {
        foreach ($account in $Accounts) {
            $index++
            Write-Progress -Id 2 -Activity $activity -Status "Prüfe Konto $index von $($Accounts.Count): $($account.DisplayName)" `
                -PercentComplete ([int](100 * ($index - 1) / $Accounts.Count))
            $userId = [string]$account.UserId
            $status = 'Skipped'
            $message = ''
            try {
                if (-not $seenIds.Add($userId)) { throw 'Konto wurde mehrfach ausgewählt.' }
                $null = Assert-DisabledAccountActionTarget -Account $account -TenantId $TenantId
                $ownedObjects = @(Get-MgUserOwnedObject -UserId $userId -All -ErrorAction Stop)
                if ($ownedObjects.Count -gt 0) {
                    Write-Warning "$($account.UPN): Besitzt $($ownedObjects.Count) Entra-Objekt(e). Diese können nach der Löschung ohne Eigentümer bleiben."
                }
                if ($PSCmdlet.ShouldProcess("$($account.DisplayName) <$($account.UPN)> [$userId]", 'Entra-Benutzer löschen')) {
                    try {
                        Remove-MgUser -UserId $userId -Confirm:$false -ErrorAction Stop
                        $status = 'Deleted'
                    } catch {
                        $status = 'Failed'
                        throw
                    }
                } else {
                    $status = 'WhatIf'
                }
            } catch {
                $message = $_.Exception.Message
            }
            [pscustomobject]@{
                DisplayName = [string]$account.DisplayName
                UPN = [string]$account.UPN
                UserId = $userId
                Status = $status
                Message = $message
            }
        }
        Write-Progress -Id 2 -Activity $activity -Status 'Abgeschlossen' -PercentComplete 100
    } finally {
        Write-Progress -Id 2 -Activity $activity -Completed
    }
}

function Invoke-SelectedDisabledAccountUnlock {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId
    )

    $ErrorActionPreference = 'Stop'
    if (-not $WhatIfPreference) { $null = Connect-DisabledAccountWriteGraph -TenantId $TenantId -Operation Unlock }
    $seenIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $activity = 'Gesperrte Entra-Konten entsperren'
    $index = 0
    try {
        foreach ($account in $Accounts) {
            $index++
            Write-Progress -Id 2 -Activity $activity -Status "Prüfe Konto $index von $($Accounts.Count): $($account.DisplayName)" `
                -PercentComplete ([int](100 * ($index - 1) / $Accounts.Count))
            $userId = [string]$account.UserId
            $status = 'Skipped'
            $message = ''
            try {
                if (-not $seenIds.Add($userId)) { throw 'Konto wurde mehrfach ausgewählt.' }
                $null = Assert-DisabledAccountActionTarget -Account $account -TenantId $TenantId
                if ($PSCmdlet.ShouldProcess("$($account.DisplayName) <$($account.UPN)> [$userId]", 'Entra-Benutzer entsperren')) {
                    try {
                        Update-MgUser -UserId $userId -AccountEnabled:$true -Confirm:$false -ErrorAction Stop
                        $status = 'Enabled'
                    } catch {
                        $status = 'Failed'
                        throw
                    }
                } else {
                    $status = 'WhatIf'
                }
            } catch {
                $message = $_.Exception.Message
            }
            [pscustomobject]@{
                DisplayName = [string]$account.DisplayName
                UPN = [string]$account.UPN
                UserId = $userId
                Status = $status
                Message = $message
            }
        }
        Write-Progress -Id 2 -Activity $activity -Status 'Abgeschlossen' -PercentComplete 100
    } finally {
        Write-Progress -Id 2 -Activity $activity -Completed
    }
}

function Invoke-DisabledAccountTui {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId
    )

    $Accounts = @($Accounts | Where-Object { [string]$_.RecipientType -ne 'SharedMailbox' })
    $cursor = 0
    $selectedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $message = ''
    while ($true) {
        Show-DisabledAccountSelection -Accounts $Accounts -SelectedIds $selectedIds -Cursor $cursor -TenantId $TenantId -Message $message
        $message = ''
        $key = Read-DisabledAccountTuiKey
        switch ($key) {
            'Escape' { return }
            'UpArrow' { if ($cursor -gt 0) { $cursor-- } }
            'DownArrow' { if ($cursor -lt $Accounts.Count - 1) { $cursor++ } }
            'Spacebar' {
                if ($Accounts.Count -gt 0) {
                    $current = $Accounts[$cursor]
                    if (-not (Test-DisabledAccountSelectable -Account $current)) {
                        $message = 'Dieses Konto ist nur zur Prüfung sichtbar. Der Exchange-Empfänger ist nicht eindeutig geklärt.'
                    } else {
                        $userId = [string]$current.UserId
                        if (-not $selectedIds.Add($userId)) { [void]$selectedIds.Remove($userId) }
                    }
                }
            }
            'Enter' {
                if ($selectedIds.Count -eq 0) {
                    $message = 'Markiere zuerst mindestens ein Konto.'
                    break
                }
                $selectedAccounts = @($Accounts | Where-Object { $selectedIds.Contains([string]$_.UserId) })
                Show-DisabledAccountConfirmation -Accounts $selectedAccounts -TenantId $TenantId
                do { $answer = Read-DisabledAccountTuiKey } while ($answer -notin @('Y', 'J', 'N', 'Escape'))
                if ($answer -in @('Y', 'J')) {
                    return Invoke-SelectedDisabledAccountDeletion -Accounts $selectedAccounts -TenantId $TenantId `
                        -WhatIf:$WhatIfPreference -Confirm:$false
                }
            }
            'U' {
                if ($selectedIds.Count -eq 0) {
                    $message = 'Markiere zuerst mindestens ein Konto.'
                    break
                }
                $selectedAccounts = @($Accounts | Where-Object { $selectedIds.Contains([string]$_.UserId) })
                Show-DisabledAccountUnlockConfirmation -Accounts $selectedAccounts -TenantId $TenantId
                do { $answer = Read-DisabledAccountTuiKey } while ($answer -notin @('Y', 'J', 'N', 'Escape'))
                if ($answer -in @('Y', 'J')) {
                    return Invoke-SelectedDisabledAccountUnlock -Accounts $selectedAccounts -TenantId $TenantId `
                        -WhatIf:$WhatIfPreference -Confirm:$false
                }
            }
        }
    }
}
