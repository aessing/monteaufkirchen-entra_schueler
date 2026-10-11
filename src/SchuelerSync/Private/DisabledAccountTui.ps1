function Read-DisabledAccountTuiKey {
    try {
        return [string][Console]::ReadKey($true).Key
    } catch {
        throw 'Die Tastaturauswahl benötigt ein interaktives Terminal.'
    }
}

function Test-DisabledAccountSelectable {
    param([Parameter(Mandatory)][object] $Account)

    if ($null -ne $Account.PSObject.Properties['MutationAllowed'] -and -not $Account.MutationAllowed) { return $false }
    $recipientType = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientType')
    $recipientStatus = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientStatus')
    if ($recipientStatus -eq 'None' -and [string]::IsNullOrWhiteSpace($recipientType)) { return $true }
    if ((Get-ComparisonPropertyValue $Account UserType) -eq 'Guest' -and
        (Get-ComparisonPropertyValue $Account MutationAllowed) -eq $true -and $recipientStatus -eq 'Unique' -and
        $recipientType -in @('MailUser','GuestMailUser')) { return $true }
    return ($recipientStatus -in @('', 'Unique') -and $recipientType -in @('UserMailbox', 'RemoteUserMailbox'))
}

function Show-DisabledAccountSelection {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.HashSet[string]] $SelectedIds,
        [Parameter(Mandatory)][int] $Cursor,
        [Parameter(Mandatory)][string] $TenantId,
        [string] $Message,
        [switch] $PersonalMode,
        [string] $Title,
        [switch] $StudentMode
    )

    Clear-Host
    if (-not $Title) { $Title = if ($PersonalMode) { 'Personalkonten' } else { 'Gesperrte Entra-Konten' } }
    Write-Host "$title ($($Accounts.Count)) | Tenant $TenantId" -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Sperrzeitpunkt und echte Person können nicht automatisch bestätigt werden.'
    if ($StudentMode) { Write-Host 'Schüler-Mitgliedskonten: B sperrt aktive Konten, Backspace/Entf löscht gesperrte, U entsperrt gezielt.' }
    elseif ($PersonalMode) { Write-Host 'Zulässige Konten und Gäste: B sperrt aktive Konten, Backspace/Entf löscht gesperrte. U gilt nur für Mitglieder.' }
    else { Write-Host 'Persönliche Benutzerpostfächer und Konten ohne Exchange-Empfänger sind auswählbar.' }
    Write-Host ''
    Write-Host "Markiert: $($SelectedIds.Count)"
    if ($Message) { Write-Host $Message -ForegroundColor Yellow }
    Write-Host ''

    $height = try { $Host.UI.RawUI.WindowSize.Height } catch { 24 }
    $pageSize = [math]::Max(5, $height - $(if ($StudentMode) { 17 } elseif ($PersonalMode) { 13 } else { 12 }))
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
        $accountStatus = if ($PersonalMode) { if ($account.AccountEnabled) { ' [Aktiv]' } else { ' [Gesperrt]' } } else { '' }
        $line = "$pointer [$mark] $($account.DisplayName) <$($account.UPN)>$accountStatus$warning"
        if ($index -eq $Cursor) { Write-Host $line -ForegroundColor Cyan }
        else { Write-Host $line }
    }
    if ($Accounts.Count -eq 0) { Write-Host 'Keine Konten in der Prüfliste.' }
    if ($Accounts.Count -gt $pageSize) { Write-Host "`nEinträge $($pageStart + 1)-$pageEnd von $($Accounts.Count)" }
    if ($Accounts.Count -gt 0) {
        $current = $Accounts[$Cursor]
        Write-Host "`nAktuell: $($current.DisplayName) | $($current.UserId)"
        $roles = @((Get-ComparisonPropertyValue $current SecAGroups) | Where-Object { $_ -like 'SEC-A-ROL-*' } | Sort-Object -Unique)
        $roleText = if ($roles.Count -gt 0) { $roles -join ', ' } else { 'Keine Rollengruppe' }
        Write-Host "Rolle: $roleText"
        if ($PersonalMode) { Write-Host "Kontotyp: $(Get-ComparisonPropertyValue $current UserType)" }
        if ($StudentMode) {
            $classes = @($current.ClassGroups | ForEach-Object { $_.DisplayName })
            Write-Host "Klasse (CLS): $(if ($classes.Count) { $classes -join ', ' } else { 'Keine CLS-Gruppe' })"
            Write-Host "Department / Jahrgang: $($current.Department) / $($current.OfficeLocation)"
            Write-Host "Klassenlehrer: $($current.Teacher)"
        }
        $lastLogin = Get-ComparisonPropertyValue -InputObject $current -Name LastLogin
        $loginText = if ($null -ne $lastLogin) {
            ([datetimeoffset]$lastLogin).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'",[Globalization.CultureInfo]::InvariantCulture)
        } else {
            $loginStatus = [string](Get-ComparisonPropertyValue -InputObject $current -Name LastLoginStatus)
            if ($loginStatus -in @('Kein Wert','Nicht abrufbar')) { $loginStatus } else { 'Kein Wert' }
        }
        Write-Host "Last Login Timestamp (UTC): $loginText"
        $recipientType = if ([string]$current.RecipientStatus -eq 'None') { 'Keiner' } elseif ([string]::IsNullOrWhiteSpace([string]$current.RecipientType)) { 'Nicht eindeutig oder nicht lesbar' } else { [string]$current.RecipientType }
        Write-Host "Exchange-Empfänger: $recipientType"
        if (@($current.ProtectionReasons).Count -gt 0) {
            Write-Host "Prüfgründe: $($current.ProtectionReasons -join ', ')" -ForegroundColor Yellow
        }
        foreach ($ownedObject in @($current.OwnedObjects)) {
            Write-Host "Besitz: $($ownedObject.Type): $($ownedObject.DisplayName) ($($ownedObject.Id))" -ForegroundColor Yellow
        }
    }
    Write-Host ''
    $commands = 'Pfeile: bewegen  Leertaste: markieren'
    if ($PersonalMode) { $commands += '  B: sperren' }
    $commands += '  U: entsperren  Backspace/Entf: löschen'
    if ($StudentMode) { $commands += '  K: Klasse/Jahrgang wechseln  L: Klassenlehrer wechseln' }
    Write-Host "$commands  Esc: zurück zur CLI" -ForegroundColor Cyan
}

function Show-DisabledAccountConfirmation {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId,
        [object[]] $SkippedAccounts = @()
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
    foreach ($account in $SkippedAccounts) { Write-Host "Übersprungen: $($account.DisplayName) <$($account.UPN)> ($($account.Message))" -ForegroundColor Yellow }
    Write-Host 'Ja (Y/J): löschen    Nein (N/Esc): zurück zur Auswahl' -ForegroundColor Red
}

function Show-DisabledAccountUnlockConfirmation {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId,
        [object[]] $SkippedAccounts = @()
    )

    Clear-Host
    Write-Host "Konten entsperren | Tenant $TenantId" -ForegroundColor Green
    Write-Host ''
    Write-Host 'Diese Benutzerkonten werden in Entra wieder für die Anmeldung aktiviert:'
    foreach ($account in $Accounts) { Write-Host "  $($account.DisplayName) <$($account.UPN)>" }
    Write-Host ''
    Write-Host 'Die Anmeldung ist danach mit dem bestehenden Passwort möglich. Es wird kein Passwort zurückgesetzt.' -ForegroundColor Yellow
    foreach ($account in $SkippedAccounts) { Write-Host "Übersprungen: $($account.DisplayName) <$($account.UPN)> ($($account.Message))" -ForegroundColor Yellow }
    Write-Host 'Ja (Y/J): entsperren    Nein (N/Esc): zurück zur Auswahl' -ForegroundColor Green
}

function Connect-DisabledAccountWriteGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $TenantId,
        [Parameter(Mandatory)][ValidateSet('Delete', 'Unlock', 'Block','Class','Teacher')][string] $Operation
    )

    $writeScope = if ($Operation -in @('Delete','Class','Teacher')) { 'User.ReadWrite.All' } else { 'User.EnableDisableAccount.All' }
    $requiredScopes = @($writeScope, 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All')
    if ($Operation -eq 'Block') { $requiredScopes += 'User.RevokeSessions.All' }
    if ($Operation -eq 'Class') { $requiredScopes += 'GroupMember.ReadWrite.All' }
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
        [Parameter(Mandatory)][string] $TenantId,
        [System.Collections.IDictionary] $PersonalConfig,
        [switch] $AllowGuest
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
    $selectedType = [string](Get-ComparisonPropertyValue $Account UserType)
    $guestAllowed = $AllowGuest -and $null -ne $PersonalConfig -and $selectedType -eq 'Guest' -and $user.UserType -eq 'Guest'
    if (($user.UserType -ne 'Member' -and -not $guestAllowed) -or ($selectedType -and $selectedType -ne $user.UserType)) { throw 'Kontotyp ist nicht mehr zulässig.' }
    if ($user.OnPremisesSyncEnabled -eq $true) { throw 'Konto wird aus lokalem AD synchronisiert.' }

    $selectedRecipientType = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientType')
    $selectedRecipientStatus = [string](Get-ComparisonPropertyValue -InputObject $Account -Name 'RecipientStatus')
    $allowedRecipientTypes = @('UserMailbox','RemoteUserMailbox')
    if ($guestAllowed) { $allowedRecipientTypes += @('MailUser','GuestMailUser') }
    if ($selectedRecipientStatus -eq 'None' -and [string]::IsNullOrWhiteSpace($selectedRecipientType)) {
        $recipients = @(Get-EXORecipient -ResultSize Unlimited -Properties @('ExternalDirectoryObjectId', 'RecipientTypeDetails') -ErrorAction Stop |
            Where-Object { ([string](Get-ComparisonPropertyValue -InputObject $_ -Name 'ExternalDirectoryObjectId')).Trim() -ieq $userId })
        if ($recipients.Count -gt 0) { throw 'Seit der Auswahl ist ein Exchange-Empfänger vorhanden. Konto erneut prüfen.' }
    } elseif ($selectedRecipientStatus -in @('', 'Unique') -and $selectedRecipientType -in $allowedRecipientTypes) {
        $recipients = @(Get-EXORecipient -ExternalDirectoryObjectId ([guid]$userId) `
            -Properties @('ExternalDirectoryObjectId', 'RecipientTypeDetails') -ErrorAction Stop)
        if ($recipients.Count -ne 1) { throw 'Exchange-Empfänger ist nicht eindeutig vorhanden.' }
        $recipientId = [string](Get-ComparisonPropertyValue -InputObject $recipients[0] -Name 'ExternalDirectoryObjectId')
        if ($recipientId -ine $userId) { throw 'Exchange-Empfänger passt nicht zur Entra-Objekt-ID.' }
        $recipientType = [string](Get-ComparisonPropertyValue -InputObject $recipients[0] -Name 'RecipientTypeDetails')
        if ($recipientType -notin $allowedRecipientTypes -or $recipientType -ne $selectedRecipientType) { throw "Exchange-Empfängertyp ist '$recipientType'." }
    } else {
        throw 'Exchange-Empfänger aus der Auswahl ist nicht eindeutig als Benutzerpostfach oder fehlend bestätigt.'
    }

    $memberships = @(Get-MgUserTransitiveMemberOf -UserId $userId -All -ErrorAction Stop)
    foreach ($member in $memberships) {
        $type = [string](Get-DisabledAccountMemberValue -Member $member -Name '@odata.type')
        if ([string]::IsNullOrWhiteSpace($type)) { throw 'Mitgliedschaftstyp nicht lesbar.' }
        if ($type -eq '#microsoft.graph.directoryRole') { throw 'Konto hat eine Verzeichnisrolle.' }
    }
    if ($null -ne $PersonalConfig) { $null = Assert-PersonalMutationTarget -UserId $userId -ExpectedUpn $Account.UPN -Config $PersonalConfig -Account $Account -AllowGuest:$guestAllowed }
    return $user
}

function Invoke-SelectedDisabledAccountDeletion {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object[]] $Accounts,
        [Parameter(Mandatory)][string] $TenantId,
        [System.Collections.IDictionary] $PersonalConfig
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
                $null = Assert-DisabledAccountActionTarget -Account $account -TenantId $TenantId -PersonalConfig $PersonalConfig -AllowGuest:($null -ne $PersonalConfig)
                $ownedObjects = @(Get-MgUserOwnedObject -UserId $userId -All -ErrorAction Stop)
                if ($ownedObjects.Count -gt 0) {
                    Write-Warning "$($account.UPN): Besitzt $($ownedObjects.Count) Entra-Objekt(e). Diese können nach der Löschung ohne Eigentümer bleiben."
                }
                if ($null -ne $PersonalConfig) { $null = Assert-PersonalMutationTarget -UserId $userId -ExpectedUpn $account.UPN -Config $PersonalConfig -Account $account -AllowGuest }
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
        [Parameter(Mandatory)][string] $TenantId,
        [System.Collections.IDictionary] $PersonalConfig
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
                $null = Assert-DisabledAccountActionTarget -Account $account -TenantId $TenantId -PersonalConfig $PersonalConfig
                if ($null -ne $PersonalConfig) { $null = Assert-PersonalMutationTarget -UserId $userId -ExpectedUpn $account.UPN -Config $PersonalConfig -Account $account }
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
        [Parameter(Mandatory)][string] $TenantId,
        [System.Collections.IDictionary] $PersonalConfig
    )

    $Accounts = @($Accounts | Where-Object { [string]$_.RecipientType -ne 'SharedMailbox' })
    $personalActions = [Collections.Generic.List[object]]::new()
    $cursor = 0
    $selectedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $message = ''
    $population = if ($null -ne $PersonalConfig -and $PersonalConfig.Contains('ManagementPopulation')) { [string]$PersonalConfig.ManagementPopulation } else { 'Personnel' }
    $title = switch ($population) { Teacher { 'Lehrer- und Ganztagskonten' } Student { 'Schülerkonten' } default { '' } }
    while ($true) {
        Show-DisabledAccountSelection -Accounts $Accounts -SelectedIds $selectedIds -Cursor $cursor -TenantId $TenantId -Message $message -PersonalMode:($null -ne $PersonalConfig) -Title $title -StudentMode:($population -eq 'Student')
        $message = ''
        $key = Read-DisabledAccountTuiKey
        if ($null -ne $PersonalConfig -and ($key -in @('B','U','Backspace','Delete') -or ($population -eq 'Student' -and $key -in @('K','L')))) {
            $selectedAccounts = @($Accounts | Where-Object { $selectedIds.Contains([string]$_.UserId) })
            if ($selectedAccounts.Count -eq 0) { $message = 'Markiere zuerst mindestens ein Konto.'; continue }
            $operation = switch ($key) { 'B' { 'Block' } 'U' { 'Unlock' } 'Backspace' { 'Delete' } 'Delete' { 'Delete' } 'K' { 'Class' } 'L' { 'Teacher' } }
            $batch = @(if ($operation -in @('Class','Teacher')) { Invoke-StudentTuiAction -Accounts $selectedAccounts -Config $PersonalConfig -Operation $operation -WhatIf:$WhatIfPreference }
                else { Invoke-PersonalTuiAction -Accounts $selectedAccounts -Config $PersonalConfig -Operation $operation -WhatIf:$WhatIfPreference })
            foreach ($action in $batch) { $personalActions.Add($action) }
            $message = @($batch | ForEach-Object { "$($_.Status): $($_.DisplayName) $($_.Message)" }) -join [Environment]::NewLine
            # Rejected or cancelled actions must leave the selection available for correction.
            if ($batch.Count -eq 0 -or @($batch | Where-Object Status -ne 'Skipped').Count -eq 0) {
                if ($batch.Count -eq 0) { $message = 'Aktion abgebrochen. Auswahl bleibt erhalten.' }
                continue
            }
            $inventory = Get-PersonalInventory -Population $population
            if ($inventory.TenantId -ine $TenantId) { throw 'Tenant hat sich während der Kontenverwaltung geändert.' }
            $Accounts = @($inventory.Accounts); $PersonalConfig = $inventory.Config
            $selectedIds.Clear(); $cursor = 0
            continue
        }
        switch ($key) {
            'Escape' { if ($null -ne $PersonalConfig) { return @($personalActions) }; return }
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
            { $_ -in @('Backspace','Delete') } {
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
