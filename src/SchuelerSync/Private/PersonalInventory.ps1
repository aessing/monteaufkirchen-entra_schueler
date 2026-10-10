function Get-PersonalConfig {
    Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/PersonalSync.psd1')
}

function Get-PersonalProtectionContext {
    param([System.Collections.IDictionary] $GroupsById, [System.Collections.IDictionary] $Config,
        [ValidateSet('Personnel','Teacher','Student')][string] $Population = 'Personnel')
    $teachers = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
    $excluded = @($teachers.Roles.School, $teachers.Roles.Ganztag, (Get-ConfiguredStudentRole))
    foreach ($definition in $excluded) {
        $group = $GroupsById[[string]$definition.Id]
        if ($null -eq $group -or [string]$group.DisplayName -cne [string]$definition.Name) { throw "Ausschlussgruppe '$($definition.Name)' nicht vollständig bestätigt." }
    }
    $protected = @()
    $warnings = @()
    foreach ($name in $Config.ProtectedGroupNames) {
        $profileMatches = @($GroupsById.Values | Where-Object DisplayName -eq $name)
        if ($profileMatches.Count -ne 1 -or ($name -eq $Config.Roles.ExterneAdmins.Name -and $profileMatches[0].Id -ine $Config.Roles.ExterneAdmins.Id)) {
            $warnings += "Schutzgruppe '$name' fehlt oder ist nicht eindeutig bestätigt. Schreibaktionen sind gesperrt."
        } else { $protected += [string]$profileMatches[0].Id }
    }
    [pscustomobject]@{
        ExcludedIds = switch ($Population) { Teacher { @([string]$excluded[2].Id) } Student { @($teachers.Roles.School.Id,$teachers.Roles.Ganztag.Id) } default { @($excluded | ForEach-Object { [string]$_.Id }) } }
        RequiredRoleIds = switch ($Population) { Teacher { @($teachers.Roles.School.Id,$teachers.Roles.Ganztag.Id) } Student { @([string]$excluded[2].Id) } default { @() } }
        ProtectedIds = @($protected)
        ProtectedNames = @($Config.ProtectedGroupNames)
        ManagedRoleIds = switch ($Population) { Teacher { @($teachers.Roles.School.Id,$teachers.Roles.Ganztag.Id) } Student { @([string]$excluded[2].Id) } default { @($Config.Roles.Values | Where-Object { $_.Name -notin $Config.ProtectedGroupNames } | ForEach-Object { [string]$_.Id }) } }
        Warnings = @($warnings)
        Complete = ($warnings.Count -eq 0)
    }
}

function Get-PersonalMembershipSafety {
    param([AllowEmptyCollection()][object[]] $Memberships, [object] $Context)
    $excluded = $false; $reasons = @()
    foreach ($member in $Memberships) {
        if ($null -eq $member) { $reasons += 'Mitgliedschaft nicht vollständig lesbar'; continue }
        $type = [string](Get-DisabledAccountMemberValue -Member $member -Name '@odata.type')
        $id = [string](Get-DisabledAccountMemberValue -Member $member -Name id)
        $name = [string](Get-DisabledAccountMemberValue -Member $member -Name displayName)
        if (-not $type -or -not $id) { $reasons += 'Mitgliedschaft nicht vollständig lesbar'; continue }
        if ($id -in $Context.ExcludedIds) { $excluded = $true }
        if ($id -in $Context.ProtectedIds -or $name -in $Context.ProtectedNames) { $reasons += "Geschützte Gruppe: $name" }
        if ($type -eq '#microsoft.graph.directoryRole') { $reasons += "Verzeichnisrolle: $name" }
    }
    [pscustomobject]@{ Excluded = $excluded; Reasons = @($reasons) }
}

function Get-AccountRecipientIndex {
    param([AllowEmptyCollection()][object[]] $Recipients)
    $byId = New-CaseInsensitiveHashtable
    $duplicateIds = New-CaseInsensitiveHashtable
    $sharedIds = New-CaseInsensitiveHashtable
    foreach ($recipient in $Recipients) {
        $id = ([string](Get-ComparisonPropertyValue $recipient ExternalDirectoryObjectId)).Trim()
        if (-not $id) { continue }
        if ([string](Get-ComparisonPropertyValue $recipient RecipientTypeDetails) -eq 'SharedMailbox') { $sharedIds[$id] = $true }
        if ($byId.ContainsKey($id)) { $duplicateIds[$id] = $true } else { $byId[$id] = $recipient }
    }
    [pscustomobject]@{ ById = $byId; DuplicateIds = $duplicateIds; SharedIds = $sharedIds }
}

function Assert-PersonalSnapshotTarget {
    param([object] $User, [object] $Snapshot)
    if (-not $Snapshot.PersonalContext.Complete) { throw ($Snapshot.PersonalContext.Warnings -join ' ') }
    $memberships = @(Get-MgUserTransitiveMemberOf -UserId $User.Id -All -ErrorAction Stop)
    $safety = Get-PersonalMembershipSafety -Memberships $memberships -Context $Snapshot.PersonalContext
    if ($safety.Excluded -or $safety.Reasons.Count -gt 0 -or $User.UserType -ne 'Member' -or
        (Get-ComparisonPropertyValue $User OnPremisesSyncEnabled) -eq $true -or $null -eq $User.AccountEnabled) {
        throw "Konto '$($User.UserPrincipalName)' ist ausgeschlossen, geschützt oder nicht vollständig prüfbar. $($safety.Reasons -join ', ')"
    }
    if ($Snapshot.RecipientIndex.DuplicateIds.ContainsKey([string]$User.Id)) { throw 'Exchange-Zuordnung ist mehrdeutig.' }
    $recipient = $Snapshot.RecipientIndex.ById[[string]$User.Id]
    if ($null -ne $recipient -and [string](Get-ComparisonPropertyValue $recipient RecipientTypeDetails) -notin @('UserMailbox','RemoteUserMailbox')) {
        throw 'Exchange-Empfänger ist kein persönliches Postfach.'
    }
}

function Assert-PersonalMutationTarget {
    param([string] $UserId, [string] $ExpectedUpn, [System.Collections.IDictionary] $Config, [object] $Account, [string[]] $AllowedRoleIds = @(), [switch] $AllowGuest)
    $context = $Config.SafetyContext
    if (-not $context.Complete) { throw ($context.Warnings -join ' ') }
    $graph = Get-MgContext
    if ([string]$graph.TenantId -ine [string]$Config.SessionTenantId) { throw 'Graph-Tenant hat sich geändert.' }
    $null = Assert-ExchangeTenant -TenantId $Config.SessionTenantId
    $user = Get-MgUser -UserId $UserId -Property @('id','displayName','givenName','surname','userPrincipalName','accountEnabled','userType','onPremisesSyncEnabled','companyName','employeeType','department','officeLocation') -ErrorAction Stop
    # Guest permission is confined to selected lifecycle actions, never profile maintenance.
    $selectedType = [string](Get-ComparisonPropertyValue $Account UserType)
    $guestAllowed = $AllowGuest -and $selectedType -eq 'Guest' -and $user.UserType -eq 'Guest' -and
        (-not $Config.Contains('ManagementPopulation') -or $Config.ManagementPopulation -ne 'Student')
    if ([string]$user.Id -ine $UserId -or [string]$user.UserPrincipalName -ine $ExpectedUpn -or
        ($user.UserType -ne 'Member' -and -not $guestAllowed) -or
        ($selectedType -and $selectedType -ne $user.UserType) -or
        $user.OnPremisesSyncEnabled -eq $true -or $null -eq $user.AccountEnabled) { throw 'Identität oder Kontotyp ist nicht mehr zulässig.' }
    if ($null -ne $Account -and [string]$user.DisplayName -cne [string]$Account.DisplayName) { throw 'Name hat sich seit der Auswahl geändert.' }
    $memberships = @(Get-MgUserTransitiveMemberOf -UserId $UserId -All -ErrorAction Stop)
    $safety = Get-PersonalMembershipSafety -Memberships $memberships -Context $context
    if ($safety.Excluded -or $safety.Reasons.Count -gt 0) { throw "Konto ist ausgeschlossen oder geschützt. $($safety.Reasons -join ', ')" }
    $direct = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    $requiredRoleIds = @(Get-ComparisonPropertyValue $context RequiredRoleIds | Where-Object { $_ })
    if ($requiredRoleIds.Count -gt 0 -and @($direct | Where-Object { $_.Id -in $requiredRoleIds }).Count -eq 0) { throw 'Konto gehört keiner für diesen Einstieg zulässigen direkten Rollengruppe mehr an.' }
    if (@($direct | Where-Object { $_.Id -in $context.ManagedRoleIds -and $_.Id -notin $AllowedRoleIds }).Count -gt $(if ($AllowedRoleIds.Count -gt 0) { 0 } else { 1 })) { throw 'Mehrere konkurrierende Personalprofile.' }
    $recipients = @(Get-EXORecipient -ResultSize Unlimited -Properties @('ExternalDirectoryObjectId','RecipientTypeDetails') -ErrorAction Stop |
        Where-Object { ([string](Get-ComparisonPropertyValue $_ ExternalDirectoryObjectId)).Trim() -ieq $UserId })
    $allowedRecipientTypes = @('UserMailbox','RemoteUserMailbox')
    if ($guestAllowed) { $allowedRecipientTypes += @('MailUser','GuestMailUser') }
    if ($recipients.Count -gt 1 -or ($recipients.Count -eq 1 -and [string](Get-ComparisonPropertyValue $recipients[0] RecipientTypeDetails) -notin $allowedRecipientTypes)) { throw 'Exchange-Empfänger ist nicht eindeutig als zulässiger Kontotyp bestätigt.' }
    if ($null -ne $Account -and $Account.RecipientStatus -eq 'Unique' -and
        ($recipients.Count -ne 1 -or [string](Get-ComparisonPropertyValue $recipients[0] RecipientTypeDetails) -ne $Account.RecipientType)) { throw 'Exchange-Zuordnung hat sich seit der Auswahl geändert.' }
    if ($null -ne $Account -and $Account.RecipientStatus -eq 'None' -and $recipients.Count -gt 0) { throw 'Seit der Auswahl ist ein Exchange-Empfänger vorhanden. Erneut prüfen.' }
    return $user
}

function Get-PersonalInventory {
    [CmdletBinding()]
    param([ValidateSet('Personnel','Teacher','Student')][string] $Population = 'Personnel')
    $activity = switch ($Population) { Teacher {'Lehrer- und Ganztagskonten prüfen'} Student {'Schülerkonten prüfen'} default {'Personalkonten prüfen'} }
    try {
        Write-Progress -Id 1 -Activity $activity -Status 'Verbinde mit Microsoft Graph ...' -PercentComplete 5
        $graph = Connect-DisabledAccountGraph
        $tenantId = [string]$graph.TenantId
        if (-not $tenantId) { throw 'Graph-Kontext enthält keine Tenant-ID.' }
        $config = Get-PersonalConfig
        if ($Population -eq 'Teacher') {
            $teacherConfig = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
            $config.ExpectedTenantId = $teacherConfig.ExpectedTenantId
        }
        if ($Population -eq 'Student') {
            $config.StudentSettings = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/SchuelerSync.psd1')
            $config.ExpectedTenantId = $config.StudentSettings.ExpectedTenantId
        }
        if ($config.ExpectedTenantId -and $tenantId -ine $config.ExpectedTenantId) { throw 'Graph-Tenant stimmt nicht mit der Konfiguration überein.' }
        Write-Progress -Id 1 -Activity $activity -Status 'Verbinde mit Exchange Online ...' -PercentComplete 15
        $null = Connect-SchuelerExchangeOnline -TenantId $tenantId
        $groups = @(Get-MgGroup -All -Property @('id','displayName','assignedLicenses','groupTypes','membershipRule') -ErrorAction Stop)
        $byId = New-CaseInsensitiveHashtable
        foreach ($group in $groups) {
            if (-not $group.Id -or $byId.ContainsKey([string]$group.Id)) { throw 'Ungültige oder doppelte Gruppen-ID.' }
            $byId[[string]$group.Id] = $group
        }
        $context = Get-PersonalProtectionContext -GroupsById $byId -Config $config -Population $Population
        $warnings = [Collections.Generic.List[string]]::new()
        foreach ($warning in $context.Warnings) { $warnings.Add($warning) }
        $fields = @('id','displayName','givenName','surname','userPrincipalName','mail','jobTitle','department','officeLocation','accountEnabled','userType','onPremisesSyncEnabled','assignedLicenses','licenseAssignmentStates')
        $lastLoginAvailable = $true
        try { $users = @(Get-MgUser -All -Property ($fields + 'signInActivity') -ErrorAction Stop) }
        catch { $lastLoginAvailable = $false; $warnings.Add("Letzte Anmeldung nicht abrufbar: $($_.Exception.Message)"); $users = @(Get-MgUser -All -Property $fields -ErrorAction Stop) }
        $usersById = New-CaseInsensitiveHashtable
        foreach ($user in $users) {
            if (-not $user.Id -or $usersById.ContainsKey([string]$user.Id)) { throw 'Benutzerbestand enthält fehlende oder doppelte Objekt-IDs.' }
            $usersById[[string]$user.Id] = $user
        }
        $index = Get-AccountRecipientIndex -Recipients @(Get-EXORecipient -ResultSize Unlimited -Properties @('ExternalDirectoryObjectId','RecipientTypeDetails') -ErrorAction Stop)
        $records = [Collections.Generic.List[object]]::new(); $excludedCount = 0; $counter = 0
        foreach ($user in @($users | Sort-Object DisplayName,UserPrincipalName)) {
            $counter++
            Write-Progress -Id 1 -Activity $activity -Status "Prüfe Konto $counter von $($users.Count): $($user.DisplayName)" -PercentComplete (30 + [int](60 * $counter / [math]::Max(1,$users.Count)))
            if (-not $user.Id -or $null -eq $user.AccountEnabled) { throw 'Benutzerbestand ist nicht vollständig lesbar.' }
            $memberships = @(Get-MgUserTransitiveMemberOf -UserId $user.Id -All -ErrorAction Stop)
            $safety = Get-PersonalMembershipSafety -Memberships $memberships -Context $context
            if (@($safety.Reasons | Where-Object { $_ -eq 'Mitgliedschaft nicht vollständig lesbar' }).Count -gt 0) { throw 'Ausschlussmitgliedschaften sind nicht vollständig lesbar.' }
            $recipient = $index.ById[[string]$user.Id]
            if ($safety.Excluded -or $index.SharedIds.ContainsKey([string]$user.Id)) { $excludedCount++; continue }
            $direct = @(Get-FreshEntraUserDirectGroup -UserId $user.Id)
            if ($Population -in @('Teacher','Student') -and @($direct | Where-Object { $_.Id -in $context.RequiredRoleIds }).Count -eq 0) { $excludedCount++; continue }
            $status = if ($index.DuplicateIds.ContainsKey([string]$user.Id)) { 'Ambiguous' } elseif ($null -eq $recipient) { 'None' } elseif (-not [string](Get-ComparisonPropertyValue $recipient RecipientTypeDetails)) { 'Unreadable' } else { 'Unique' }
            $issues = @($safety.Reasons)
            $owned = @()
            try { $owned = @(Get-MgUserOwnedObject -UserId $user.Id -All -ErrorAction Stop) }
            catch { $issues += 'Besitz von Entra-Objekten nicht prüfbar' }
            $jobs = if ($Population -in @('Teacher','Student')) { @($direct | Where-Object { $_.Id -in $context.RequiredRoleIds } | ForEach-Object { $_.Id }) }
                else { @($config.Roles.Keys | Where-Object { $config.Roles[$_].Id -in @($direct | ForEach-Object { $_.Id }) }) }
            $jobs = @($jobs)
            if ($jobs.Count -gt 1) { $issues += 'Mehrere konkurrierende Personalprofile' }
            $licenseReason = 'Nicht prüfbar'
            try { $licenseReason = (Test-LehrerLicenseState -User $user -DirectGroups @($direct | ForEach-Object { $byId[[string]$_.Id] })).Reason }
            catch { $issues += "Lizenzprüfung: $($_.Exception.Message)" }
            $record = ConvertTo-DisabledAccountCheckRecord -User $user -Memberships $memberships -OwnedObjects $owned -Recipient $recipient -RecipientStatus $status -CheckErrors $issues -LastLoginAvailable:$lastLoginAvailable -AllowActive
            $job = if ($Population -eq 'Teacher') { [string](Get-ComparisonPropertyValue $user JobTitle) } elseif ($jobs.Count -eq 1) { [string]$jobs[0] } else { '' }
            $allowedRecipientTypes = @('UserMailbox','RemoteUserMailbox')
            if ($user.UserType -eq 'Guest') { $allowedRecipientTypes += @('MailUser','GuestMailUser') }
            $allowed = $context.Complete -and $user.UserType -in @('Member','Guest') -and $user.OnPremisesSyncEnabled -ne $true -and $safety.Reasons.Count -eq 0 -and $jobs.Count -lt 2 -and $status -in @('None','Unique') -and ($status -eq 'None' -or $record.RecipientType -in $allowedRecipientTypes)
            if ($Population -eq 'Student' -and $user.UserType -ne 'Member') { $allowed = $false }
            $record | Add-Member -NotePropertyMembers @{
                AccountEnabled = $user.AccountEnabled; UserType = [string]$user.UserType; GivenName = [string]$user.GivenName; Surname = [string]$user.Surname
                Mail = [string]$user.Mail; Job = $job; LicenseReason = $licenseReason; MutationAllowed = [bool]$allowed
            }
            if ($Population -eq 'Student') {
                $managerAvailable = $true; $managerId = ''; $teacher = 'Keiner'
                try {
                    $managerId = [string](Get-UserManagerId -Snapshot ([pscustomobject]@{ManagerByUserId=(New-CaseInsensitiveHashtable)}) -UserId $user.Id)
                    if ($managerId) {
                        $manager = $usersById[$managerId]
                        $teacher = if ($null -ne $manager) { "$($manager.DisplayName) <$($manager.UserPrincipalName)>" } else { "Objekt-ID: $managerId" }
                    }
                } catch { $managerAvailable = $false; $teacher = 'Nicht abrufbar' }
                $record | Add-Member -NotePropertyMembers @{
                    ClassGroups = @($direct | Where-Object { $_.DisplayName.StartsWith($config.StudentSettings.ClassGroupPrefix,[StringComparison]::OrdinalIgnoreCase) })
                    Department = [string](Get-ComparisonPropertyValue $user Department); OfficeLocation = [string](Get-ComparisonPropertyValue $user OfficeLocation)
                    ManagerId = $managerId; ManagerAvailable = $managerAvailable; Teacher = $teacher
                }
            }
            if (-not $context.Complete) { $record.ProtectionReasons = @($record.ProtectionReasons) + @($context.Warnings) }
            $records.Add($record)
        }
        $config.SessionTenantId = $tenantId; $config.SafetyContext = $context; $config.ManagementPopulation = $Population
        [pscustomobject]@{ Mode = 'List'; TenantId = $tenantId; CheckedAt = [datetimeoffset]::UtcNow; Accounts = @($records); ExcludedCount = $excludedCount; Warnings = @($warnings); Config = $config }
    } finally { Write-Progress -Id 1 -Activity $activity -Completed }
}

function Write-PersonalInventoryReport {
    param([object] $Inventory, [string] $OutputFile)
    $text = "Personalkonten ($($Inventory.Accounts.Count)) | Ausgeschlossen: $($Inventory.ExcludedCount) | Tenant $($Inventory.TenantId)"
    Write-Host "`n$text" -ForegroundColor Cyan
    $lines = [Collections.Generic.List[string]]::new(); $lines.Add($text)
    foreach ($account in $Inventory.Accounts) {
        $row = [ordered]@{
            DisplayName = $account.DisplayName; UPN = $account.UPN; UserId = $account.UserId; AccountEnabled = $account.AccountEnabled
            Job = $account.Job; 'Exchange-Empfänger' = "$($account.RecipientType) ($($account.RecipientStatus))"
            'SEC-A Gruppen' = $account.SecAGroups -join ', '; 'Last Login' = if ($account.LastLogin) { $account.LastLogin.ToString('u') } else { $account.LastLoginStatus }
            Lizenzentscheidung = $account.LicenseReason
            Besitzobjekte = @($account.OwnedObjects | ForEach-Object { "$($_.Type): $($_.DisplayName) ($($_.Id))" }) -join ', '
            'Schutz/Prüfung' = $account.ProtectionReasons -join ', '
        }
        $line = ([pscustomobject]$row | Format-List | Out-String).TrimEnd()
        Write-Information $line -InformationAction Continue; $lines.Add(''); $lines.Add($line)
    }
    foreach ($warning in $Inventory.Warnings) { Write-Warning $warning; $lines.Add($warning) }
    if ($OutputFile -and -not $WhatIfPreference) { Write-ComparisonReportFile -Path $OutputFile -Text ($lines -join [Environment]::NewLine) -Confirm:$false }
}
