function Resolve-UpdateSelection {
    param([switch] $Update, [switch] $CreateNewUsers, [switch] $UpdateUsers, [switch] $DisableUsers, [switch] $RevokeSessions)
    $selective = $CreateNewUsers -or $UpdateUsers -or $DisableUsers -or $RevokeSessions
    if ($selective -and -not $Update) { throw 'Aktionsschalter erfordern -Update.' }
    $all = $Update -and -not $selective
    [pscustomobject]@{
        CreateNewUsers = [bool]($all -or $CreateNewUsers); UpdateUsers = [bool]($all -or $UpdateUsers)
        DisableUsers = [bool]($all -or $DisableUsers); RevokeSessions = [bool]($all -or $RevokeSessions)
        ConfigureAllActiveExchange = [bool]$all
    }
}

function Assert-NewEntraStudentAttribute {
    param([Parameter(Mandatory)][string] $UserId, [Parameter(Mandatory)][object] $Desired)
    $fields = @('DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'Mail', 'MailNickname', 'Department', 'OfficeLocation', 'CompanyName', 'EmployeeType', 'UsageLocation', 'AgeGroup', 'ConsentProvidedForMinor')
    $user = Get-MgUser -UserId $UserId -Property @($fields + 'AccountEnabled') -ErrorAction Stop
    if ($null -eq $user.AccountEnabled -or [bool]$user.AccountEnabled) { throw "Neues Konto '$UserId' ist nicht nachweislich deaktiviert." }
    foreach ($field in $fields) {
        if (-not (Test-EntraMutationValueEqual -Field $field -Current (Get-ComparisonPropertyValue $user $field) -Desired (Get-ComparisonPropertyValue $Desired $field))) {
            throw "Attribut '$field' konnte für das neue Konto '$UserId' nicht verifiziert werden."
        }
    }
    return $true
}

function Get-StudentGroupParameter {
    param([Parameter(Mandatory)][object] $Entry, [Parameter(Mandatory)][object] $Snapshot, [Parameter(Mandatory)][System.Collections.IDictionary] $Config)
    $names = $Entry.DesiredState.RequiredGroupNames
    @{
        RoleGroup = $Snapshot.GroupsByDisplayName[$names[0]]
        LicenseGroup = $Snapshot.GroupsByDisplayName[$names[1]]
        ClassGroup = $Snapshot.GroupsByDisplayName[$names[2]]
        RoleGroupPrefix = $Config.RoleGroupPrefix
        ClassGroupPrefix = $Config.ClassGroupPrefix
    }
}

function Test-StudentUpnRename {
    param([Parameter(Mandatory)][object] $Entry)
    return @($Entry.Differences | Where-Object {
        $_.Area -eq 'Entra' -and $_.Field -eq 'UserPrincipalName' -and $_.Action -eq 'Set'
    }).Count -gt 0
}

function Get-EntraStudentCurrentIdentity {
    param([Parameter(Mandatory)][string] $UserId)
    $user = Get-MgUser -UserId $UserId -Property @('id', 'userPrincipalName', 'mail') -ErrorAction Stop
    if ([string]$user.Id -ine $UserId -or [string]::IsNullOrWhiteSpace([string]$user.UserPrincipalName)) {
        throw "Aktuelle UPN konnte für Entra-Objekt '$UserId' nicht verifiziert werden."
    }
    return $user
}

function Save-StudentIdentityCheckpoint {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object] $Entry,
        [Parameter(Mandatory)][string] $UserId,
        [Parameter(Mandatory)][string] $UserPrincipalName,
        [Parameter(Mandatory)][string] $Mail,
        [Parameter(Mandatory)][string] $File,
        [Parameter(Mandatory)][object] $WorkbookState
    )
    if (-not $PSCmdlet.ShouldProcess($File, "Persist stable identity for Entra object '$UserId'")) {
        throw 'Identitätssicherung wurde nicht bestätigt. Entra-Umbenennung wird nicht fortgesetzt.'
    }
    $identityUpdate = [pscustomobject]@{
        RowNumber = [int]$Entry.Student.RowNumber
        EntraObjectId = $UserId
        UPN = $UserPrincipalName
        Mail = $Mail
    }
    $written = Write-StudentWorkbookUpdate -Path $File -Updates @($identityUpdate) -ExpectedSourceHash $WorkbookState.SourceHash -Confirm:$false
    $verified = Read-StudentWorkbook -Path $File
    if ($verified.SourceHash -cne $written.SourceHash) { throw 'Die Schülerdatei wurde während der Identitätssicherung verändert.' }
    $rows = @($verified.Students | Where-Object RowNumber -eq $identityUpdate.RowNumber)
    $originalPassword = [string](Get-ComparisonPropertyValue $Entry.Student Password)
    if ($rows.Count -ne 1 -or $rows[0].EntraObjectId -ine $UserId -or $rows[0].StoredUpn -ine $UserPrincipalName -or
        $rows[0].StoredMail -ine $Mail -or $rows[0].Password -cne $originalPassword) {
        throw "Identitätssicherung konnte für Zeile $($identityUpdate.RowNumber) nicht verifiziert werden."
    }
    $WorkbookState.SourceHash = $verified.SourceHash
}

function Save-StudentWorkbookIdentityBatch {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [Parameter(Mandatory)][string] $File,
        [Parameter(Mandatory)][object] $WorkbookState,
        [object[]] $Secrets
    )

    $updates = [Collections.Generic.List[object]]::new()
    foreach ($entry in $Entries) {
        $userId = [string](Get-ComparisonPropertyValue -InputObject $entry.User -Name Id)
        try {
            $identity = Get-EntraStudentCurrentIdentity -UserId $userId
            $upn = ([string](Get-ComparisonPropertyValue -InputObject $identity -Name UserPrincipalName)).Trim()
            $mail = ([string](Get-ComparisonPropertyValue -InputObject $identity -Name Mail)).Trim()
            if ([string]::IsNullOrWhiteSpace($mail)) {
                throw "Aktuelle Mail-Adresse konnte für Entra-Objekt '$userId' nicht verifiziert werden."
            }
            $storedObjectId = ([string](Get-ComparisonPropertyValue -InputObject $entry.Student -Name EntraObjectId)).Trim()
            $storedUpn = ([string](Get-ComparisonPropertyValue -InputObject $entry.Student -Name StoredUpn)).Trim()
            $storedMail = ([string](Get-ComparisonPropertyValue -InputObject $entry.Student -Name StoredMail)).Trim()
            if ([string]::Equals($storedObjectId, $userId, [StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals($storedUpn, $upn, [StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals($storedMail, $mail, [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
            $updates.Add([pscustomobject]@{
                    RowNumber = [int]$entry.Student.RowNumber
                    EntraObjectId = $userId
                    UPN = $upn
                    Mail = $mail
                    OriginalPassword = [string](Get-ComparisonPropertyValue -InputObject $entry.Student -Name Password)
                })
        } catch {
            New-StudentActionResult -UserId $userId -Phase WorkbookIdentity -Status Failed -Message $_.Exception.Message `
                -Secrets $Secrets -RecoveryCommand (Get-StudentRecoveryCommand -File $File)
        }
    }
    if ($updates.Count -eq 0) { return }

    if (-not $PSCmdlet.ShouldProcess($File, 'Persist current Entra identities for matched students')) {
        foreach ($update in $updates) {
            New-StudentActionResult -UserId $update.EntraObjectId -UserPrincipalName $update.UPN -Phase WorkbookIdentity `
                -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
        }
        return
    }

    try {
        $written = Write-StudentWorkbookUpdate -Path $File -Updates @($updates) -ExpectedSourceHash $WorkbookState.SourceHash -Confirm:$false
        $verified = Read-StudentWorkbook -Path $File
        if ($verified.SourceHash -cne $written.SourceHash) {
            throw 'Die Schülerdatei wurde während der Identitätsrückschreibung verändert.'
        }
        foreach ($update in $updates) {
            $rows = @($verified.Students | Where-Object RowNumber -eq $update.RowNumber)
            if ($rows.Count -ne 1 -or $rows[0].EntraObjectId -ine $update.EntraObjectId -or
                $rows[0].StoredUpn -ine $update.UPN -or $rows[0].StoredMail -ine $update.Mail -or
                $rows[0].Password -cne $update.OriginalPassword) {
                throw "Identitätsrückschreibung konnte für Zeile $($update.RowNumber) nicht verifiziert werden."
            }
        }
        $WorkbookState.SourceHash = $verified.SourceHash
        foreach ($update in $updates) {
            New-StudentActionResult -UserId $update.EntraObjectId -UserPrincipalName $update.UPN -Phase WorkbookIdentity -Status Succeeded
        }
    } catch {
        foreach ($update in $updates) {
            New-StudentActionResult -UserId $update.EntraObjectId -UserPrincipalName $update.UPN -Phase WorkbookIdentity -Status Failed `
                -Message $_.Exception.Message -Secrets $Secrets -RecoveryCommand (Get-StudentRecoveryCommand -File $File)
        }
    }
}

function Invoke-StudentCreateBatch {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [Parameter(Mandatory)][string] $File,
        [Parameter(Mandatory)][object] $WorkbookState,
        [AllowEmptyCollection()][Collections.Generic.HashSet[string]] $UsedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    )
    $ErrorActionPreference = 'Stop'
    $pending = [Collections.Generic.List[object]]::new()
    $recovery = Get-StudentRecoveryCommand -File $File
    foreach ($entry in $Entries) {
        $upn = $entry.DesiredState.UserPrincipalName
        $userId = ''
        $phase = 'Create'
        if (-not $PSCmdlet.ShouldProcess($upn, 'Create, configure and persist new student')) {
            New-StudentActionResult -UserPrincipalName $upn -Phase Create -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
            continue
        }
        try {
            Assert-StudentWorkbookVersion -Path $File -ExpectedSourceHash $WorkbookState.SourceHash
            $initialPassword = New-StudentPassword -UsedPasswords $UsedPasswords
            [void]$UsedPasswords.Add($initialPassword)
            $created = New-StudentWithPasswordRetry -Desired $entry.DesiredState -InitialPassword $initialPassword -UsedPasswords $UsedPasswords -Confirm:$false
            if ($null -eq $created) { throw 'Neuanlage wurde nicht bestätigt.' }
            [void]$UsedPasswords.Add($created.Password)
            $userId = [string]$created.User.Id
            if ([string]::IsNullOrWhiteSpace($userId)) { throw 'Neuanlage lieferte keine Objekt-ID.' }
            $phase = 'Attributes'
            $attributesVerified = Assert-NewEntraStudentAttribute -UserId $userId -Desired $entry.DesiredState
            $phase = 'Manager'
            $manager = Set-EntraStudentManager -UserId $userId -DesiredManagerId $entry.DesiredState.ManagerId -Confirm:$false
            if (-not $manager.Verified) { throw 'Manager wurde nicht verifiziert.' }
            $phase = 'Groups'
            $groupParameters = Get-StudentGroupParameter -Entry $entry -Snapshot $Snapshot -Config $Config
            $groups = Sync-EntraStudentGroup -UserId $userId @groupParameters -CurrentDirectGroups @() -Confirm:$false
            if (-not $groups.Verified) { throw 'Pflichtgruppen wurden nicht verifiziert.' }
            $pending.Add([pscustomobject]@{
                RowNumber = [int]$entry.Student.RowNumber; Password = $created.Password; EntraObjectId = $userId; UPN = $upn; Mail = $upn
                AttributesVerified = [bool]$attributesVerified; ManagerVerified = [bool]$manager.Verified; GroupsVerified = [bool]$groups.Verified
            })
        } catch {
            New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase $phase -Status Failed -Secrets @($UsedPasswords) `
                -Message "$($_.Exception.Message) Neu angelegte Konten bleiben deaktiviert. Objekt-ID der Excel-Zeile zuordnen und fehlende Schritte sowie Passwortwiederherstellung administrativ prüfen." -RecoveryCommand $recovery
        } finally {
            $initialPassword = $null
            $created = $null
        }
    }
    if ($pending.Count -eq 0) { return }
    try {
        if (-not $PSCmdlet.ShouldProcess($File, 'Persist new student credentials and identities with a recoverable workbook commit')) {
            throw 'Excel-Rückschreibung wurde nicht bestätigt. Neue Konten bleiben deaktiviert.'
        }
        $written = Write-StudentWorkbookUpdate -Path $File -Updates @($pending) -ExpectedSourceHash $WorkbookState.SourceHash -Confirm:$false
        $verifiedWorkbook = Read-StudentWorkbook -Path $File
        if ($verifiedWorkbook.SourceHash -cne $written.SourceHash) { throw 'Die Schülerdatei wurde während der Passwort-Rückschreibung verändert.' }
        # Verify the entire batch before enabling any account.
        foreach ($candidate in $pending) {
            $rows = @($verifiedWorkbook.Students | Where-Object RowNumber -eq $candidate.RowNumber)
            if ($rows.Count -ne 1 -or $rows[0].EntraObjectId -ne $candidate.EntraObjectId -or $rows[0].StoredUpn -ine $candidate.UPN -or
                $rows[0].StoredMail -ine $candidate.Mail -or $rows[0].Password -cne $candidate.Password) {
                throw "Excel-Verifikation fehlgeschlagen für Zeile $($candidate.RowNumber)."
            }
        }
        $WorkbookState.SourceHash = $verifiedWorkbook.SourceHash
    } catch {
        foreach ($candidate in $pending) {
            New-StudentActionResult -UserId $candidate.EntraObjectId -UserPrincipalName $candidate.UPN -Phase Workbook -Status Failed `
                -Message "$($_.Exception.Message) Konto bleibt deaktiviert. Excel-Zuordnung und Passwort administrativ wiederherstellen, dann Pflichtzustand prüfen und aktivieren." -Secrets @($UsedPasswords) -RecoveryCommand $recovery
        }
        return
    }
    foreach ($candidate in $pending) {
        try {
            $enabled = Enable-EntraStudent -UserId $candidate.EntraObjectId -WorkbookVerified -GroupsVerified:$candidate.GroupsVerified `
                -ManagerVerified:$candidate.ManagerVerified -AttributesVerified:$candidate.AttributesVerified -Confirm:$false
            if (-not $enabled.Verified) { throw 'Aktivierung wurde nicht verifiziert.' }
            New-StudentActionResult -UserId $candidate.EntraObjectId -UserPrincipalName $candidate.UPN -Phase Create -Status Succeeded
        } catch {
            New-StudentActionResult -UserId $candidate.EntraObjectId -UserPrincipalName $candidate.UPN -Phase Enable -Status Failed `
                -Message $_.Exception.Message -Secrets @($UsedPasswords) -RecoveryCommand $recovery
        }
        $candidate.Password = $null
    }
}

function Invoke-StudentUpdate {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [Parameter(Mandatory)][string] $File,
        [Parameter(Mandatory)][object] $WorkbookState,
        [object[]] $Secrets
    )
    $ErrorActionPreference = 'Stop'
    foreach ($entry in $Entries) {
        $userId = [string]$entry.User.Id
        $upn = [string]$entry.User.UserPrincipalName
        $phase = 'Update'
        if (-not $PSCmdlet.ShouldProcess($upn, 'Apply listed student differences')) {
            New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Update -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
            continue
        }
        try {
            Assert-StudentWorkbookVersion -Path $File -ExpectedSourceHash $WorkbookState.SourceHash
            $renamesUpn = Test-StudentUpnRename -Entry $entry
            if ($renamesUpn) {
                $phase = 'WorkbookIdentity'
                # Persist the stable ID before Graph can partially apply a rename.
                Save-StudentIdentityCheckpoint -Entry $entry -UserId $userId -UserPrincipalName $upn `
                    -Mail ([string]$entry.User.Mail) -File $File -WorkbookState $WorkbookState -Confirm:$false
            }
            $phase = 'Attributes'
            try {
                $attributes = Set-EntraStudentAttribute -UserId $userId -Desired $entry.DesiredState -Differences $entry.Differences -Confirm:$false
            } finally {
                if ($renamesUpn) {
                    $phase = 'WorkbookIdentity'
                    # Graph may have accepted the rename even when its verification failed.
                    $currentIdentity = Get-EntraStudentCurrentIdentity -UserId $userId
                    $upn = [string]$currentIdentity.UserPrincipalName
                    Save-StudentIdentityCheckpoint -Entry $entry -UserId $userId -UserPrincipalName $upn `
                        -Mail ([string]$currentIdentity.Mail) -File $File -WorkbookState $WorkbookState -Confirm:$false
                    $phase = 'Attributes'
                }
            }
            if (-not $attributes.Verified) { throw 'Attribute wurden nicht verifiziert.' }
            $upn = [string]$entry.DesiredState.UserPrincipalName
            if (@($entry.Differences | Where-Object Area -eq Manager).Count -gt 0) {
                $phase = 'Manager'
                $manager = Set-EntraStudentManager -UserId $userId -CurrentManagerId (Get-UserManagerId -Snapshot $Snapshot -UserId $userId) -DesiredManagerId $entry.DesiredState.ManagerId -Confirm:$false
                if (-not $manager.Verified) { throw 'Manager wurde nicht verifiziert.' }
            }
            if (@($entry.Differences | Where-Object Area -eq Group).Count -gt 0) {
                $phase = 'Groups'
                $groupParameters = Get-StudentGroupParameter -Entry $entry -Snapshot $Snapshot -Config $Config
                $groups = Sync-EntraStudentGroup -UserId $userId @groupParameters -CurrentDirectGroups @(Get-UserDirectGroup -Snapshot $Snapshot -UserId $userId) -Confirm:$false
                if (-not $groups.Verified) { throw 'Pflichtgruppen wurden nicht verifiziert.' }
            }
            New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Update -Status Succeeded
        } catch {
            New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message -Secrets $Secrets -RecoveryCommand (Get-StudentRecoveryCommand -File $File)
        }
    }
}

function Invoke-StudentDeparture {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [switch] $DisableUsers, [switch] $RevokeSessions,
        [string] $File, [object[]] $Secrets,
        [string] $RecoveryCommand
    )
    $ErrorActionPreference = 'Stop'
    $recovery = if (-not [string]::IsNullOrWhiteSpace($RecoveryCommand)) {
        $RecoveryCommand
    } elseif (-not [string]::IsNullOrWhiteSpace($File)) {
        Get-StudentRecoveryCommand -File $File
    } else { '' }
    foreach ($entry in $Entries) {
        foreach ($phase in 'Disable', 'RevokeSessions') {
            if (($phase -eq 'Disable' -and -not $DisableUsers) -or ($phase -eq 'RevokeSessions' -and -not $RevokeSessions)) { continue }
            $id = [string]$entry.User.Id
            $upn = [string]$entry.User.UserPrincipalName
            if ($phase -eq 'Disable' -and $entry.User.AccountEnabled -eq $false) {
                New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Compliant
                continue
            }
            if (-not $PSCmdlet.ShouldProcess($upn, $phase)) {
                New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
                continue
            }
            try {
                if ($phase -eq 'Disable') {
                    $result = Disable-EntraStudent -UserId $id -Confirm:$false
                    if (-not $result.Verified) { throw 'Deaktivierung wurde nicht verifiziert.' }
                } else {
                    $result = Revoke-EntraStudentSession -UserId $id -Selected -Confirm:$false
                    $success = if ($result -is [bool]) { $result } else { Get-ComparisonPropertyValue $result Value }
                    if ($success -ne $true) { throw 'Graph hat den Sitzungswiderruf nicht bestätigt.' }
                }
                New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Succeeded
            } catch {
                New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message -Secrets $Secrets -RecoveryCommand $recovery
            }
        }
    }
}
