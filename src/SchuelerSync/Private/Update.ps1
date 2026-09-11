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

function Assert-NewEntraStudentAttributes {
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

function Get-StudentGroupParameters {
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

function Invoke-StudentCreateBatch {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [Parameter(Mandatory)][string] $File,
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
            $initialPassword = New-StudentPassword -UsedPasswords $UsedPasswords
            [void]$UsedPasswords.Add($initialPassword)
            $created = New-StudentWithPasswordRetry -Desired $entry.DesiredState -InitialPassword $initialPassword -UsedPasswords $UsedPasswords -Confirm:$false
            if ($null -eq $created) { throw 'Neuanlage wurde nicht bestätigt.' }
            [void]$UsedPasswords.Add($created.Password)
            $userId = [string]$created.User.Id
            if ([string]::IsNullOrWhiteSpace($userId)) { throw 'Neuanlage lieferte keine Objekt-ID.' }
            $phase = 'Attributes'
            $attributesVerified = Assert-NewEntraStudentAttributes -UserId $userId -Desired $entry.DesiredState
            $phase = 'Manager'
            $manager = Set-EntraStudentManager -UserId $userId -DesiredManagerId $entry.DesiredState.ManagerId -Confirm:$false
            if (-not $manager.Verified) { throw 'Manager wurde nicht verifiziert.' }
            $phase = 'Groups'
            $groupParameters = Get-StudentGroupParameters -Entry $entry -Snapshot $Snapshot -Config $Config
            $groups = Sync-EntraStudentGroups -UserId $userId @groupParameters -CurrentDirectGroups @() -Confirm:$false
            if (-not $groups.Verified) { throw 'Pflichtgruppen wurden nicht verifiziert.' }
            $pending.Add([pscustomobject]@{
                RowNumber = [int]$entry.Student.RowNumber; Password = $created.Password; EntraObjectId = $userId; UPN = $upn
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
        if (-not $PSCmdlet.ShouldProcess($File, 'Persist new student credentials and identities atomically')) {
            throw 'Excel-Rückschreibung wurde nicht bestätigt. Neue Konten bleiben deaktiviert.'
        }
        $null = Write-StudentWorkbookUpdates -Path $File -Updates @($pending) -Confirm:$false
        $verifiedWorkbook = Read-StudentWorkbook -Path $File
        # Verify the entire batch before enabling any account.
        foreach ($candidate in $pending) {
            $rows = @($verifiedWorkbook.Students | Where-Object RowNumber -eq $candidate.RowNumber)
            if ($rows.Count -ne 1 -or $rows[0].EntraObjectId -ne $candidate.EntraObjectId -or $rows[0].StoredUpn -ine $candidate.UPN -or $rows[0].Password -cne $candidate.Password) {
                throw "Excel-Verifikation fehlgeschlagen für Zeile $($candidate.RowNumber)."
            }
        }
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

function Invoke-StudentUpdates {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [Parameter(Mandatory)][string] $File,
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
            $phase = 'Attributes'
            $attributes = Set-EntraStudentAttributes -UserId $userId -Desired $entry.DesiredState -Differences $entry.Differences -Confirm:$false
            if (-not $attributes.Verified) { throw 'Attribute wurden nicht verifiziert.' }
            $upn = [string]$entry.DesiredState.UserPrincipalName
            if (@($entry.Differences | Where-Object Area -eq Manager).Count -gt 0) {
                $phase = 'Manager'
                $manager = Set-EntraStudentManager -UserId $userId -CurrentManagerId (Get-UserManagerId -Snapshot $Snapshot -UserId $userId) -DesiredManagerId $entry.DesiredState.ManagerId -Confirm:$false
                if (-not $manager.Verified) { throw 'Manager wurde nicht verifiziert.' }
            }
            if (@($entry.Differences | Where-Object Area -eq Group).Count -gt 0) {
                $phase = 'Groups'
                $groupParameters = Get-StudentGroupParameters -Entry $entry -Snapshot $Snapshot -Config $Config
                $groups = Sync-EntraStudentGroups -UserId $userId @groupParameters -CurrentDirectGroups @(Get-UserDirectGroups -Snapshot $Snapshot -UserId $userId) -Confirm:$false
                if (-not $groups.Verified) { throw 'Pflichtgruppen wurden nicht verifiziert.' }
            }
            New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Update -Status Succeeded
        } catch {
            New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message -Secrets $Secrets -RecoveryCommand (Get-StudentRecoveryCommand -File $File)
        }
    }
}

function Invoke-StudentDepartures {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Entries,
        [switch] $DisableUsers, [switch] $RevokeSessions,
        [Parameter(Mandatory)][string] $File, [object[]] $Secrets
    )
    $ErrorActionPreference = 'Stop'
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
                    $result = Revoke-EntraStudentSessions -UserId $id -Selected -Confirm:$false
                    $success = if ($result -is [bool]) { $result } else { Get-ComparisonPropertyValue $result Value }
                    if ($success -ne $true) { throw 'Graph hat den Sitzungswiderruf nicht bestätigt.' }
                }
                New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Succeeded
            } catch {
                New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message -Secrets $Secrets -RecoveryCommand (Get-StudentRecoveryCommand -File $File)
            }
        }
    }
}
