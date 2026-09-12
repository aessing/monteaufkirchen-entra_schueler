function Invoke-ManualStudentAdd {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object] $Entry,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config
    )

    $ErrorActionPreference = 'Stop'
    $upn = [string]$Entry.DesiredState.UserPrincipalName
    $userId = ''
    $phase = 'Add'
    $usedPasswords = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)

    if (-not $PSCmdlet.ShouldProcess($upn, 'Create, configure and enable one student')) {
        return New-StudentActionResult -UserPrincipalName $upn -Phase Add -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
    }

    try {
        $phase = 'Password'
        $initialPassword = New-StudentPassword -UsedPasswords $usedPasswords
        [void]$usedPasswords.Add($initialPassword)
        $created = New-StudentWithPasswordRetry -Desired $Entry.DesiredState -InitialPassword $initialPassword -UsedPasswords $usedPasswords -Confirm:$false
        if ($null -eq $created) { throw 'Neuanlage wurde nicht bestätigt.' }
        [void]$usedPasswords.Add([string]$created.Password)
        $userId = [string]$created.User.Id
        if ([string]::IsNullOrWhiteSpace($userId)) { throw 'Neuanlage lieferte keine Objekt-ID.' }

        $phase = 'Attributes'
        $attributesVerified = Assert-NewEntraStudentAttribute -UserId $userId -Desired $Entry.DesiredState

        $phase = 'Manager'
        $manager = Set-EntraStudentManager -UserId $userId -DesiredManagerId $Entry.DesiredState.ManagerId -Confirm:$false
        if (-not $manager.Verified) { throw 'Manager wurde nicht verifiziert.' }

        $phase = 'Groups'
        $groupParameters = Get-StudentGroupParameter -Entry $Entry -Snapshot $Snapshot -Config $Config
        $groups = Sync-EntraStudentGroup -UserId $userId @groupParameters -CurrentDirectGroups @() -Confirm:$false
        if (-not $groups.Verified) { throw 'Pflichtgruppen wurden nicht verifiziert.' }

        $phase = 'Enable'
        $enabled = Enable-EntraStudent -UserId $userId -WorkbookVerified -GroupsVerified:$groups.Verified `
            -ManagerVerified:$manager.Verified -AttributesVerified:$attributesVerified -Confirm:$false
        if (-not $enabled.Verified) { throw 'Aktivierung wurde nicht verifiziert.' }

        # Manual additions have no workbook credential escrow. Reveal the password only after every required state was verified.
        Write-Information "Einmaliges Startpasswort für $upn`: $($created.Password)" -InformationAction Continue
        return New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Add -Status Succeeded
    } catch {
        $recoveryGivenName = ([string](Get-ComparisonPropertyValue $Entry.Student GivenName)).Replace("'", "''")
        $recoverySurname = ([string](Get-ComparisonPropertyValue $Entry.Student Surname)).Replace("'", "''")
        $recoveryClass = ([string](Get-ComparisonPropertyValue $Entry.Student ClassName)).Replace("'", "''")
        $recoveryTeacher = ([string](Get-ComparisonPropertyValue $Entry.Student Teacher)).Replace("'", "''")
        $recovery = ".\Sync-SchuelerEntra.ps1 -Add -Vorname '$recoveryGivenName' -Nachname '$recoverySurname' -Klasse '$recoveryClass' -Klassenlehrer '$recoveryTeacher'"
        if (-not [string]::IsNullOrWhiteSpace($userId)) {
            $recovery += " -EntraObjectId '$($userId.Replace("'", "''"))'"
        }
        if ($null -ne $created -and -not [string]::IsNullOrWhiteSpace([string]$created.Password)) {
            Write-Information "Einmaliges Startpasswort für $upn`: $($created.Password)" -InformationAction Continue
        }
        return New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase $phase -Status Failed -Secrets @($usedPasswords) `
            -Message "$($_.Exception.Message) Der aktuelle Kontostatus ist unbekannt. Prüfe das Entra-Objekt anhand der Objekt-ID und führe anschließend den angegebenen Wiederanlauf aus." `
            -RecoveryCommand $recovery
    } finally {
        $initialPassword = $null
        $created = $null
    }
}

function Invoke-ManualStudentUpdate {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object] $Entry,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config
    )

    $ErrorActionPreference = 'Stop'
    $userId = [string]$Entry.User.Id
    $upn = [string]$Entry.User.UserPrincipalName
    $phase = 'Add'
    $needsEnable = (Get-ComparisonPropertyValue $Entry.User AccountEnabled) -eq $false
    if (@($Entry.Differences).Count -eq 0 -and -not $needsEnable) {
        return New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Add -Status Compliant
    }
    if (-not $PSCmdlet.ShouldProcess($upn, 'Update the existing student with the supplied manual values')) {
        return New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Add -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
    }

    try {
        $attributesVerified = $true
        $managerVerified = $true
        $groupsVerified = $true
        if (@($Entry.Differences | Where-Object Area -eq Entra).Count -gt 0) {
            $phase = 'Attributes'
            $attributes = Set-EntraStudentAttribute -UserId $userId -Desired $Entry.DesiredState -Differences $Entry.Differences -Confirm:$false
            if (-not $attributes.Verified) { throw 'Attribute wurden nicht verifiziert.' }
            $attributesVerified = $attributes.Verified
        }
        if (@($Entry.Differences | Where-Object Area -eq Manager).Count -gt 0) {
            $phase = 'Manager'
            $manager = Set-EntraStudentManager -UserId $userId -CurrentManagerId (Get-UserManagerId -Snapshot $Snapshot -UserId $userId) `
                -DesiredManagerId $Entry.DesiredState.ManagerId -Confirm:$false
            if (-not $manager.Verified) { throw 'Manager wurde nicht verifiziert.' }
            $managerVerified = $manager.Verified
        }
        if (@($Entry.Differences | Where-Object Area -eq Group).Count -gt 0) {
            $phase = 'Groups'
            $groupParameters = Get-StudentGroupParameter -Entry $Entry -Snapshot $Snapshot -Config $Config
            $groups = Sync-EntraStudentGroup -UserId $userId @groupParameters -CurrentDirectGroups @(Get-UserDirectGroup -Snapshot $Snapshot -UserId $userId) -Confirm:$false
            if (-not $groups.Verified) { throw 'Pflichtgruppen wurden nicht verifiziert.' }
            $groupsVerified = $groups.Verified
        }
        if ($needsEnable) {
            $phase = 'Enable'
            $enabled = Enable-EntraStudent -UserId $userId -WorkbookVerified -GroupsVerified:$groupsVerified `
                -ManagerVerified:$managerVerified -AttributesVerified:$attributesVerified -Confirm:$false
            if (-not $enabled.Verified) { throw 'Aktivierung wurde nicht verifiziert.' }
        }
        return New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase Add -Status Succeeded
    } catch {
        $recoveryGivenName = ([string](Get-ComparisonPropertyValue $Entry.Student GivenName)).Replace("'", "''")
        $recoverySurname = ([string](Get-ComparisonPropertyValue $Entry.Student Surname)).Replace("'", "''")
        $recoveryClass = ([string](Get-ComparisonPropertyValue $Entry.Student ClassName)).Replace("'", "''")
        $recoveryTeacher = ([string](Get-ComparisonPropertyValue $Entry.Student Teacher)).Replace("'", "''")
        return New-StudentActionResult -UserId $userId -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message `
            -RecoveryCommand ".\Sync-SchuelerEntra.ps1 -Add -Vorname '$recoveryGivenName' -Nachname '$recoverySurname' -Klasse '$recoveryClass' -Klassenlehrer '$recoveryTeacher' -EntraObjectId '$($userId.Replace("'", "''"))'"
    }
}
