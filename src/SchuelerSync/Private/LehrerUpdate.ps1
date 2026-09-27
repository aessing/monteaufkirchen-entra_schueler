function New-LehrerPassword {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Creates an in-memory password without changing external state.')]
    param([Collections.Generic.HashSet[string]] $UsedPasswords)
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghjkmnpqrstuvwxyz', '23456789', '!#%+?@')
    $alphabet = $sets -join ''
    for ($attempt = 0; $attempt -lt 100; $attempt++) {
        $characters = [Collections.Generic.List[char]]::new()
        foreach ($set in $sets) { $characters.Add($set[[Security.Cryptography.RandomNumberGenerator]::GetInt32($set.Length)]) }
        while ($characters.Count -lt 12) {
            $characters.Add($alphabet[[Security.Cryptography.RandomNumberGenerator]::GetInt32($alphabet.Length)])
        }
        for ($index = $characters.Count - 1; $index -gt 0; $index--) {
            $other = [Security.Cryptography.RandomNumberGenerator]::GetInt32($index + 1)
            $swap = $characters[$index]; $characters[$index] = $characters[$other]; $characters[$other] = $swap
        }
        $password = -join $characters.ToArray()
        if ($null -eq $UsedPasswords -or -not $UsedPasswords.Contains($password)) { return $password }
    }
    throw 'Es konnte kein eindeutiges Initialpasswort erzeugt werden.'
}

function New-DisabledEntraLehrer {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'Password',
        Justification = 'Graph requires an in-memory password. It is never logged or returned publicly.')]
    param([Parameter(Mandatory)][object] $Desired, [Parameter(Mandatory)][string] $Password)
    $body = @{
        AccountEnabled = $false; DisplayName = $Desired.DisplayName
        GivenName = $Desired.GivenName; Surname = $Desired.Surname
        UserPrincipalName = $Desired.UserPrincipalName; Mail = $Desired.Mail
        MailNickname = $Desired.MailNickname; JobTitle = $Desired.JobTitle
        CompanyName = $Desired.CompanyName; EmployeeType = $Desired.EmployeeType
        UsageLocation = $Desired.UsageLocation; AgeGroup = $Desired.AgeGroup
        ConsentProvidedForMinor = $Desired.ConsentProvidedForMinor
        PasswordProfile = @{ Password = $Password; ForceChangePasswordNextSignIn = $true }
    }
    if ($PSCmdlet.ShouldProcess($Desired.UserPrincipalName, 'Create disabled personnel account')) {
        return New-MgUser -BodyParameter $body -ErrorAction Stop
    }
}

function Assert-LehrerAttribute {
    param([string] $UserId, [object] $Desired, [switch] $MustBeDisabled)
    $fields = @('displayName','givenName','surname','jobTitle','companyName','employeeType','usageLocation',
        'ageGroup','consentProvidedForMinor','accountEnabled','userPrincipalName')
    $user = Get-MgUser -UserId $UserId -Property $fields -ErrorAction Stop
    if ($MustBeDisabled -and ($null -eq $user.AccountEnabled -or $user.AccountEnabled)) {
        throw "Neues Personalkonto '$UserId' ist nicht nachweislich deaktiviert."
    }
    foreach ($field in @('DisplayName','GivenName','Surname','JobTitle','CompanyName','EmployeeType',
            'UsageLocation','AgeGroup','ConsentProvidedForMinor','UserPrincipalName')) {
        if (-not (Test-EntraMutationValueEqual -Field $field -Current (Get-ComparisonPropertyValue $user $field) -Desired $Desired.$field)) {
            throw "Personalattribut '$field' konnte für '$UserId' nicht verifiziert werden."
        }
    }
    return $true
}

function Set-LehrerAttribute {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([string] $UserId, [object] $Desired, [object[]] $Differences)
    $body = @{}
    $allowed = @('DisplayName','GivenName','Surname','JobTitle','CompanyName','EmployeeType',
        'UsageLocation','AgeGroup','ConsentProvidedForMinor')
    foreach ($difference in $Differences) {
        if ($difference.Area -eq 'Entra' -and $difference.Action -eq 'Set' -and $difference.Field -in $allowed) {
            $body[[string]$difference.Field] = $Desired.($difference.Field)
        }
    }
    if ($body.Count -eq 0) { return $true }
    if (-not $PSCmdlet.ShouldProcess($UserId, 'Set personnel attributes')) { return $false }
    Update-MgUser -UserId $UserId -BodyParameter $body -ErrorAction Stop
    return Assert-LehrerAttribute -UserId $UserId -Desired $Desired
}

function Assert-LehrerRecoveryIdentity {
    param([string] $UserId, [object] $Person, [System.Collections.IDictionary] $Config)
    $user = Get-MgUser -UserId $UserId -Property @('id','givenName','surname','companyName','employeeType','userType','accountEnabled') -ErrorAction Stop
    $direct = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    $studentRole = Get-ConfiguredStudentRole
    if ($user.Id -ine $UserId -or $user.UserType -ne 'Member' -or @($direct | Where-Object Id -eq $studentRole.Id).Count -gt 0 -or
        (Get-NormalizedStudentNameKey -Student $user) -cne (Get-NormalizedStudentNameKey -Student $Person)) {
        throw "Wiederanlauf-ID '$UserId' gehört nicht zur angegebenen Person."
    }
    if (@($direct | Where-Object { $_.Id -in @($Config.Roles.School.Id, $Config.Roles.Ganztag.Id) }).Count -eq 0) {
        if ($user.AccountEnabled -ne $false -or [string]$Person.EntraObjectId -ine $UserId) {
            throw "Konto '$UserId' gehört keiner Personalrolle mehr an. Wiederanlauf erfordert ein gesperrtes Konto und eine explizite Objekt-ID."
        }
        $personnelProfile = Get-LehrerProfile -Job $Person.Job -Config $Config
        if ($user.CompanyName -cne $personnelProfile.CompanyName -or $user.EmployeeType -cne $personnelProfile.EmployeeType) {
            throw "Wiederanlauf-ID '$UserId' hat kein passendes Personalprofil."
        }
    }
    return $user
}

function Sync-LehrerGroup {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([string] $UserId, [object] $Desired, [string] $RoleGroupPrefix = 'SEC-A-ROL-')
    $role = Get-WritablePersonnelRole -GroupId $Desired.Role.Id
    $direct = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    $hasRole = @($direct | Where-Object Id -eq $role.Id).Count -gt 0
    if (-not $hasRole) {
        if (@($role.AssignedLicenses | Where-Object { $null -ne $_ }).Count -gt 0) { throw "Zielrolle '$($role.DisplayName)' vergibt Lizenzen." }
        if (Get-EntraGroupIsDynamic -Group $role) { throw "Zielrolle '$($role.DisplayName)' ist dynamisch." }
        if ($PSCmdlet.ShouldProcess($role.DisplayName, "Add personnel '$UserId' to role")) {
            New-MgGroupMemberByRef -GroupId $role.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" } -ErrorAction Stop
        } else { return $false }
    }
    $afterAdd = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    if (@($afterAdd | Where-Object Id -eq $role.Id).Count -ne 1) { throw "Zielrolle für '$UserId' wurde nicht bestätigt." }
    foreach ($group in $afterAdd) {
        if ($group.Id -eq $role.Id -or -not $group.DisplayName.StartsWith($RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $null = Get-WritablePersonnelRole -GroupId $group.Id
        if ($PSCmdlet.ShouldProcess($group.DisplayName, "Remove competing role from personnel '$UserId'")) {
            Remove-MgGroupMemberDirectoryObjectByRef -GroupId $group.Id -DirectoryObjectId $UserId -ErrorAction Stop
        } else { return $false }
    }
    $final = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    if (@($final | Where-Object Id -eq $role.Id).Count -ne 1 -or
        @($final | Where-Object { $_.Id -ne $role.Id -and $_.DisplayName.StartsWith($RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0) {
        throw "Rollen für '$UserId' wurden nicht vollständig verifiziert."
    }
    return $true
}

function Add-LehrerLicenseGroup {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([string] $UserId, [object] $Desired, [switch] $NewAccount)
    $group = $Desired.LicenseGroup
    $direct = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    if (@($direct | Where-Object Id -eq $group.Id).Count -gt 0) { return $true }
    if (-not $NewAccount) {
        $fresh = Get-MgUser -UserId $UserId -Property @('id','assignedLicenses','licenseAssignmentStates') -ErrorAction Stop
        $decision = Test-LehrerLicenseState -User $fresh -DirectGroups $direct
        if (-not $decision.Add) { return $true }
    }
    if (-not $PSCmdlet.ShouldProcess($group.DisplayName, "Add personnel '$UserId' to license group")) { return $false }
    New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" } -ErrorAction Stop
    if (@(Get-FreshEntraUserDirectGroup -UserId $UserId | Where-Object Id -eq $group.Id).Count -ne 1) {
        throw "Lizenzgruppe für '$UserId' wurde nicht bestätigt."
    }
    return $true
}

function Invoke-LehrerCreate {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([object] $Entry, [System.Collections.IDictionary] $Config,
        [string] $File, [object] $WorkbookState, [Collections.Generic.HashSet[string]] $UsedPasswords)
    $upn = [string]$Entry.DesiredState.UserPrincipalName
    if (-not $PSCmdlet.ShouldProcess($upn, 'Create personnel account')) {
        return New-StudentActionResult -UserPrincipalName $upn -Phase Create -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
    }
    $phase = 'Create'; $id = ''; $password = ''; $created = $null; $revealed = $false; $checkpointWritten = $false
    try {
        if ($File) { Assert-StudentWorkbookVersion -Path $File -ExpectedSourceHash $WorkbookState.SourceHash }
        for ($attempt = 1; $attempt -le 6; $attempt++) {
            $password = New-LehrerPassword -UsedPasswords $UsedPasswords
            [void]$UsedPasswords.Add($password)
            try {
                $created = New-DisabledEntraLehrer -Desired $Entry.DesiredState -Password $password -Confirm:$false
                break
            } catch {
                if ($attempt -ge 6 -or -not (Test-EntraPasswordPolicyRejection -ErrorRecord $_)) { throw }
            }
        }
        $id = [string]$created.Id
        if (-not $id) { throw 'Neuanlage lieferte keine Objekt-ID.' }
        $phase = 'Workbook'
        if ($File) {
            $written = Write-StudentWorkbookUpdate -Path $File -ExpectedSourceHash $WorkbookState.SourceHash -Kind Teacher `
                -Updates @([pscustomobject]@{ RowNumber = [int]$Entry.Person.RowNumber; Password = $password
                    EntraObjectId = $id; UPN = $upn; Mail = $upn }) -Confirm:$false
            $WorkbookState.SourceHash = $written.SourceHash
            $checkpointWritten = $true
        }
        $phase = 'Attributes'
        $null = Assert-LehrerAttribute -UserId $id -Desired $Entry.DesiredState -MustBeDisabled
        $phase = 'Groups'
        if (-not (Sync-LehrerGroup -UserId $id -Desired $Entry.DesiredState -RoleGroupPrefix $Config.RoleGroupPrefix -Confirm:$false)) {
            throw 'Rolle konnte nicht verifiziert werden.'
        }
        if (-not (Add-LehrerLicenseGroup -UserId $id -Desired $Entry.DesiredState -NewAccount -Confirm:$false)) {
            throw 'Lizenzgruppe konnte nicht verifiziert werden.'
        }
        $phase = 'Enable'
        $enabled = Enable-EntraStudent -UserId $id -WorkbookVerified -GroupsVerified -ManagerVerified -AttributesVerified -Confirm:$false
        if (-not $enabled.Verified) { throw 'Aktivierung wurde nicht bestätigt.' }
        if (-not $File) { Write-Information "Einmaliges Startpasswort für $upn`: $password" -InformationAction Continue; $revealed = $true }
        return New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase Create -Status Succeeded
    } catch {
        if (-not $File -and $id -and -not $revealed) { Write-Information "Einmaliges Startpasswort für $upn`: $password" -InformationAction Continue }
        $recovery = if ($File -and ($checkpointWritten -or -not $id)) { ".\Sync-LehrerEntra.ps1 -File '$($File.Replace("'", "''"))' -Update -CreateNewUsers -UpdateUsers" }
            else {
                $command = ".\Sync-LehrerEntra.ps1 -Add -Vorname '$($Entry.Person.GivenName.Replace("'", "''"))' -Nachname '$($Entry.Person.Surname.Replace("'", "''"))' -Job '$($Entry.Person.Job)'"
                if ($id) { $command += " -EntraObjectId '$id'" }
                $command
            }
        $message = $_.Exception.Message
        if ($File -and $id -and -not $checkpointWritten) {
            $message += ' Passwort und Identität konnten nicht in Excel gesichert werden. Das Konto bleibt gesperrt. Vor dem manuellen Wiederanlauf das Startpasswort administrativ zurücksetzen und sicher übergeben. Anschließend die Objekt-ID in Excel nachtragen.'
        }
        return New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Failed -Message $message -Secrets @($UsedPasswords) -RecoveryCommand $recovery
    } finally { $password = $null; $created = $null }
}

function Invoke-LehrerUpdate {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([object] $Entry, [System.Collections.IDictionary] $Config)
    $id = [string]$Entry.User.Id; $upn = [string]$Entry.User.UserPrincipalName
    if (-not $PSCmdlet.ShouldProcess($upn, 'Update personnel account')) {
        return New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase Update -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })
    }
    $phase = 'Identity'
    try {
        $null = Assert-LehrerRecoveryIdentity -UserId $id -Person $Entry.Person -Config $Config
        $phase = 'Attributes'
        if (-not (Set-LehrerAttribute -UserId $id -Desired $Entry.DesiredState -Differences $Entry.Differences -Confirm:$false)) { throw 'Attribute nicht verifiziert.' }
        $phase = 'Groups'
        if (-not (Sync-LehrerGroup -UserId $id -Desired $Entry.DesiredState -RoleGroupPrefix $Config.RoleGroupPrefix -Confirm:$false)) { throw 'Rollen nicht verifiziert.' }
        if ($Entry.DesiredState.AddLicenseGroup -and -not (Add-LehrerLicenseGroup -UserId $id -Desired $Entry.DesiredState -Confirm:$false)) { throw 'Lizenzgruppe nicht verifiziert.' }
        if ($Entry.User.AccountEnabled -eq $false) {
            $phase = 'Enable'
            $enabled = Enable-EntraStudent -UserId $id -WorkbookVerified -GroupsVerified -ManagerVerified -AttributesVerified -Confirm:$false
            if (-not $enabled.Verified) { throw 'Reaktivierung nicht bestätigt.' }
        }
        return New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase Update -Status Succeeded
    } catch {
        return New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message `
            -RecoveryCommand ".\Sync-LehrerEntra.ps1 -Add -Vorname '$($Entry.Person.GivenName.Replace("'", "''"))' -Nachname '$($Entry.Person.Surname.Replace("'", "''"))' -Job '$($Entry.Person.Job)' -EntraObjectId '$id'"
    }
}
