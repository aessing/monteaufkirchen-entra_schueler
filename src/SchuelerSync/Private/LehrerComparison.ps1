function Get-LehrerProfile {
    param([Parameter(Mandatory)][string] $Job, [Parameter(Mandatory)][System.Collections.IDictionary] $Config)
    $codes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $primaryJobs = [Collections.Generic.List[string]]::new()
    $additionalJobs = [Collections.Generic.List[string]]::new()
    foreach ($part in ($Job -split ',')) {
        $code = $part.Trim().ToUpperInvariant()
        if (-not $code) { throw "Ungültige Jobliste '$Job'." }
        if ($Config.Jobs.ContainsKey($code)) {
            if ($codes.Add($code)) { $primaryJobs.Add($code) }
        } elseif ($code -in $Config.AdditionalJobs) {
            if ($codes.Add($code)) { $additionalJobs.Add($code) }
        } else { throw "Unbekannter Personaljob '$code'." }
    }
    if ($primaryJobs.Count -eq 0) { throw "Zusatzjob '$Job' benötigt mindestens einen Hauptjob." }
    $profileKeys = @($primaryJobs | ForEach-Object { [string]$Config.Jobs[$_] } | Select-Object -Unique)
    if ($profileKeys.Count -ne 1) { throw "Jobs '$Job' gehören unterschiedlichen Profilen an." }
    $key = $profileKeys[0]
    $normalized = (@($primaryJobs | Sort-Object) + @($additionalJobs | Sort-Object)) -join ', '
    $personnelProfile = $Config.Profiles[$key]
    $exchange = @{} + $Config.Exchange
    $exchange.AddressBookPolicy = $personnelProfile.AddressBookPolicy
    $exchange.CustomAttribute1 = "$($personnelProfile.CompanyName) - $($personnelProfile.EmployeeType)"
    [pscustomobject]@{
        Job = $normalized
        Key = $key
        CompanyName = $personnelProfile.CompanyName
        EmployeeType = $personnelProfile.EmployeeType
        Role = $Config.Roles[$key]
        LicenseGroupName = $personnelProfile.LicenseGroupName
        Exchange = $exchange
    }
}

function New-LehrerIssue {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Creates an in-memory issue without changing external state.')]
    param([string] $Area, [string] $Field, [string] $Message, [AllowNull()][object] $Person)
    [pscustomobject]@{ Area = $Area; Field = $Field; Message = $Message; RowNumber = Get-ComparisonPropertyValue $Person RowNumber }
}

function Get-LehrerDirectGroup {
    param([object] $Snapshot, [string] $UserId)
    foreach ($group in @(Get-UserDirectGroup -Snapshot $Snapshot -UserId $UserId)) {
        $known = $Snapshot.GroupsById[[string]$group.Id]
        if ($null -eq $known) { throw "Direkte Gruppe '$($group.Id)' fehlt im Gruppensnapshot." }
        $known
    }
}

function Get-LehrerSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary] $Config, [switch] $SkipLicenseGroupValidation)
    $userProperties = @('id','displayName','givenName','surname','userPrincipalName','mail','mailNickname',
        'proxyAddresses','companyName','employeeType','jobTitle','usageLocation','ageGroup',
        'consentProvidedForMinor','legalAgeGroupClassification','accountEnabled','userType',
        'assignedLicenses','licenseAssignmentStates')
    $groups = @(Get-MgGroup -All -Property @('id','displayName','groupTypes','membershipRule','assignedLicenses') -ErrorAction Stop)
    $users = @(Get-MgUser -All -Property $userProperties -ErrorAction Stop)
    $groupsById = New-CaseInsensitiveHashtable
    $groupMatches = New-CaseInsensitiveHashtable
    foreach ($group in $groups) {
        if ([string]::IsNullOrWhiteSpace([string]$group.Id)) { throw 'Graph lieferte eine Gruppe ohne ID.' }
        $groupsById[[string]$group.Id] = $group
        $name = [string]$group.DisplayName
        if (-not $groupMatches.ContainsKey($name)) { $groupMatches[$name] = @() }
        $groupMatches[$name] = @($groupMatches[$name]) + $group
    }
    $required = @($Config.Roles.School, $Config.Roles.Ganztag)
    foreach ($role in $required) {
        $group = $groupsById[[string]$role.Id]
        if ($null -eq $group -or -not [string]::Equals([string]$group.DisplayName, [string]$role.Name, [StringComparison]::OrdinalIgnoreCase) -or
            @($groupMatches[[string]$role.Name]).Count -ne 1 -or (Get-EntraGroupIsDynamic -Group $group)) {
            throw "Konfigurierte Personalrolle '$($role.Name)' fehlt, ist mehrdeutig oder dynamisch."
        }
    }
    $licenseGroups = New-CaseInsensitiveHashtable
    foreach ($personnelProfile in $(if ($SkipLicenseGroupValidation) { @() } else { $Config.Profiles.Values })) {
        $name = [string]$personnelProfile.LicenseGroupName
        $found = @($groupMatches[$name])
        if ($found.Count -ne 1 -or (Get-EntraGroupIsDynamic -Group $found[0])) {
            throw "Lizenzgruppe '$name' fehlt, ist mehrdeutig oder dynamisch."
        }
        if ($null -eq $found[0].PSObject.Properties['AssignedLicenses'] -or
            @($found[0].AssignedLicenses | Where-Object { $null -ne $_ }).Count -eq 0) {
            throw "Lizenzgruppe '$name' hat keine nachweisbare Lizenzzuweisung."
        }
        $licenseGroups[$name] = $found[0]
    }
    $usersById = New-CaseInsensitiveHashtable
    $usersByUpn = New-CaseInsensitiveHashtable
    $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $addressOwners = New-CaseInsensitiveHashtable
    foreach ($user in $users) {
        $id = [string]$user.Id
        if ([string]::IsNullOrWhiteSpace($id) -or $usersById.ContainsKey($id)) { throw "Ungültige oder doppelte Benutzer-ID '$id'." }
        $usersById[$id] = $user
        $upn = [string]$user.UserPrincipalName
        if ($upn) {
            if ($usersByUpn.ContainsKey($upn)) { throw "Doppelter UPN '$upn'." }
            $usersByUpn[$upn] = $user
            Add-EntraReservedAddress -ReservedAddresses $reserved -AddressOwners $addressOwners -Address $upn -OwnerId $id
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$user.Mail)) {
            Add-EntraReservedAddress -ReservedAddresses $reserved -AddressOwners $addressOwners -Address $user.Mail -OwnerId $id
        }
        foreach ($address in @($user.ProxyAddresses)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$address)) {
                Add-EntraReservedAddress -ReservedAddresses $reserved -AddressOwners $addressOwners -Address $address -OwnerId $id -IsProxyAddress
            }
        }
    }
    $memberIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $rolesByUserId = New-CaseInsensitiveHashtable
    foreach ($role in $required) {
        foreach ($member in @(Get-MgGroupMember -GroupId $role.Id -All -ErrorAction Stop)) {
            if ($member.Id -and $usersById.ContainsKey([string]$member.Id)) {
                $id = [string]$member.Id
                [void]$memberIds.Add($id)
                if (-not $rolesByUserId.ContainsKey($id)) { $rolesByUserId[$id] = @() }
                $rolesByUserId[$id] = @($rolesByUserId[$id]) + [string]$role.Name
            }
        }
    }
    $studentRole = Get-ConfiguredStudentRole
    $studentRoleId = [string]$studentRole.Id
    if (-not $groupsById.ContainsKey($studentRoleId) -or $groupsById[$studentRoleId].DisplayName -cne $studentRole.Name) {
        throw 'Konfigurierte Schülerrolle fehlt oder ihr Name stimmt nicht überein.'
    }
    $studentIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($groupsById.ContainsKey($studentRoleId)) {
        foreach ($member in @(Get-MgGroupMember -GroupId $studentRoleId -All -ErrorAction Stop)) {
            if ($member.Id) { [void]$studentIds.Add([string]$member.Id) }
        }
    }
    [pscustomobject]@{
        UsersById = $usersById; UsersByUpn = $usersByUpn; GroupsById = $groupsById
        GroupMatchesByDisplayName = $groupMatches; LicenseGroups = $licenseGroups
        PersonnelMemberIds = $memberIds; PersonnelRolesByUserId = $rolesByUserId; StudentMemberIds = $studentIds
        ReservedAddresses = $reserved; AddressOwners = $addressOwners
        DirectGroupsByUserId = New-CaseInsensitiveHashtable
        TransitiveGroupsByUserId = New-CaseInsensitiveHashtable
    }
}

function Test-LehrerLicenseState {
    param([AllowNull()][object] $User, [AllowEmptyCollection()][object[]] $DirectGroups)
    if ($null -eq $User) { return [pscustomobject]@{ Add = $true; Reason = 'Neuanlage' } }
    foreach ($propertyName in 'AssignedLicenses','LicenseAssignmentStates') {
        if ($null -eq $User.PSObject.Properties[$propertyName] -or $null -eq $User.$propertyName) {
            throw "Lizenzdaten '$propertyName' für '$($User.Id)' sind unvollständig."
        }
    }
    $assigned = @($User.AssignedLicenses)
    $states = @($User.LicenseAssignmentStates)
    if (@($states | Where-Object {
        $errorValue = ([string]$_.Error).Trim()
        ($errorValue -and $errorValue -ine 'None') -or $_.State -in @('Error','ActiveWithError')
    }).Count -gt 0) {
        throw "Lizenzzuweisung für '$($User.Id)' meldet einen Fehler."
    }
    if ($assigned.Count -gt 0 -or $states.Count -gt 0) {
        return [pscustomobject]@{ Add = $false; Reason = 'Bestehende Lizenz bleibt erhalten' }
    }
    foreach ($group in $DirectGroups) {
        $groupLicenses = Get-ComparisonPropertyValue -InputObject $group -Name AssignedLicenses
        if ($null -eq $groupLicenses -and $group.DisplayName -like 'SEC-A-LIC-*') {
            throw "Lizenzdaten der Gruppe '$($group.DisplayName)' sind unvollständig."
        }
        if (@($groupLicenses | Where-Object { $null -ne $_ }).Count -gt 0) {
            return [pscustomobject]@{ Add = $false; Reason = 'Lizenzverarbeitung durch bestehende Gruppe ausstehend' }
        }
    }
    return [pscustomobject]@{ Add = $true; Reason = 'Nachweislich ohne Lizenz' }
}

function Resolve-LehrerIdentity {
    param([object] $Person, [object] $Snapshot, [switch] $Recovery, [System.Collections.IDictionary] $Config)
    $objectId = ([string]$Person.EntraObjectId).Trim()
    $upn = ([string]$Person.StoredUpn).Trim()
    $user = $null
    if ($objectId) {
        $user = $Snapshot.UsersById[$objectId]
        if ($null -eq $user) { throw "EntraObjectId '$objectId' wurde nicht gefunden." }
    }
    if ($upn) {
        $byUpn = $Snapshot.UsersByUpn[$upn]
        if ($null -eq $byUpn -or ($null -ne $user -and $byUpn.Id -ine $user.Id)) {
            throw "UPN '$upn' verweist nicht eindeutig auf die angegebene Objekt-ID."
        }
        $user = $byUpn
    }
    if ($null -eq $user -and -not $objectId -and -not $upn) {
        $key = Get-NormalizedStudentNameKey -Student $Person
        $hits = @($Snapshot.UsersById.Values | Where-Object {
            try { (Get-NormalizedStudentNameKey -Student $_) -ceq $key } catch { $false }
        })
        if ($hits.Count -gt 1) { throw "Name '$($Person.NameMitRufname)' ist im Verzeichnis mehrdeutig." }
        if ($hits.Count -eq 1) { $user = $hits[0] }
    }
    if ($null -ne $user) {
        if ($Snapshot.StudentMemberIds.Contains([string]$user.Id) -or $user.UserType -eq 'Guest') {
            throw "Benutzer '$($user.Id)' ist Schüler oder Gast."
        }
        if (-not $Snapshot.PersonnelMemberIds.Contains([string]$user.Id)) {
            $checkpoint = $objectId -and $null -ne $Config -and
                (Get-ComparisonPropertyValue $user AccountEnabled) -eq $false -and
                -not [string]::IsNullOrWhiteSpace([string](Get-ComparisonPropertyValue $Person Password))
            if ($checkpoint) {
                $recoveryProfile = Get-LehrerProfile -Job $Person.Job -Config $Config
                $checkpoint = $user.CompanyName -ceq $recoveryProfile.CompanyName -and $user.EmployeeType -ceq $recoveryProfile.EmployeeType
            }
            if (-not ($Recovery -and $objectId) -and -not $checkpoint) {
                throw "Benutzer '$($user.Id)' gehört keiner Personalrolle an."
            }
        }
        if ((Get-NormalizedStudentNameKey -Student $Person) -cne (Get-NormalizedStudentNameKey -Student $user)) {
            throw "Name für Benutzer '$($user.Id)' stimmt nicht mit Excel überein."
        }
    }
    return $user
}

function Compare-LehrerDirectory {
    param([object[]] $People, [object] $Snapshot, [System.Collections.IDictionary] $Config,
        [object] $ExchangeAddressOwners, [switch] $Recovery)
    $result = [pscustomobject]@{
        NewPeople = @(); Departures = @(); ChangedPeople = @(); ExistingPeople = @(); Warnings = @(); Errors = @()
    }
    $claims = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $identities = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $claimedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($address in $Snapshot.ReservedAddresses) { [void]$reserved.Add($address) }
    $owners = New-CaseInsensitiveHashtable
    foreach ($key in $Snapshot.AddressOwners.Keys) { $owners[$key] = @($Snapshot.AddressOwners[$key]) }
    if ($null -ne $ExchangeAddressOwners) {
        foreach ($key in $ExchangeAddressOwners.Keys) {
            if (-not $owners.ContainsKey($key)) { $owners[$key] = @() }
            $owners[$key] = @($owners[$key]) + @($ExchangeAddressOwners[$key])
            [void]$reserved.Add([string]$key)
        }
    }
    $personIndex = 0
    foreach ($person in $People) {
        $personIndex++
        $personPercent = if ($People.Count -eq 0) { 0 } else { [int](100 * $personIndex / $People.Count) }
        Write-Progress -Id 2 -ParentId 1 -Activity 'Lehrerdetails prüfen' `
            -Status "Prüfe Person $personIndex von $($People.Count): $($person.NameMitRufname)" -PercentComplete $personPercent
        try {
            $personnelProfile = Get-LehrerProfile -Job $person.Job -Config $Config
            $nameKey = Get-NormalizedStudentNameKey -Student $person
            if (-not $identities.Add($nameKey)) { throw "Doppelter Name '$($person.NameMitRufname)' in der Datei." }
            $user = Resolve-LehrerIdentity -Person $person -Snapshot $Snapshot -Recovery:$Recovery -Config $Config
            if ($null -ne $user -and -not $claims.Add([string]$user.Id)) { throw "Mehrere Zeilen beanspruchen Benutzer '$($user.Id)'." }
            if ($null -ne $user -and $null -eq (Get-ComparisonPropertyValue $user AccountEnabled)) {
                throw "Kontostatus für Benutzer '$($user.Id)' ist unbekannt."
            }
            $role = $Snapshot.GroupsById[[string]$personnelProfile.Role.Id]
            $licenseGroup = $Snapshot.LicenseGroups[[string]$personnelProfile.LicenseGroupName]
            if ($null -eq $role -or $null -eq $licenseGroup) { throw 'Konfigurierte Personal- oder Lizenzgruppe fehlt im Snapshot.' }
            if (@($role.AssignedLicenses | Where-Object { $null -ne $_ }).Count -gt 0) {
                throw "Zielrolle '$($role.DisplayName)' vergibt selbst Lizenzen."
            }
            $direct = if ($null -eq $user) { @() } else { @(Get-LehrerDirectGroup -Snapshot $Snapshot -UserId $user.Id) }
            $license = Test-LehrerLicenseState -User $user -DirectGroups $direct
            if ($null -eq $user) {
                $selected = Select-AvailableUpn -GivenName $person.GivenName -Surname $person.Surname -Domain $Config.Domain `
                    -UsedAddresses $reserved -AddressOwners $owners
                $upn = $selected.Upn
            } else {
                $upn = [string]$user.UserPrincipalName
                if (-not $upn) { throw "Bestehender Benutzer '$($user.Id)' hat keinen UPN." }
                [void]$claimedIds.Add([string]$user.Id)
            }
            $desired = [pscustomobject]@{
                DisplayName = "$($person.GivenName) $($person.Surname)"
                GivenName = $person.GivenName; Surname = $person.Surname
                UserPrincipalName = $upn; Mail = $upn; MailNickname = ($upn -split '@')[0]
                JobTitle = $personnelProfile.Job; CompanyName = $personnelProfile.CompanyName
                EmployeeType = $personnelProfile.EmployeeType; UsageLocation = 'DE'
                AgeGroup = 'Adult'; ConsentProvidedForMinor = 'NotRequired'
                LegalAgeGroupClassification = 'Adult'
                Role = $role; LicenseGroup = $licenseGroup; AddLicenseGroup = $license.Add
                LicenseReason = $license.Reason; Exchange = $personnelProfile.Exchange
            }
            $differences = [Collections.Generic.List[object]]::new()
            if ($null -ne $user) {
                foreach ($field in 'DisplayName','GivenName','Surname','JobTitle','CompanyName','EmployeeType','UsageLocation','AgeGroup','ConsentProvidedForMinor') {
                    $difference = New-StateDifference -Area Entra -Field $field -Current (Get-ComparisonPropertyValue $user $field) -Desired $desired.$field
                    if ($null -ne $difference) { $differences.Add($difference) }
                }
                $legal = New-StateDifference -Area Entra -Field LegalAgeGroupClassification -Current $user.LegalAgeGroupClassification -Desired 'Adult' -Action Compare
                if ($null -ne $legal) { $differences.Add($legal) }
                if ($user.AccountEnabled -eq $false) {
                    $differences.Add((New-StateDifference -Area Entra -Field AccountEnabled -Current $false -Desired $true -Action Set))
                }
                $roleIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($group in $direct) {
                    [void]$roleIds.Add([string]$group.Id)
                    if ($group.DisplayName.StartsWith($Config.RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase) -and
                        $group.Id -ine $role.Id) {
                        if (@($Snapshot.GroupsById[[string]$group.Id].AssignedLicenses | Where-Object { $null -ne $_ }).Count -gt 0) {
                            throw "Rolle '$($group.DisplayName)' vergibt selbst Lizenzen. Rollenwechsel würde den Bestandsschutz verletzen."
                        }
                        if (Get-EntraGroupIsDynamic -Group $group) { throw "Dynamische Rolle '$($group.DisplayName)' kann nicht entfernt werden." }
                        $differences.Add((New-StateDifference -Area Group -Field Membership -Current $group.DisplayName -Desired $null -Action Remove))
                    }
                }
                if (-not $roleIds.Contains([string]$role.Id)) {
                    if (@($role.AssignedLicenses | Where-Object { $null -ne $_ }).Count -gt 0) { throw "Zielrolle '$($role.DisplayName)' vergibt selbst Lizenzen." }
                    $differences.Add((New-StateDifference -Area Group -Field Membership -Current $null -Desired $role.DisplayName -Action Add))
                }
                if ($license.Add -and -not $roleIds.Contains([string]$licenseGroup.Id)) {
                    $differences.Add((New-StateDifference -Area LicenseGroup -Field Membership -Current $null -Desired $licenseGroup.DisplayName -Action Add))
                }
            }
            $entry = [pscustomobject]@{ Person = $person; User = $user; DesiredState = $desired
                Differences = [object[]]@($differences); LicenseReason = $license.Reason }
            if ($null -eq $user) { $result.NewPeople += $entry }
            elseif ($differences.Count -gt 0) { $result.ChangedPeople += $entry }
            else { $result.ExistingPeople += $entry }
        } catch {
            $result.Errors += New-LehrerIssue -Area Preflight -Field Identity -Message $_.Exception.Message -Person $person
        }
    }
    Write-Progress -Id 2 -ParentId 1 -Activity 'Lehrerdetails prüfen' -Completed
    if ($result.Errors.Count -gt 0) {
        $result.NewPeople = @(); $result.Departures = @()
    } elseif (-not $Recovery) {
        foreach ($id in $Snapshot.PersonnelMemberIds) {
            if ($claimedIds.Contains($id)) { continue }
            $user = $Snapshot.UsersById[$id]
            if ($null -eq $user) {
                $result.Errors += New-LehrerIssue -Area Preflight -Field Departure -Message "Personalrollenmitglied '$id' fehlt im Snapshot."
            } elseif ($null -eq (Get-ComparisonPropertyValue $user AccountEnabled)) {
                $result.Errors += New-LehrerIssue -Area Preflight -Field Departure -Message "Kontostatus für Personalrollenmitglied '$id' ist unbekannt."
            } elseif ($user.AccountEnabled -eq $false) {
                continue
            } else {
                $result.Departures += [pscustomobject]@{ Person = $null; User = $user; DesiredState = $null; Differences = @()
                    SourceRoles = @($Snapshot.PersonnelRolesByUserId[$id]) }
            }
        }
        if ($result.Errors.Count -gt 0) { $result.Departures = @() }
    }
    return $result
}
