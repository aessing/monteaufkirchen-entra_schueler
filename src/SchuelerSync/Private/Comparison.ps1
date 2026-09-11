function Get-ComparisonPropertyValue {
    param(
        [AllowNull()][object] $InputObject,
        [Parameter(Mandatory)][string] $Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-NormalizedStudentNameKey {
    param([Parameter(Mandatory)][object] $Student)

    $givenName = ConvertTo-UpnToken ([string](Get-ComparisonPropertyValue -InputObject $Student -Name GivenName))
    $surname = ConvertTo-UpnToken ([string](Get-ComparisonPropertyValue -InputObject $Student -Name Surname))
    if ([string]::IsNullOrWhiteSpace($givenName) -or [string]::IsNullOrWhiteSpace($surname)) {
        throw 'Vor- und Nachname müssen einen nicht leeren normalisierten Namen ergeben.'
    }
    return "$givenName|$surname"
}

function New-ComparisonIssue {
    param(
        [Parameter(Mandatory)][ValidateSet('Warning', 'Error')][string] $Severity,
        [Parameter(Mandatory)][string] $Area,
        [Parameter(Mandatory)][string] $Field,
        [AllowNull()][object] $Current,
        [AllowNull()][object] $Desired,
        [Parameter(Mandatory)][string] $Message,
        [AllowNull()][object] $Student,
        [AllowNull()][object] $User
    )

    [pscustomobject]@{
        Severity = $Severity
        Area = $Area
        Field = $Field
        Current = $Current
        Desired = $Desired
        Message = $Message
        RowNumber = Get-ComparisonPropertyValue -InputObject $Student -Name RowNumber
        Student = $Student
        User = $User
    }
}

function Resolve-StudentIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Student,
        [Parameter(Mandatory)][object] $Snapshot
    )

    $objectId = ([string](Get-ComparisonPropertyValue -InputObject $Student -Name EntraObjectId)).Trim()
    if (-not [string]::IsNullOrWhiteSpace($objectId)) {
        $user = $Snapshot.UsersById[$objectId]
        if ($null -eq $user) {
            throw "EntraObjectId '$objectId' verweist auf keinen vorhandenen Entra-Benutzer."
        }
        return [pscustomobject]@{
            Method = 'EntraObjectId'
            User = $user
            Warnings = @()
        }
    }

    $storedUpn = ([string](Get-ComparisonPropertyValue -InputObject $Student -Name StoredUpn)).Trim()
    if (-not [string]::IsNullOrWhiteSpace($storedUpn)) {
        $matches = @($Snapshot.UsersById.Values | Where-Object {
                [string]::Equals(
                    ([string](Get-ComparisonPropertyValue -InputObject $_ -Name UserPrincipalName)).Trim(),
                    $storedUpn,
                    [StringComparison]::OrdinalIgnoreCase
                )
            })
        if ($matches.Count -ne 1) {
            throw "StoredUpn '$storedUpn' verweist nicht eindeutig auf einen vorhandenen Entra-Benutzer ($($matches.Count) Treffer)."
        }
        return [pscustomobject]@{
            Method = 'StoredUpn'
            User = $matches[0]
            Warnings = @()
        }
    }

    $nameKey = Get-NormalizedStudentNameKey -Student $Student
    $studentMatches = @()
    $outsideMatches = @()
    foreach ($user in @($Snapshot.UsersById.Values)) {
        $userId = ([string](Get-ComparisonPropertyValue -InputObject $user -Name Id)).Trim()
        $userNameKey = $null
        try {
            $userNameKey = Get-NormalizedStudentNameKey -Student $user
        } catch {
            continue
        }
        if (-not [string]::Equals($nameKey, $userNameKey, [StringComparison]::Ordinal)) {
            continue
        }
        if ($Snapshot.StudentRoleMemberIds.Contains($userId)) {
            $studentMatches += $user
        } else {
            $outsideMatches += $user
        }
    }

    if ($studentMatches.Count -gt 1) {
        throw "Der normalisierte Name '$nameKey' ist innerhalb der Schüler-Rollengruppe nicht eindeutig ($($studentMatches.Count) Treffer)."
    }
    if ($studentMatches.Count -eq 1) {
        return [pscustomobject]@{
            Method = 'UniqueStudentName'
            User = $studentMatches[0]
            Warnings = @()
        }
    }
    if ($outsideMatches.Count -gt 0) {
        throw "Der normalisierte Name '$nameKey' gehört zu $($outsideMatches.Count) Entra-Benutzer(n) außerhalb der Schüler-Rollengruppe."
    }

    return [pscustomobject]@{
        Method = 'None'
        User = $null
        Warnings = @()
    }
}

function New-StudentDesiredState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Student,
        [Parameter(Mandatory)][string] $SelectedUpn,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [Parameter(Mandatory)][object] $Manager
    )

    $givenName = [string](Get-ComparisonPropertyValue -InputObject $Student -Name GivenName)
    $surname = [string](Get-ComparisonPropertyValue -InputObject $Student -Name Surname)
    $className = [string](Get-ComparisonPropertyValue -InputObject $Student -Name ClassName)
    $upn = $SelectedUpn.Trim().ToLowerInvariant()
    [pscustomobject]@{
        DisplayName = "$givenName $surname"
        GivenName = $givenName
        Surname = $surname
        UserPrincipalName = $upn
        Mail = $upn
        MailNickname = ($upn -split '@')[0]
        Department = $className
        OfficeLocation = Get-OfficeLocation $className
        CompanyName = $Config.CompanyName
        EmployeeType = $Config.EmployeeType
        UsageLocation = $Config.UsageLocation
        AgeGroup = $Config.AgeGroup
        ConsentProvidedForMinor = $Config.ConsentProvidedForMinor
        LegalAgeGroupClassification = $Config.LegalAgeGroupClassification
        ManagerId = [string](Get-ComparisonPropertyValue -InputObject $Manager -Name Id)
        RequiredGroupNames = @(
            $Config.StudentRoleGroup.Name
            $Config.LicenseGroupName
            "$($Config.ClassGroupPrefix)$className"
        )
    }
}

function New-StateDifference {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Area,
        [Parameter(Mandatory)][string] $Field,
        [AllowNull()][object] $Current,
        [AllowNull()][object] $Desired,
        [ValidateSet('Add', 'Remove', 'Set', 'Compare')][string] $Action = 'Set'
    )

    $isEqual = $false
    if ($null -eq $Current -and $null -eq $Desired) {
        $isEqual = $true
    } elseif ($null -eq $Current -or $null -eq $Desired) {
        $isEqual = $false
    } elseif ($Area -eq 'Group' -or $Field -in @('UserPrincipalName', 'Mail', 'MailNickname')) {
        $isEqual = [string]::Equals(
            ([string]$Current).Trim(),
            ([string]$Desired).Trim(),
            [StringComparison]::OrdinalIgnoreCase
        )
    } elseif ($Field -match '(?i)Id$') {
        $isEqual = [string]::Equals([string]$Current, [string]$Desired, [StringComparison]::Ordinal)
    } elseif ($Current -is [string] -or $Desired -is [string]) {
        $isEqual = [string]::Equals(
            ([string]$Current).Trim(),
            ([string]$Desired).Trim(),
            [StringComparison]::Ordinal
        )
    } else {
        $isEqual = [object]::Equals($Current, $Desired)
    }

    if ($isEqual) { return }
    [pscustomobject]@{
        Area = $Area
        Field = $Field
        Current = $Current
        Desired = $Desired
        Action = $Action
    }
}

function Add-ComparisonAddressOwner {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary] $AddressOwners,
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]] $ReservedAddresses,
        [Parameter(Mandatory)][string] $Address,
        [AllowNull()][object] $Owner
    )

    $normalized = $Address.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($normalized)) { return }
    [void] $ReservedAddresses.Add($normalized)
    $owners = @($AddressOwners[$normalized])
    $ownerIds = @(Get-AddressOwnerIds $Owner)
    foreach ($ownerRecord in @($Owner)) {
        if ($null -eq $ownerRecord) { continue }
        $externalIdProperty = $ownerRecord.PSObject.Properties['ExternalDirectoryObjectId']
        if ($null -ne $externalIdProperty -and
            -not [string]::IsNullOrWhiteSpace([string]$externalIdProperty.Value)) {
            $ownerIds += [string]$externalIdProperty.Value
        }
    }
    foreach ($ownerId in $ownerIds) {
        if (@($owners | Where-Object {
                    [string]::Equals([string]$_, [string]$ownerId, [StringComparison]::OrdinalIgnoreCase)
                }).Count -eq 0) {
            $owners += [string]$ownerId
        }
    }
    $AddressOwners[$normalized] = $owners
}

function Test-AddressOwnedOnlyBy {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary] $AddressOwners,
        [Parameter(Mandatory)][string] $Address,
        [Parameter(Mandatory)][string] $OwnerId
    )

    $owners = @(Get-AddressOwnerIds $AddressOwners[$Address])
    if ($owners.Count -eq 0) { return $false }
    foreach ($owner in $owners) {
        if (-not [string]::Equals([string]$owner, $OwnerId, [StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
    }
    return $true
}

function Test-ValidStudentUpnCandidate {
    param(
        [Parameter(Mandatory)][object] $Student,
        [Parameter(Mandatory)][string] $UserPrincipalName,
        [Parameter(Mandatory)][string] $Domain
    )

    foreach ($candidate in @(Get-UpnCandidates -GivenName $Student.GivenName -Surname $Student.Surname -Domain $Domain)) {
        if ([string]::Equals($candidate, $UserPrincipalName.Trim(), [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    $localPart = "$(ConvertTo-UpnToken $Student.GivenName)$(ConvertTo-UpnToken $Student.Surname)"
    $pattern = '^{0}(?<suffix>[2-9][0-9]*)@{1}$' -f [regex]::Escape($localPart), [regex]::Escape($Domain.Trim())
    return $UserPrincipalName.Trim() -match "(?i)$pattern"
}

function Get-ComparisonGroupMatches {
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $DisplayName
    )

    $index = Get-ComparisonPropertyValue -InputObject $Snapshot -Name GroupMatchesByDisplayName
    if ($null -ne $index) {
        return @($index[$DisplayName])
    }
    $singleIndex = Get-ComparisonPropertyValue -InputObject $Snapshot -Name GroupsByDisplayName
    if ($null -ne $singleIndex -and $null -ne $singleIndex[$DisplayName]) {
        return @($singleIndex[$DisplayName])
    }
    return @()
}

function Test-ComparisonGroupDynamic {
    param([Parameter(Mandatory)][object] $Group)

    $value = Get-ComparisonPropertyValue -InputObject $Group -Name IsDynamic
    if ($null -ne $value) { return [bool]$value }
    $isDynamic = Get-EntraGroupIsDynamic -Group $Group
    return $isDynamic
}

function Test-ManagedStudentGroupName {
    param(
        [Parameter(Mandatory)][string] $DisplayName,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config
    )

    return $DisplayName.StartsWith([string]$Config.RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        $DisplayName.StartsWith([string]$Config.ClassGroupPrefix, [StringComparison]::OrdinalIgnoreCase)
}

function Compare-StudentDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Students,
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][System.Collections.IDictionary] $Config,
        [System.Collections.IDictionary] $ExchangeAddressOwners = @{}
    )

    $errors = [Collections.Generic.List[object]]::new()
    $warnings = [Collections.Generic.List[object]]::new()
    $newStudents = [Collections.Generic.List[object]]::new()
    $departures = [Collections.Generic.List[object]]::new()
    $changedStudents = [Collections.Generic.List[object]]::new()
    $existingStudents = [Collections.Generic.List[object]]::new()
    $records = [Collections.Generic.List[object]]::new()
    $invalidIndices = [Collections.Generic.HashSet[int]]::new()
    $hasIdentityAmbiguity = $false

    $reservedAddresses = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $addressOwners = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($address in @($Snapshot.ReservedAddresses)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$address)) {
            Add-ComparisonAddressOwner -AddressOwners $addressOwners -ReservedAddresses $reservedAddresses -Address ([string]$address) -Owner $Snapshot.AddressOwners[$address]
        }
    }
    foreach ($address in @($Snapshot.AddressOwners.Keys)) {
        Add-ComparisonAddressOwner -AddressOwners $addressOwners -ReservedAddresses $reservedAddresses -Address ([string]$address) -Owner $Snapshot.AddressOwners[$address]
    }
    foreach ($address in @($ExchangeAddressOwners.Keys)) {
        Add-ComparisonAddressOwner -AddressOwners $addressOwners -ReservedAddresses $reservedAddresses -Address ([string]$address) -Owner $ExchangeAddressOwners[$address]
    }
    foreach ($address in @($addressOwners.Keys)) {
        $uniqueOwners = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($ownerId in @(Get-AddressOwnerIds $addressOwners[$address])) {
            if (-not [string]::IsNullOrWhiteSpace([string]$ownerId)) { [void]$uniqueOwners.Add([string]$ownerId) }
        }
        if ($uniqueOwners.Count -gt 1) {
            $hasIdentityAmbiguity = $true
            $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field AddressOwnership -Current @($uniqueOwners) -Desired 'exactly one owner' -Message "Die Adresse '$address' gehört mehreren Verzeichnisobjekten."))
        }
    }

    $objectIdRows = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    $storedUpnRows = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    $nameRows = [hashtable]::new([StringComparer]::Ordinal)
    for ($index = 0; $index -lt $Students.Count; $index++) {
        $student = $Students[$index]
        $objectId = ([string](Get-ComparisonPropertyValue -InputObject $student -Name EntraObjectId)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($objectId)) { $objectIdRows[$objectId] = @($objectIdRows[$objectId]) + $index }
        $storedUpn = ([string](Get-ComparisonPropertyValue -InputObject $student -Name StoredUpn)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($storedUpn)) { $storedUpnRows[$storedUpn] = @($storedUpnRows[$storedUpn]) + $index }
        try {
            $nameKey = Get-NormalizedStudentNameKey -Student $student
            $nameRows[$nameKey] = @($nameRows[$nameKey]) + $index
        } catch {
            [void]$invalidIndices.Add($index)
            $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field NormalizedName -Current $null -Desired 'eindeutiger normalisierter Name' -Message $_.Exception.Message -Student $student))
        }
    }
    foreach ($definition in @(
            [pscustomobject]@{ Field = 'EntraObjectId'; Values = $objectIdRows }
            [pscustomobject]@{ Field = 'StoredUpn'; Values = $storedUpnRows }
            [pscustomobject]@{ Field = 'NormalizedName'; Values = $nameRows }
        )) {
        foreach ($key in @($definition.Values.Keys)) {
            $indices = @($definition.Values[$key])
            if ($indices.Count -le 1) { continue }
            $hasIdentityAmbiguity = $true
            foreach ($index in $indices) { [void]$invalidIndices.Add([int]$index) }
            $rows = @($indices | ForEach-Object { Get-ComparisonPropertyValue -InputObject $Students[$_] -Name RowNumber })
            $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field $definition.Field -Current $key -Desired 'eindeutig pro Excel-Zeile' -Message "Doppelte Excel-Identität '$key' in Zeilen $($rows -join ', ')."))
        }
    }

    for ($index = 0; $index -lt $Students.Count; $index++) {
        if ($invalidIndices.Contains($index)) { continue }
        $student = $Students[$index]
        try {
            $identity = Resolve-StudentIdentity -Student $student -Snapshot $Snapshot
            $records.Add([pscustomobject]@{
                    Index = $index
                    Student = $student
                    Identity = $identity
                    IsValid = $true
                })
        } catch {
            $hasIdentityAmbiguity = $true
            [void]$invalidIndices.Add($index)
            $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field Identity -Current $null -Desired 'eindeutige Entra-Zuordnung' -Message $_.Exception.Message -Student $student))
        }
    }

    $claims = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $records) {
        if ($null -eq $record.Identity.User) { continue }
        $userId = [string](Get-ComparisonPropertyValue -InputObject $record.Identity.User -Name Id)
        $claims[$userId] = @($claims[$userId]) + $record
    }
    foreach ($userId in @($claims.Keys)) {
        $claimants = @($claims[$userId])
        if ($claimants.Count -le 1) { continue }
        $hasIdentityAmbiguity = $true
        foreach ($record in $claimants) {
            $record.IsValid = $false
            [void]$invalidIndices.Add([int]$record.Index)
        }
        $rows = @($claimants | ForEach-Object { $_.Student.RowNumber })
        $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field ResolvedEntraObjectId -Current $userId -Desired 'höchstens eine Excel-Zeile' -Message "Mehrere Excel-Zeilen ($($rows -join ', ')) beanspruchen dasselbe Entra-Objekt '$userId'."))
    }

    $claimedUserIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $records) {
        if (-not $record.IsValid -or $null -eq $record.Identity.User) { continue }
        [void]$claimedUserIds.Add([string]$record.Identity.User.Id)
    }

    foreach ($record in @($records | Sort-Object Index)) {
        if (-not $record.IsValid) { continue }
        $student = $record.Student
        $user = $record.Identity.User
        $userId = if ($null -eq $user) { '' } else { [string]$user.Id }

        if ($null -eq $user) {
            $outsideOwners = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($candidate in @(Get-UpnCandidates -GivenName $student.GivenName -Surname $student.Surname -Domain $Config.Domain)) {
                foreach ($ownerId in @(Get-AddressOwnerIds $addressOwners[$candidate])) {
                    if (([string]$ownerId).StartsWith('excel-row-', [StringComparison]::OrdinalIgnoreCase)) {
                        continue
                    }
                    if (-not $Snapshot.StudentRoleMemberIds.Contains([string]$ownerId)) {
                        [void]$outsideOwners.Add([string]$ownerId)
                    }
                }
            }
            if ($outsideOwners.Count -gt 0) {
                $hasIdentityAmbiguity = $true
                $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field UserPrincipalName -Current @($outsideOwners) -Desired 'freie Schüleradresse' -Message 'Mindestens ein UPN-Kandidat gehört einem Verzeichnisobjekt außerhalb der Schüler-Rollengruppe.' -Student $student))
                continue
            }
        }

        try {
            [void](Get-OfficeLocation ([string]$student.ClassName))
        } catch {
            $errors.Add((New-ComparisonIssue -Severity Error -Area Entra -Field OfficeLocation -Current $null -Desired $student.ClassName -Message $_.Exception.Message -Student $student -User $user))
            continue
        }

        $manager = $null
        try {
            $manager = Resolve-StudentManager -Snapshot $Snapshot -Teacher ([string]$student.Teacher) -Student "$($student.GivenName) $($student.Surname)"
        } catch {
            $errors.Add((New-ComparisonIssue -Severity Error -Area Manager -Field ManagerId -Current $null -Desired $student.Teacher -Message $_.Exception.Message -Student $student -User $user))
            continue
        }

        $requiredGroupNames = @(
            [string]$Config.StudentRoleGroup.Name
            [string]$Config.LicenseGroupName
            "$($Config.ClassGroupPrefix)$($student.ClassName)"
        )
        $groupValidationFailed = $false
        foreach ($groupName in $requiredGroupNames) {
            $matches = @(Get-ComparisonGroupMatches -Snapshot $Snapshot -DisplayName $groupName)
            if ($matches.Count -ne 1) {
                $groupValidationFailed = $true
                $description = if ($matches.Count -eq 0) { 'fehlt' } else { "ist mehrdeutig ($($matches.Count) Treffer)" }
                $errors.Add((New-ComparisonIssue -Severity Error -Area Group -Field $groupName -Current $matches.Count -Desired 1 -Message "Die Pflichtgruppe '$groupName' $description." -Student $student -User $user))
                continue
            }
            if (Test-ComparisonGroupDynamic -Group $matches[0]) {
                $groupValidationFailed = $true
                $errors.Add((New-ComparisonIssue -Severity Error -Area Group -Field $groupName -Current 'Dynamic' -Desired 'StaticAssigned' -Message "Die Pflichtgruppe '$groupName' ist dynamisch und kann nicht als statisches Ziel verwaltet werden." -Student $student -User $user))
            }
        }
        if ($groupValidationFailed) { continue }

        $selectedUpn = $null
        $selection = $null
        $namesUnchanged = $false
        if ($null -ne $user) {
            try {
                $namesUnchanged = [string]::Equals(
                    (Get-NormalizedStudentNameKey -Student $student),
                    (Get-NormalizedStudentNameKey -Student $user),
                    [StringComparison]::Ordinal
                )
            } catch {
                $namesUnchanged = $false
            }
            $currentUpn = ([string](Get-ComparisonPropertyValue -InputObject $user -Name UserPrincipalName)).Trim()
            if ($namesUnchanged -and
                -not [string]::IsNullOrWhiteSpace($currentUpn) -and
                (Test-ValidStudentUpnCandidate -Student $student -UserPrincipalName $currentUpn -Domain $Config.Domain) -and
                (Test-AddressOwnedOnlyBy -AddressOwners $addressOwners -Address $currentUpn -OwnerId $userId)) {
                $selectedUpn = $currentUpn.ToLowerInvariant()
            }
        }
        if ([string]::IsNullOrWhiteSpace($selectedUpn)) {
            $selectionReserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($address in $reservedAddresses) {
                if (-not [string]::IsNullOrWhiteSpace($userId) -and
                    (Test-AddressOwnedOnlyBy -AddressOwners $addressOwners -Address $address -OwnerId $userId)) {
                    continue
                }
                [void]$selectionReserved.Add($address)
            }
            $selection = Select-AvailableUpn -GivenName $student.GivenName -Surname $student.Surname -Domain $Config.Domain -UsedAddresses $selectionReserved -AddressOwners $addressOwners -CurrentObjectId $userId
            $selectedUpn = $selection.Upn
        }

        $reservationOwner = if ([string]::IsNullOrWhiteSpace($userId)) { "excel-row-$($student.RowNumber)" } else { $userId }
        Add-ComparisonAddressOwner -AddressOwners $addressOwners -ReservedAddresses $reservedAddresses -Address $selectedUpn -Owner $reservationOwner
        $firstCandidate = @(Get-UpnCandidates -GivenName $student.GivenName -Surname $student.Surname -Domain $Config.Domain)[0]
        if (-not [string]::Equals($selectedUpn, $firstCandidate, [StringComparison]::OrdinalIgnoreCase)) {
            $collisions = if ($null -eq $selection) { @() } else { @($selection.Collisions) }
            $warnings.Add((New-ComparisonIssue -Severity Warning -Area Entra -Field UserPrincipalName -Current $collisions -Desired $selectedUpn -Message "UPN-Fallback auf '$selectedUpn' nach $($collisions.Count) aktueller Kollision(en)." -Student $student -User $user))
        }

        $desired = New-StudentDesiredState -Student $student -SelectedUpn $selectedUpn -Config $Config -Manager $manager
        if ($null -eq $user) {
            $newStudents.Add([pscustomobject]@{
                    Student = $student
                    User = $null
                    DesiredState = $desired
                    IdentityMethod = $record.Identity.Method
                    Differences = @()
                })
            continue
        }

        $differences = [Collections.Generic.List[object]]::new()
        foreach ($field in @(
                'DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'Mail', 'MailNickname',
                'Department', 'OfficeLocation', 'CompanyName', 'EmployeeType', 'UsageLocation',
                'AgeGroup', 'ConsentProvidedForMinor'
            )) {
            if ($field -eq 'MailNickname' -and $null -eq $user.PSObject.Properties[$field]) {
                continue
            }
            $difference = New-StateDifference -Area Entra -Field $field -Current (Get-ComparisonPropertyValue -InputObject $user -Name $field) -Desired (Get-ComparisonPropertyValue -InputObject $desired -Name $field)
            if ($null -ne $difference) { $differences.Add($difference) }
        }
        if ([string]::IsNullOrWhiteSpace([string](Get-ComparisonPropertyValue -InputObject $user -Name Mail))) {
            $warnings.Add((New-ComparisonIssue -Severity Warning -Area Entra -Field Mail -Current $null -Desired $desired.Mail -Message 'Das Entra-Attribut mail fehlt.' -Student $student -User $user))
        }
        $legalDifference = New-StateDifference -Area Entra -Field LegalAgeGroupClassification -Current (Get-ComparisonPropertyValue -InputObject $user -Name LegalAgeGroupClassification) -Desired $desired.LegalAgeGroupClassification -Action Compare
        if ($null -ne $legalDifference) {
            $differences.Add($legalDifference)
            $warnings.Add((New-ComparisonIssue -Severity Warning -Area Entra -Field LegalAgeGroupClassification -Current $legalDifference.Current -Desired $legalDifference.Desired -Message 'legalAgeGroupClassification wird nur verglichen und nicht geschrieben.' -Student $student -User $user))
        }

        $currentManagerId = Get-UserManagerId -Snapshot $Snapshot -UserId $userId
        $managerDifference = New-StateDifference -Area Manager -Field ManagerId -Current $currentManagerId -Desired $desired.ManagerId
        if ($null -ne $managerDifference) { $differences.Add($managerDifference) }

        $directGroups = @(Get-UserDirectGroups -Snapshot $Snapshot -UserId $userId)
        $transitiveGroups = @(Get-UserTransitiveGroups -Snapshot $Snapshot -UserId $userId)
        foreach ($requiredGroupName in $desired.RequiredGroupNames) {
            $isDirectMember = @($directGroups | Where-Object {
                    [string]::Equals([string]$_.DisplayName, [string]$requiredGroupName, [StringComparison]::OrdinalIgnoreCase)
                }).Count -gt 0
            if (-not $isDirectMember) {
                $differences.Add((New-StateDifference -Area Group -Field Membership -Current $null -Desired $requiredGroupName -Action Add))
            }
        }
        foreach ($group in $directGroups) {
            $groupName = [string](Get-ComparisonPropertyValue -InputObject $group -Name DisplayName)
            if (-not (Test-ManagedStudentGroupName -DisplayName $groupName -Config $Config)) { continue }
            $isDesired = @($desired.RequiredGroupNames | Where-Object {
                    [string]::Equals([string]$_, $groupName, [StringComparison]::OrdinalIgnoreCase)
                }).Count -gt 0
            if ($isDesired) { continue }
            $isDynamic = Test-ComparisonGroupDynamic -Group $group
            $isInherited = [bool](Get-ComparisonPropertyValue -InputObject $group -Name IsInherited)
            $isRemovableValue = Get-ComparisonPropertyValue -InputObject $group -Name IsRemovable
            $isRemovable = if ($null -eq $isRemovableValue) { -not $isDynamic -and -not $isInherited } else { [bool]$isRemovableValue }
            if ($isRemovable -and -not $isDynamic -and -not $isInherited) {
                $differences.Add((New-StateDifference -Area Group -Field Membership -Current $groupName -Desired $null -Action Remove))
            } else {
                $warnings.Add((New-ComparisonIssue -Severity Warning -Area Group -Field ManagedGroup -Current $groupName -Desired $null -Message "Die direkte verwaltete Gruppe '$groupName' ist dynamisch oder geerbt und wird nicht entfernt." -Student $student -User $user))
            }
        }
        foreach ($group in $transitiveGroups) {
            $groupName = [string](Get-ComparisonPropertyValue -InputObject $group -Name DisplayName)
            if (-not (Test-ManagedStudentGroupName -DisplayName $groupName -Config $Config)) { continue }
            $isDesired = @($desired.RequiredGroupNames | Where-Object {
                    [string]::Equals([string]$_, $groupName, [StringComparison]::OrdinalIgnoreCase)
                }).Count -gt 0
            if ($isDesired) { continue }
            $isDirect = @($directGroups | Where-Object {
                    [string]::Equals([string]$_.Id, [string]$group.Id, [StringComparison]::OrdinalIgnoreCase)
                }).Count -gt 0
            if ($isDirect) { continue }
            $warnings.Add((New-ComparisonIssue -Severity Warning -Area Group -Field ManagedGroup -Current $groupName -Desired $null -Message "Die geerbte verwaltete Gruppe '$groupName' wird nicht entfernt." -Student $student -User $user))
        }

        $entry = [pscustomobject]@{
            Student = $student
            User = $user
            DesiredState = $desired
            IdentityMethod = $record.Identity.Method
            Differences = [object[]]@($differences)
        }
        if ($differences.Count -gt 0) {
            $changedStudents.Add($entry)
        } else {
            $existingStudents.Add($entry)
        }
    }

    if (-not $hasIdentityAmbiguity) {
        foreach ($memberId in @($Snapshot.StudentRoleMemberIds)) {
            if ($claimedUserIds.Contains([string]$memberId)) { continue }
            $user = $Snapshot.UsersById[[string]$memberId]
            if ($null -eq $user) {
                $errors.Add((New-ComparisonIssue -Severity Error -Area Identity -Field Departure -Current $memberId -Desired 'vorhandener Entra-Benutzer' -Message "Das direkte Schüler-Rollenmitglied '$memberId' fehlt im Benutzer-Snapshot."))
                continue
            }
            $departures.Add([pscustomobject]@{
                    Student = $null
                    User = $user
                    DesiredState = $null
                    IdentityMethod = 'UnclaimedStudentRoleMember'
                    Differences = @()
                })
        }
    } else {
        $newStudents.Clear()
        $departures.Clear()
    }

    return [pscustomobject]@{
        NewStudents = [object[]]@($newStudents)
        Departures = [object[]]@($departures)
        ChangedStudents = [object[]]@($changedStudents)
        ExistingStudents = [object[]]@($existingStudents)
        Warnings = [object[]]@($warnings)
        Errors = [object[]]@($errors)
    }
}

function Add-ExchangeComparison {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Comparison,
        [Parameter(Mandatory)][string] $UserId,
        [Parameter(Mandatory)][object] $MailboxState
    )

    $entry = @($Comparison.ChangedStudents | Where-Object {
            [string]::Equals([string]$_.User.Id, $UserId, [StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1)
    $wasExisting = $false
    if ($entry.Count -eq 0) {
        $entry = @($Comparison.ExistingStudents | Where-Object {
                [string]::Equals([string]$_.User.Id, $UserId, [StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1)
        $wasExisting = $entry.Count -eq 1
    }
    if ($entry.Count -ne 1) {
        throw "Für Entra-Benutzer '$UserId' existiert kein eindeutiger aktiver Vergleichseintrag."
    }
    $studentEntry = $entry[0]

    $existsValue = Get-ComparisonPropertyValue -InputObject $MailboxState -Name Exists
    $status = [string](Get-ComparisonPropertyValue -InputObject $MailboxState -Name Status)
    $mailboxNotReady = ($null -ne $existsValue -and -not [bool]$existsValue) -or
        [string]::Equals($status, 'MailboxNotReady', [StringComparison]::OrdinalIgnoreCase)
    if ($mailboxNotReady) {
        $warning = New-ComparisonIssue -Severity Warning -Area Exchange -Field Mailbox -Current $null -Desired 'Bereit' -Message "EXO-Konfiguration ausstehend für '$($studentEntry.DesiredState.UserPrincipalName)'." -Student $studentEntry.Student -User $studentEntry.User
        $Comparison.Warnings = [object[]]@($Comparison.Warnings) + $warning
        return $Comparison
    }

    $exchangeDifferences = [Collections.Generic.List[object]]::new()
    foreach ($difference in @((Get-ComparisonPropertyValue -InputObject $MailboxState -Name Differences))) {
        if ($null -eq $difference) { continue }
        $action = [string](Get-ComparisonPropertyValue -InputObject $difference -Name Action)
        if ([string]::IsNullOrWhiteSpace($action)) { $action = 'Set' }
        $normalized = New-StateDifference -Area Exchange -Field ([string](Get-ComparisonPropertyValue -InputObject $difference -Name Field)) -Current (Get-ComparisonPropertyValue -InputObject $difference -Name Current) -Desired (Get-ComparisonPropertyValue -InputObject $difference -Name Desired) -Action $action
        if ($null -ne $normalized) { $exchangeDifferences.Add($normalized) }
    }
    if ($exchangeDifferences.Count -eq 0) { return $Comparison }

    $studentEntry.Differences = [object[]]@($studentEntry.Differences) + [object[]]@($exchangeDifferences)
    if ($wasExisting) {
        $Comparison.ExistingStudents = [object[]]@($Comparison.ExistingStudents | Where-Object {
                -not [string]::Equals([string]$_.User.Id, $UserId, [StringComparison]::OrdinalIgnoreCase)
            })
        $Comparison.ChangedStudents = [object[]]@($Comparison.ChangedStudents) + $studentEntry
    }
    return $Comparison
}
