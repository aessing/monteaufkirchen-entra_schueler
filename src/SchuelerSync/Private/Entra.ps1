function New-CaseInsensitiveHashtable {
    return [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
}

function Test-GraphScopeSet {
    param(
        [AllowNull()][string[]] $GrantedScopes,
        [Parameter(Mandatory)][string[]] $RequiredScopes
    )

    $granted = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($scope in @($GrantedScopes)) {
        if (-not [string]::IsNullOrWhiteSpace($scope)) {
            [void] $granted.Add($scope)
        }
    }
    return @($RequiredScopes | Where-Object { -not $granted.Contains($_) })
}

function Connect-SchuelerGraph {
    [CmdletBinding()]
    param()

    $requiredScopes = @(
        'User.ReadWrite.All'
        'User-Mail.ReadWrite.All'
        'Group.Read.All'
        'GroupMember.ReadWrite.All'
        'User.RevokeSessions.All'
    )

    $context = Get-MgContext
    $contextScopes = if ($null -eq $context) { @() } else { @($context.Scopes) }
    $missingScopes = @(Test-GraphScopeSet -GrantedScopes $contextScopes -RequiredScopes $requiredScopes)
    if ($missingScopes.Count -eq 0) {
        return $context
    }

    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    $verifiedContext = Get-MgContext
    $verifiedScopes = if ($null -eq $verifiedContext) { @() } else { @($verifiedContext.Scopes) }
    $missingScopes = @(Test-GraphScopeSet -GrantedScopes $verifiedScopes -RequiredScopes $requiredScopes)
    if ($missingScopes.Count -gt 0) {
        throw "Microsoft Graph context lacks required scopes: $($missingScopes -join ', ')."
    }

    return $verifiedContext
}

function Get-EntraGroupIsDynamic {
    param([Parameter(Mandatory)][object] $Group)

    $hasDynamicType = @($Group.GroupTypes | Where-Object {
            [string]::Equals([string]$_, 'DynamicMembership', [StringComparison]::OrdinalIgnoreCase)
        }).Count -gt 0
    return $hasDynamicType -or -not [string]::IsNullOrWhiteSpace([string]$Group.MembershipRule)
}

function Set-EntraGroupMetadata {
    param(
        [Parameter(Mandatory)][object] $Group,
        [bool] $IsInherited = $false
    )

    $isDynamic = Get-EntraGroupIsDynamic -Group $Group
    Add-Member -InputObject $Group -MemberType NoteProperty -Name IsDynamic -Value $isDynamic -Force
    Add-Member -InputObject $Group -MemberType NoteProperty -Name IsInherited -Value $IsInherited -Force
    Add-Member -InputObject $Group -MemberType NoteProperty -Name IsRemovable -Value (-not $isDynamic -and -not $IsInherited) -Force
    return $Group
}

function Add-EntraReservedAddress {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]] $ReservedAddresses,
        [Parameter(Mandatory)][System.Collections.IDictionary] $AddressOwners,
        [Parameter(Mandatory)][string] $Address,
        [Parameter(Mandatory)][string] $OwnerId,
        [switch] $IsProxyAddress
    )

    $value = $Address.Trim()
    if ($IsProxyAddress) {
        if ($value -notmatch '(?i)^smtp:') { return }
        $value = ($value -replace '(?i)^smtp:', '').Trim()
    }
    if ([string]::IsNullOrWhiteSpace($value)) { return }

    $normalized = $value.ToLowerInvariant()
    [void] $ReservedAddresses.Add($normalized)
    $owners = @($AddressOwners[$normalized])
    if (@($owners | Where-Object {
                [string]::Equals([string]$_, $OwnerId, [StringComparison]::OrdinalIgnoreCase)
            }).Count -eq 0) {
        $AddressOwners[$normalized] = @($owners + $OwnerId)
    }
}

function Resolve-EntraSnapshotGroup {
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $DisplayName
    )

    $groups = @($Snapshot.GroupMatchesByDisplayName[$DisplayName])
    if ($groups.Count -ne 1) {
        throw "Expected exactly one Entra group named '$DisplayName', found $($groups.Count)."
    }
    return $groups[0]
}

function Get-EntraSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Collections.IDictionary] $Config)

    $userProperties = @(
        'id', 'displayName', 'givenName', 'surname', 'userPrincipalName', 'mail', 'mailNickname', 'proxyAddresses',
        'department', 'officeLocation', 'companyName', 'employeeType', 'usageLocation', 'ageGroup',
        'consentProvidedForMinor', 'legalAgeGroupClassification', 'accountEnabled', 'userType'
    )
    $groupProperties = @(
        'id', 'displayName', 'groupTypes', 'membershipRule', 'membershipRuleProcessingState'
    )

    $users = @(Get-MgUser -All -Property $userProperties)
    $groups = @(Get-MgGroup -All -Property $groupProperties)

    $usersById = New-CaseInsensitiveHashtable
    $usersByUpn = New-CaseInsensitiveHashtable
    $reservedAddresses = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $addressOwners = New-CaseInsensitiveHashtable
    foreach ($user in $users) {
        $userId = ([string]$user.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($userId)) {
            throw 'Microsoft Graph returned a user without an id.'
        }
        if ($usersById.ContainsKey($userId)) {
            throw "Microsoft Graph returned duplicate user id '$userId'."
        }
        $usersById[$userId] = $user

        $upn = ([string]$user.UserPrincipalName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            if ($usersByUpn.ContainsKey($upn)) {
                throw "Microsoft Graph returned duplicate user principal name '$upn'."
            }
            $usersByUpn[$upn] = $user
            Add-EntraReservedAddress -ReservedAddresses $reservedAddresses -AddressOwners $addressOwners -Address $upn -OwnerId $userId
        }
        $mail = ([string]$user.Mail).Trim()
        if (-not [string]::IsNullOrWhiteSpace($mail)) {
            Add-EntraReservedAddress -ReservedAddresses $reservedAddresses -AddressOwners $addressOwners -Address $mail -OwnerId $userId
        }
        foreach ($proxyAddress in @($user.ProxyAddresses)) {
            if ($null -ne $proxyAddress) {
                Add-EntraReservedAddress -ReservedAddresses $reservedAddresses -AddressOwners $addressOwners -Address ([string]$proxyAddress) -OwnerId $userId -IsProxyAddress
            }
        }
    }

    $groupsById = New-CaseInsensitiveHashtable
    $groupMatchesByDisplayName = New-CaseInsensitiveHashtable
    foreach ($group in $groups) {
        $groupId = ([string]$group.Id).Trim()
        if ([string]::IsNullOrWhiteSpace($groupId)) {
            throw 'Microsoft Graph returned a group without an id.'
        }
        if ($groupsById.ContainsKey($groupId)) {
            throw "Microsoft Graph returned duplicate group id '$groupId'."
        }
        $group = Set-EntraGroupMetadata -Group $group
        $groupsById[$groupId] = $group
        $displayName = ([string]$group.DisplayName).Trim()
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            $groupMatchesByDisplayName[$displayName] = @($groupMatchesByDisplayName[$displayName]) + $group
        }
    }

    $groupsByDisplayName = New-CaseInsensitiveHashtable
    foreach ($displayName in $groupMatchesByDisplayName.Keys) {
        $matches = @($groupMatchesByDisplayName[$displayName])
        if ($matches.Count -eq 1) {
            $groupsByDisplayName[$displayName] = $matches[0]
        }
    }

    $roleConfig = $Config.StudentRoleGroup
    $roleGroupId = ([string]$roleConfig.Id).Trim()
    $roleGroupName = ([string]$roleConfig.Name).Trim()
    if ([string]::IsNullOrWhiteSpace($roleGroupId) -or [string]::IsNullOrWhiteSpace($roleGroupName)) {
        throw 'StudentRoleGroup must define both Id and Name.'
    }
    $configuredRoleGroup = $groupsById[$roleGroupId]
    if ($null -eq $configuredRoleGroup -or -not [string]::Equals(
            [string]$configuredRoleGroup.DisplayName, $roleGroupName, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Configured student role group '$roleGroupId' does not match '$roleGroupName'."
    }
    $resolvedRoleGroup = Resolve-EntraSnapshotGroup -Snapshot ([pscustomobject]@{ GroupMatchesByDisplayName = $groupMatchesByDisplayName }) -DisplayName $roleGroupName
    if (-not [string]::Equals([string]$resolvedRoleGroup.Id, $roleGroupId, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Configured student role group name '$roleGroupName' resolves to a different id."
    }

    $licenseGroupName = ([string]$Config.LicenseGroupName).Trim()
    if ([string]::IsNullOrWhiteSpace($licenseGroupName)) {
        throw 'LicenseGroupName must not be empty.'
    }
    [void](Resolve-EntraSnapshotGroup -Snapshot ([pscustomobject]@{ GroupMatchesByDisplayName = $groupMatchesByDisplayName }) -DisplayName $licenseGroupName)

    $classGroupPrefix = ([string]$Config.ClassGroupPrefix).Trim()
    if (-not [string]::IsNullOrWhiteSpace($classGroupPrefix)) {
        foreach ($displayName in $groupMatchesByDisplayName.Keys) {
            if ($displayName.StartsWith($classGroupPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                [void](Resolve-EntraSnapshotGroup -Snapshot ([pscustomobject]@{ GroupMatchesByDisplayName = $groupMatchesByDisplayName }) -DisplayName $displayName)
            }
        }
    }

    $studentRoleMemberIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($member in @(Get-MgGroupMember -GroupId $roleGroupId -All)) {
        $memberId = ([string]$member.Id).Trim()
        if (-not [string]::IsNullOrWhiteSpace($memberId)) {
            [void] $studentRoleMemberIds.Add($memberId)
        }
    }

    return [pscustomobject]@{
        UsersById = $usersById
        UsersByUpn = $usersByUpn
        StudentRoleMemberIds = $studentRoleMemberIds
        GroupsById = $groupsById
        GroupsByDisplayName = $groupsByDisplayName
        ReservedAddresses = $reservedAddresses
        AddressOwners = $addressOwners
        DirectGroupsByUserId = New-CaseInsensitiveHashtable
        TransitiveGroupsByUserId = New-CaseInsensitiveHashtable
        ManagerByUserId = New-CaseInsensitiveHashtable
        GroupMatchesByDisplayName = $groupMatchesByDisplayName
    }
}

function Get-UserDirectGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $UserId
    )

    if (-not $Snapshot.DirectGroupsByUserId.ContainsKey($UserId)) {
        $groups = @(
            Get-MgUserMemberOfAsGroup -UserId $UserId -All |
                ForEach-Object { Set-EntraGroupMetadata -Group $_ }
        )
        $Snapshot.DirectGroupsByUserId[$UserId] = $groups
    }
    return @($Snapshot.DirectGroupsByUserId[$UserId])
}

function Get-UserTransitiveGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $UserId
    )

    if (-not $Snapshot.TransitiveGroupsByUserId.ContainsKey($UserId)) {
        $directGroupIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($group in @(Get-UserDirectGroups -Snapshot $Snapshot -UserId $UserId)) {
            [void] $directGroupIds.Add([string]$group.Id)
        }
        $groups = @(
            Get-MgUserTransitiveMemberOfAsGroup -UserId $UserId -All |
                ForEach-Object {
                    Set-EntraGroupMetadata -Group $_ -IsInherited (-not $directGroupIds.Contains([string]$_.Id))
                }
        )
        $Snapshot.TransitiveGroupsByUserId[$UserId] = $groups
    }
    return @($Snapshot.TransitiveGroupsByUserId[$UserId])
}

function Get-UserManagerId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $UserId
    )

    if (-not $Snapshot.ManagerByUserId.ContainsKey($UserId)) {
        $manager = Get-MgUserManager -UserId $UserId
        $Snapshot.ManagerByUserId[$UserId] = if ($null -eq $manager) { $null } else { [string]$manager.Id }
    }
    return $Snapshot.ManagerByUserId[$UserId]
}

function Resolve-StudentManager {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $Teacher,
        [string] $Student
    )

    $value = $Teacher.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw 'Kein eindeutiger Manager: die Lehrkraft ist leer.'
    }

    $users = @($Snapshot.UsersById.Values)
    if ($value.Contains('@')) {
        $matches = @($users | Where-Object {
                [string]::Equals(([string]$_.UserPrincipalName).Trim(), $value, [StringComparison]::OrdinalIgnoreCase) -or
                [string]::Equals(([string]$_.Mail).Trim(), $value, [StringComparison]::OrdinalIgnoreCase)
            })
    } else {
        $matches = @($users | Where-Object {
                [string]::Equals(([string]$_.DisplayName).Trim(), $value, [StringComparison]::OrdinalIgnoreCase)
            })
    }

    if ($matches.Count -ne 1) {
        $studentSuffix = if ([string]::IsNullOrWhiteSpace($Student)) { '' } else { " for student '$Student'" }
        throw "Kein eindeutiger Manager$studentSuffix für '$value' gefunden ($($matches.Count) Treffer)."
    }
    return $matches[0]
}
