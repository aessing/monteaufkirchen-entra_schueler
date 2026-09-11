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
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]] $ReservedAddresses,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.IDictionary] $AddressOwners,
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
    $owners = @()
    if ($AddressOwners.Contains($normalized)) { $owners = @($AddressOwners[$normalized]) }
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

    $groups = @()
    if ($Snapshot.GroupMatchesByDisplayName.Contains($DisplayName)) {
        $groups = @($Snapshot.GroupMatchesByDisplayName[$DisplayName])
    }
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
            if (-not $groupMatchesByDisplayName.ContainsKey($displayName)) {
                $groupMatchesByDisplayName[$displayName] = @()
            }
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

function Test-EntraManagerNotFoundError {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord] $ErrorRecord)

    # Graph documents HTTP 404 for an unassigned manager. SDK generations expose
    # that status in different properties, so do not infer it from translated text.
    $exception = $ErrorRecord.Exception
    $notFound = $false
    while ($null -ne $exception) {
        $responseProperty = $exception.PSObject.Properties['Response']
        $response = if ($null -eq $responseProperty) { $null } else { $responseProperty.Value }
        foreach ($source in @($exception, $response)) {
            if ($null -eq $source) { continue }
            foreach ($name in @('StatusCode', 'ResponseStatusCode', 'HttpStatusCode')) {
                $property = $source.PSObject.Properties[$name]
                if ($null -eq $property -or $null -eq $property.Value) { continue }
                $status = 0
                if ($property.Value -is [Net.HttpStatusCode]) {
                    $status = [int]$property.Value
                } elseif (-not [int]::TryParse([string]$property.Value, [ref]$status)) {
                    continue
                }
                if ($status -ge 100 -and $status -le 599) {
                    if ($status -ne 404) { return $false }
                    $notFound = $true
                }
            }
        }
        $exception = $exception.InnerException
    }
    return $notFound
}

function Get-UserManagerId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object] $Snapshot,
        [Parameter(Mandatory)][string] $UserId
    )

    if (-not $Snapshot.ManagerByUserId.ContainsKey($UserId)) {
        try {
            $manager = Get-MgUserManager -UserId $UserId -ErrorAction Stop
        } catch {
            if (-not (Test-EntraManagerNotFoundError -ErrorRecord $_)) { throw }
            $Snapshot.ManagerByUserId[$UserId] = $null
            return $null
        }
        if ($null -eq $manager -or [string]::IsNullOrWhiteSpace([string]$manager.Id)) {
            throw "Die Manager-Abfrage für '$UserId' lieferte keine verifizierbare Objekt-ID."
        }
        $Snapshot.ManagerByUserId[$UserId] = [string]$manager.Id
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

function Get-EntraMutationPropertyValue {
    param(
        [AllowNull()][object] $InputObject,
        [Parameter(Mandatory)][string] $Name
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Test-EntraMutationValueEqual {
    param(
        [Parameter(Mandatory)][string] $Field,
        [AllowNull()][object] $Current,
        [AllowNull()][object] $Desired
    )

    if ($null -eq $Current -and $null -eq $Desired) { return $true }
    if ($null -eq $Current -or $null -eq $Desired) { return $false }
    if ($Field -in @('UserPrincipalName', 'Mail', 'MailNickname')) {
        return [string]::Equals(
            ([string]$Current).Trim(),
            ([string]$Desired).Trim(),
            [StringComparison]::OrdinalIgnoreCase
        )
    }
    if ($Current -is [string] -or $Desired -is [string]) {
        return [string]::Equals(
            ([string]$Current).Trim(),
            ([string]$Desired).Trim(),
            [StringComparison]::Ordinal
        )
    }
    return [object]::Equals($Current, $Desired)
}

function Test-EntraPasswordPolicyRejection {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord] $ErrorRecord)

    $messageParts = [Collections.Generic.List[string]]::new()
    if ($null -ne $ErrorRecord.Exception) {
        $messageParts.Add([string]$ErrorRecord.Exception.Message)
        foreach ($name in @('StatusCode', 'ResponseStatusCode')) {
            $property = $ErrorRecord.Exception.PSObject.Properties[$name]
            if ($null -ne $property) { $messageParts.Add([string]$property.Value) }
        }
    }
    if ($null -ne $ErrorRecord.ErrorDetails) {
        $messageParts.Add([string]$ErrorRecord.ErrorDetails.Message)
    }
    $message = $messageParts -join ' '
    $isBadRequest = $message -match '(?i)(Request_BadRequest|BadRequest|status\s*code\s*400|\b400\b)'
    $identifiesPassword = $message -match '(?i)(PasswordProfile\.Password|password)'
    $identifiesPolicy = $message -match '(?i)(policy|policies|complexity|requirements?|does not comply)'
    return $isBadRequest -and $identifiesPassword -and $identifiesPolicy
}

function New-DisabledEntraStudent {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    # New-MgUser requires an in-memory plain string. It is not logged and is persisted only to the protected workbook.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingPlainTextForPassword',
        'Password',
        Justification = 'New-MgUser requires an in-memory plain string that is not logged and is persisted only to the protected workbook.'
    )]
    param(
        [Parameter(Mandatory)][object] $Desired,
        [Parameter(Mandatory)][string] $Password
    )

    $userPrincipalName = [string](Get-EntraMutationPropertyValue -InputObject $Desired -Name UserPrincipalName)
    if ([string]::IsNullOrWhiteSpace($userPrincipalName) -or -not $userPrincipalName.Contains('@')) {
        throw 'The desired user principal name must contain a local part and domain.'
    }
    $body = @{
        AccountEnabled = $false
        DisplayName = (Get-EntraMutationPropertyValue -InputObject $Desired -Name DisplayName)
        GivenName = (Get-EntraMutationPropertyValue -InputObject $Desired -Name GivenName)
        Surname = (Get-EntraMutationPropertyValue -InputObject $Desired -Name Surname)
        UserPrincipalName = $userPrincipalName
        MailNickname = ($userPrincipalName -split '@')[0]
        Mail = (Get-EntraMutationPropertyValue -InputObject $Desired -Name Mail)
        Department = (Get-EntraMutationPropertyValue -InputObject $Desired -Name Department)
        OfficeLocation = (Get-EntraMutationPropertyValue -InputObject $Desired -Name OfficeLocation)
        CompanyName = (Get-EntraMutationPropertyValue -InputObject $Desired -Name CompanyName)
        EmployeeType = (Get-EntraMutationPropertyValue -InputObject $Desired -Name EmployeeType)
        UsageLocation = (Get-EntraMutationPropertyValue -InputObject $Desired -Name UsageLocation)
        AgeGroup = (Get-EntraMutationPropertyValue -InputObject $Desired -Name AgeGroup)
        ConsentProvidedForMinor = (Get-EntraMutationPropertyValue -InputObject $Desired -Name ConsentProvidedForMinor)
        PasswordProfile = @{
            Password = $Password
            ForceChangePasswordNextSignIn = $false
        }
    }

    if ($PSCmdlet.ShouldProcess($userPrincipalName, 'Create disabled Entra student')) {
        return New-MgUser -BodyParameter $body -ErrorAction Stop
    }
}

function New-StudentWithPasswordRetry {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][object] $Desired,
        [Parameter(Mandatory)][string] $InitialPassword,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]] $UsedPasswords
    )

    $nestedShouldProcessParameters = @{}
    foreach ($commonParameter in @('WhatIf', 'Confirm')) {
        if ($PSBoundParameters.ContainsKey($commonParameter)) {
            $nestedShouldProcessParameters[$commonParameter] = $PSBoundParameters[$commonParameter]
        }
    }
    $password = $InitialPassword
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            $user = New-DisabledEntraStudent -Desired $Desired -Password $password -ErrorAction Stop @nestedShouldProcessParameters
            if ($null -eq $user) { return }
            [void]$UsedPasswords.Add($password)
            return [pscustomobject]@{
                User = $user
                Password = $password
                Attempts = $attempt
            }
        } catch {
            if ($attempt -ge 6 -or -not (Test-EntraPasswordPolicyRejection -ErrorRecord $_)) {
                throw
            }
            [void]$UsedPasswords.Add($password)
            $password = New-StudentPassword -UsedPasswords $UsedPasswords
        }
    }
}

function Set-EntraStudentAttributes {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string] $UserId,
        [Parameter(Mandatory)][object] $Desired,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $Differences
    )

    $writableFields = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($field in @(
            'DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'Mail', 'MailNickname',
            'Department', 'OfficeLocation', 'CompanyName', 'EmployeeType', 'UsageLocation',
            'AgeGroup', 'ConsentProvidedForMinor'
        )) {
        $writableFields[$field] = $field
    }

    $body = @{}
    foreach ($difference in @($Differences)) {
        $area = [string](Get-EntraMutationPropertyValue -InputObject $difference -Name Area)
        $action = [string](Get-EntraMutationPropertyValue -InputObject $difference -Name Action)
        $field = [string](Get-EntraMutationPropertyValue -InputObject $difference -Name Field)
        if (-not [string]::Equals($area, 'Entra', [StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not [string]::Equals($action, 'Set', [StringComparison]::OrdinalIgnoreCase)) { continue }
        if ([string]::Equals($field, 'LegalAgeGroupClassification', [StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not $writableFields.ContainsKey($field)) { continue }

        $canonicalField = [string]$writableFields[$field]
        $body[$canonicalField] = Get-EntraMutationPropertyValue -InputObject $Desired -Name $canonicalField
    }

    if ($body.Count -eq 0) {
        return [pscustomobject]@{
            UserId = $UserId
            Changed = $false
            Verified = $true
            Fields = @()
            User = $null
        }
    }

    if (-not $PSCmdlet.ShouldProcess($UserId, "Update Entra student attributes: $($body.Keys -join ', ')")) {
        return [pscustomobject]@{
            UserId = $UserId
            Changed = $false
            Verified = $false
            Fields = [string[]]@($body.Keys)
            User = $null
        }
    }
    Update-MgUser -UserId $UserId -BodyParameter $body -ErrorAction Stop

    $properties = [string[]]@($body.Keys)
    $user = Get-MgUser -UserId $UserId -Property $properties -ErrorAction Stop
    foreach ($field in $properties) {
        $actual = Get-EntraMutationPropertyValue -InputObject $user -Name $field
        if (-not (Test-EntraMutationValueEqual -Field $field -Current $actual -Desired $body[$field])) {
            throw "Entra attribute '$field' could not be verified for user '$UserId'."
        }
    }
    return [pscustomobject]@{
        UserId = $UserId
        Changed = $true
        Verified = $true
        Fields = $properties
        User = $user
    }
}

function Set-EntraStudentManager {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string] $UserId,
        [AllowNull()][string] $CurrentManagerId,
        [Parameter(Mandatory)][string] $DesiredManagerId
    )

    $changed = -not [string]::Equals($CurrentManagerId, $DesiredManagerId, [StringComparison]::OrdinalIgnoreCase)
    if ($changed) {
        $body = @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/users/$DesiredManagerId"
        }
        if (-not $PSCmdlet.ShouldProcess($UserId, "Set Entra manager to '$DesiredManagerId'")) {
            return [pscustomobject]@{
                UserId = $UserId
                ManagerId = $CurrentManagerId
                Changed = $false
                Verified = $false
            }
        }
        Set-MgUserManagerByRef -UserId $UserId -BodyParameter $body -ErrorAction Stop
    }

    $manager = Get-MgUserManager -UserId $UserId -ErrorAction Stop
    $actualManagerId = [string](Get-EntraMutationPropertyValue -InputObject $manager -Name Id)
    if (-not [string]::Equals($actualManagerId, $DesiredManagerId, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Entra manager '$DesiredManagerId' could not be verified for user '$UserId'."
    }
    return [pscustomobject]@{
        UserId = $UserId
        ManagerId = $actualManagerId
        Changed = $changed
        Verified = $true
    }
}

function Get-FreshEntraUserDirectGroups {
    param([Parameter(Mandatory)][string] $UserId)

    return @(
        Get-MgUserMemberOfAsGroup -UserId $UserId -All -ErrorAction Stop |
            ForEach-Object {
                $inheritedProperty = $_.PSObject.Properties['IsInherited']
                $isInherited = $null -ne $inheritedProperty -and [bool]$inheritedProperty.Value
                Set-EntraGroupMetadata -Group $_ -IsInherited $isInherited
            }
    )
}

function Sync-EntraStudentGroups {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string] $UserId,
        [Parameter(Mandatory)][object] $RoleGroup,
        [Parameter(Mandatory)][object] $LicenseGroup,
        [Parameter(Mandatory)][object] $ClassGroup,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]] $CurrentDirectGroups,
        [string] $RoleGroupPrefix = 'SEC-A-ROL-',
        [string] $ClassGroupPrefix = 'SEC-A-CLS-'
    )

    $requiredGroups = @($RoleGroup, $LicenseGroup, $ClassGroup)
    $requiredIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($group in $requiredGroups) {
        $groupId = ([string](Get-EntraMutationPropertyValue -InputObject $group -Name Id)).Trim()
        $displayName = ([string](Get-EntraMutationPropertyValue -InputObject $group -Name DisplayName)).Trim()
        if ([string]::IsNullOrWhiteSpace($groupId) -or [string]::IsNullOrWhiteSpace($displayName)) {
            throw 'Every mandatory Entra group must define both Id and DisplayName.'
        }
        if (Get-EntraGroupIsDynamic -Group $group) {
            throw "Mandatory target group '$displayName' is dynamic and cannot be assigned directly."
        }
        [void]$requiredIds.Add($groupId)
    }
    if ($requiredIds.Count -ne 3) {
        throw 'Role, license and class target groups must have three distinct IDs.'
    }

    $currentIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($group in @($CurrentDirectGroups)) {
        $groupId = ([string](Get-EntraMutationPropertyValue -InputObject $group -Name Id)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($groupId)) { [void]$currentIds.Add($groupId) }
    }

    $memberReference = @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
    }
    $addedGroupIds = [Collections.Generic.List[string]]::new()
    $writeSkipped = $false
    foreach ($group in $requiredGroups) {
        $groupId = [string](Get-EntraMutationPropertyValue -InputObject $group -Name Id)
        if ($currentIds.Contains($groupId)) { continue }
        $displayName = [string](Get-EntraMutationPropertyValue -InputObject $group -Name DisplayName)
        if ($PSCmdlet.ShouldProcess($displayName, "Add Entra student '$UserId' to group")) {
            New-MgGroupMemberByRef -GroupId $groupId -BodyParameter $memberReference -ErrorAction Stop
            $addedGroupIds.Add($groupId)
        } else {
            $writeSkipped = $true
        }
    }
    if ($writeSkipped) {
        return [pscustomobject]@{
            UserId = $UserId
            AddedGroupIds = [string[]]@($addedGroupIds)
            RemovedGroupIds = @()
            DirectGroups = [object[]]@($CurrentDirectGroups)
            Verified = $false
        }
    }

    $afterAdds = @(Get-FreshEntraUserDirectGroups -UserId $UserId)
    $afterAddIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($group in $afterAdds) {
        $groupId = ([string](Get-EntraMutationPropertyValue -InputObject $group -Name Id)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($groupId)) { [void]$afterAddIds.Add($groupId) }
    }
    $missingTargetIds = @($requiredIds | Where-Object { -not $afterAddIds.Contains($_) })
    if ($missingTargetIds.Count -gt 0) {
        throw "Entra mandatory target groups could not be verified for user '$UserId': $($missingTargetIds -join ', ')."
    }

    $removedGroupIds = [Collections.Generic.List[string]]::new()
    $removalSkipped = $false
    foreach ($group in $afterAdds) {
        $groupId = [string](Get-EntraMutationPropertyValue -InputObject $group -Name Id)
        if ($requiredIds.Contains($groupId)) { continue }
        $displayName = [string](Get-EntraMutationPropertyValue -InputObject $group -Name DisplayName)
        $isManaged = $displayName.StartsWith($RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase) -or
            $displayName.StartsWith($ClassGroupPrefix, [StringComparison]::OrdinalIgnoreCase)
        $isInherited = [bool](Get-EntraMutationPropertyValue -InputObject $group -Name IsInherited)
        $isDynamic = Get-EntraGroupIsDynamic -Group $group
        if (-not $isManaged -or $isInherited -or $isDynamic) { continue }

        if ($PSCmdlet.ShouldProcess($displayName, "Remove Entra student '$UserId' from competing managed group")) {
            Remove-MgGroupMemberByRef -GroupId $groupId -DirectoryObjectId $UserId -ErrorAction Stop
            $removedGroupIds.Add($groupId)
        } else {
            $removalSkipped = $true
        }
    }
    if ($removalSkipped) {
        return [pscustomobject]@{
            UserId = $UserId
            AddedGroupIds = [string[]]@($addedGroupIds)
            RemovedGroupIds = [string[]]@($removedGroupIds)
            DirectGroups = [object[]]@($afterAdds)
            Verified = $false
        }
    }

    $finalGroups = @(Get-FreshEntraUserDirectGroups -UserId $UserId)
    $finalIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($group in $finalGroups) {
        $groupId = ([string](Get-EntraMutationPropertyValue -InputObject $group -Name Id)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($groupId)) { [void]$finalIds.Add($groupId) }
    }
    $missingFinalTargetIds = @($requiredIds | Where-Object { -not $finalIds.Contains($_) })
    if ($missingFinalTargetIds.Count -gt 0) {
        throw "Final Entra mandatory groups could not be verified for user '$UserId': $($missingFinalTargetIds -join ', ')."
    }

    $unexpectedManagedGroups = @($finalGroups | Where-Object {
            $displayName = [string](Get-EntraMutationPropertyValue -InputObject $_ -Name DisplayName)
            $groupId = [string](Get-EntraMutationPropertyValue -InputObject $_ -Name Id)
            ($displayName.StartsWith($RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase) -or
                $displayName.StartsWith($ClassGroupPrefix, [StringComparison]::OrdinalIgnoreCase)) -and
                -not $requiredIds.Contains($groupId)
        })
    if ($unexpectedManagedGroups.Count -gt 0) {
        $unexpectedNames = @($unexpectedManagedGroups | ForEach-Object { $_.DisplayName })
        throw "Entra exact managed group state could not be verified for user '$UserId': $($unexpectedNames -join ', ')."
    }

    return [pscustomobject]@{
        UserId = $UserId
        AddedGroupIds = [string[]]@($addedGroupIds)
        RemovedGroupIds = [string[]]@($removedGroupIds)
        DirectGroups = [object[]]@($finalGroups)
        Verified = $true
    }
}

function Enable-EntraStudent {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string] $UserId,
        [switch] $WorkbookVerified,
        [switch] $GroupsVerified,
        [switch] $ManagerVerified,
        [switch] $AttributesVerified
    )

    if (-not ($WorkbookVerified -and $GroupsVerified -and $ManagerVerified -and $AttributesVerified)) {
        throw "Cannot enable Entra student '$UserId' without all verified prerequisites."
    }
    if (-not $PSCmdlet.ShouldProcess($UserId, 'Enable Entra student')) {
        return [pscustomobject]@{ UserId = $UserId; AccountEnabled = $false; Changed = $false; Verified = $false }
    }
    Update-MgUser -UserId $UserId -AccountEnabled:$true -ErrorAction Stop

    $user = Get-MgUser -UserId $UserId -Property @('id', 'accountEnabled') -ErrorAction Stop
    if (-not [bool](Get-EntraMutationPropertyValue -InputObject $user -Name AccountEnabled)) {
        throw "Enabled state could not be verified for Entra user '$UserId'."
    }
    return [pscustomobject]@{
        UserId = $UserId
        AccountEnabled = $true
        Changed = $true
        Verified = $true
        User = $user
    }
}

function Disable-EntraStudent {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([Parameter(Mandatory)][string] $UserId)

    if (-not $PSCmdlet.ShouldProcess($UserId, 'Disable Entra student')) {
        return [pscustomobject]@{ UserId = $UserId; AccountEnabled = $null; Changed = $false; Verified = $false }
    }
    Update-MgUser -UserId $UserId -AccountEnabled:$false -ErrorAction Stop

    $user = Get-MgUser -UserId $UserId -Property @('id', 'accountEnabled') -ErrorAction Stop
    $accountEnabled = Get-EntraMutationPropertyValue -InputObject $user -Name AccountEnabled
    if ($null -eq $accountEnabled -or [bool]$accountEnabled) {
        throw "Disabled state could not be verified for Entra user '$UserId'."
    }
    return [pscustomobject]@{
        UserId = $UserId
        AccountEnabled = $false
        Changed = $true
        Verified = $true
        User = $user
    }
}

function Revoke-EntraStudentSessions {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string] $UserId,
        [switch] $Selected
    )

    if (-not $Selected) { return }
    if ($PSCmdlet.ShouldProcess($UserId, 'Revoke Entra student sign-in sessions')) {
        return Revoke-MgUserSignInSession -UserId $UserId -ErrorAction Stop
    }
}
