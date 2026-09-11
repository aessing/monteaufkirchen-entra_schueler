function ConvertTo-UpnToken {
    param([Parameter(Mandatory)][string] $Value)
    $mapped = $Value.Trim().ToLowerInvariant().
        Replace('ä', 'ae').Replace('ö', 'oe').Replace('ü', 'ue').Replace('ß', 'ss')
    $decomposed = $mapped.Normalize([Text.NormalizationForm]::FormD)
    $builder = [Text.StringBuilder]::new()
    foreach ($character in $decomposed.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($character) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void] $builder.Append($character)
        }
    }
    return ($builder.ToString().Normalize([Text.NormalizationForm]::FormC) -replace '[^a-z0-9]', '')
}

function Get-OfficeLocation {
    param([Parameter(Mandatory)][string] $ClassName)
    $matches = [regex]::Matches($ClassName, '(?i)(?<office>g[1-4]|m[1-3]|o[1-2]|a[1-2])(?=_|$)')
    if ($matches.Count -ne 1) {
        throw "Klasse '$ClassName' enthält keine eindeutige erlaubte Office Location."
    }
    return $matches[0].Groups['office'].Value.ToUpperInvariant()
}

function Get-UpnCandidates {
    param(
        [Parameter(Mandatory)][string] $GivenName,
        [Parameter(Mandatory)][string] $Surname,
        [Parameter(Mandatory)][string] $Domain
    )
    $givenToken = ConvertTo-UpnToken $GivenName
    $surnameToken = ConvertTo-UpnToken $Surname
    if ([string]::IsNullOrEmpty($givenToken) -or [string]::IsNullOrEmpty($surnameToken)) {
        throw 'Vor- und Nachname müssen nach der Normalisierung nicht leer sein.'
    }
    for ($length = 1; $length -le $givenToken.Length; $length++) {
        '{0}{1}@{2}' -f $givenToken.Substring(0, $length), $surnameToken, $Domain.Trim()
    }
}

function Get-AddressOwnerIds {
    param([AllowNull()][object] $Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [string] -or $Value -is [Guid]) { return @([string]$Value) }
    if ($Value -is [System.Collections.IDictionary]) {
        return @($Value.Values | ForEach-Object { Get-AddressOwnerIds $_ })
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [hashtable]) {
        return @($Value | ForEach-Object { Get-AddressOwnerIds $_ })
    }
    $ids = @()
    foreach ($name in 'CurrentObjectId', 'GraphObjectId', 'ExchangeObjectId', 'ObjectId', 'OwnerId', 'OwnerIds', 'Ids') {
        $property = $Value.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            $ids += @(Get-AddressOwnerIds $property.Value)
        }
    }
    return @($ids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-UsedAddress {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]] $UsedAddresses,
        [Parameter(Mandatory)][string] $Candidate
    )
    foreach ($address in $UsedAddresses) {
        if ([string]::Equals([string]$address, $Candidate, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Select-AvailableUpn {
    param(
        [Parameter(Mandatory)][string] $GivenName,
        [Parameter(Mandatory)][string] $Surname,
        [Parameter(Mandatory)][string] $Domain,
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]] $UsedAddresses,
        [Parameter(Mandatory)][System.Collections.IDictionary] $AddressOwners,
        [AllowNull()][string] $CurrentObjectId
    )
    $candidates = @(Get-UpnCandidates $GivenName $Surname $Domain)
    $collisions = [Collections.Generic.List[string]]::new()
    $isAvailable = {
        param($candidate)
        if (Test-UsedAddress -UsedAddresses $UsedAddresses -Candidate $candidate) { return $false }
        $ownerEntry = $null
        foreach ($key in $AddressOwners.Keys) {
            if ([string]::Equals([string]$key, $candidate, [StringComparison]::OrdinalIgnoreCase)) {
                $ownerEntry = $AddressOwners[$key]
                break
            }
        }
        if ($null -eq $ownerEntry) { return $true }
        $owners = @(Get-AddressOwnerIds $ownerEntry)
        if ([string]::IsNullOrWhiteSpace($CurrentObjectId) -or $owners.Count -eq 0) { return $false }
        foreach ($owner in $owners) {
            if (-not [string]::Equals($owner, $CurrentObjectId, [StringComparison]::OrdinalIgnoreCase)) { return $false }
        }
        return $true
    }
    foreach ($candidate in $candidates) {
        if (& $isAvailable $candidate) {
            [void] $UsedAddresses.Add($candidate)
            return [pscustomobject]@{ Upn = $candidate; WasFallback = $false; CollisionCount = $collisions.Count; Collisions = [string[]]$collisions }
        }
        $collisions.Add($candidate)
    }
    $fullGiven = ConvertTo-UpnToken $GivenName
    $surnameToken = ConvertTo-UpnToken $Surname
    for ($suffix = 2; $true; $suffix++) {
        $candidate = '{0}{1}{2}@{3}' -f $fullGiven, $surnameToken, $suffix, $Domain.Trim()
        if (& $isAvailable $candidate) {
            [void] $UsedAddresses.Add($candidate)
            return [pscustomobject]@{ Upn = $candidate; WasFallback = $true; CollisionCount = $collisions.Count; Collisions = [string[]]$collisions }
        }
        $collisions.Add($candidate)
    }
}

$script:StudentPasswordWords = @(
    'Lego', 'Wald', 'Mond', 'Tiger', 'Wiese', 'Wolke', 'Zebra', 'Biene', 'Pizza', 'Sonne', 'Blume', 'Apfel', 'Panda', 'Koala',
    'Garten', 'Rakete', 'Schule', 'Pferde', 'Bienen', 'Kuchen'
)

function New-StudentPassword {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.HashSet[string]] $UsedPasswords,
        [scriptblock] $RandomIndexScriptBlock
    )
    $pairs = @()
    for ($left = 0; $left -lt $script:StudentPasswordWords.Count; $left++) {
        for ($right = 0; $right -lt $script:StudentPasswordWords.Count; $right++) {
            if (($script:StudentPasswordWords[$left].Length + $script:StudentPasswordWords[$right].Length) -eq 10) {
                $pairs += ,@($script:StudentPasswordWords[$left], $script:StudentPasswordWords[$right])
            }
        }
    }
    for ($attempt = 0; $attempt -lt 100; $attempt++) {
        if ($null -eq $RandomIndexScriptBlock) {
            $pairIndex = [Security.Cryptography.RandomNumberGenerator]::GetInt32(0, $pairs.Count)
        } else { $pairIndex = & $RandomIndexScriptBlock $pairs.Count }
        $pair = $pairs[$pairIndex % $pairs.Count]
        $number = [Security.Cryptography.RandomNumberGenerator]::GetInt32(10, 100)
        $password = '{0}{1}{2:00}' -f $pair[0], $pair[1], $number
        if ($UsedPasswords.Add($password)) { return $password }
    }
    throw 'Nach 100 Versuchen konnte kein eindeutiges Schülerpasswort erzeugt werden.'
}
