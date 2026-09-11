BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $repoRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force
}

Describe 'Stable student identity and directory comparison' {
    InModuleScope SchuelerSync {
        BeforeAll {
            $script:ComparisonTestConfig = @{
                Domain = 'monteaufkirchen.com'
                CompanyName = 'Montessori Schule Aufkirchen'
                EmployeeType = 'Schüler'
                UsageLocation = 'DE'
                AgeGroup = 'Minor'
                ConsentProvidedForMinor = 'Granted'
                LegalAgeGroupClassification = 'MinorWithParentalConsent'
                StudentRoleGroup = @{
                    Id = 'group-role-student'
                    Name = 'SEC-A-ROL-Schule_Schüler'
                }
                LicenseGroupName = 'SEC-A-LIC-O365A1Student'
                RoleGroupPrefix = 'SEC-A-ROL-'
                ClassGroupPrefix = 'SEC-A-CLS-'
            }

            function New-TestStudent {
                param(
                    [string] $GivenName = 'Maria',
                    [string] $Surname = 'Müller',
                    [string] $ClassName = 'JK1-3g2_1',
                    [string] $Teacher = 'Lea Lehrerin',
                    [AllowNull()][string] $EntraObjectId = '',
                    [AllowNull()][string] $StoredUpn = '',
                    [int] $RowNumber = 2
                )
                [pscustomobject]@{
                    RowNumber = $RowNumber
                    NameMitRufname = "$Surname, $GivenName"
                    GivenName = $GivenName
                    Surname = $Surname
                    ClassName = $ClassName
                    Teacher = $Teacher
                    Password = ''
                    EntraObjectId = $EntraObjectId
                    StoredUpn = $StoredUpn
                }
            }

            function New-TestUser {
                param(
                    [Parameter(Mandatory)][string] $Id,
                    [string] $GivenName = 'Maria',
                    [string] $Surname = 'Müller',
                    [string] $UserPrincipalName = 'mmueller@monteaufkirchen.com',
                    [AllowNull()][string] $Mail = '__USE_UPN__',
                    [string] $Department = 'JK1-3g2_1',
                    [string] $OfficeLocation = 'G2',
                    [string] $ManagerId = 'teacher-id'
                )
                [pscustomobject]@{
                    Id = $Id
                    DisplayName = "$GivenName $Surname"
                    GivenName = $GivenName
                    Surname = $Surname
                    UserPrincipalName = $UserPrincipalName
                    Mail = if ($Mail -eq '__USE_UPN__') { $UserPrincipalName } else { $Mail }
                    MailNickname = if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) { '' } else { ($UserPrincipalName -split '@')[0] }
                    ProxyAddresses = @()
                    Department = $Department
                    OfficeLocation = $OfficeLocation
                    CompanyName = 'Montessori Schule Aufkirchen'
                    EmployeeType = 'Schüler'
                    UsageLocation = 'DE'
                    AgeGroup = 'Minor'
                    ConsentProvidedForMinor = 'Granted'
                    LegalAgeGroupClassification = 'MinorWithParentalConsent'
                    AccountEnabled = $true
                    UserType = 'Member'
                    TestManagerId = $ManagerId
                }
            }

            function New-TestGroup {
                param(
                    [Parameter(Mandatory)][string] $Id,
                    [Parameter(Mandatory)][string] $Name,
                    [bool] $IsDynamic = $false,
                    [bool] $IsInherited = $false
                )
                [pscustomobject]@{
                    Id = $Id
                    DisplayName = $Name
                    GroupTypes = if ($IsDynamic) { @('DynamicMembership') } else { @() }
                    MembershipRule = if ($IsDynamic) { '(user.department -eq "G2")' } else { $null }
                    IsDynamic = $IsDynamic
                    IsInherited = $IsInherited
                    IsRemovable = -not $IsDynamic -and -not $IsInherited
                }
            }

            function New-TestSnapshot {
                param(
                    [object[]] $Users = @(),
                    [string[]] $RoleMemberIds = @(),
                    [object[]] $Groups = @(),
                    [hashtable] $DirectGroups = @{},
                    [hashtable] $TransitiveGroups = @{},
                    [hashtable] $AddressOwners = @{}
                )
                if ($Groups.Count -eq 0) {
                    $Groups = @(
                        New-TestGroup -Id group-role-student -Name 'SEC-A-ROL-Schule_Schüler'
                        New-TestGroup -Id group-license -Name 'SEC-A-LIC-O365A1Student'
                        New-TestGroup -Id group-class-g2 -Name 'SEC-A-CLS-JK1-3g2_1'
                    )
                }
                $usersById = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                $usersByUpn = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                $reserved = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $owners = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($key in $AddressOwners.Keys) {
                    $owners[$key] = @($AddressOwners[$key])
                    [void] $reserved.Add([string]$key)
                }
                $managerByUserId = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($user in $Users) {
                    $usersById[$user.Id] = $user
                    if (-not [string]::IsNullOrWhiteSpace([string]$user.UserPrincipalName)) {
                        $usersByUpn[$user.UserPrincipalName] = $user
                        [void] $reserved.Add($user.UserPrincipalName)
                        if (-not $owners.ContainsKey($user.UserPrincipalName)) { $owners[$user.UserPrincipalName] = @() }
                        $owners[$user.UserPrincipalName] = @($owners[$user.UserPrincipalName]) + $user.Id
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$user.Mail)) {
                        [void] $reserved.Add($user.Mail)
                        if (-not $owners.ContainsKey($user.Mail)) { $owners[$user.Mail] = @() }
                        $owners[$user.Mail] = @($owners[$user.Mail]) + $user.Id
                    }
                    $managerByUserId[$user.Id] = $user.TestManagerId
                }
                $groupsById = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                $groupsByName = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                $matchesByName = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($group in $Groups) {
                    $groupsById[$group.Id] = $group
                    if (-not $matchesByName.ContainsKey($group.DisplayName)) { $matchesByName[$group.DisplayName] = @() }
                    $matchesByName[$group.DisplayName] = @($matchesByName[$group.DisplayName]) + $group
                    if (@($matchesByName[$group.DisplayName]).Count -eq 1) {
                        $groupsByName[$group.DisplayName] = $group
                    } else {
                        $groupsByName.Remove($group.DisplayName)
                    }
                }
                $roleMembers = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($id in $RoleMemberIds) { [void] $roleMembers.Add($id) }
                $directByUser = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                $transitiveByUser = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($key in $DirectGroups.Keys) { $directByUser[$key] = @($DirectGroups[$key]) }
                foreach ($key in $TransitiveGroups.Keys) { $transitiveByUser[$key] = @($TransitiveGroups[$key]) }
                foreach ($user in $Users) {
                    if (-not $directByUser.ContainsKey($user.Id)) { $directByUser[$user.Id] = @() }
                    if (-not $transitiveByUser.ContainsKey($user.Id)) { $transitiveByUser[$user.Id] = @($directByUser[$user.Id]) }
                }
                [pscustomobject]@{
                    UsersById = $usersById
                    UsersByUpn = $usersByUpn
                    StudentRoleMemberIds = $roleMembers
                    GroupsById = $groupsById
                    GroupsByDisplayName = $groupsByName
                    GroupMatchesByDisplayName = $matchesByName
                    ReservedAddresses = $reserved
                    AddressOwners = $owners
                    DirectGroupsByUserId = $directByUser
                    TransitiveGroupsByUserId = $transitiveByUser
                    ManagerByUserId = $managerByUserId
                }
            }

            function Get-RequiredTestGroups {
                @(
                    New-TestGroup -Id group-role-student -Name 'SEC-A-ROL-Schule_Schüler'
                    New-TestGroup -Id group-license -Name 'SEC-A-LIC-O365A1Student'
                    New-TestGroup -Id group-class-g2 -Name 'SEC-A-CLS-JK1-3g2_1'
                )
            }
        }

        AfterAll {
            Remove-Variable -Name ComparisonTestConfig -Scope Script -ErrorAction SilentlyContinue
        }

        Context 'identity precedence' {
            It 'matches by object ID before stored UPN and name' {
                $objectWinner = New-TestUser -Id object-id-wins -UserPrincipalName object@monteaufkirchen.com
                $storedWinner = New-TestUser -Id stored-upn-loses -UserPrincipalName stored@monteaufkirchen.com
                $nameWinner = New-TestUser -Id name-loses -UserPrincipalName name@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($objectWinner, $storedWinner, $nameWinner) -RoleMemberIds @('name-loses')
                $student = New-TestStudent -EntraObjectId object-id-wins -StoredUpn stored@monteaufkirchen.com

                $match = Resolve-StudentIdentity -Student $student -Snapshot $snapshot

                $match.Method | Should -Be 'EntraObjectId'
                $match.User.Id | Should -Be 'object-id-wins'
                $match.Warnings | Should -BeNullOrEmpty
            }

            It 'uses stored UPN when no object ID exists' {
                $storedWinner = New-TestUser -Id stored-upn-wins -UserPrincipalName stored@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($storedWinner) -RoleMemberIds @()
                $student = New-TestStudent -StoredUpn ' STORED@MONTEAUFKIRCHEN.COM '

                $match = Resolve-StudentIdentity -Student $student -Snapshot $snapshot

                $match.Method | Should -Be 'StoredUpn'
                $match.User.Id | Should -Be 'stored-upn-wins'
            }

            It 'uses a unique normalized name only inside the student role group' {
                $studentUser = New-TestUser -Id student-id -GivenName 'Mária' -Surname 'Müller'
                $tenantUser = New-TestUser -Id tenant-id -GivenName 'Andere' -Surname 'Person' -UserPrincipalName andere@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($studentUser, $tenantUser) -RoleMemberIds @('student-id')
                $student = New-TestStudent -GivenName 'Maria' -Surname 'Müller'

                $match = Resolve-StudentIdentity -Student $student -Snapshot $snapshot

                $match.Method | Should -Be 'UniqueStudentName'
                $match.User.Id | Should -Be 'student-id'
            }

            It 'does not fall through from a stale object ID' {
                $nameMatch = New-TestUser -Id name-match
                $snapshot = New-TestSnapshot -Users @($nameMatch) -RoleMemberIds @('name-match')
                $student = New-TestStudent -EntraObjectId stale-id

                { Resolve-StudentIdentity -Student $student -Snapshot $snapshot } |
                    Should -Throw '*EntraObjectId*stale-id*'
            }

            It 'does not fall through from a stale stored UPN' {
                $nameMatch = New-TestUser -Id name-match
                $snapshot = New-TestSnapshot -Users @($nameMatch) -RoleMemberIds @('name-match')
                $student = New-TestStudent -StoredUpn stale@monteaufkirchen.com

                { Resolve-StudentIdentity -Student $student -Snapshot $snapshot } |
                    Should -Throw '*StoredUpn*stale@monteaufkirchen.com*'
            }

            It 'rejects duplicate normalized names inside the student role group' {
                $first = New-TestUser -Id first -GivenName 'Mária' -Surname Müller -UserPrincipalName first@monteaufkirchen.com
                $second = New-TestUser -Id second -GivenName Maria -Surname Müller -UserPrincipalName second@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($first, $second) -RoleMemberIds @('first', 'second')

                { Resolve-StudentIdentity -Student (New-TestStudent -GivenName Maria -Surname Müller) -Snapshot $snapshot } |
                    Should -Throw '*normalisierte Name*nicht eindeutig*'
            }
        }

        Context 'preflight collision handling' {
            It 'keeps one unique row and one Entra claim out of duplicate errors' {
                $required = @(Get-RequiredTestGroups)
                $user = New-TestUser -Id student-id
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups $required -DirectGroups @{
                    'student-id' = @($required)
                }

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id -StoredUpn mmueller@monteaufkirchen.com
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors | Should -BeNullOrEmpty
                @($result.Errors | Where-Object { $_.Field -in @('EntraObjectId', 'StoredUpn', 'NormalizedName', 'ResolvedEntraObjectId') }) |
                    Should -BeNullOrEmpty
                $result.ExistingStudents.Count | Should -Be 1
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'reports a stale explicit object ID and suppresses new and departure actions' {
                $nameMatch = New-TestUser -Id name-match
                $snapshot = New-TestSnapshot -Users @($nameMatch) -RoleMemberIds @('name-match')

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId stale-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'Identity'
                ($result.Errors.Message -join ' ') | Should -Match 'EntraObjectId.*stale-id'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'reports a duplicate normalized student-role name without destructive categories' {
                $first = New-TestUser -Id first -GivenName 'Mária' -Surname Müller -UserPrincipalName first@monteaufkirchen.com
                $second = New-TestUser -Id second -GivenName Maria -Surname Müller -UserPrincipalName second@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($first, $second) -RoleMemberIds @('first', 'second')

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -GivenName Maria -Surname Müller
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'Identity'
                ($result.Errors.Message -join ' ') | Should -Match 'normalisierte Name.*nicht eindeutig'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'reports duplicate Excel object IDs, stored UPNs, and normalized names' {
                $user = New-TestUser -Id student-id -UserPrincipalName stored@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user) -RoleMemberIds @('student-id')
                $students = @(
                    New-TestStudent -EntraObjectId student-id -StoredUpn stored@monteaufkirchen.com -RowNumber 2
                    New-TestStudent -EntraObjectId STUDENT-ID -StoredUpn STORED@MONTEAUFKIRCHEN.COM -RowNumber 3
                )

                $result = Compare-StudentDirectory -Students $students -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'EntraObjectId'
                $result.Errors.Field | Should -Contain 'StoredUpn'
                $result.Errors.Field | Should -Contain 'NormalizedName'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'suppresses departures when a nonempty row cannot produce a normalized name' {
                $user = New-TestUser -Id student-id
                $snapshot = New-TestSnapshot -Users @($user) -RoleMemberIds @('student-id')

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -GivenName '---' -Surname '...' -RowNumber 2
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'NormalizedName'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'reports one address owned by multiple directory objects without destructive categories' {
                $first = New-TestUser -Id first -UserPrincipalName first@monteaufkirchen.com
                $second = New-TestUser -Id second -UserPrincipalName second@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($first, $second) -RoleMemberIds @('first', 'second') -AddressOwners @{
                    'shared@monteaufkirchen.com' = @('first', 'second')
                }

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId first
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'AddressOwnership'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'merges Exchange address owners before detecting multiple owners' {
                $user = New-TestUser -Id student-id -UserPrincipalName shared@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user) -RoleMemberIds @('student-id')

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig -ExchangeAddressOwners @{
                    'SHARED@MONTEAUFKIRCHEN.COM' = [pscustomobject]@{ ExternalDirectoryObjectId = 'other-id' }
                }

                $result.Errors.Field | Should -Contain 'AddressOwnership'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'reports a candidate address owned by a non-student tenant user' {
                $staff = New-TestUser -Id staff-id -GivenName Alice -Surname Admin -UserPrincipalName mmueller@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($staff) -RoleMemberIds @()

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'UserPrincipalName'
                $result.Errors.Message | Should -Match 'außerhalb der Schüler-Rollengruppe'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }

            It 'reports two Excel rows resolving to one Entra ID' {
                $user = New-TestUser -Id shared-id -UserPrincipalName shared@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user) -RoleMemberIds @('shared-id')
                $students = @(
                    New-TestStudent -GivenName Maria -Surname Müller -EntraObjectId shared-id -RowNumber 2
                    New-TestStudent -GivenName Marie -Surname Mueller -StoredUpn shared@monteaufkirchen.com -RowNumber 3
                )

                $result = Compare-StudentDirectory -Students $students -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'ResolvedEntraObjectId'
                $result.NewStudents | Should -BeNullOrEmpty
                $result.Departures | Should -BeNullOrEmpty
            }
        }

        Context 'desired state and differences' {
            It 'preserves Unicode names and creates every required desired value' {
                $student = New-TestStudent -GivenName 'Änne' -Surname 'Weiß'
                $manager = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com

                $desired = New-StudentDesiredState -Student $student -SelectedUpn aweiss@monteaufkirchen.com -Config $script:ComparisonTestConfig -Manager $manager

                $desired.DisplayName | Should -Be 'Änne Weiß'
                $desired.GivenName | Should -Be 'Änne'
                $desired.Surname | Should -Be 'Weiß'
                $desired.UserPrincipalName | Should -Be 'aweiss@monteaufkirchen.com'
                $desired.Mail | Should -Be 'aweiss@monteaufkirchen.com'
                $desired.MailNickname | Should -Be 'aweiss'
                $desired.Department | Should -Be 'JK1-3g2_1'
                $desired.OfficeLocation | Should -Be 'G2'
                $desired.CompanyName | Should -Be 'Montessori Schule Aufkirchen'
                $desired.EmployeeType | Should -Be 'Schüler'
                $desired.UsageLocation | Should -Be 'DE'
                $desired.AgeGroup | Should -Be 'Minor'
                $desired.ConsentProvidedForMinor | Should -Be 'Granted'
                $desired.LegalAgeGroupClassification | Should -Be 'MinorWithParentalConsent'
                $desired.ManagerId | Should -Be teacher-id
                $desired.RequiredGroupNames | Should -Be @(
                    'SEC-A-ROL-Schule_Schüler'
                    'SEC-A-LIC-O365A1Student'
                    'SEC-A-CLS-JK1-3g2_1'
                )
            }

            It 'compares text ordinally after trimming, addresses case-insensitively, and IDs exactly' {
                New-StateDifference -Area Entra -Field DisplayName -Current ' Maria Müller ' -Desired 'Maria Müller' |
                    Should -BeNullOrEmpty
                New-StateDifference -Area Entra -Field UserPrincipalName -Current 'MMUELLER@MONTEAUFKIRCHEN.COM' -Desired 'mmueller@monteaufkirchen.com' |
                    Should -BeNullOrEmpty
                (New-StateDifference -Area Manager -Field ManagerId -Current 'ABC' -Desired 'abc').Action |
                    Should -Be 'Set'
            }

            It 'keeps a valid current UPN for unchanged names and recalculates it after a name change' {
                $groups = @(Get-RequiredTestGroups)
                $current = New-TestUser -Id student-id -GivenName Maria -Surname Müller -UserPrincipalName mamueller@monteaufkirchen.com
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $direct = @{
                    'student-id' = @($groups)
                }
                $snapshot = New-TestSnapshot -Users @($current, $teacher) -RoleMemberIds @('student-id') -Groups $groups -DirectGroups $direct

                $unchanged = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig
                $changedName = Compare-StudentDirectory -Students @(
                    New-TestStudent -GivenName Marianne -Surname Müller -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig
                $changedSurname = Compare-StudentDirectory -Students @(
                    New-TestStudent -GivenName Maria -Surname Schmidt -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $unchanged.ExistingStudents[0].DesiredState.UserPrincipalName | Should -Be 'mamueller@monteaufkirchen.com'
                $changedName.ChangedStudents[0].DesiredState.UserPrincipalName | Should -Be 'mmueller@monteaufkirchen.com'
                $changedName.ChangedStudents[0].DesiredState.MailNickname | Should -Be 'mmueller'
                $changedSurname.ChangedStudents[0].DesiredState.UserPrincipalName | Should -Be 'mschmidt@monteaufkirchen.com'
                $changedSurname.ChangedStudents[0].DesiredState.MailNickname | Should -Be 'mschmidt'
            }

            It 'preserves every numbered UPN candidate boundary from 2 through 100 without old collisions' {
                $required = @(Get-RequiredTestGroups)
                foreach ($suffix in 2, 9, 10, 19, 100) {
                    $upn = "mariamueller$suffix@monteaufkirchen.com"
                    $user = New-TestUser -Id "student-$suffix" -UserPrincipalName $upn -ManagerId "teacher-$suffix"
                    $teacher = New-TestUser -Id "teacher-$suffix" -GivenName Lea -Surname Lehrerin -UserPrincipalName "lea.$suffix@monteaufkirchen.com"
                    $directGroups = @{}
                    $directGroups[$user.Id] = @($required)
                    $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @($user.Id) -Groups $required -DirectGroups $directGroups

                    $result = Compare-StudentDirectory -Students @(
                        New-TestStudent -EntraObjectId $user.Id
                    ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                    $result.Errors | Should -BeNullOrEmpty
                    $result.ExistingStudents.Count | Should -Be 1
                    $result.ExistingStudents[0].DesiredState.UserPrincipalName | Should -Be $upn
                    @($result.ExistingStudents[0].Differences | Where-Object Field -eq UserPrincipalName) |
                        Should -BeNullOrEmpty
                }
            }

            It 'reserves selected addresses in workbook order and reports the fallback' {
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($teacher) -RoleMemberIds @()

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -GivenName Maria -Surname Müller -RowNumber 2
                    New-TestStudent -GivenName Markus -Surname Müller -RowNumber 3
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.NewStudents[0].DesiredState.UserPrincipalName | Should -Be 'mmueller@monteaufkirchen.com'
                $result.NewStudents[1].DesiredState.UserPrincipalName | Should -Be 'mamueller@monteaufkirchen.com'
                $result.Warnings.Field | Should -Contain 'UserPrincipalName'
                $result.Warnings.Desired | Should -Contain 'mamueller@monteaufkirchen.com'
            }

            It 'marks attribute, manager, and safe group deviations as field-level changes' {
                $required = @(Get-RequiredTestGroups)
                $oldRole = New-TestGroup -Id old-role -Name 'SEC-A-ROL-AlteRolle'
                $oldClass = New-TestGroup -Id old-class -Name 'SEC-A-CLS-AlteKlasse'
                $user = New-TestUser -Id student-id -Department Alt -ManagerId old-teacher
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups @($required + $oldRole + $oldClass) -DirectGroups @{
                    'student-id' = @($required[0], $oldRole, $oldClass)
                }

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig
                $entry = $result.ChangedStudents[0]

                $entry.Differences.Field | Should -Contain 'Department'
                $entry.Differences.Field | Should -Contain 'ManagerId'
                @($entry.Differences | Where-Object { $_.Area -eq 'Group' -and $_.Action -eq 'Add' }).Desired |
                    Should -Contain 'SEC-A-LIC-O365A1Student'
                @($entry.Differences | Where-Object { $_.Area -eq 'Group' -and $_.Action -eq 'Add' }).Desired |
                    Should -Contain 'SEC-A-CLS-JK1-3g2_1'
                @($entry.Differences | Where-Object { $_.Area -eq 'Group' -and $_.Action -eq 'Remove' }).Current |
                    Should -Be @('SEC-A-ROL-AlteRolle', 'SEC-A-CLS-AlteKlasse')
            }

            It 'warns for dynamic and inherited managed deviations without removal differences' {
                $required = @(Get-RequiredTestGroups)
                $dynamic = New-TestGroup -Id dynamic-role -Name 'SEC-A-ROL-Dynamisch' -IsDynamic $true
                $inherited = New-TestGroup -Id inherited-class -Name 'SEC-A-CLS-Geerbt' -IsInherited $true
                $user = New-TestUser -Id student-id
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups @($required + $dynamic + $inherited) -DirectGroups @{
                    'student-id' = @($required + $dynamic)
                } -TransitiveGroups @{
                    'student-id' = @($required + $dynamic + $inherited)
                }

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                @($result.Warnings | Where-Object Field -eq 'ManagedGroup').Count | Should -Be 2
                @($result.ExistingStudents[0].Differences | Where-Object Action -eq Remove) | Should -BeNullOrEmpty
            }

            It 'uses a comparison-only legal-age difference and warns when mail is missing' {
                $required = @(Get-RequiredTestGroups)
                $user = New-TestUser -Id student-id -Mail $null
                $user.LegalAgeGroupClassification = 'Undefined'
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups $required -DirectGroups @{
                    'student-id' = @($required)
                }

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig
                $legal = $result.ChangedStudents[0].Differences | Where-Object Field -eq LegalAgeGroupClassification

                $legal.Action | Should -Be 'Compare'
                $result.Warnings.Field | Should -Contain 'LegalAgeGroupClassification'
                $result.Warnings.Field | Should -Contain 'Mail'
            }

            It 'treats a missing mail nickname as a field-level Set difference' {
                $required = @(Get-RequiredTestGroups)
                $user = New-TestUser -Id student-id
                $user.PSObject.Properties.Remove('MailNickname')
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups $required -DirectGroups @{
                    'student-id' = @($required)
                }

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig
                $difference = $result.ChangedStudents[0].Differences | Where-Object Field -eq MailNickname

                $difference.Area | Should -Be 'Entra'
                $difference.Current | Should -BeNullOrEmpty
                $difference.Desired | Should -Be 'mmueller'
                $difference.Action | Should -Be 'Set'
            }

            It 'keeps new, departure, changed, and existing categories mutually exclusive' {
                $groups = @(Get-RequiredTestGroups)
                $existing = New-TestUser -Id existing-id -GivenName Emil -Surname Eins -UserPrincipalName eeins@monteaufkirchen.com
                $changed = New-TestUser -Id changed-id -GivenName Carla -Surname Zwei -UserPrincipalName czwei@monteaufkirchen.com -Department Alt
                $departure = New-TestUser -Id departure-id -GivenName Dora -Surname Drei -UserPrincipalName ddrei@monteaufkirchen.com
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($existing, $changed, $departure, $teacher) -RoleMemberIds @('existing-id', 'changed-id', 'departure-id') -Groups $groups -DirectGroups @{
                    'existing-id' = @($groups)
                    'changed-id' = @($groups)
                    'departure-id' = @($groups)
                }
                $students = @(
                    New-TestStudent -GivenName Emil -Surname Eins -EntraObjectId existing-id -RowNumber 2
                    New-TestStudent -GivenName Carla -Surname Zwei -EntraObjectId changed-id -RowNumber 3
                    New-TestStudent -GivenName Nora -Surname Neu -RowNumber 4
                )

                $result = Compare-StudentDirectory -Students $students -Snapshot $snapshot -Config $script:ComparisonTestConfig
                $categoryIds = @(
                    @($result.NewStudents | ForEach-Object { "row:$($_.Student.RowNumber)" })
                    @($result.Departures | ForEach-Object { "id:$($_.User.Id)" })
                    @($result.ChangedStudents | ForEach-Object { "id:$($_.User.Id)" })
                    @($result.ExistingStudents | ForEach-Object { "id:$($_.User.Id)" })
                )

                $result.NewStudents.Count | Should -Be 1
                $result.Departures.User.Id | Should -Be departure-id
                $result.ChangedStudents.User.Id | Should -Be changed-id
                $result.ExistingStudents.User.Id | Should -Be existing-id
                ($categoryIds | Select-Object -Unique).Count | Should -Be $categoryIds.Count
            }
        }

        Context 'mandatory group validation and Exchange merge' {
            It 'reports an invalid office location as a field-level error' {
                $user = New-TestUser -Id student-id
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id')

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -ClassName 'JK1-3x9_1' -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'OfficeLocation'
                $result.Errors.Desired | Should -Contain 'JK1-3x9_1'
            }

            It 'reports an ambiguous teacher as a field-level manager error' {
                $user = New-TestUser -Id student-id
                $teacherA = New-TestUser -Id teacher-a -GivenName Lea -Surname Lehrerin -UserPrincipalName lea.a@monteaufkirchen.com
                $teacherB = New-TestUser -Id teacher-b -GivenName Lea -Surname Lehrerin -UserPrincipalName lea.b@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacherA, $teacherB) -RoleMemberIds @('student-id')

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'ManagerId'
                ($result.Errors.Message -join ' ') | Should -Match 'Kein eindeutiger Manager'
            }

            It 'reports a missing mandatory target group' {
                $role = New-TestGroup -Id group-role-student -Name 'SEC-A-ROL-Schule_Schüler'
                $license = New-TestGroup -Id group-license -Name 'SEC-A-LIC-O365A1Student'
                $user = New-TestUser -Id student-id
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups @($role, $license)

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $issue = @($result.Errors | Where-Object Field -eq 'SEC-A-CLS-JK1-3g2_1')
                $issue.Count | Should -Be 1
                $issue[0].Severity | Should -Be 'Error'
                $issue[0].Area | Should -Be 'Group'
                $issue[0].Current | Should -Be 0
                $issue[0].Desired | Should -Be 1
                $issue[0].Message | Should -Match 'fehlt'
            }

            It 'reports missing, ambiguous, and dynamic mandatory target groups' {
                $role = New-TestGroup -Id group-role-student -Name 'SEC-A-ROL-Schule_Schüler'
                $licenseA = New-TestGroup -Id license-a -Name 'SEC-A-LIC-O365A1Student'
                $licenseB = New-TestGroup -Id license-b -Name 'SEC-A-LIC-O365A1Student'
                $dynamicClass = New-TestGroup -Id class-dynamic -Name 'SEC-A-CLS-JK1-3g2_1' -IsDynamic $true
                $user = New-TestUser -Id student-id
                $teacher = New-TestUser -Id teacher-id -GivenName Lea -Surname Lehrerin -UserPrincipalName lea@monteaufkirchen.com
                $snapshot = New-TestSnapshot -Users @($user, $teacher) -RoleMemberIds @('student-id') -Groups @($role, $licenseA, $licenseB, $dynamicClass)

                $result = Compare-StudentDirectory -Students @(
                    New-TestStudent -EntraObjectId student-id
                ) -Snapshot $snapshot -Config $script:ComparisonTestConfig

                $result.Errors.Field | Should -Contain 'SEC-A-LIC-O365A1Student'
                $result.Errors.Field | Should -Contain 'SEC-A-CLS-JK1-3g2_1'
                ($result.Errors.Message -join ' ') | Should -Match 'mehrdeutig'
                ($result.Errors.Message -join ' ') | Should -Match 'dynamisch'
            }

            It 'moves an existing student to changed for Exchange field differences' {
                $entry = [pscustomobject]@{
                    Student = New-TestStudent -EntraObjectId student-id
                    User = New-TestUser -Id student-id
                    DesiredState = [pscustomobject]@{ UserPrincipalName = 'mmueller@monteaufkirchen.com' }
                    IdentityMethod = 'EntraObjectId'
                    Differences = @()
                }
                $comparison = [pscustomobject]@{
                    NewStudents = @()
                    Departures = @()
                    ChangedStudents = @()
                    ExistingStudents = @($entry)
                    Warnings = @()
                    Errors = @()
                }

                $result = Add-ExchangeComparison -Comparison $comparison -UserId student-id -MailboxState ([pscustomobject]@{
                    Exists = $true
                    Differences = @(
                        [pscustomobject]@{ Field = 'AddressBookPolicy'; Current = 'Alt'; Desired = 'Neu'; Action = 'Set' }
                    )
                })

                $result.ExistingStudents | Should -BeNullOrEmpty
                $result.ChangedStudents.Count | Should -Be 1
                $result.ChangedStudents[0].Differences[0].Area | Should -Be Exchange
                $result.ChangedStudents[0].Differences[0].Field | Should -Be AddressBookPolicy
            }

            It 'warns for a missing mailbox without fabricating Exchange differences' {
                $entry = [pscustomobject]@{
                    Student = New-TestStudent -EntraObjectId student-id
                    User = New-TestUser -Id student-id
                    DesiredState = [pscustomobject]@{ UserPrincipalName = 'mmueller@monteaufkirchen.com' }
                    IdentityMethod = 'EntraObjectId'
                    Differences = @()
                }
                $comparison = [pscustomobject]@{
                    NewStudents = @()
                    Departures = @()
                    ChangedStudents = @()
                    ExistingStudents = @($entry)
                    Warnings = @()
                    Errors = @()
                }

                $result = Add-ExchangeComparison -Comparison $comparison -UserId student-id -MailboxState ([pscustomobject]@{
                    Exists = $false
                    Status = 'MailboxNotReady'
                    Differences = @()
                })

                $result.ExistingStudents.Count | Should -Be 1
                $result.ChangedStudents | Should -BeNullOrEmpty
                $result.Warnings.Message | Should -Match 'EXO-Konfiguration ausstehend'
                $result.ExistingStudents[0].Differences | Should -BeNullOrEmpty
            }
        }
    }
}
