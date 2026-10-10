[CmdletBinding()]
param(
    [ValidateSet('List','Object','Export','Compare','Update','Add','Remove','Exchange','Tui','Invalid')][string] $Mode='List',
    [string] $File,
    [ValidateSet('Personnel','Student','Teacher')][string] $Population = 'Personnel',
    [switch] $Preview,
    [switch] $DisabledFixture,
    [switch] $GuestFixture,
    [switch] $ReviewFixture
)
# This standalone process substitutes every remote command. Connectors and writes
# throw so an incomplete fixture cannot fall through to a real tenant.
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $repo 'src/SchuelerSync/SchuelerSync.psd1') -ErrorAction Stop
$global:PersonalCliConfig=Import-PowerShellDataFile (Join-Path $repo 'config/PersonalSync.psd1')
$teachers=Import-PowerShellDataFile (Join-Path $repo 'config/LehrerSync.psd1')
$students=Import-PowerShellDataFile (Join-Path $repo 'config/SchuelerSync.psd1')
$global:PersonalCliGroups=@{}
foreach($role in @($global:PersonalCliConfig.Roles.Values)+@($teachers.Roles.Values)+@($students.StudentRoleGroup)+@(
    @{Id='bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';Name='SEC-A-ADM-GlobalAdministrators'},
    @{Id='cccccccc-cccc-cccc-cccc-cccccccccccc';Name='SEC-A-LIC-O365A1Faculty'})) {
    $global:PersonalCliGroups[[string]$role.Id]=[pscustomobject]@{Id=$role.Id;DisplayName=$role.Name;OdataType='#microsoft.graph.group';GroupTypes=@();MembershipRule=$null;AssignedLicenses=@()}
}
$global:PersonalCliGroups['cccccccc-cccc-cccc-cccc-cccccccccccc'].AssignedLicenses=@([pscustomobject]@{SkuId='synthetic-sku'})
$global:PersonalCliUser=[pscustomobject]@{
    Id='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';DisplayName='Lea Testwald';GivenName='Lea';Surname='Testwald'
    UserPrincipalName='ltestwald@example.invalid';Mail='ltestwald@example.invalid';MailNickname='ltestwald';ProxyAddresses=@()
    AccountEnabled=$true;UserType='Member';OnPremisesSyncEnabled=$false;JobTitle='Schule Sekretariat'
    CompanyName='Montessori Schule Aufkirchen';EmployeeType='Sekretariat';UsageLocation='DE';AgeGroup='Adult'
    ConsentProvidedForMinor='NotRequired';LegalAgeGroupClassification='Adult';AssignedLicenses=@([pscustomobject]@{SkuId='synthetic-sku'})
    LicenseAssignmentStates=@();SignInActivity=[pscustomobject]@{LastSuccessfulSignInDateTime='2026-01-02T12:30:00+02:00'}
}
if ($DisabledFixture) { $global:PersonalCliUser.AccountEnabled=$false }
if ($GuestFixture) { $global:PersonalCliUser.UserType='Guest' }
$global:PersonalCliRoleIds = @($global:PersonalCliConfig.Roles.Schule_Sekretariat.Id)
$global:PersonalCliUsers = @($global:PersonalCliUser)
$global:PersonalCliReviewMemberships = $null
if ($ReviewFixture) {
    if ($Population -ne 'Personnel') { throw 'ReviewFixture ist nur für Personal vorgesehen.' }
    $global:PersonalCliReviewMemberships = @{$global:PersonalCliUser.Id=$global:PersonalCliRoleIds}
    foreach ($definition in @(
        @{Id='11111111-1111-1111-1111-111111111111';GivenName='Eva';Surname='Adminwald';Job='ExterneAdmins';Type='Member'},
        @{Id='22222222-2222-2222-2222-222222222222';GivenName='';Surname='';Job='';Type='Member'},
        @{Id='33333333-3333-3333-3333-333333333333';GivenName='Nora';Surname='Gastwald';Job='ExterneBenutzer';Type='Guest'}
    )) {
        $copy=$global:PersonalCliUser | Select-Object *
        $copy.Id=$definition.Id; $copy.GivenName=$definition.GivenName; $copy.Surname=$definition.Surname
        $copy.DisplayName="Synthetic $($definition.Id.Substring(0,4))"
        $copy.UserPrincipalName="$($definition.Id.Substring(0,4))@example.invalid"; $copy.Mail=$copy.UserPrincipalName
        $copy.JobTitle=$definition.Job; $copy.UserType=$definition.Type
        $global:PersonalCliUsers += $copy
        $global:PersonalCliReviewMemberships[$copy.Id]=if ($definition.Job) {@($global:PersonalCliConfig.Roles[$definition.Job].Id)} else {@()}
    }
}
if ($Population -eq 'Teacher') {
    $global:PersonalCliRoleIds = @($teachers.Roles.School.Id)
    $global:PersonalCliUser.JobTitle = 'PA, JAS'
}
if ($Population -eq 'Student') {
    $global:PersonalCliRoleIds = @($students.StudentRoleGroup.Id)
    $global:PersonalCliGroups['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'] = [pscustomobject]@{
        Id='eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';DisplayName='SEC-A-CLS-JK1-3g2_1';GroupTypes=@();MembershipRule=$null;AssignedLicenses=@()}
    $global:PersonalCliGroups['dddddddd-dddd-dddd-dddd-dddddddddddd'] = [pscustomobject]@{
        Id='dddddddd-dddd-dddd-dddd-dddddddddddd';DisplayName='SEC-A-CLS-JK4-6m1_1';GroupTypes=@();MembershipRule=$null;AssignedLicenses=@()}
    $global:PersonalCliGroups['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'] | Add-Member -NotePropertyName OdataType -NotePropertyValue '#microsoft.graph.group'
    $global:PersonalCliUser | Add-Member -NotePropertyMembers @{Department='JK1-3g2_1';OfficeLocation='g2'}
    foreach ($teacher in @(@{Id='ffffffff-ffff-ffff-ffff-ffffffffffff';Name='Nora Lehrwald';UPN='teacher@example.invalid'},@{Id='11111111-1111-1111-1111-111111111111';Name='Zoe Lehrwald';UPN='second-teacher@example.invalid'})) {
        $global:PersonalCliUsers += [pscustomobject]@{Id=$teacher.Id;DisplayName=$teacher.Name;UserPrincipalName=$teacher.UPN
            GivenName='Nora';Surname='Lehrwald';Mail=$teacher.UPN;AccountEnabled=$true;UserType='Member';OnPremisesSyncEnabled=$false
            AssignedLicenses=@();LicenseAssignmentStates=@();SignInActivity=$null}
    }
}
function global:Get-MgContext {
    [pscustomobject]@{TenantId='synthetic-tenant';Account='operator@example.invalid';Scopes=@(
        'User.ReadWrite.All','User-Mail.ReadWrite.All','Group.Read.All','GroupMember.ReadWrite.All','User.RevokeSessions.All',
        'User.Read.All','Directory.Read.All','AuditLog.Read.All','User.EnableDisableAccount.All')}
}
function global:Get-ConnectionInformation { [pscustomobject]@{Name='ExchangeOnline';State='Connected';TenantID='synthetic-tenant'} }
function global:Get-MgGroup { [CmdletBinding()]param($GroupId,[switch]$All,$Property)
    if($GroupId) {$global:PersonalCliGroups[[string]$GroupId]} else {$global:PersonalCliGroups.Values}
}
function global:Get-MgUser { [CmdletBinding()]param($UserId,[switch]$All,$Property)
    if (-not $UserId) { $global:PersonalCliUsers } else { $global:PersonalCliUsers | Where-Object { $UserId -in @($_.Id,$_.UserPrincipalName) } }
}
function global:Get-MgGroupMember { [CmdletBinding()]param($GroupId,[switch]$All)
    if ($null -ne $global:PersonalCliReviewMemberships) { $global:PersonalCliUsers | Where-Object {$GroupId -in $global:PersonalCliReviewMemberships[$_.Id]}; return }
    if($GroupId -in $global:PersonalCliRoleIds) {$global:PersonalCliUser}
    elseif ($global:PersonalCliUsers.Count -gt 1 -and $GroupId -eq '103a1c4c-036b-40bc-bbe8-e3f7a866cb01') { $global:PersonalCliUsers | Where-Object Id -ne $global:PersonalCliUser.Id }
}
function global:Get-MgUserMemberOfAsGroup { [CmdletBinding()]param($UserId,[switch]$All,$Property)
    if ($null -ne $global:PersonalCliReviewMemberships) { foreach ($id in $global:PersonalCliReviewMemberships[$UserId]) {$global:PersonalCliGroups[$id]}; return }
    if ($UserId -ne $global:PersonalCliUser.Id) { $global:PersonalCliGroups['103a1c4c-036b-40bc-bbe8-e3f7a866cb01']; return }
    foreach ($id in $global:PersonalCliRoleIds) { $global:PersonalCliGroups[$id] }
    if ($global:PersonalCliGroups.ContainsKey('eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee')) { $global:PersonalCliGroups['eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'] }
}
function global:Get-MgUserManager { [CmdletBinding()]param($UserId) [pscustomobject]@{Id='ffffffff-ffff-ffff-ffff-ffffffffffff'} }
function global:Get-MgUserTransitiveMemberOf { [CmdletBinding()]param($UserId,[switch]$All)
    if ($null -ne $global:PersonalCliReviewMemberships) { foreach ($id in $global:PersonalCliReviewMemberships[$UserId]) {$global:PersonalCliGroups[$id]}; return }
    if ($UserId -ne $global:PersonalCliUser.Id) { $global:PersonalCliGroups['103a1c4c-036b-40bc-bbe8-e3f7a866cb01']; return }
    foreach ($id in $global:PersonalCliRoleIds) { $global:PersonalCliGroups[$id] }
}
function global:Get-MgUserOwnedObject { [CmdletBinding()]param($UserId,[switch]$All) }
function global:Get-EXORecipient { [CmdletBinding()]param($ResultSize,$Properties,$ExternalDirectoryObjectId)
    [pscustomobject]@{ExternalDirectoryObjectId=$global:PersonalCliUser.Id;RecipientTypeDetails=$(if ($global:PersonalCliUser.UserType -eq 'Guest') {'GuestMailUser'} else {'UserMailbox'})}
}
function global:Get-Recipient { [CmdletBinding()]param($ResultSize)
    [pscustomobject]@{ExternalDirectoryObjectId=$global:PersonalCliUser.Id;PrimarySmtpAddress=$global:PersonalCliUser.UserPrincipalName;EmailAddresses=@()}
}
$global:PersonalCliMailbox=[pscustomobject]@{
    AddressBookPolicy='MON-EXO-ABP-Schule_Sekretariat';CustomAttribute1='Montessori Schule Aufkirchen - Sekretariat'
    AuditEnabled=$true;AuditLogAgeLimit=[timespan]::FromDays(365);RetainDeletedItemsFor=[timespan]::FromDays(30)
    RoleAssignmentPolicy='MON-EXO-UserRoles-Default';SharingPolicy='MON-EXO-Sharing-Default';RetentionPolicy='MON-EXO-Retention-Default'
    PersistedCapabilities=@();AuditDelegate=@();AuditOwner=@();AuditAdmin=@()
}
$audit=& (Get-Module SchuelerSync) { Get-RequiredExchangeAuditAction -Mailbox $global:PersonalCliMailbox }
foreach($field in @('AuditDelegate','AuditOwner','AuditAdmin')) {$global:PersonalCliMailbox.$field=$audit.$field}
function global:Get-Mailbox { [CmdletBinding()]param($Identity) $global:PersonalCliMailbox }
function global:Get-CASMailbox { [CmdletBinding()]param($Identity)
    [pscustomobject]@{ActiveSyncEnabled=$false;ImapEnabled=$false;MAPIEnabled=$true;OWAEnabled=$true;OWAforDevicesEnabled=$false
        OwaMailboxPolicy='MON-EXO-OWA-Default';PopEnabled=$false;SmtpClientAuthenticationDisabled=$true}
}
foreach($name in @('Get-AddressBookPolicy','Get-RoleAssignmentPolicy','Get-SharingPolicy','Get-RetentionPolicy','Get-OwaMailboxPolicy')) {
    Set-Item "function:global:$name" {[CmdletBinding()]param($Identity) [pscustomobject]@{Name=$Identity}}
}
foreach($name in @('Connect-MgGraph','Connect-ExchangeOnline','New-MgUser','Update-MgUser','Remove-MgUser',
    'New-MgGroupMemberByRef','Remove-MgGroupMemberDirectoryObjectByRef','Revoke-MgUserSignInSession','Set-MgUserManagerByRef','Set-Mailbox','Set-CASMailbox')) {
    Set-Item "function:global:$name" {throw 'Remote operation forbidden in synthetic CLI harness.'}
}
$entryName = switch ($Population) { Student { 'Sync-SchuelerEntra.ps1' } Teacher { 'Sync-LehrerEntra.ps1' } default { 'Sync-PersonalEntra.ps1' } }
$entry=Join-Path $repo $entryName
$previewParameters=@{}
if($Preview) {$previewParameters.WhatIf=$true}
$global:LASTEXITCODE=0
switch($Mode) {
    List { $result=& $entry -List -PassThru @previewParameters }
    Object { $result=& $entry -PassThru @previewParameters }
    Export {
        $exportParameters = @{}
        if ($Population -eq 'Personnel') { $exportParameters.PassThru = $true }
        $result=& $entry -ListToExcel -File $File @exportParameters @previewParameters
    }
    Compare { $result=& $entry -File $File }
    Update { $result=& $entry -File $File -Update -WhatIf }
    Add { $result=& $entry -Add -Vorname Nora -Nachname Testwald -Job Unterstützung -WhatIf }
    Remove { $result=& $entry -Remove -UPN 'ltestwald@example.invalid' -WhatIf }
    Exchange { $result=& $entry -ConfigureExchangeOnlineOnly -UserPrincipalName 'ltestwald@example.invalid' -WhatIf }
    Tui { $result=& $entry -ManageAccounts -WhatIf }
    Invalid { $result=& $entry -Add -Vorname Lea -Nachname Testwald -Job ExterneAdmins -WhatIf }
}

$result | ConvertTo-Json -Depth 6
if ($null -ne $result -and $null -ne $result.PSObject.Properties['HasErrors'] -and $result.HasErrors) { exit 1 }
