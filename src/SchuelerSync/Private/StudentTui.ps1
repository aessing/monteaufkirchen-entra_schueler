function Assert-StudentClassTarget {
    param([object] $Target, [System.Collections.IDictionary] $Config)
    $group = Get-WritablePersonnelRole -GroupId $Target.Id
    if ([string]$group.DisplayName -cne [string]$Target.DisplayName -or
        -not $group.DisplayName.StartsWith($Config.StudentSettings.ClassGroupPrefix,[StringComparison]::OrdinalIgnoreCase)) { throw 'CLS-Zielgruppe hat sich geändert.' }
    $matchingGroups = @(Get-MgGroup -All -Property @('id','displayName') -ErrorAction Stop | Where-Object DisplayName -eq $Target.DisplayName)
    if ($matchingGroups.Count -ne 1 -or $matchingGroups[0].Id -ine $Target.Id) { throw 'CLS-Zielgruppe ist nicht eindeutig.' }
    $class = $group.DisplayName.Substring($Config.StudentSettings.ClassGroupPrefix.Length)
    [pscustomobject]@{Id=$group.Id;DisplayName=$group.DisplayName;Department=$class;OfficeLocation=(Get-OfficeLocation $class)}
}

function Assert-StudentTeacherTarget {
    param([object] $Target)
    $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
    $role = Get-MgGroup -GroupId $config.Roles.School.Id -Property @('id','displayName') -ErrorAction Stop
    if ($role.Id -ine $config.Roles.School.Id -or $role.DisplayName -cne $config.Roles.School.Name) { throw 'Schul-Pädagogikrolle ist nicht bestätigt.' }
    $user = Get-MgUser -UserId $Target.Id -Property @('id','displayName','userPrincipalName','accountEnabled','userType') -ErrorAction Stop
    if ($user.Id -ine $Target.Id -or $user.UserPrincipalName -ine $Target.UPN -or $user.DisplayName -cne $Target.DisplayName -or
        $user.AccountEnabled -ne $true -or $user.UserType -ne 'Member') { throw 'Klassenlehrer ist nicht mehr eindeutig als aktives Mitgliedskonto bestätigt.' }
    $direct = @(Get-FreshEntraUserDirectGroup -UserId $Target.Id)
    if (@($direct | Where-Object Id -eq $role.Id).Count -ne 1 -or @($direct | Where-Object Id -eq (Get-ConfiguredStudentRole).Id).Count -gt 0) { throw 'Klassenlehrer gehört nicht eindeutig zum Schul-Pädagogikteam.' }
    return $user
}

function Get-StudentTuiOption {
    param([System.Collections.IDictionary] $Config, [ValidateSet('Class','Teacher')][string] $Operation)
    if ($Operation -eq 'Class') {
        $groups = @(Get-MgGroup -All -Property @('id','displayName','assignedLicenses','groupTypes','membershipRule') -ErrorAction Stop)
        foreach ($group in @($groups | Where-Object { $_.DisplayName.StartsWith($Config.StudentSettings.ClassGroupPrefix,[StringComparison]::OrdinalIgnoreCase) } | Sort-Object DisplayName)) {
            try {
                $target = Assert-StudentClassTarget -Target $group -Config $Config
                $target | Add-Member -NotePropertyName Label -NotePropertyValue "$($target.DisplayName) (Jahrgang $($target.OfficeLocation))"
                $target
            } catch { Write-Warning "$($group.DisplayName): $($_.Exception.Message)" }
        }
    } else {
        $teachers = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
        $options = @()
        foreach ($member in @(Get-MgGroupMember -GroupId $teachers.Roles.School.Id -All -ErrorAction Stop)) {
            $type = [string](Get-DisabledAccountMemberValue $member '@odata.type')
            if ($type -and $type -ne '#microsoft.graph.user') { continue }
            $id = [string](Get-DisabledAccountMemberValue $member id)
            if (-not $id) { throw 'Lehrerkandidat ohne Objekt-ID.' }
            $user = Get-MgUser -UserId $id -Property @('id','displayName','userPrincipalName','accountEnabled','userType') -ErrorAction Stop
            if ($user.AccountEnabled -ne $true -or $user.UserType -ne 'Member') { continue }
            $target = [pscustomobject]@{Id=$id;DisplayName=$user.DisplayName;UPN=$user.UserPrincipalName;Label="$($user.DisplayName) <$($user.UserPrincipalName)>"}
            $null = Assert-StudentTeacherTarget -Target $target
            $options += $target
        }
        $options | Sort-Object Label,Id -Unique
    }
}

function Select-StudentTuiTarget {
    param([AllowEmptyCollection()][object[]] $Options, [string] $Title)
    $cursor = 0
    while ($true) {
        Clear-Host
        Write-Host $Title -ForegroundColor Cyan
        if ($Options.Count -eq 0) {
            Write-Host 'Keine zulässigen Ziele verfügbar.'
            Write-Host 'Pfeile: bewegen  Enter: Ziel wählen  Esc: abbrechen' -ForegroundColor Cyan
            return
        }
        $height = try { $Host.UI.RawUI.WindowSize.Height } catch { 24 }
        $pageSize = [math]::Max(5,$height - 5)
        $start = [int]([math]::Floor($cursor / $pageSize) * $pageSize)
        for ($index=$start; $index -lt [math]::Min($Options.Count,$start+$pageSize); $index++) {
            Write-Host "$(if ($index -eq $cursor) {'>'} else {' '}) $($Options[$index].Label)"
        }
        Write-Host ''
        Write-Host 'Pfeile: bewegen  Enter: Ziel wählen  Esc: abbrechen' -ForegroundColor Cyan
        switch (Read-DisabledAccountTuiKey) {
            Escape { return }
            UpArrow { if ($cursor -gt 0) { $cursor-- } }
            DownArrow { if ($cursor -lt $Options.Count-1) { $cursor++ } }
            Enter { return $Options[$cursor] }
        }
    }
}

function Assert-StudentAssignmentAccount {
    param([object] $Account, [System.Collections.IDictionary] $Config)
    if ($Config.ManagementPopulation -ne 'Student' -or -not $Account.MutationAllowed) { throw 'Auswahl ist kein zulässiges Schülerkonto.' }
    $user = Assert-PersonalMutationTarget -UserId $Account.UserId -ExpectedUpn $Account.UPN -Config $Config -Account $Account
    Assert-CurrentStudentMembership -UserId $Account.UserId
    if ($user.AccountEnabled -ne $Account.AccountEnabled) { throw 'Kontostatus hat sich seit der Auswahl geändert.' }
    return $user
}

function Assert-StudentClassState {
    param([object] $Account, [System.Collections.IDictionary] $Config, [AllowEmptyCollection()][string[]] $ClassIds,
        [AllowEmptyString()][string] $Department, [AllowEmptyString()][string] $OfficeLocation)
    $user = Assert-StudentAssignmentAccount -Account $Account -Config $Config
    $current = @(Get-FreshEntraUserDirectGroup -UserId $Account.UserId | Where-Object { $_.DisplayName.StartsWith($Config.StudentSettings.ClassGroupPrefix,[StringComparison]::OrdinalIgnoreCase) })
    if ((@($current | ForEach-Object {$_.Id} | Sort-Object) -join ',') -ine (@($ClassIds | Sort-Object) -join ',') -or
        [string](Get-ComparisonPropertyValue $user Department) -cne $Department -or [string](Get-ComparisonPropertyValue $user OfficeLocation) -cne $OfficeLocation) { throw 'Klasse oder Jahrgang hat sich seit der Auswahl geändert.' }
}

function Invoke-StudentAssignment {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    param([object[]] $Accounts, [System.Collections.IDictionary] $Config, [ValidateSet('Class','Teacher')][string] $Operation, [object] $Target)
    if (-not $WhatIfPreference) { $null = Connect-DisabledAccountWriteGraph -TenantId $Config.SessionTenantId -Operation $Operation }
    $activity = if ($Operation -eq 'Class') { 'Klasse/Jahrgang wechseln' } else { 'Klassenlehrer wechseln' }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $index = 0
    try {
        foreach ($account in $Accounts) {
            $index++; $status='Skipped'; $message=''; $changed=$false
            Write-Progress -Id 2 -Activity $activity -Status "Prüfe $index von $($Accounts.Count): $($account.DisplayName)" -PercentComplete ([int](100*($index-1)/$Accounts.Count))
            try {
                if (-not $seen.Add([string]$account.UserId)) { throw 'Konto wurde mehrfach ausgewählt.' }
                $null = Assert-StudentAssignmentAccount -Account $account -Config $Config
                if ($Operation -eq 'Class') {
                    $destination = Assert-StudentClassTarget -Target $Target -Config $Config
                    $expectedIds = @($account.ClassGroups | ForEach-Object { [string]$_.Id })
                    $department = [string]$account.Department; $office = [string]$account.OfficeLocation
                    Assert-StudentClassState -Account $account -Config $Config -ClassIds $expectedIds -Department $department -OfficeLocation $office
                    foreach ($group in $account.ClassGroups) {
                        $fresh = Get-WritablePersonnelRole -GroupId $group.Id
                        if ($fresh.DisplayName -cne $group.DisplayName) { throw 'Bisherige CLS-Gruppe hat sich geändert.' }
                    }
                    if ($expectedIds.Count -eq 1 -and $expectedIds[0] -ieq $destination.Id -and $department -ceq $destination.Department -and $office -ceq $destination.OfficeLocation) { $status='Compliant' }
                    elseif ($PSCmdlet.ShouldProcess("$($account.DisplayName) <$($account.UPN)> → $($destination.DisplayName)",$activity)) {
                        $status='Failed'
                        if ($destination.Id -notin $expectedIds) {
                            $null = Assert-StudentClassTarget -Target $Target -Config $Config
                            Assert-StudentClassState -Account $account -Config $Config -ClassIds $expectedIds -Department $department -OfficeLocation $office
                            New-MgGroupMemberByRef -GroupId $destination.Id -BodyParameter @{'@odata.id'="https://graph.microsoft.com/v1.0/directoryObjects/$($account.UserId)"} -ErrorAction Stop
                            $changed=$true; $expectedIds+= [string]$destination.Id
                            Assert-StudentClassState -Account $account -Config $Config -ClassIds $expectedIds -Department $department -OfficeLocation $office
                        }
                        $desired = [pscustomobject]@{Department=$destination.Department;OfficeLocation=$destination.OfficeLocation}
                        $differences = @(foreach ($field in @('Department','OfficeLocation')) {
                            New-StateDifference -Area Entra -Field $field -Current $(if ($field -eq 'Department') {$department} else {$office}) -Desired $desired.$field
                        })
                        $null = Assert-StudentClassTarget -Target $Target -Config $Config
                        Assert-StudentClassState -Account $account -Config $Config -ClassIds $expectedIds -Department $department -OfficeLocation $office
                        if ($differences.Count -gt 0) { $changed=$true }
                        $attributes = Set-EntraStudentAttribute -UserId $account.UserId -Desired $desired -Differences $differences -Confirm:$false
                        if (-not $attributes.Verified) { throw 'Klassenattribute wurden nicht bestätigt.' }
                        $department=$destination.Department; $office=$destination.OfficeLocation
                        foreach ($old in @($account.ClassGroups | Where-Object Id -ne $destination.Id)) {
                            $null = Assert-StudentClassTarget -Target $Target -Config $Config
                            Assert-StudentClassState -Account $account -Config $Config -ClassIds $expectedIds -Department $department -OfficeLocation $office
                            $fresh = Get-WritablePersonnelRole -GroupId $old.Id
                            if ($fresh.DisplayName -cne $old.DisplayName) { throw 'Bisherige CLS-Gruppe hat sich geändert.' }
                            Remove-MgGroupMemberDirectoryObjectByRef -GroupId $old.Id -DirectoryObjectId $account.UserId -ErrorAction Stop
                            $changed=$true; $expectedIds=@($expectedIds | Where-Object { $_ -ine $old.Id })
                        }
                        Assert-StudentClassState -Account $account -Config $Config -ClassIds @([string]$destination.Id) -Department $department -OfficeLocation $office
                        $status='Succeeded'
                    } else { $status=if ($WhatIfPreference) {'WhatIf'} else {'Skipped'} }
                } else {
                    $null = Assert-StudentTeacherTarget -Target $Target
                    if (-not $account.ManagerAvailable) { throw 'Bisheriger Klassenlehrer ist nicht prüfbar.' }
                    $current = [string](Get-UserManagerId -Snapshot ([pscustomobject]@{ManagerByUserId=(New-CaseInsensitiveHashtable)}) -UserId $account.UserId)
                    if ($current -ine $account.ManagerId) { throw 'Klassenlehrer hat sich seit der Auswahl geändert.' }
                    if ($current -ieq $Target.Id) { $status='Compliant' }
                    elseif ($PSCmdlet.ShouldProcess("$($account.DisplayName) <$($account.UPN)> → $($Target.Label)",$activity)) {
                        $null = Assert-StudentAssignmentAccount -Account $account -Config $Config
                        $null = Assert-StudentTeacherTarget -Target $Target
                        $freshManager = [string](Get-UserManagerId -Snapshot ([pscustomobject]@{ManagerByUserId=(New-CaseInsensitiveHashtable)}) -UserId $account.UserId)
                        if ($freshManager -ine $current) { throw 'Klassenlehrer hat sich vor dem Schreiben geändert.' }
                        $status='Failed'; $changed=$true
                        $result = Set-EntraStudentManager -UserId $account.UserId -CurrentManagerId $current -DesiredManagerId $Target.Id -Confirm:$false
                        if (-not $result.Verified) { throw 'Klassenlehrer wurde nicht bestätigt.' }
                        $status='Succeeded'
                    } else { $status=if ($WhatIfPreference) {'WhatIf'} else {'Skipped'} }
                }
            } catch { $message=$_.Exception.Message; if ($changed) {$message='Teiländerung möglich. Bestand neu prüfen. '+$message} }
            [pscustomobject]@{DisplayName=$account.DisplayName;UPN=$account.UPN;UserId=$account.UserId;Status=$status;Message=$message;Operation=$Operation;TargetId=$Target.Id;Changed=$changed}
        }
    } finally { Write-Progress -Id 2 -Activity $activity -Completed }
}

function Invoke-StudentTuiAction {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess','',Justification='Delegates confirmed mutations to Invoke-StudentAssignment, forwarding WhatIf.')]
    [CmdletBinding(SupportsShouldProcess)]
    param([object[]] $Accounts, [System.Collections.IDictionary] $Config, [ValidateSet('Class','Teacher')][string] $Operation)
    $title = if ($Operation -eq 'Class') {'Klasse/Jahrgang wählen'} else {'Klassenlehrer wählen (Schul-Pädagogikteam)'}
    $target = Select-StudentTuiTarget -Options @(Get-StudentTuiOption -Config $Config -Operation $Operation) -Title $title
    if ($null -eq $target) { return }
    Clear-Host
    Write-Host "$title | Tenant $($Config.SessionTenantId)" -ForegroundColor Yellow
    foreach ($account in $Accounts) {
        $previous = if ($Operation -eq 'Class') { @($account.ClassGroups | ForEach-Object {$_.DisplayName}) -join ', ' } else { $account.Teacher }
        Write-Host "$($account.DisplayName) <$($account.UPN)>: $previous → $($target.Label)"
    }
    Write-Host 'Y/J: bestätigen  N/Esc: abbrechen. Exceldateien werden nicht aktualisiert.'
    do { $answer=Read-DisabledAccountTuiKey } while ($answer -notin @('Y','J','N','Escape'))
    if ($answer -in @('Y','J')) { Invoke-StudentAssignment -Accounts $Accounts -Config $Config -Operation $Operation -Target $target -WhatIf:$WhatIfPreference -Confirm:$false }
}
