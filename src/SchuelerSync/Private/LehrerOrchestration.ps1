function Assert-LehrerExchangePolicy {
    param([System.Collections.IDictionary] $Config)
    foreach ($personnelProfile in $Config.Profiles.Values) {
        if (-not (Get-AddressBookPolicy -Identity $personnelProfile.AddressBookPolicy -ErrorAction Stop)) {
            throw "Adressbuchrichtlinie '$($personnelProfile.AddressBookPolicy)' fehlt."
        }
    }
    foreach ($definition in @(
            @{ Name = $Config.Exchange.RoleAssignmentPolicy; Command = 'Get-RoleAssignmentPolicy' }
            @{ Name = $Config.Exchange.SharingPolicy; Command = 'Get-SharingPolicy' }
            @{ Name = $Config.Exchange.RetentionPolicy; Command = 'Get-RetentionPolicy' }
            @{ Name = $Config.Exchange.OwaMailboxPolicy; Command = 'Get-OwaMailboxPolicy' }
        )) {
        if (-not (& $definition.Command -Identity $definition.Name -ErrorAction Stop)) {
            throw "Exchange-Richtlinie '$($definition.Name)' fehlt."
        }
    }
}

function Wait-LehrerMailbox {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([Parameter(Mandatory)][System.Collections.IDictionary] $Targets,
        [int] $MaxRetries = 5, [int] $RetryDelaySeconds = 60,
        [scriptblock] $SleepAction = { param($seconds) Start-Sleep -Seconds $seconds })
    $pending = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($upn in $Targets.Keys) { [void]$pending.Add([string]$upn) }
    $ready = [Collections.Generic.List[object]]::new()
    $failed = [Collections.Generic.List[object]]::new()
    try {
        for ($attempt = 0; $attempt -le $MaxRetries -and $pending.Count -gt 0; $attempt++) {
            if ($attempt -gt 0) {
                Write-Progress -Id 3 -ParentId 1 -Activity 'Personalpostfächer bereitstellen' `
                    -Status "Warte $RetryDelaySeconds Sekunden vor Versuch $($attempt + 1) von $($MaxRetries + 1), $($pending.Count) Postfach/Postfächer ausstehend ..." `
                    -PercentComplete ([int](100 * $attempt / ($MaxRetries + 1)))
                & $SleepAction $RetryDelaySeconds
            }
            Write-Progress -Id 3 -ParentId 1 -Activity 'Personalpostfächer bereitstellen' `
                -Status "Prüfe Versuch $($attempt + 1) von $($MaxRetries + 1), $($pending.Count) Postfach/Postfächer ausstehend ..." `
                -PercentComplete ([int](100 * ($attempt + 1) / ($MaxRetries + 1)))
            foreach ($upn in @($pending)) {
                try {
                    $config = $Targets[$upn]
                    $state = Get-StudentMailboxState -UserPrincipalName $upn -Config $config
                    if (-not $state.Exists) { continue }
                    $configuration = if ($null -ne $state.PSObject.Properties['Differences'] -and @($state.Differences).Count -eq 0) {
                        New-ExchangeConfigurationResult -UserPrincipalName $upn -Status Compliant -Changed:$false `
                            -Differences @() -RemainingDifferences @() -ErrorMessage $null
                    } elseif ($WhatIfPreference) {
                        Set-StudentMailboxConfiguration -UserPrincipalName $upn -Config $config -MailboxState $state -WhatIf
                    } elseif ($PSCmdlet.ShouldProcess($upn, 'Configure personnel mailbox')) {
                        Set-StudentMailboxConfiguration -UserPrincipalName $upn -Config $config -MailboxState $state -Confirm:$false
                    } else { $null }
                    if ($null -ne $configuration -and $configuration.Status -eq 'Failed') {
                        $failed.Add([pscustomobject]@{ UserPrincipalName = $upn; Error = $configuration.Error })
                    } else {
                        $ready.Add([pscustomobject]@{ UserPrincipalName = $upn; Configuration = $configuration })
                    }
                    [void]$pending.Remove($upn)
                } catch {
                    $failed.Add([pscustomobject]@{ UserPrincipalName = $upn; Error = $_.Exception.Message })
                    [void]$pending.Remove($upn)
                }
            }
            if ($WhatIfPreference) { break }
        }
    } finally { Write-Progress -Id 3 -ParentId 1 -Activity 'Personalpostfächer bereitstellen' -Completed }
    return [pscustomobject]@{ Ready = @($ready); Missing = @($pending); Failed = @($failed) }
}

function ConvertTo-LehrerExchangeAction {
    param([object] $Batch, [switch] $WhatIfMode)
    foreach ($item in $Batch.Ready) {
        $status = if ($null -eq $item.Configuration) { 'Skipped' } else { [string]$item.Configuration.Status }
        if ($WhatIfMode -and $status -eq 'Planned') { $status = 'WhatIf' }
        New-StudentActionResult -UserPrincipalName $item.UserPrincipalName -Phase Exchange -Status $status
    }
    foreach ($upn in $Batch.Missing) {
        New-StudentActionResult -UserPrincipalName $upn -Phase Exchange -Status $(if ($WhatIfMode) { 'WhatIf' } else { 'Pending' }) `
            -Message 'EXO-Konfiguration ausstehend.' -RecoveryCommand ".\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN '$($upn.Replace("'", "''"))'"
    }
    foreach ($item in $Batch.Failed) {
        New-StudentActionResult -UserPrincipalName $item.UserPrincipalName -Phase Exchange -Status Failed -Message $item.Error `
            -RecoveryCommand ".\Sync-LehrerEntra.ps1 -ConfigureExchangeOnlineOnly -UPN '$($item.UserPrincipalName.Replace("'", "''"))'"
    }
}

function ConvertTo-SafeLehrerComparison {
    param([object] $Comparison)
    $safe = [ordered]@{}
    foreach ($category in 'NewPeople','Departures','ChangedPeople','ExistingPeople') {
        $safe[$category] = @($Comparison.$category | ForEach-Object {
            [pscustomobject]@{
                RowNumber = Get-ComparisonPropertyValue $_.Person RowNumber
                NameMitRufname = Get-ComparisonPropertyValue $_.Person NameMitRufname
                UserId = Get-ComparisonPropertyValue $_.User Id
                DisplayName = if ($null -ne $_.DesiredState) { $_.DesiredState.DisplayName } else { $_.User.DisplayName }
                UserPrincipalName = if ($null -ne $_.DesiredState) { $_.DesiredState.UserPrincipalName } else { $_.User.UserPrincipalName }
                Job = if ($null -ne $_.Person) { Get-ComparisonPropertyValue $_.Person Job } else { Get-ComparisonPropertyValue $_.User JobTitle }
                LicenseReason = if ($null -ne $_.DesiredState) { $_.DesiredState.LicenseReason } else { 'Unverändert' }
                Role = if ($null -ne $_.DesiredState) { $_.DesiredState.Role.DisplayName } else { @($_.SourceRoles) -join ', ' }
                AddressBookPolicy = if ($null -ne $_.DesiredState) { $_.DesiredState.Exchange.AddressBookPolicy } else { '' }
                CustomAttribute1 = if ($null -ne $_.DesiredState) { $_.DesiredState.Exchange.CustomAttribute1 } else { '' }
                AccountEnabled = Get-ComparisonPropertyValue $_.User AccountEnabled
                Differences = @($_.Differences | ForEach-Object {
                    [pscustomobject]@{ Area = $_.Area; Field = $_.Field; Current = $_.Current; Desired = $_.Desired; Action = $_.Action }
                })
            }
        })
    }
    $safe.Warnings = @($Comparison.Warnings | ForEach-Object { [pscustomobject]@{ Code = "$($_.Area).$($_.Field)"; Message = $_.Message; RowNumber = $_.RowNumber } })
    $safe.Errors = @($Comparison.Errors | ForEach-Object { [pscustomobject]@{ Code = "$($_.Area).$($_.Field)"; Message = $_.Message; RowNumber = $_.RowNumber } })
    return [pscustomobject]$safe
}

function Write-LehrerReport {
    param([object] $Comparison, [object[]] $Actions, [string[]] $HeaderLines, [string] $OutputFile,
        [switch] $ActionsOnly)
    $tables = [ordered]@{
        'Neuzugänge' = @($Comparison.NewPeople)
        'Abgänge' = @($Comparison.Departures)
        'Änderungen' = @($Comparison.ChangedPeople)
        'Bestehende' = @($Comparison.ExistingPeople)
        'Warnungen und Fehler' = @($Comparison.Warnings) + @($Comparison.Errors)
        'Aktionsergebnisse' = @($Actions)
    }
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($line in $HeaderLines) { if ($line) { $lines.Add($line) } }
    foreach ($title in $tables.Keys) {
        if ($lines.Count -gt 0) { $lines.Add('') }
        $lines.Add("$title ($(@($tables[$title]).Count))")
        $showInTerminal = -not $ActionsOnly -or $title -eq 'Aktionsergebnisse'
        if ($showInTerminal) {
            Write-Host ''
            Write-Host "$title ($(@($tables[$title]).Count))" -ForegroundColor Cyan
        }
        if (@($tables[$title]).Count -eq 0) { continue }
        if ($title -eq 'Abgänge') {
            $rows = @($tables[$title] | Select-Object DisplayName,UserPrincipalName,Role,AccountEnabled,UserId)
        } elseif ($title -eq 'Änderungen') {
            $rows = @($tables[$title] | ForEach-Object {
                $entry = $_
                foreach ($difference in $entry.Differences) {
                    [pscustomobject]@{ Name = $entry.NameMitRufname; Job = $entry.Job; Area = $difference.Area
                        Field = $difference.Field; Current = $difference.Current; Desired = $difference.Desired; Action = $difference.Action }
                }
            })
        } elseif ($title -eq 'Bestehende') {
            $rows = @($tables[$title] | Select-Object NameMitRufname,UserId,DisplayName,UserPrincipalName,Job)
        } else { $rows = @($tables[$title]) }
        $formatted = ($rows | Format-Table -AutoSize -Wrap | Out-String -Width 220).TrimEnd()
        if ($showInTerminal) { Write-Information $formatted -InformationAction Continue }
        $lines.Add($formatted)
    }
    if ($OutputFile) {
        Write-ComparisonReportFile -Path $OutputFile -Text ($lines -join [Environment]::NewLine) -Confirm:$false
    }
}

function Invoke-LehrerSync {
    [CmdletBinding(DefaultParameterSetName = 'Sync', SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(ParameterSetName = 'Sync')][ValidateScript({ [IO.Path]::GetExtension($_) -ieq '.xlsx' })]
        [string] $File = (Join-Path $script:SchuelerSyncRepositoryRoot 'Lehrer.xlsx'),
        [Parameter(ParameterSetName = 'Sync')][switch] $Update,
        [Parameter(ParameterSetName = 'Sync')][switch] $CreateNewUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $UpdateUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $DisableUsers,
        [Parameter(ParameterSetName = 'Sync')][switch] $RevokeSessions,
        [Parameter(ParameterSetName = 'Add', Mandatory)][switch] $Add,
        [Parameter(ParameterSetName = 'Add', Mandatory)][string] $Vorname,
        [Parameter(ParameterSetName = 'Add', Mandatory)][string] $Nachname,
        [Parameter(ParameterSetName = 'Add', Mandatory)]
        [ValidatePattern('(?i)^\s*(?:L|CO-L|PA|OGTS|JAS)(?:\s*,\s*(?:L|CO-L|PA|OGTS|JAS))*\s*$')][string] $Job,
        [Parameter(ParameterSetName = 'Add')][string] $EntraObjectId,
        [Parameter(ParameterSetName = 'Remove', Mandatory)][switch] $Remove,
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)][switch] $ConfigureExchangeOnlineOnly,
        [Parameter(ParameterSetName = 'Remove', Mandatory)]
        [Parameter(ParameterSetName = 'ExchangeOnly', Mandatory)]
        [Alias('UPN','UserPrincipalName')][string[]] $Mail,
        [string] $OutputFile
    )
    $ErrorActionPreference = 'Stop'
    if ($OutputFile) {
        $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
        Assert-ComparisonReportPath -Path $OutputFile
    }
    $progressActivity = 'Lehrerabgleich'
    try {
    $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/LehrerSync.psd1')
    $comparison = [pscustomobject]@{ NewPeople = @(); Departures = @(); ChangedPeople = @(); ExistingPeople = @(); Warnings = @(); Errors = @() }
    $actions = [Collections.Generic.List[object]]::new()
    $header = [Collections.Generic.List[string]]::new()
    $targets = New-CaseInsensitiveHashtable
    $workbook = $null
    $snapshot = $null
    $common = @{ WhatIf = [bool]$WhatIfPreference }
    if ($PSBoundParameters.ContainsKey('Confirm')) { $common.Confirm = $PSBoundParameters.Confirm }
    $mode = if ($Add) { 'Add' } elseif ($Remove) { 'Remove' } elseif ($ConfigureExchangeOnlineOnly) { 'ExchangeOnly' } elseif ($Update) { 'Update' } else { 'Compare' }
    try {
        if ($mode -eq 'Update' -or $mode -eq 'Compare') {
            $selection = Resolve-UpdateSelection -Update:$Update -CreateNewUsers:$CreateNewUsers -UpdateUsers:$UpdateUsers -DisableUsers:$DisableUsers -RevokeSessions:$RevokeSessions
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Lese Excel-Datei ...' -PercentComplete 5
            $File = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($File)
            $workbook = Read-StudentWorkbook -Path $File -Kind Teacher
            if ($Update -and @($workbook.Students).Count -eq 0) { throw 'Leere Personalliste darf keinen Update-Lauf auslösen.' }
        }
        Write-Progress -Id 1 -Activity $progressActivity -Status 'Verbinde mit Microsoft Entra ID ...' -PercentComplete 10
        $context = Connect-SchuelerGraph
        $tenantId = [string](Get-ComparisonPropertyValue $context TenantId)
        if (-not $tenantId -or ($config.ExpectedTenantId -and $tenantId -ine [string]$config.ExpectedTenantId)) {
            throw 'Graph-Tenant konnte nicht bestätigt werden.'
        }
        $header.Add("Entra-Tenant: $tenantId | Konto: $([string](Get-ComparisonPropertyValue $context Account))")
        Write-Progress -Id 1 -Activity $progressActivity -Status 'Lade Entra-Benutzer, Gruppen und Personalrollen ...' -PercentComplete 20
        $snapshot = Get-LehrerSnapshot -Config $config -SkipLicenseGroupValidation:($mode -in @('Remove','ExchangeOnly'))
        if ($mode -eq 'Remove' -or $mode -eq 'ExchangeOnly') {
            if ($mode -eq 'Remove' -and $Mail.Count -ne 1) { throw '-Remove akzeptiert genau einen UPN.' }
            foreach ($address in $Mail) {
                $upn = $address.Trim().ToLowerInvariant()
                $user = $snapshot.UsersByUpn[$upn]
                if ($null -eq $user -or -not $snapshot.PersonnelMemberIds.Contains([string]$user.Id)) {
                    throw "'$upn' gehört keiner direkten Personalrolle an."
                }
                if ($mode -eq 'Remove') {
                    $comparison.Departures += [pscustomobject]@{ Person = $null; User = $user; DesiredState = $null; Differences = @()
                        SourceRoles = @($snapshot.PersonnelRolesByUserId[[string]$user.Id]) }
                } else {
                    $personnelProfile = Get-LehrerProfile -Job ([string]$user.JobTitle) -Config $config
                    $direct = @(Get-UserDirectGroup -Snapshot $snapshot -UserId $user.Id)
                    if (@($direct | Where-Object Id -eq $personnelProfile.Role.Id).Count -ne 1 -or
                        @($direct | Where-Object { $_.DisplayName.StartsWith($config.RoleGroupPrefix, [StringComparison]::OrdinalIgnoreCase) }).Count -ne 1) {
                        throw "Personalrolle und JobTitle von '$upn' widersprechen sich."
                    }
                    $targets[$upn] = $personnelProfile.Exchange
                }
            }
            if ($mode -eq 'ExchangeOnly') {
                Write-Progress -Id 1 -Activity $progressActivity -Status 'Verbinde mit Exchange Online ...' -PercentComplete 30
                $null = Connect-SchuelerExchangeOnline -TenantId $tenantId
            }
        } else {
            $people = if ($mode -eq 'Add') {
                @([pscustomobject]@{ RowNumber = 0; NameMitRufname = "$Nachname, $Vorname"; GivenName = $Vorname.Trim(); Surname = $Nachname.Trim()
                    Job = $Job.Trim().ToUpperInvariant(); EntraObjectId = [string]$EntraObjectId; StoredUpn = ''; StoredMail = ''; Password = '' })
            } else { @($workbook.Students) }
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Verbinde mit Exchange Online ...' -PercentComplete 30
            $null = Connect-SchuelerExchangeOnline -TenantId $tenantId
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Lade Exchange-Empfängeradressen ...' -PercentComplete 35
            $recipients = Get-ExchangeRecipientAddress
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Vergleiche Personaldaten ...' -PercentComplete 40
            $comparison = Compare-LehrerDirectory -People $people -Snapshot $snapshot -Config $config `
                -ExchangeAddressOwners $recipients.AddressOwners -Recovery:($mode -eq 'Add' -and [bool]$EntraObjectId)
            if ($mode -eq 'Add') { $comparison.Departures = @() }
            $mailboxEntries = @($comparison.ChangedPeople) + @($comparison.ExistingPeople)
            $mailboxIndex = 0
            foreach ($entry in $mailboxEntries) {
                $mailboxIndex++
                $mailboxPercent = 55 + [int](20 * $mailboxIndex / $mailboxEntries.Count)
                Write-Progress -Id 1 -Activity $progressActivity `
                    -Status "Prüfe Exchange-Postfach $mailboxIndex von $($mailboxEntries.Count): $($entry.User.UserPrincipalName)" -PercentComplete $mailboxPercent
                try {
                    $state = Get-StudentMailboxState -UserPrincipalName $entry.User.UserPrincipalName -Config $entry.DesiredState.Exchange
                    if (-not $state.Exists) {
                        $comparison.Warnings += New-LehrerIssue -Area Exchange -Field Mailbox -Message "Postfach '$($entry.User.UserPrincipalName)' ist ausstehend." -Person $entry.Person
                    } elseif (@($state.Differences).Count -gt 0) {
                        $entry.Differences = @($entry.Differences) + @($state.Differences | ForEach-Object {
                            New-StateDifference -Area Exchange -Field $_.Field -Current $_.Current -Desired $_.Desired -Action $(if ($_.Action) { $_.Action } else { 'Set' })
                        })
                        if (@($comparison.ExistingPeople | Where-Object { $_.User.Id -eq $entry.User.Id }).Count -gt 0) {
                            $comparison.ExistingPeople = @($comparison.ExistingPeople | Where-Object { $_.User.Id -ne $entry.User.Id })
                            $comparison.ChangedPeople += $entry
                        }
                    }
                } catch {
                    $comparison.Errors += New-LehrerIssue -Area Exchange -Field Preflight -Message $_.Exception.Message -Person $entry.Person
                }
            }
        }
        if ($mode -ne 'Remove') {
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Prüfe Richtlinien und Arbeitsmappe ...' -PercentComplete 78
            Assert-LehrerExchangePolicy -Config $config
            if ($Update -and -not $WhatIfPreference -and $workbook) {
                $identityEntries = @($comparison.ExistingPeople) + @($comparison.ChangedPeople)
                $needsIdentityBackfill = $selection.UpdateUsers -and @($identityEntries | Where-Object {
                    $_.Person.EntraObjectId -ine $_.User.Id -or $_.Person.StoredUpn -ine $_.User.UserPrincipalName -or
                        $_.Person.StoredMail -ine $(if ($_.User.Mail) { $_.User.Mail } else { $_.User.UserPrincipalName })
                }).Count -gt 0
                if (($selection.CreateNewUsers -and @($comparison.NewPeople).Count -gt 0) -or $needsIdentityBackfill) {
                    Assert-WorkbookSafeForPasswordWrite -Path $File -Kind Teacher
                }
                Assert-StudentWorkbookVersion -Path $File -ExpectedSourceHash $workbook.SourceHash
            }
        }
    } catch {
        $comparison.Errors += New-LehrerIssue -Area Preflight -Field Inventory -Message $_.Exception.Message
    }
    if ($snapshot) {
        $activeMembers = @($snapshot.PersonnelMemberIds | Where-Object { $snapshot.UsersById[$_].AccountEnabled -eq $true }).Count
        $header.Add("Personalrollen: $($snapshot.PersonnelMemberIds.Count) direkte Mitglieder | Aktiv: $activeMembers | Abgänge: $(@($comparison.Departures).Count)")
        $header.Add("Quellrollen: $($config.Roles.School.Name), $($config.Roles.Ganztag.Name)")
    }
    $comparisonPrinted = $false
    if ($mode -eq 'Update' -and @($comparison.Errors).Count -eq 0) {
        Write-Progress -Id 1 -Activity $progressActivity -Status 'Zeige Vergleich vor den Aktionen ...' -PercentComplete 79
        Write-LehrerReport -Comparison (ConvertTo-SafeLehrerComparison -Comparison $comparison) -Actions @() -HeaderLines @($header)
        $comparisonPrinted = $true
    }
    if (@($comparison.Errors).Count -eq 0) {
        if ($mode -eq 'Update' -or $mode -eq 'Add') {
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Verarbeite Personalkonten ...' -PercentComplete 80
            $createEntries = if ($mode -eq 'Add') { @($comparison.NewPeople) } elseif ($selection.CreateNewUsers) { @($comparison.NewPeople) } else { @() }
            $updateEntries = if ($mode -eq 'Add') { @($comparison.ChangedPeople) + @($comparison.ExistingPeople | Where-Object { $_.User.AccountEnabled -eq $false }) }
                elseif ($selection.UpdateUsers) { @($comparison.ChangedPeople) } else { @() }
            $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            if ($workbook) { foreach ($person in $workbook.Students) { if ($person.Password) { [void]$used.Add($person.Password) } } }
            foreach ($entry in $createEntries) {
                $action = Invoke-LehrerCreate -Entry $entry -Config $config -File $(if ($workbook) { $File } else { '' }) -WorkbookState $workbook -UsedPasswords $used @common
                $actions.Add($action)
                if ($action.Status -eq 'Succeeded') { $targets[$action.UserPrincipalName] = $entry.DesiredState.Exchange }
            }
            foreach ($entry in $updateEntries) {
                $action = Invoke-LehrerUpdate -Entry $entry -Config $config @common
                $actions.Add($action)
                if ($action.Status -eq 'Succeeded') { $targets[$action.UserPrincipalName] = $entry.DesiredState.Exchange }
            }
            if ($mode -eq 'Add') {
                foreach ($entry in $comparison.ExistingPeople) {
                    if ($entry.User.AccountEnabled -eq $false) { continue }
                    $actions.Add((New-StudentActionResult -UserId $entry.User.Id -UserPrincipalName $entry.User.UserPrincipalName -Phase Update -Status Compliant))
                    $targets[[string]$entry.User.UserPrincipalName] = $entry.DesiredState.Exchange
                }
            }
            if ($workbook -and $selection.UpdateUsers -and -not $WhatIfPreference) {
                $declinedIds = @($actions | Where-Object { $_.Status -in @('Skipped', 'Failed') } | ForEach-Object { $_.UserId })
                $backfill = @(@($comparison.ExistingPeople) + @($comparison.ChangedPeople) | Where-Object { $_.User.Id -notin $declinedIds } | Where-Object {
                    $_.Person.EntraObjectId -ine $_.User.Id -or $_.Person.StoredUpn -ine $_.User.UserPrincipalName -or
                        $_.Person.StoredMail -ine $(if ($_.User.Mail) { $_.User.Mail } else { $_.User.UserPrincipalName })
                } | ForEach-Object {
                    [pscustomobject]@{ RowNumber = [int]$_.Person.RowNumber; EntraObjectId = [string]$_.User.Id
                        UPN = [string]$_.User.UserPrincipalName; Mail = $(if ($_.User.Mail) { [string]$_.User.Mail } else { [string]$_.User.UserPrincipalName }) }
                })
                if ($backfill.Count -gt 0 -and @($actions | Where-Object Status -eq Failed).Count -eq 0 -and
                    $PSCmdlet.ShouldProcess($File, 'Personalidentitäten in Excel nachtragen')) {
                    try {
                        $written = Write-StudentWorkbookUpdate -Path $File -ExpectedSourceHash $workbook.SourceHash -Kind Teacher -Updates $backfill -Confirm:$false
                        $workbook.SourceHash = $written.SourceHash
                    } catch { $actions.Add((New-StudentActionResult -Phase Workbook -Status Failed -Message $_.Exception.Message)) }
                }
            }
            if ($mode -eq 'Update') {
                foreach ($entry in $comparison.Departures) {
                    foreach ($phase in @('Disable','RevokeSessions')) {
                        if (($phase -eq 'Disable' -and -not $selection.DisableUsers) -or ($phase -eq 'RevokeSessions' -and -not $selection.RevokeSessions)) { continue }
                        $upn = [string]$entry.User.UserPrincipalName; $id = [string]$entry.User.Id
                        if (-not $PSCmdlet.ShouldProcess($upn, $phase)) {
                            $actions.Add((New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })))
                            continue
                        }
                        try {
                            $fresh = @(Get-FreshEntraUserDirectGroup -UserId $id)
                            if (@($fresh | Where-Object { $_.Id -in @($config.Roles.School.Id, $config.Roles.Ganztag.Id) }).Count -eq 0) {
                                throw 'Personalkonto gehört keiner direkten Personalrolle mehr an.'
                            }
                            if ($phase -eq 'Disable') {
                                $done = Disable-EntraStudent -UserId $id -Confirm:$false
                                if (-not $done.Verified) { throw 'Deaktivierung nicht bestätigt.' }
                            } else {
                                $response = Revoke-EntraStudentSession -UserId $id -Selected -Confirm:$false
                                if ($response -is [bool]) { $revoked = $response } else { $revoked = Get-ComparisonPropertyValue $response Value }
                                if ($revoked -ne $true) { throw 'Sitzungswiderruf nicht bestätigt.' }
                            }
                            $actions.Add((New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Succeeded))
                        } catch { $actions.Add((New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message)) }
                    }
                }
                if ($selection.ConfigureAllActiveExchange) {
                    foreach ($entry in @($comparison.ChangedPeople) + @($comparison.ExistingPeople)) { $targets[[string]$entry.User.UserPrincipalName] = $entry.DesiredState.Exchange }
                }
            }
        } elseif ($mode -eq 'Remove') {
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Deaktiviere Personalkonto und widerrufe Sitzungen ...' -PercentComplete 80
            $entry = $comparison.Departures[0]
            $id = [string]$entry.User.Id; $upn = [string]$entry.User.UserPrincipalName
            foreach ($phase in @('Disable','RevokeSessions')) {
                if (-not $PSCmdlet.ShouldProcess($upn, $phase)) {
                    $actions.Add((New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status $(if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' })))
                    continue
                }
                try {
                    $fresh = Get-FreshEntraUserDirectGroup -UserId $id
                    if (@($fresh | Where-Object { $_.Id -in @($config.Roles.School.Id, $config.Roles.Ganztag.Id) }).Count -eq 0) { throw 'Personalkonto gehört keiner direkten Personalrolle mehr an.' }
                    if ($phase -eq 'Disable') { $done = Disable-EntraStudent -UserId $id -Confirm:$false; if (-not $done.Verified) { throw 'Deaktivierung nicht bestätigt.' } }
                    else {
                        $response = Revoke-EntraStudentSession -UserId $id -Selected -Confirm:$false
                        if ($response -is [bool]) { $revoked = $response } else { $revoked = Get-ComparisonPropertyValue $response Value }
                        if ($revoked -ne $true) { throw 'Sitzungswiderruf nicht bestätigt.' }
                    }
                    $actions.Add((New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Succeeded))
                } catch { $actions.Add((New-StudentActionResult -UserId $id -UserPrincipalName $upn -Phase $phase -Status Failed -Message $_.Exception.Message)) }
            }
        }
        if ($mode -eq 'ExchangeOnly' -or $targets.Count -gt 0) {
            Write-Progress -Id 1 -Activity $progressActivity -Status 'Prüfe und konfiguriere Exchange-Postfächer ...' -PercentComplete 90
            if ($mode -ne 'ExchangeOnly' -and $WhatIfPreference) {
                foreach ($upn in $targets.Keys) { $actions.Add((New-StudentActionResult -UserPrincipalName $upn -Phase Exchange -Status WhatIf)) }
            } else {
                try {
                    $null = Connect-SchuelerExchangeOnline -TenantId $tenantId
                    $batch = Wait-LehrerMailbox -Targets $targets -MaxRetries $config.Exchange.MaxMailboxRetries -RetryDelaySeconds $config.Exchange.RetryDelaySeconds @common
                    foreach ($action in @(ConvertTo-LehrerExchangeAction -Batch $batch -WhatIfMode:$WhatIfPreference)) { $actions.Add($action) }
                } catch { $actions.Add((New-StudentActionResult -Phase Exchange -Status Failed -Message $_.Exception.Message)) }
            }
        }
    }
    $safe = ConvertTo-SafeLehrerComparison -Comparison $comparison
    if ($OutputFile) { $OutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile) }
    Write-Progress -Id 1 -Activity $progressActivity -Status 'Erzeuge Ergebnisbericht ...' -PercentComplete 95
    Write-LehrerReport -Comparison $safe -Actions @($actions) -HeaderLines @($header) -OutputFile $OutputFile -ActionsOnly:$comparisonPrinted
    [pscustomobject]@{ Mode = $mode; Comparison = $safe; Actions = @($actions)
        HasErrors = (@($comparison.Errors).Count -gt 0 -or @($actions | Where-Object Status -eq Failed).Count -gt 0) }
    } finally { Write-Progress -Id 1 -Activity $progressActivity -Completed }
}
