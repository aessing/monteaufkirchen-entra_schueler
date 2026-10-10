function Invoke-PersonalTuiAction {
    [CmdletBinding(SupportsShouldProcess)]
    param([object[]] $Accounts, [System.Collections.IDictionary] $Config, [ValidateSet('Block','Unlock','Delete')][string] $Operation)
    $eligible = @(); $skipped = @()
    foreach ($account in $Accounts) {
        $guestUnlock = $Operation -eq 'Unlock' -and (Get-ComparisonPropertyValue $account UserType) -eq 'Guest'
        if ($account.MutationAllowed -and -not $guestUnlock -and $account.AccountEnabled -eq ($Operation -eq 'Block')) { $eligible += $account }
        else {
            $reason = if (-not $account.MutationAllowed) { 'Konto ist geschützt. Prüfe die angezeigten Prüfgründe.' }
                elseif ($guestUnlock) { 'Gäste können hier nicht entsperrt werden.' }
                elseif ($Operation -eq 'Delete' -and $account.AccountEnabled -eq $true) { 'Konto ist aktiv. Zuerst mit B sperren, danach erneut markieren und mit Backspace/Entf löschen.' }
                elseif ($Operation -eq 'Block' -and $account.AccountEnabled -eq $false) { 'Konto ist bereits gesperrt. Mit Backspace/Entf löschen oder als Mitglied mit U entsperren.' }
                elseif ($Operation -eq 'Unlock' -and $account.AccountEnabled -eq $true) { 'Konto ist bereits aktiv.' }
                else { 'Kontostatus ist für diese Aktion nicht eindeutig lesbar.' }
            $skipped += [pscustomobject]@{ DisplayName = $account.DisplayName; UPN = $account.UPN; UserId = $account.UserId; Status = 'Skipped'; Message = $reason }
        }
    }
    if ($eligible.Count -eq 0) { return $skipped }
    if ($Operation -eq 'Delete') { Show-DisabledAccountConfirmation -Accounts $eligible -TenantId $Config.SessionTenantId -SkippedAccounts $skipped }
    elseif ($Operation -eq 'Unlock') { Show-DisabledAccountUnlockConfirmation -Accounts $eligible -TenantId $Config.SessionTenantId -SkippedAccounts $skipped }
    else {
        Clear-Host
        Write-Host "Konten sperren | Tenant $($Config.SessionTenantId)" -ForegroundColor Yellow
        foreach ($account in $eligible) { Write-Host "  $($account.DisplayName) <$($account.UPN)>" }
        Write-Host 'Die Anmeldung wird gesperrt und die Sitzungen werden widerrufen. Konto, Gruppen und Lizenzen bleiben erhalten.'
        foreach ($account in $skipped) { Write-Host "Übersprungen: $($account.DisplayName) <$($account.UPN)> ($($account.Message))" -ForegroundColor Yellow }
        Write-Host 'Ja (Y/J): sperren    Nein (N/Esc): zurück zur Auswahl'
    }
    do { $answer = Read-DisabledAccountTuiKey } while ($answer -notin @('Y','J','N','Escape'))
    if ($answer -notin @('Y','J')) { return }
    $skipped
    if ($Operation -eq 'Delete') {
        Invoke-SelectedDisabledAccountDeletion -Accounts $eligible -TenantId $Config.SessionTenantId -PersonalConfig $Config -WhatIf:$WhatIfPreference -Confirm:$false
    } elseif ($Operation -eq 'Unlock') {
        Invoke-SelectedDisabledAccountUnlock -Accounts $eligible -TenantId $Config.SessionTenantId -PersonalConfig $Config -WhatIf:$WhatIfPreference -Confirm:$false
    } else {
        Invoke-PersonalAccountBlock -Accounts $eligible -Config $Config -WhatIf:$WhatIfPreference -Confirm:$false
    }
}

function Invoke-PersonalAccountBlock {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([object[]] $Accounts, [System.Collections.IDictionary] $Config)
    if (-not $WhatIfPreference) { $null = Connect-DisabledAccountWriteGraph -TenantId $Config.SessionTenantId -Operation Block }
    $activity = if ($Config.Contains('ManagementPopulation') -and $Config.ManagementPopulation -eq 'Teacher') { 'Lehrer- und Ganztagskonten sperren' }
        elseif ($Config.Contains('ManagementPopulation') -and $Config.ManagementPopulation -eq 'Student') { 'Schülerkonten sperren' } else { 'Personalkonten sperren' }
    $index = 0
    $seenIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        foreach ($account in $Accounts) {
            $index++; $status = 'Skipped'; $message = ''; $disabled = $false
            Write-Progress -Id 2 -Activity $activity -Status "Prüfe Konto $index von $($Accounts.Count): $($account.DisplayName)" -PercentComplete ([int](100 * ($index - 1) / $Accounts.Count))
            try {
                if (-not $seenIds.Add([string]$account.UserId)) { throw 'Konto wurde mehrfach ausgewählt.' }
                $user = Assert-PersonalMutationTarget -UserId $account.UserId -ExpectedUpn $account.UPN -Config $Config -Account $account -AllowGuest
                if ($user.AccountEnabled -ne $true) { throw 'Konto ist nicht mehr aktiv.' }
                if ($PSCmdlet.ShouldProcess("$($account.DisplayName) <$($account.UPN)>", 'Sperren und Sitzungen widerrufen')) {
                    $status = 'Failed'
                    $response = Disable-EntraStudent -UserId $account.UserId -Confirm:$false
                    if (-not $response.Verified) { throw 'Sperrung wurde nicht bestätigt.' }
                    $disabled = $true
                    $null = Assert-PersonalMutationTarget -UserId $account.UserId -ExpectedUpn $account.UPN -Config $Config -Account $account -AllowGuest
                    $revoked = Revoke-EntraStudentSession -UserId $account.UserId -Selected -Confirm:$false
                    if ($revoked -isnot [bool]) { $revoked = Get-ComparisonPropertyValue $revoked Value }
                    if ($revoked -ne $true) { throw 'Konto ist gesperrt, Sitzungswiderruf wurde nicht bestätigt.' }
                    $status = 'Disabled'
                } else { $status = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' } }
            } catch { $message = $_.Exception.Message; if ($disabled) { $message = 'Konto bereits gesperrt. ' + $message } }
            [pscustomobject]@{ DisplayName = $account.DisplayName; UPN = $account.UPN; UserId = $account.UserId; Status = $status; Message = $message; AccountDisabled = $disabled }
        }
        Write-Progress -Id 2 -Activity $activity -Status 'Abgeschlossen' -PercentComplete 100
    } finally { Write-Progress -Id 2 -Activity $activity -Completed }
}
