[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $PassThru,
    [switch] $List
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'src/SchuelerSync/SchuelerSync.psd1') -Force

$result = Invoke-DisabledAccountCheck -ErrorAction Stop
if ($null -eq $result) { throw 'Die Kontenprüfung wurde ohne Ergebnis beendet.' }
if (-not $List -and -not $PassThru) {
    if ([Console]::IsInputRedirected) {
        throw 'Die Tastaturauswahl benötigt ein interaktives Terminal. Für die lesende Ausgabe -List verwenden.'
    }
    $actions = @(Invoke-DisabledAccountTui -Accounts @($result.Accounts) -TenantId $result.TenantId -WhatIf:$WhatIfPreference)
    if ($actions.Count -gt 0) {
        Write-Host ''
        Write-Host "Aktionsergebnisse ($($actions.Count))" -ForegroundColor Cyan
    }
    foreach ($action in $actions) {
        $detail = if ($action.Message) { ": $($action.Message)" } else { '' }
        Write-Information "$($action.Status): $($action.DisplayName) <$($action.UPN)>$detail" -InformationAction Continue
    }
    return
}

Write-Host ''
$withoutPersonalMailbox = @($result.Accounts | Where-Object { $_.RecipientType -notin @('UserMailbox', 'RemoteUserMailbox') }).Count
Write-Host "Gesperrte Entra-Konten ($($result.Accounts.Count)), ohne persönliches Postfach: $withoutPersonalMailbox, Tenant $($result.TenantId)" -ForegroundColor Cyan
Write-Host ''
Write-Information 'Sperrzeitpunkt und 90-Tage-Frist: unbekannt. Diese Prüfung gibt keine Löschung frei.' -InformationAction Continue
Write-Host ''
foreach ($account in $result.Accounts) {
    $ownedObjectLines = @(
        foreach ($ownedObject in $account.OwnedObjects) {
            $name = if ([string]::IsNullOrWhiteSpace($ownedObject.DisplayName)) { 'Name nicht lesbar' } else { $ownedObject.DisplayName }
            "$($ownedObject.Type): $name (ID: $($ownedObject.Id))"
        }
    )
    [pscustomobject]@{
        DisplayName = $account.DisplayName
        UPN = $account.UPN
        UserId = $account.UserId
        'Exchange-Empfänger' = if ($account.RecipientType) { $account.RecipientType } elseif ($account.RecipientStatus -eq 'None') { 'Keiner' } elseif ($account.RecipientStatus -eq 'Ambiguous') { 'Nicht eindeutig' } else { 'Typ nicht lesbar' }
        'SEC-A Gruppen' = if ($account.SecAGroups.Count) { $account.SecAGroups -join ', ' } else { 'Keine ermittelt' }
        'Last Login' = if ($null -ne $account.LastLogin) { $account.LastLogin.ToUniversalTime().ToString('yyyy-MM-dd HH:mm') + ' UTC' } else { $account.LastLoginStatus }
        'Besitzobjekte' = if ($ownedObjectLines.Count) { $ownedObjectLines -join "`n" } else { 'Keine ermittelt' }
        'Schutz/Prüfung' = if ($account.ProtectionReasons.Count) { $account.ProtectionReasons -join ', ' } else { $account.ReviewStatus }
    } | Format-List | Out-String | Write-Information -InformationAction Continue
}
if ($result.Warnings.Count -gt 0) {
    Write-Host ''
    Write-Host "Warnungen ($($result.Warnings.Count))" -ForegroundColor Yellow
    foreach ($warning in $result.Warnings) { Write-Warning $warning }
}
if ($PassThru) { $result }
