function Protect-StudentMessage {
    param([AllowNull()][string] $Message, [AllowNull()][object[]] $Secrets)
    if ($null -eq $Message) { return '' }
    foreach ($secret in @($Secrets)) {
        if (-not [string]::IsNullOrEmpty([string]$secret)) { $Message = $Message.Replace([string]$secret, '[REDACTED]') }
    }
    return $Message
}

function Get-StudentRecoveryCommand {
    param([string] $File, [string] $UserPrincipalName, [switch] $ExchangeOnly)
    if ($ExchangeOnly) {
        return ".\Sync-SchuelerEntra.ps1 -ConfigureExchangeOnlineOnly -Mail '$($UserPrincipalName.Replace("'", "''"))'"
    }
    return ".\Sync-SchuelerEntra.ps1 -File '$($File.Replace("'", "''"))'"
}

function New-StudentActionResult {
    param(
        [string] $UserId, [string] $UserPrincipalName,
        [Parameter(Mandatory)][string] $Phase,
        [Parameter(Mandatory)][string] $Status,
        [string] $Message, [string] $RecoveryCommand,
        [AllowNull()][object[]] $Secrets
    )
    [pscustomobject]@{
        UserId = $UserId; UserPrincipalName = $UserPrincipalName; Phase = $Phase; Status = $Status
        Message = Protect-StudentMessage -Message $Message -Secrets $Secrets
        RecoveryCommand = $RecoveryCommand
    }
}

function ConvertTo-SafeStudentComparison {
    param([Parameter(Mandatory)][object] $Comparison, [string] $File, [object[]] $Secrets)
    $safe = [ordered]@{}
    foreach ($category in 'NewStudents', 'Departures', 'ChangedStudents', 'ExistingStudents') {
        $safe[$category] = @($Comparison.$category | ForEach-Object {
            [pscustomobject]@{
                RowNumber = Get-ComparisonPropertyValue $_.Student RowNumber
                NameMitRufname = Get-ComparisonPropertyValue $_.Student NameMitRufname
                UserId = Get-ComparisonPropertyValue $_.User Id
                DisplayName = if ($null -ne $_.DesiredState) { $_.DesiredState.DisplayName } else { $_.User.DisplayName }
                UserPrincipalName = if ($null -ne $_.DesiredState) { $_.DesiredState.UserPrincipalName } else { $_.User.UserPrincipalName }
                ClassName = Get-ComparisonPropertyValue $_.Student ClassName
                AccountEnabled = Get-ComparisonPropertyValue $_.User AccountEnabled
                Differences = @($_.Differences | ForEach-Object {
                    [pscustomobject]@{ Area = $_.Area; Field = $_.Field; Current = $_.Current; Desired = $_.Desired; Action = $_.Action }
                })
            }
        })
    }
    foreach ($category in 'Warnings', 'Errors') {
        $safe[$category] = @($Comparison.$category | ForEach-Object {
            $upn = [string](Get-ComparisonPropertyValue $_.User UserPrincipalName)
            [pscustomobject]@{
                Code = "$($_.Area).$($_.Field)"
                Student = [string](Get-ComparisonPropertyValue $_.Student NameMitRufname)
                Message = Protect-StudentMessage -Message $_.Message -Secrets $Secrets
                RecoveryCommand = if ($_.Area -eq 'Exchange' -and $_.Field -eq 'Mailbox' -and $upn) {
                    Get-StudentRecoveryCommand -UserPrincipalName $upn -ExchangeOnly
                } else { Get-StudentRecoveryCommand -File $File }
            }
        })
    }
    [pscustomobject]$safe
}

function Write-StudentComparisonReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object] $Comparison, [AllowEmptyCollection()][object[]] $Actions = @(), [switch] $ActionsOnly)
    $tables = [ordered]@{
        'Neuzugänge' = @($Comparison.NewStudents | Select-Object -Property @('RowNumber', 'NameMitRufname', 'DisplayName', 'ClassName', 'UserPrincipalName'))
        'Abgänge' = @($Comparison.Departures | Select-Object -Property @('UserId', 'DisplayName', 'UserPrincipalName', 'AccountEnabled'))
        'Änderungen' = @($Comparison.ChangedStudents | ForEach-Object {
            $entry = $_
            foreach ($difference in $entry.Differences) {
                [pscustomobject]@{ NameMitRufname = $entry.NameMitRufname; Area = $difference.Area; Field = $difference.Field; Current = $difference.Current; Desired = $difference.Desired; Action = $difference.Action }
            }
        })
        'Bestehende' = @($Comparison.ExistingStudents | Select-Object -Property @('NameMitRufname', 'UserId', 'DisplayName', 'UserPrincipalName', 'ClassName'))
        'Warnungen und Fehler' = @($Comparison.Warnings) + @($Comparison.Errors)
        'Aktionsergebnisse' = @($Actions)
    }
    foreach ($title in $tables.Keys) {
        if ($ActionsOnly -and $title -ne 'Aktionsergebnisse') { continue }
        Write-Information "$title ($(@($tables[$title]).Count))" -InformationAction Continue
        if (@($tables[$title]).Count -gt 0) {
            Write-Information ($tables[$title] | Format-Table -AutoSize -Wrap | Out-String -Width 220) -InformationAction Continue
        }
    }
}

function ConvertTo-ExchangeActionResults {
    param([Parameter(Mandatory)][object] $Batch, [switch] $WhatIfMode)
    foreach ($ready in $Batch.Ready) {
        $status = [string]$ready.Configuration.Status
        if ($WhatIfMode -and $status -eq 'Planned') { $status = 'WhatIf' }
        New-StudentActionResult -UserPrincipalName $ready.UserPrincipalName -Phase Exchange -Status $status
    }
    foreach ($missing in $Batch.Missing) {
        New-StudentActionResult -UserPrincipalName $missing -Phase Exchange -Status $(if ($WhatIfMode) { 'WhatIf' } else { 'Pending' }) `
            -Message 'EXO-Konfiguration ausstehend.' -RecoveryCommand (Get-StudentRecoveryCommand -UserPrincipalName $missing -ExchangeOnly)
    }
    foreach ($failure in $Batch.Failed) {
        New-StudentActionResult -UserPrincipalName $failure.UserPrincipalName -Phase Exchange -Status Failed `
            -Message $failure.Error -RecoveryCommand (Get-StudentRecoveryCommand -UserPrincipalName $failure.UserPrincipalName -ExchangeOnly)
    }
}
