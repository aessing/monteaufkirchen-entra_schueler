$script:MandatoryStudentHeaders = @(
    'Name mit Rufname',
    'Vorname',
    'Nachname',
    'Klassen',
    'Klassenlehrer'
)

$script:ManagedStudentHeaders = @('Passwort', 'EntraObjectId', 'UPN')

function ConvertTo-WorkbookCellText {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value) { return '' }
    return ([string] $Value).Trim()
}

function Get-WorksheetHeaderMap {
    param([Parameter(Mandatory)][object] $Worksheet)

    $headers = @{}
    if ($null -eq $Worksheet.Dimension) {
        return $headers
    }

    for ($column = 1; $column -le $Worksheet.Dimension.End.Column; $column++) {
        $header = ConvertTo-WorkbookCellText $Worksheet.Cells[1, $column].Value
        if (-not [string]::IsNullOrWhiteSpace($header) -and -not $headers.ContainsKey($header)) {
            $headers[$header] = $column
        }
    }
    return $headers
}

function Resolve-StudentWorkbookPath {
    param([Parameter(Mandatory)][string] $Path)

    if ([IO.Path]::GetExtension($Path) -ine '.xlsx') {
        throw "Die Schülerdatei muss die Erweiterung .xlsx haben: '$Path'."
    }
    if (-not [IO.File]::Exists($Path)) {
        throw "Die Schülerdatei existiert nicht oder ist keine Datei: '$Path'."
    }

    return [IO.Path]::GetFullPath($Path)
}

function Read-StudentWorkbook {
    param([Parameter(Mandatory)][string] $Path)

    $resolvedPath = Resolve-StudentWorkbookPath -Path $Path
    $package = $null
    try {
        $package = Open-ExcelPackage -Path $resolvedPath
        $candidates = @()
        foreach ($worksheet in $package.Workbook.Worksheets) {
            $headers = Get-WorksheetHeaderMap -Worksheet $worksheet
            $hasRequiredHeaders = $true
            foreach ($requiredHeader in $script:MandatoryStudentHeaders) {
                if (-not $headers.ContainsKey($requiredHeader)) {
                    $hasRequiredHeaders = $false
                    break
                }
            }
            if ($hasRequiredHeaders) {
                $candidates += [pscustomobject]@{
                    Worksheet = $worksheet
                    Headers = $headers
                }
            }
        }

        if ($candidates.Count -eq 0) {
            throw 'Keine Arbeitsmappe enthält alle fünf Pflichtspalten für Schülerdaten.'
        }
        if ($candidates.Count -gt 1) {
            throw 'Mehrere Arbeitsblätter enthalten alle fünf Pflichtspalten für Schülerdaten.'
        }

        $candidate = $candidates[0]
        $worksheet = $candidate.Worksheet
        $headers = $candidate.Headers
        $students = @()
        $lastRow = if ($null -eq $worksheet.Dimension) { 1 } else { $worksheet.Dimension.End.Row }
        for ($row = 2; $row -le $lastRow; $row++) {
            $values = @{}
            foreach ($requiredHeader in $script:MandatoryStudentHeaders) {
                $values[$requiredHeader] = ConvertTo-WorkbookCellText $worksheet.Cells[$row, $headers[$requiredHeader]].Value
            }
            $filledCount = @($values.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
            if ($filledCount -eq 0) {
                continue
            }
            if ($filledCount -ne $script:MandatoryStudentHeaders.Count) {
                throw "Unvollständige Schülerdaten in Arbeitsblatt '$($worksheet.Name)', Zeile $row."
            }

            $students += [pscustomobject]@{
                RowNumber = $row
                NameMitRufname = $values['Name mit Rufname']
                GivenName = $values['Vorname']
                Surname = $values['Nachname']
                ClassName = $values['Klassen']
                Teacher = $values['Klassenlehrer']
                Password = if ($headers.ContainsKey('Passwort')) {
                    ConvertTo-WorkbookCellText $worksheet.Cells[$row, $headers['Passwort']].Value
                } else { '' }
                EntraObjectId = if ($headers.ContainsKey('EntraObjectId')) {
                    ConvertTo-WorkbookCellText $worksheet.Cells[$row, $headers['EntraObjectId']].Value
                } else { '' }
                StoredUpn = if ($headers.ContainsKey('UPN')) {
                    ConvertTo-WorkbookCellText $worksheet.Cells[$row, $headers['UPN']].Value
                } else { '' }
            }
        }

        return [pscustomobject]@{
            Path = $resolvedPath
            WorksheetName = $worksheet.Name
            Headers = [string[]] @($headers.Keys)
            Students = [pscustomobject[]] $students
        }
    } finally {
        if ($null -ne $package) {
            $package.Dispose()
        }
    }
}

function Get-WorkbookGitRoot {
    param([Parameter(Mandatory)][string] $Path)

    if ($null -eq (Get-Command git -ErrorAction SilentlyContinue)) {
        return $null
    }

    $directory = Split-Path -Parent $Path
    $gitRoot = & git -C $directory rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -ne 0) {
        return $null
    }
    return [IO.Path]::GetFullPath(([string] $gitRoot).Trim())
}

function Assert-WorkbookSafeForPasswordWrite {
    param(
        [Parameter(Mandatory)][string] $Path,
        [switch] $SkipGitSafetyCheck
    )

    $resolvedPath = Resolve-StudentWorkbookPath -Path $Path
    $lockHandle = $null
    try {
        $lockHandle = [IO.File]::Open(
            $resolvedPath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::None
        )
    } catch {
        throw "Die Schülerdatei ist gesperrt und kann nicht sicher geschrieben werden: '$resolvedPath'."
    } finally {
        if ($null -ne $lockHandle) {
            $lockHandle.Dispose()
        }
    }

    if ($SkipGitSafetyCheck) {
        return
    }

    $gitRoot = Get-WorkbookGitRoot -Path $resolvedPath
    if ($null -eq $gitRoot) {
        return
    }

    $relativePath = [IO.Path]::GetRelativePath($gitRoot, $resolvedPath)
    if ($relativePath.StartsWith('..') -or [IO.Path]::IsPathRooted($relativePath)) {
        return
    }

    & git -C $gitRoot check-ignore -q -- $relativePath
    if ($LASTEXITCODE -ne 0) {
        throw "Die Schülerdatei innerhalb des Git-Worktrees muss ignoriert sein: '$relativePath'."
    }
    & git -C $gitRoot ls-files --error-unmatch -- $relativePath 2>$null
    if ($LASTEXITCODE -eq 0) {
        throw "Die Schülerdatei innerhalb des Git-Worktrees darf nicht versioniert sein: '$relativePath'."
    }
}

function Get-RequiredWorkbookUpdateValue {
    param(
        [Parameter(Mandatory)][object] $Update,
        [Parameter(Mandatory)][string] $PropertyName
    )

    $property = $Update.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string] $property.Value)) {
        throw "Eine Excel-Rückschreibung benötigt einen nicht leeren Wert für '$PropertyName'."
    }
    return ConvertTo-WorkbookCellText $property.Value
}

function Write-StudentWorkbookUpdates {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)][object[]] $Updates,
        [switch] $SkipGitSafetyCheck
    )

    if (-not $PSCmdlet.ShouldProcess($Path, 'Back up and atomically persist student workbook updates')) { return }
    $context = Read-StudentWorkbook -Path $Path
    $updatesByRow = @{}
    foreach ($update in $Updates) {
        $rowProperty = $update.PSObject.Properties['RowNumber']
        if ($null -eq $rowProperty -or $rowProperty.Value -isnot [int] -or $rowProperty.Value -lt 2) {
            throw 'Jede Excel-Rückschreibung benötigt eine gültige RowNumber ab 2.'
        }
        $rowNumber = [int] $rowProperty.Value
        if ($updatesByRow.ContainsKey($rowNumber)) {
            throw "Mehrere Excel-Rückschreibungen für Zeile $rowNumber sind nicht zulässig."
        }
        if ($null -eq ($context.Students | Where-Object RowNumber -eq $rowNumber)) {
            throw "Excel-Rückschreibung verweist auf keine gültige Schülerzeile: $rowNumber."
        }
        $updatesByRow[$rowNumber] = [pscustomobject]@{
            RowNumber = $rowNumber
            Password = if ($null -eq $update.PSObject.Properties['Password']) { '' } else {
                ConvertTo-WorkbookCellText $update.PSObject.Properties['Password'].Value
            }
            EntraObjectId = Get-RequiredWorkbookUpdateValue -Update $update -PropertyName 'EntraObjectId'
            UPN = Get-RequiredWorkbookUpdateValue -Update $update -PropertyName 'UPN'
        }
    }

    Assert-WorkbookSafeForPasswordWrite -Path $context.Path -SkipGitSafetyCheck:$SkipGitSafetyCheck

    $directory = Split-Path -Parent $context.Path
    $baseName = [IO.Path]::GetFileNameWithoutExtension($context.Path)
    $extension = [IO.Path]::GetExtension($context.Path)
    $backupTimestamp = Get-Date
    $backupPath = Join-Path $directory ('{0}.backup-{1}{2}' -f $baseName, $backupTimestamp.ToString('yyyyMMdd-HHmmss'), $extension)
    while (Test-Path -LiteralPath $backupPath) {
        $backupTimestamp = $backupTimestamp.AddSeconds(1)
        $backupPath = Join-Path $directory ('{0}.backup-{1}{2}' -f $baseName, $backupTimestamp.ToString('yyyyMMdd-HHmmss'), $extension)
    }
    $temporaryPath = Join-Path $directory ('.{0}.{1}.tmp{2}' -f $baseName, [guid]::NewGuid().Guid, $extension)
    $package = $null
    try {
        Copy-Item -LiteralPath $context.Path -Destination $backupPath -ErrorAction Stop
        Copy-Item -LiteralPath $context.Path -Destination $temporaryPath -ErrorAction Stop

        $package = Open-ExcelPackage -Path $temporaryPath
        $worksheet = $package.Workbook.Worksheets[$context.WorksheetName]
        $headers = Get-WorksheetHeaderMap -Worksheet $worksheet
        $rightmostColumn = if ($null -eq $worksheet.Dimension) { 0 } else { $worksheet.Dimension.End.Column }
        foreach ($managedHeader in $script:ManagedStudentHeaders) {
            if (-not $headers.ContainsKey($managedHeader)) {
                $rightmostColumn++
                $worksheet.Cells[1, $rightmostColumn].Value = $managedHeader
                $headers[$managedHeader] = $rightmostColumn
            }
        }

        foreach ($update in $updatesByRow.Values) {
            if (-not [string]::IsNullOrWhiteSpace($update.Password)) {
                $worksheet.Cells[$update.RowNumber, $headers['Passwort']].Value = $update.Password
            }
            $worksheet.Cells[$update.RowNumber, $headers['EntraObjectId']].Value = $update.EntraObjectId
            $worksheet.Cells[$update.RowNumber, $headers['UPN']].Value = $update.UPN
        }
        $package.Save()
        $package.Dispose()
        $package = $null

        $verificationPackage = $null
        try {
            $verificationPackage = Open-ExcelPackage -Path $temporaryPath
            $verificationWorksheet = $verificationPackage.Workbook.Worksheets[$context.WorksheetName]
            $verificationHeaders = Get-WorksheetHeaderMap -Worksheet $verificationWorksheet
            foreach ($managedHeader in $script:ManagedStudentHeaders) {
                if (-not $verificationHeaders.ContainsKey($managedHeader)) {
                    throw "Temporäre Schülerdatei enthält die verwaltete Spalte '$managedHeader' nicht."
                }
            }
            foreach ($update in $updatesByRow.Values) {
                $storedObjectId = ConvertTo-WorkbookCellText $verificationWorksheet.Cells[$update.RowNumber, $verificationHeaders['EntraObjectId']].Value
                $storedUpn = ConvertTo-WorkbookCellText $verificationWorksheet.Cells[$update.RowNumber, $verificationHeaders['UPN']].Value
                if ($storedObjectId -ne $update.EntraObjectId -or $storedUpn -ne $update.UPN) {
                    throw "Temporäre Schülerdatei konnte die Entra-Daten für Zeile $($update.RowNumber) nicht verifizieren."
                }
                if (-not [string]::IsNullOrWhiteSpace($update.Password)) {
                    $storedPassword = ConvertTo-WorkbookCellText $verificationWorksheet.Cells[$update.RowNumber, $verificationHeaders['Passwort']].Value
                    if ($storedPassword -ne $update.Password) {
                        throw "Temporäre Schülerdatei konnte das Passwort für Zeile $($update.RowNumber) nicht verifizieren."
                    }
                }
            }
        } finally {
            if ($null -ne $verificationPackage) {
                $verificationPackage.Dispose()
            }
        }

        [IO.File]::Move($temporaryPath, $context.Path, $true)
        return [pscustomobject]@{
            Path = $context.Path
            BackupPath = $backupPath
        }
    } catch {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
        throw
    } finally {
        if ($null -ne $package) {
            $package.Dispose()
        }
    }
}
