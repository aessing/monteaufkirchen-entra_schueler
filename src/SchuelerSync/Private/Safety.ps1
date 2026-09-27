function Get-ConfiguredStudentRole {
    $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot 'config/SchuelerSync.psd1')
    if (-not $config.StudentRoleGroup.Id -or -not $config.StudentRoleGroup.Name) { throw 'Schülerrolle ist nicht vollständig konfiguriert.' }
    return $config.StudentRoleGroup
}

function Assert-CurrentStudentMembership {
    param([Parameter(Mandatory)][string] $UserId)
    $role = Get-ConfiguredStudentRole
    $direct = @(Get-FreshEntraUserDirectGroup -UserId $UserId)
    if (@($direct | Where-Object Id -eq $role.Id).Count -ne 1) {
        throw "Konto '$UserId' gehört nicht mehr zur konfigurierten Schülerrolle. Erneuten Vergleich starten."
    }
}

function Get-WritablePersonnelRole {
    param([Parameter(Mandatory)][string] $GroupId)
    $group = Get-MgGroup -GroupId $GroupId -Property @('id','displayName','assignedLicenses','groupTypes','membershipRule') -ErrorAction Stop
    if ([string]$group.Id -ine $GroupId -or $null -eq $group.PSObject.Properties['AssignedLicenses'] -or $null -eq $group.AssignedLicenses) {
        throw "Aktuelle Lizenzdaten der Rolle '$GroupId' fehlen."
    }
    if (@($group.AssignedLicenses | Where-Object { $null -ne $_ }).Count -gt 0) { throw "Rolle '$($group.DisplayName)' vergibt Lizenzen." }
    if ($null -eq $group.PSObject.Properties['GroupTypes'] -or (Get-EntraGroupIsDynamic -Group $group)) {
        throw "Rolle '$GroupId' ist dynamisch oder ihre Änderbarkeit ist unbekannt."
    }
    return $group
}

function Assert-ComparisonReportPath {
    param([Parameter(Mandatory)][string] $Path)
    if ([IO.Path]::GetExtension($Path) -ine '.txt') { throw 'Berichte benötigen eine neue .txt-Datei als Ziel.' }
    $fullPath = [IO.Path]::GetFullPath($Path)
    # Do not follow links into another file or repository when writing confidential reports.
    $part = $fullPath
    while ($part) {
        $item = Get-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        if ($null -ne $item -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Berichtspfad darf keine symbolischen Links enthalten: '$part'."
        }
        $part = Split-Path -Parent $part
    }
    if (Test-Path -LiteralPath $fullPath) { throw "Berichtsziel existiert bereits und wird nicht überschrieben: '$fullPath'." }
    if ($null -eq (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Git wird zur Prüfung des vertraulichen Berichtspfads benötigt.' }
    Assert-WorkbookArtifactGitSafety -Path $fullPath -Label 'vertrauliche Berichtsdatei'
}

function Write-ComparisonReportFile {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string] $Path, [Parameter(Mandatory)][AllowEmptyString()][string] $Text)
    Assert-ComparisonReportPath -Path $Path
    if (-not $PSCmdlet.ShouldProcess($Path, 'Neuen vertraulichen Bericht schreiben')) { return }
    $parent = Split-Path -Parent ([IO.Path]::GetFullPath($Path))
    $null = [IO.Directory]::CreateDirectory($parent)
    Assert-ComparisonReportPath -Path $Path
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($Text + [Environment]::NewLine)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
}
