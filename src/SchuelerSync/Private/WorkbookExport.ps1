function Assert-WorkbookExportPath {
    param([string] $Path)
    if ([IO.Path]::GetExtension($Path) -ine '.xlsx') { throw 'Der Export benötigt eine neue .xlsx-Datei.' }
    $part = [IO.Path]::GetFullPath($Path)
    while ($part) {
        $item = Get-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        if ($null -ne $item -and ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Exportpfad darf keine symbolischen Links enthalten.' }
        $part = Split-Path -Parent $part
    }
    if (Test-Path -LiteralPath $Path) { throw "Exportziel existiert bereits: '$Path'." }
    if ($null -eq (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Git wird zur Prüfung des Exportpfads benötigt.' }
    Assert-WorkbookArtifactGitSafety -Path $Path
}

function Export-AccountWorkbook {
    [CmdletBinding(SupportsShouldProcess)]
    param([object] $Inventory, [string] $Path, [ValidateSet('Student','Teacher','Personnel')][string] $Kind = 'Personnel')
    $Path = [IO.Path]::GetFullPath($Path)
    Assert-WorkbookExportPath -Path $Path
    if (-not $PSCmdlet.ShouldProcess($Path, 'Neue Bestands-Arbeitsmappe erstellen')) { return }
    $parent = Split-Path -Parent $Path
    $null = [IO.Directory]::CreateDirectory($parent)
    $temporary = Join-Path $parent ('.personal-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    Assert-WorkbookExportPath -Path $temporary
    $headers = $(if ($Kind -eq 'Student') { @($script:MandatoryStudentHeaders) } else { @($script:MandatoryTeacherHeaders) }) + @($script:ManagedStudentHeaders) + @('Last Login Timestamp')
    $package = $null
    try {
        Import-Module ImportExcel -ErrorAction Stop 3>$null
        $package = [OfficeOpenXml.ExcelPackage]::new()
        $sheet = $package.Workbook.Worksheets.Add('Tabelle1')
        for ($column = 1; $column -le $headers.Count; $column++) { $sheet.Cells[1,$column].Value = $headers[$column - 1] }
        $rows = @(
            foreach ($account in $Inventory.Accounts) {
                $name = if ($account.GivenName -and $account.Surname) { "$($account.Surname), $($account.GivenName)" } else { $account.DisplayName }
                $lastLogin = Get-ComparisonPropertyValue -InputObject $account -Name LastLogin
                $timestamp = if ($null -eq $lastLogin) { '' } else {
                    ([datetimeoffset]$lastLogin).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'",[Globalization.CultureInfo]::InvariantCulture)
                }
                if ($Kind -eq 'Student') {
                    ,@($name, $account.GivenName, $account.Surname, $account.ClassName, $account.Teacher, '', $account.UserId, $account.UPN, $account.Mail, $timestamp)
                } else {
                    $job = if ($Kind -eq 'Personnel') { ConvertTo-PersonalJobDisplayName -Job $account.Job } else { $account.Job }
                    ,@($name, $account.GivenName, $account.Surname, $job, '', $account.UserId, $account.UPN, $account.Mail, $timestamp)
                }
            }
        )
        for ($row = 0; $row -lt $rows.Count; $row++) {
            for ($column = 0; $column -lt $headers.Count; $column++) {
                $sheet.Cells[($row + 2),($column + 1)].Value = [string]$rows[$row][$column]
                $sheet.Cells[($row + 2),($column + 1)].Style.Numberformat.Format = '@'
            }
        }
        $table = $sheet.Tables.Add($sheet.Cells[1,1,[math]::Max(1,$rows.Count + 1),$headers.Count], 'Tabelle1')
        $table.TableStyle = 'Medium14'; $table.ShowHeader = $true; $table.ShowRowStripes = $true
        $widths = if ($Kind -eq 'Student') { @(21.664,11.164,15,32,32,11,38,38,38,26) } else { @(21.664,11.164,15,32,11,38,38,38,26) }
        for ($column = 1; $column -le $headers.Count; $column++) { $sheet.Column($column).Width = $widths[$column - 1] }
        $package.SaveAs([IO.FileInfo]::new($temporary)); $package.Dispose(); $package = $null
        $package = Open-StudentWorkbookPackage -Path $temporary
        $sheet = $package.Workbook.Worksheets['Tabelle1']
        # EPPlus exposes the serialized style through StyleName when reopening a table.
        if ($package.Workbook.Worksheets.Count -ne 1 -or $sheet.Tables.Count -ne 1 -or
            $sheet.Tables[0].Name -cne 'Tabelle1' -or $sheet.Tables[0].StyleName -cne 'TableStyleMedium14' -or
            $sheet.Dimension.End.Row -ne [math]::Max(1,$rows.Count + 1)) { throw 'Exportstruktur konnte nicht verifiziert werden.' }
        for ($column = 0; $column -lt $headers.Count; $column++) {
            if ([string]$sheet.Cells[1,($column + 1)].Value -cne $headers[$column]) { throw 'Exportüberschriften sind nicht korrekt.' }
        }
        for ($row = 0; $row -lt $rows.Count; $row++) {
            for ($column = 0; $column -lt $headers.Count; $column++) {
                $cell = $sheet.Cells[($row + 2),($column + 1)]
                if ([string]$cell.Value -cne [string]$rows[$row][$column] -or $cell.Formula) { throw 'Exportwerte konnten nicht als unveränderte Texte verifiziert werden.' }
            }
        }
        $package.Dispose(); $package = $null
        Assert-WorkbookExportPath -Path $Path
        Assert-WorkbookArtifactGitSafety -Path $temporary
        [IO.File]::Move($temporary,$Path)
        [pscustomobject]@{ Path = $Path; Count = $rows.Count; Status = 'Succeeded' }
    } finally {
        if ($null -ne $package) { $package.Dispose() }
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}

function Get-RoleWorkbookInventory {
    [CmdletBinding()]
    param([ValidateSet('Student','Teacher')][string] $Kind)
    $configFile = if ($Kind -eq 'Student') { 'SchuelerSync.psd1' } else { 'LehrerSync.psd1' }
    $config = Import-PowerShellDataFile (Join-Path $script:SchuelerSyncRepositoryRoot "config/$configFile")
    $graph = Connect-DisabledAccountGraph -RequiredScopes @('User.Read.All','Group.Read.All','AuditLog.Read.All')
    $tenantId = [string](Get-ComparisonPropertyValue $graph TenantId)
    if (-not $tenantId) { throw 'Graph-Kontext enthält keine Tenant-ID.' }
    if ($config.ExpectedTenantId -and $tenantId -ine $config.ExpectedTenantId) { throw 'Graph-Tenant stimmt nicht mit der Konfiguration überein.' }
    $groups = @(Get-MgGroup -All -Property @('id','displayName') -ErrorAction Stop)
    $roles = if ($Kind -eq 'Student') { @($config.StudentRoleGroup) } else { @($config.Roles.School,$config.Roles.Ganztag) }
    foreach ($role in $roles) {
        $roleMatches = @($groups | Where-Object DisplayName -eq $role.Name)
        if ($roleMatches.Count -ne 1 -or $roleMatches[0].Id -ine $role.Id) { throw "Exportrolle '$($role.Name)' fehlt oder ist mehrdeutig." }
    }
    $warnings = [Collections.Generic.List[string]]::new()
    $fields = @('id','displayName','givenName','surname','userPrincipalName','mail',
        'jobTitle','accountEnabled','userType','onPremisesSyncEnabled')
    $lastLoginAvailable = $true
    try { $users = @(Get-MgUser -All -Property ($fields + 'signInActivity') -ErrorAction Stop) }
    catch {
        $lastLoginAvailable = $false
        $warnings.Add("Letzte Anmeldung nicht abrufbar: $($_.Exception.Message)")
        $users = @(Get-MgUser -All -Property $fields -ErrorAction Stop)
    }
    $usersById = New-CaseInsensitiveHashtable
    foreach ($user in $users) {
        $id = [string](Get-ComparisonPropertyValue $user Id)
        if (-not $id -or $usersById.ContainsKey($id)) { throw 'Benutzerbestand enthält fehlende oder doppelte Objekt-IDs.' }
        $usersById[$id] = $user
    }
    $rolesByUser = New-CaseInsensitiveHashtable
    foreach ($role in $roles) {
        foreach ($member in @(Get-MgGroupMember -GroupId $role.Id -All -ErrorAction Stop)) {
            $type = [string](Get-DisabledAccountMemberValue -Member $member -Name '@odata.type')
            if ($type -and $type -ne '#microsoft.graph.user') { continue }
            $id = [string](Get-DisabledAccountMemberValue -Member $member -Name id)
            if (-not $id -or -not $usersById.ContainsKey($id)) { throw "Mitglied der Exportrolle '$($role.Name)' fehlt im Benutzerbestand." }
            if (-not $rolesByUser.ContainsKey($id)) { $rolesByUser[$id] = @() }
            $rolesByUser[$id] = @($rolesByUser[$id]) + [string]$role.Id
        }
    }
    $null = Connect-SchuelerExchangeOnline -TenantId $tenantId
    $recipients = Get-AccountRecipientIndex -Recipients @(Get-EXORecipient -ResultSize Unlimited -Properties @('ExternalDirectoryObjectId','RecipientTypeDetails') -ErrorAction Stop)
    $snapshot = [pscustomobject]@{ ManagerByUserId = (New-CaseInsensitiveHashtable) }
    $accounts = [Collections.Generic.List[object]]::new()
    $excludedCount = 0
    foreach ($id in $rolesByUser.Keys) {
        if ($recipients.SharedIds.ContainsKey($id)) { $excludedCount++; continue }
        $user = $usersById[$id]
        $login = Get-EntraLastLogin -User $user -LastLoginAvailable:$lastLoginAvailable
        $account = [pscustomobject]@{
            UserId = $id
            DisplayName = [string](Get-ComparisonPropertyValue $user DisplayName)
            GivenName = [string](Get-ComparisonPropertyValue $user GivenName)
            Surname = [string](Get-ComparisonPropertyValue $user Surname)
            UPN = [string](Get-ComparisonPropertyValue $user UserPrincipalName)
            Mail = [string](Get-ComparisonPropertyValue $user Mail)
            LastLogin = $login.LastLogin; LastLoginStatus = $login.LastLoginStatus
            Job = ''; ClassName = ''; Teacher = ''
        }
        $reasons = [Collections.Generic.List[string]]::new()
        if (-not $account.GivenName -or -not $account.Surname -or -not $account.UPN) { $reasons.Add('Name oder UPN unvollständig') }
        if ((Get-ComparisonPropertyValue $user UserType) -ne 'Member' -or
            (Get-ComparisonPropertyValue $user OnPremisesSyncEnabled) -eq $true) { $reasons.Add('Gastkonto oder synchronisiertes Konto') }
        if ((Get-ComparisonPropertyValue $user AccountEnabled) -eq $false) { $reasons.Add('Konto ist gesperrt, bisheriger Schüler-/Lehrer-Sync kann es reaktivieren') }
        if ($recipients.DuplicateIds.ContainsKey($id)) { $reasons.Add('Exchange-Zuordnung mehrdeutig') }
        if ($Kind -eq 'Teacher') {
            try {
                $jobProfile = Get-LehrerProfile -Job ([string](Get-ComparisonPropertyValue $user JobTitle)) -Config $config
                if (@($rolesByUser[$id] | Select-Object -Unique).Count -ne 1 -or $jobProfile.Role.Id -notin $rolesByUser[$id]) {
                    throw 'Job und direkte Lehrer-/Ganztagrolle widersprechen sich.'
                }
                $account.Job = $jobProfile.Job
            } catch { $reasons.Add('Job nicht eindeutig bestätigt, Jobzelle bleibt leer') }
        } else {
            try {
                $direct = @(Get-FreshEntraUserDirectGroup -UserId $id)
                $classes = @($direct | Where-Object { $_.DisplayName.StartsWith($config.ClassGroupPrefix,[StringComparison]::OrdinalIgnoreCase) })
                if ($classes.Count -ne 1) { throw 'Keine eindeutige direkte Klassengruppe.' }
                $className = $classes[0].DisplayName.Substring($config.ClassGroupPrefix.Length)
                if (-not $className) { throw 'Klassengruppe enthält keinen Klassennamen.' }
                $account.ClassName = $className
            } catch { $reasons.Add('Klasse nicht eindeutig lesbar, Klassenzelle bleibt leer') }
            try {
                $managerId = Get-UserManagerId -Snapshot $snapshot -UserId $id
                if (-not $managerId -or -not $usersById.ContainsKey($managerId)) { throw 'Klassenlehrer fehlt im Benutzerbestand.' }
                $account.Teacher = [string](Get-ComparisonPropertyValue $usersById[$managerId] UserPrincipalName)
                if (-not $account.Teacher) { throw 'Klassenlehrer hat keinen UPN.' }
            } catch { $reasons.Add('Klassenlehrer nicht lesbar, Klassenlehrerzelle bleibt leer') }
        }
        if ($reasons.Count -gt 0) {
            $warnings.Add("$($account.UPN) ($id): Exportzeile vor einem schreibenden Import prüfen. $($reasons -join '. ').")
        }
        $accounts.Add($account)
    }
    [pscustomobject]@{ Accounts = @($accounts | Sort-Object DisplayName,UPN); Warnings = @($warnings); TenantId = $tenantId; ExcludedCount = $excludedCount }
}

function Invoke-RoleWorkbookExport {
    [CmdletBinding(SupportsShouldProcess)]
    param([ValidateSet('Student','Teacher')][string] $Kind, [string] $File)
    $ErrorActionPreference = 'Stop'
    $label = if ($Kind -eq 'Student') { 'Schüler' } else { 'Lehrer' }
    $activity = "$label nach Excel exportieren"
    $File = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($File)
    # Validate before connecting, so an existing workbook never causes tenant access.
    Assert-WorkbookExportPath -Path $File
    $common = @{ WhatIf = [bool]$WhatIfPreference }
    if ($PSBoundParameters.ContainsKey('Confirm')) { $common.Confirm = $PSBoundParameters.Confirm }
    try {
        Write-Progress -Id 1 -Activity $activity -Status 'Lade Rollen, Benutzer und Exchange-Empfänger ...' -PercentComplete 10
        $inventory = Get-RoleWorkbookInventory -Kind $Kind
        Write-Progress -Id 1 -Activity $activity -Status 'Prüfe und schreibe neue Arbeitsmappe ...' -PercentComplete 80
        foreach ($warning in $inventory.Warnings) { Write-Warning $warning }
        $export = Export-AccountWorkbook -Inventory $inventory -Path $File -Kind $Kind @common
        $status = if ($null -eq $export) { 'Skipped' } else { 'Succeeded' }
        if ($null -eq $export -and $WhatIfPreference) { $status = 'WhatIf' }
        if ($null -ne $export) { Write-Information "${label}export: $($export.Count) Konten, $($export.Path)" -InformationAction Continue }
        [pscustomobject]@{ Mode = 'Export'; Path = $File; Count = $inventory.Accounts.Count; Status = $status
            TenantId = $inventory.TenantId; Warnings = $inventory.Warnings; HasErrors = $false }
    } finally { Write-Progress -Id 1 -Activity $activity -Completed }
}
