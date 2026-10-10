@{
    RootModule = 'SchuelerSync.psm1'
    ModuleVersion = '0.3.0'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Invoke-SchuelerSync', 'Invoke-LehrerSync', 'Invoke-PersonalSync', 'Invoke-DisabledAccountCheck', 'Invoke-DisabledAccountTui')
}
