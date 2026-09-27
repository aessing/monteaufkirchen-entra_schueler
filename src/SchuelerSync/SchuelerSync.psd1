@{
    RootModule = 'SchuelerSync.psm1'
    ModuleVersion = '0.2.0'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Invoke-SchuelerSync', 'Invoke-LehrerSync', 'Invoke-DisabledAccountCheck', 'Invoke-DisabledAccountTui')
}
