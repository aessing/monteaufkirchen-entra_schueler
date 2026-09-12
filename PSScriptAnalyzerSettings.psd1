@{
    Severity = @('Error', 'Warning')
    IncludeDefaultRules = $true
    Rules = @{
        PSAvoidUsingPlainTextForPassword = @{ Enable = $true }
        PSAvoidUsingWriteHost = @{ Enable = $true }
        PSUseShouldProcessForStateChangingFunctions = @{ Enable = $true }
    }
}
