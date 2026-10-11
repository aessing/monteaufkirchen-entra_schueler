@{
    PersonnelMode = $true
    Domain = 'montessori-erding.de'
    ExpectedTenantId = $null
    EntryPoint = 'Sync-PersonalEntra.ps1'
    WorkbookKind = 'Personnel'
    RoleGroupPrefix = 'SEC-A-ROL-'
    ProtectedJobs = @('ExterneAdmins')
    ProtectedGroupNames = @('SEC-A-ROL-ExterneAdmins', 'SEC-A-ADM-GlobalAdministrators')
    Roles = @{
        Schule_Sekretariat = @{ Name = 'SEC-A-ROL-Schule_Sekretariat'; Id = '534f94ed-420d-4794-a374-20c572def8cb' }
        Schule_Vertretungskräfte = @{ Name = 'SEC-A-ROL-Schule_Vertretungskräfte'; Id = '4ed4f87b-374f-4679-ac11-5e51212e9a91' }
        Kinderhaus_PädagogischesTeam = @{ Name = 'SEC-A-ROL-Kinderhaus_PädagogischesTeam'; Id = 'cfc6d358-875d-441b-b45a-0d9734b7d788' }
        Kinderhaus_Sekretariat = @{ Name = 'SEC-A-ROL-Kinderhaus_Sekretariat'; Id = '43165f2c-3ce5-471e-a7be-e16b66c03406' }
        Kinderhaus_Vertretungskräfte = @{ Name = 'SEC-A-ROL-Kinderhaus_Vertretungskräfte'; Id = 'd2e7e424-83d2-481d-9226-bf529bc9640a' }
        Unterstützung = @{ Name = 'SEC-A-ROL-Unterstützung'; Id = 'a4a00442-bc04-48ee-b659-cb3bde194630' }
        ExterneBenutzer = @{ Name = 'SEC-A-ROL-ExterneBenutzer'; Id = '9c620843-0071-4d44-ba23-318e19cd09a0' }
        ExterneAdmins = @{ Name = 'SEC-A-ROL-ExterneAdmins'; Id = 'a246dfde-eccf-489b-ac1c-8d92f67145ca' }
    }
    Jobs = @{
        Schule_Sekretariat = 'Schule_Sekretariat'
        Schule_Vertretungskräfte = 'Schule_Vertretungskräfte'
        Kinderhaus_PädagogischesTeam = 'Kinderhaus_PädagogischesTeam'
        Kinderhaus_Sekretariat = 'Kinderhaus_Sekretariat'
        Kinderhaus_Vertretungskräfte = 'Kinderhaus_Vertretungskräfte'
        Unterstützung = 'Unterstützung'
        ExterneBenutzer = 'ExterneBenutzer'
        ExterneAdmins = 'ExterneAdmins'
    }
    AdditionalJobs = @()
    Profiles = @{
        Schule_Sekretariat = @{ CompanyName = 'Montessori Schule Aufkirchen'; EmployeeType = 'Sekretariat'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-Schule_Sekretariat' }
        Schule_Vertretungskräfte = @{ CompanyName = 'Montessori Schule Aufkirchen'; EmployeeType = 'Vertretungskräfte'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-Schule_Vertretungskräfte' }
        Kinderhaus_PädagogischesTeam = @{ CompanyName = 'Montessori Kinderhaus Erding'; EmployeeType = 'Pädagogisches Team'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-Kinderhaus_PädagogischesTeam' }
        Kinderhaus_Sekretariat = @{ CompanyName = 'Montessori Kinderhaus Erding'; EmployeeType = 'Sekretariat'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-Kinderhaus_Sekretariat' }
        Kinderhaus_Vertretungskräfte = @{ CompanyName = 'Montessori Kinderhaus Erding'; EmployeeType = 'Vertretungskräfte'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-Kinderhaus_Vertretungskräfte' }
        Unterstützung = @{ CompanyName = 'Montessori Verein Landkreis Erding e.V.'; EmployeeType = 'Unterstützung'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-Unterstützung' }
        ExterneBenutzer = @{ CompanyName = 'Montessori Verein Landkreis Erding e.V.'; EmployeeType = 'Externe Benutzer'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = 'MON-EXO-ABP-ExterneBenutzer' }
        ExterneAdmins = @{ CompanyName = 'Montessori Verein Landkreis Erding e.V.'; EmployeeType = 'Externe Admins'; LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'; AddressBookPolicy = $null }
    }
    Exchange = @{
        AuditLogAgeLimitDays = 365
        RetainDeletedItemsForDays = 30
        RoleAssignmentPolicy = 'MON-EXO-UserRoles-Default'
        SharingPolicy = 'MON-EXO-Sharing-Default'
        RetentionPolicy = 'MON-EXO-Retention-Default'
        OwaMailboxPolicy = 'MON-EXO-OWA-Default'
        MaxMailboxRetries = 5
        RetryDelaySeconds = 60
    }
}
