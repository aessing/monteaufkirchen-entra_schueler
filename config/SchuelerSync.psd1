@{
    Domain = 'monteaufkirchen.com'
    ExpectedTenantId = $null
    CompanyName = 'Montessori Schule Aufkirchen'
    EmployeeType = 'Schüler'
    AgeGroup = 'Minor'
    ConsentProvidedForMinor = 'Granted'
    LegalAgeGroupClassification = 'MinorWithParentalConsent'
    UsageLocation = 'DE'
    PasswordLength = 12
    StudentRoleGroup = @{
        Name = 'SEC-A-ROL-Schule_Schüler'
        Id = 'cebc1326-1174-4126-ba84-7a8960850e0a'
    }
    LicenseGroupName = 'SEC-A-LIC-O365A1Student'
    RoleGroupPrefix = 'SEC-A-ROL-'
    ClassGroupPrefix = 'SEC-A-CLS-'
    KnownRoleGroups = @(
        @{ Name = 'SEC-A-ROL-Schule_PädagogischesTeam'; Id = '103a1c4c-036b-40bc-bbe8-e3f7a866cb01' }
        @{ Name = 'SEC-A-ROL-Schule_Schüler'; Id = 'cebc1326-1174-4126-ba84-7a8960850e0a' }
        @{ Name = 'SEC-A-ROL-Schule_Sekretariat'; Id = '534f94ed-420d-4794-a374-20c572def8cb' }
        @{ Name = 'SEC-A-ROL-Schule_Vertretungskräfte'; Id = '4ed4f87b-374f-4679-ac11-5e51212e9a91' }
        @{ Name = 'SEC-A-ROL-Kinderhaus_PädagogischesTeam'; Id = 'cfc6d358-875d-441b-b45a-0d9734b7d788' }
        @{ Name = 'SEC-A-ROL-Kinderhaus_Sekretariat'; Id = '43165f2c-3ce5-471e-a7be-e16b66c03406' }
        @{ Name = 'SEC-A-ROL-Kinderhaus_Vertretungskräfte'; Id = 'd2e7e424-83d2-481d-9226-bf529bc9640a' }
        @{ Name = 'SEC-A-ROL-Ganztag'; Id = '3997db35-2914-4b0f-92b4-08038ef8f84a' }
        @{ Name = 'SEC-A-ROL-Unterstützung'; Id = 'a4a00442-bc04-48ee-b659-cb3bde194630' }
        @{ Name = 'SEC-A-ROL-ExterneBenutzer'; Id = '9c620843-0071-4d44-ba23-318e19cd09a0' }
        @{ Name = 'SEC-A-ROL-ExterneAdmins'; Id = 'a246dfde-eccf-489b-ac1c-8d92f67145ca' }
    )
    Exchange = @{
        AddressBookPolicy = 'MON-EXO-ABP-Schule_Schüler'
        CustomAttribute1 = 'Montessori Schule Aufkirchen - Schüler'
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
