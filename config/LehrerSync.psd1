@{
    Domain = 'monteaufkirchen.com'
    ExpectedTenantId = $null
    RoleGroupPrefix = 'SEC-A-ROL-'
    Roles = @{
        School = @{ Name = 'SEC-A-ROL-Schule_PädagogischesTeam'; Id = '103a1c4c-036b-40bc-bbe8-e3f7a866cb01' }
        Ganztag = @{ Name = 'SEC-A-ROL-Ganztag'; Id = '3997db35-2914-4b0f-92b4-08038ef8f84a' }
    }
    Jobs = @{
        L = 'School'
        'CO-L' = 'School'
        PA = 'School'
        OGTS = 'Ganztag'
    }
    AdditionalJobs = @('JAS')
    Profiles = @{
        School = @{
            CompanyName = 'Montessori Schule Aufkirchen'
            EmployeeType = 'Pädagogisches Team'
            LicenseGroupName = 'SEC-A-LIC-M365A3Faculty'
            AddressBookPolicy = 'MON-EXO-ABP-Schule_PädagogischesTeam'
        }
        Ganztag = @{
            CompanyName = 'Montessori Verein Landkreis Erding e.V.'
            EmployeeType = 'Ganztag'
            LicenseGroupName = 'SEC-A-LIC-O365A1Faculty'
            AddressBookPolicy = 'MON-EXO-ABP-Ganztag'
        }
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
