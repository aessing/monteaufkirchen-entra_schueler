BeforeAll {
    $repoRoot = Split-Path $PSScriptRoot -Parent
    $entryPoint = Join-Path $repoRoot 'Sync-SchuelerEntra.ps1'
    $readmePath = Join-Path $repoRoot 'README.md'
    $guidePath = Join-Path $repoRoot 'docs/ENTRA-SCHUELER-SYNC.md'
    $operationsPath = Join-Path $repoRoot 'docs/BETRIEB.md'
    $testingPath = Join-Path $repoRoot 'docs/TESTING.md'
    $heroPath = Join-Path $repoRoot 'docs/assets/entra-schueler-sync-hero.png'
    $command = Get-Command $entryPoint -ErrorAction Stop
    $readme = Get-Content -LiteralPath $readmePath -Raw
    $guide = Get-Content -LiteralPath $guidePath -Raw
    $operations = Get-Content -LiteralPath $operationsPath -Raw
    $testing = Get-Content -LiteralPath $testingPath -Raw
}

Describe 'Documentation contract' {
    It 'derives every documented public selector from the executable command' {
        foreach ($parameter in @(
                'File', 'Update', 'CreateNewUsers', 'DisableUsers', 'UpdateUsers',
                'RevokeSessions', 'ConfigureExchangeOnlineOnly', 'Mail', 'OutputFile',
                'Add', 'Remove', 'Vorname', 'Nachname', 'Klasse', 'Klassenlehrer', 'EntraObjectId', 'WhatIf',
                'Confirm', 'Verbose'
            )) {
            $command.Parameters.Keys | Should -Contain $parameter
        }
        $command.Parameters['Mail'].Aliases | Should -Contain 'UPN'
        $command.Parameters['Mail'].Aliases | Should -Contain 'UserPrincipalName'
    }

    It 'ships the generic hero and all operator documents' {
        Test-Path -LiteralPath $heroPath -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $guidePath -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $operationsPath -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $testingPath -PathType Leaf | Should -BeTrue
    }

    It 'keeps the README examples aligned with the command surface' {
        $readme | Should -Match ([regex]::Escape('docs/assets/entra-schueler-sync-hero.png'))
        $readme | Should -Match ([regex]::Escape('.\Sync-SchuelerEntra.ps1'))
        $readme | Should -Match ([regex]::Escape('-File'))
        $readme | Should -Match ([regex]::Escape('-Update'))
        $readme | Should -Match ([regex]::Escape('-CreateNewUsers'))
        $readme | Should -Match ([regex]::Escape('-UpdateUsers'))
        $readme | Should -Match ([regex]::Escape('-DisableUsers'))
        $readme | Should -Match ([regex]::Escape('-RevokeSessions'))
        $readme | Should -Match ([regex]::Escape('-WhatIf'))
        $readme | Should -Match ([regex]::Escape('-ConfigureExchangeOnlineOnly'))
        $readme | Should -Match ([regex]::Escape('-Mail'))
        $readme | Should -Match ([regex]::Escape('-OutputFile'))
        $readme | Should -Match ([regex]::Escape('-Add'))
        $readme | Should -Match ([regex]::Escape('-Remove'))
    }

    It 'keeps the recommended report directory out of Git' {
        $reportPath = Join-Path $repoRoot 'Berichte/Schueler-Abgleich.txt'
        $null = & git -C $repoRoot check-ignore --no-index $reportPath
        $LASTEXITCODE | Should -Be 0
    }

    It 'documents the exact workbook schema and managed columns' {
        foreach ($header in @('Name mit Rufname', 'Vorname', 'Nachname', 'Klassen', 'Klassenlehrer', 'Passwort', 'EntraObjectId', 'UPN')) {
            $guide | Should -Match ([regex]::Escape($header))
        }
        $guide | Should -Match 'Kopfzeile|Zeile 1'
        $guide | Should -Match 'genau ein Arbeitsblatt'
        $guide | Should -Match 'vollständig leer|leere Zeilen'
        $guide | Should -Match ([regex]::Escape('JK1-3g2_1'))
    }

    It 'documents runtime modules, Graph scopes, roles and privacy boundaries' {
        foreach ($module in @(
                'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users',
                'Microsoft.Graph.Users.Actions', 'Microsoft.Graph.Groups',
                'ExchangeOnlineManagement', 'ImportExcel', 'Pester', 'PSScriptAnalyzer'
            )) {
            $guide | Should -Match ([regex]::Escape($module))
        }
        foreach ($scope in @(
                'User.ReadWrite.All', 'User-Mail.ReadWrite.All', 'Group.Read.All',
                'GroupMember.ReadWrite.All', 'User.RevokeSessions.All'
            )) {
            $guide | Should -Match ([regex]::Escape($scope))
        }
        foreach ($role in @('User Administrator', 'Groups Administrator', 'Exchange Administrator')) {
            $guide | Should -Match ([regex]::Escape($role))
        }
        $guide | Should -Match 'sechs\s+Pr.fungen|sofort.+f.nf'
        $guide | Should -Match '60\s+Sekunden'
        $guide | Should -Match ([regex]::Escape('legalAgeGroupClassification'))
        $guide | Should -Match 'schreibgesch.tzt'
        $guide | Should -Match ([regex]::Escape('/*.xlsx'))
        $guide | Should -Match 'Passw.rter.+nicht.+Ausgabe|keine.+Passw.rter.+Ausgabe'
    }

    It 'documents every managed Exchange mailbox and CAS value' {
        foreach ($value in @(
                'Montessori Schule Aufkirchen - Schüler',
                'MON-EXO-ABP-Schule_Schüler',
                'MON-EXO-UserRoles-Default',
                'MON-EXO-Sharing-Default',
                'MON-EXO-Retention-Default',
                'MON-EXO-OWA-Default',
                'CommunicationsCompliance', 'MailItemsAccessed',
                'ActiveSyncEnabled', 'ImapEnabled', 'MAPIEnabled', 'OWAEnabled',
                'OWAforDevicesEnabled', 'PopEnabled', 'SmtpClientAuthenticationDisabled'
            )) {
            $guide | Should -Match ([regex]::Escape($value))
        }
        foreach ($auditAction in @(
                'Create', 'FolderBind', 'HardDelete', 'Move', 'MoveToDeletedItems',
                'SendAs', 'SendOnBehalf', 'SoftDelete', 'Update',
                'UpdateFolderPermissions', 'UpdateInboxRules', 'MailboxLogin',
                'UpdateCalendarDelegation', 'Copy'
            )) {
            $guide | Should -Match ([regex]::Escape($auditAction))
        }
        $guide | Should -Match 'AuditLogAgeLimit.+365'
        $guide | Should -Match 'RetainDeletedItemsFor.+30'
    }

    It 'warns that departure actions target the complete configured student role population' {
        foreach ($document in @($operations, $testing)) {
            $document | Should -Match 'alle.+Mitglieder.+SEC-A-ROL-Schule_Sch.ler'
            $document | Should -Match 'vollst.ndige.+Sch.lerpopulation'
            $document | Should -Match 'Arbeitsmappenausschnitt|Testklasse.+nicht.+isoliert'
        }
    }
}
