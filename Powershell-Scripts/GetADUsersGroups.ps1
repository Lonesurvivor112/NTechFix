<#
Requires: RSAT Active Directory module
Run in a PowerShell session with domain creds that can read users & groups.
#>

[CmdletBinding()]
param(
    # Base container to search users under (sub-OUs included)
    [string]$BaseOU = "OU=Users,OU=Contoso Corp,DC=contoso,DC=com",

    # One or more OU DNs to exclude (everything under these OUs will be skipped)
    [string[]]$ExcludeOUs = @(
        "OU=No-Longer-Employed-By-CC,OU=Users,OU=Contoso Corp,DC=contoso,DC=com"
    ),

    # Include nested group membership (effective membership). Omit for direct-only.
    [switch]$Transitive,

    # Optional: export path
    [string]$ExportCsvPath = ".\AD_UserGroups_$(if($Transitive){'Transitive'}else{'Direct'}).csv"
)

Import-Module ActiveDirectory -ErrorAction Stop

# Build a single case-insensitive regex that matches any excluded OU DN fragment
$excludePattern = ($ExcludeOUs | ForEach-Object { [regex]::Escape($_) }) -join '|'

Write-Host "Querying users under $BaseOU (excluding $($ExcludeOUs -join '; '))..." -ForegroundColor Cyan

# Pull users (subtree), include memberOf for fast direct lookups
$users = Get-ADUser -SearchBase $BaseOU `
                    -LDAPFilter "(&(objectCategory=person)(objectClass=user))" `
                    -SearchScope Subtree `
                    -Properties memberOf, displayName, samAccountName, distinguishedName

# Exclude any users whose DN contains one of the excluded OU paths (covers sub-OUs too)
$users = $users | Where-Object { $_.DistinguishedName -notmatch $excludePattern }

if (-not $users) {
    Write-Warning "No users found after applying OU exclusion."
    return
}

Write-Host ("Found {0} user(s). Collecting {1} memberships..." -f $users.Count, $(if($Transitive){"TRANSITIVE"}else{"DIRECT"})) -ForegroundColor Cyan

$rows = New-Object System.Collections.Generic.List[object]

foreach ($u in $users) {
    if ($Transitive) {
        # Includes nested groups
        $groups = Get-ADPrincipalGroupMembership -Identity $u.SamAccountName |
                  Sort-Object Name
    }
    else {
        # Direct groups from memberOf (may be $null)
        $groups = @()
        if ($u.memberOf) {
            foreach ($gdn in $u.memberOf) {
                try {
                    $groups += Get-ADGroup -Identity $gdn -Properties samAccountName, name, groupCategory, groupScope, distinguishedName
                } catch {
                    Write-Verbose "Could not resolve group DN '$gdn' for user $($u.SamAccountName): $($_.Exception.Message)"
                }
            }
            $groups = $groups | Sort-Object Name
        }
    }

    if ($groups -and $groups.Count -gt 0) {
        foreach ($g in $groups) {
            $rows.Add([pscustomobject]@{
                SamAccountName   = $u.SamAccountName
                UserDisplayName  = $u.DisplayName
                UserDN           = $u.DistinguishedName
                GroupSam         = $g.SamAccountName
                GroupName        = $g.Name
                GroupDN          = $g.DistinguishedName
                GroupScope       = $g.GroupScope
                GroupCategory    = $g.GroupCategory
                MembershipScope  = $(if($Transitive){"Transitive"}else{"Direct"})
            })
        }
    }
    else {
        # User with no groups in the chosen scope
        $rows.Add([pscustomobject]@{
            SamAccountName   = $u.SamAccountName
            UserDisplayName  = $u.DisplayName
            UserDN           = $u.DistinguishedName
            GroupSam         = $null
            GroupName        = $null
            GroupDN          = $null
            GroupScope       = $null
            GroupCategory    = $null
            MembershipScope  = $(if($Transitive){"Transitive"}else{"Direct"})
        })
    }
}

# Output & export
$report = $rows | Sort-Object SamAccountName, GroupName, GroupDN

Write-Host ("Report rows: {0}" -f $report.Count) -ForegroundColor Green
$report | Format-Table -AutoSize

if ($ExportCsvPath) {
    $report | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV exported to: $ExportCsvPath" -ForegroundColor Magenta
}
