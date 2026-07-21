<#
.SYNOPSIS
    Export users and their AD group memberships with error logging.
#>

Import-Module ActiveDirectory

# --- Config ---
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir    = Join-Path -Path $PSScriptRoot -ChildPath 'logs'
$null      = New-Item -ItemType Directory -Path $logDir -Force

$failCsv   = Join-Path $logDir "GroupResolveFailures_$timestamp.csv"
$outCsv    = Join-Path $logDir "UserGroups_$timestamp.csv"

# --- Initialize collections ---
$Results  = [System.Collections.Generic.List[object]]::new()
$Failures = [System.Collections.Generic.List[object]]::new()

# --- Helper: Extract domain from DN ---
function Get-DomainFromDN {
    param([string]$DN)
    $dcParts = ($DN -split ',') | Where-Object { $_ -like 'DC=*' } | ForEach-Object { $_.Substring(3) }
    if ($dcParts) { $dcParts -join '.' } else { $null }
}

# --- Get all users (adjust filter as needed) ---
$users = Get-ADUser -Filter * -Properties memberOf,SamAccountName,DistinguishedName |
    Where-Object { $_.DistinguishedName -notlike "*OU=No-Longer-Employed-By-CC,OU=Users,OU=Contoso Corp,DC=contoso,DC=com*" }

foreach ($u in $users) {
    $memberOf = @($u.memberOf) | Where-Object { $_ }
    if ($memberOf.Count -eq 0) {
        # User has no groups
        $Results.Add([pscustomobject]@{
            UserSAM  = $u.SamAccountName
            Group    = '<none>'
        }) | Out-Null
        continue
    }

    foreach ($gdn in $memberOf) {
        $domainFqdn = Get-DomainFromDN -DN $gdn
        try {
            $params = @{
                Identity    = $gdn
                Properties  = 'samAccountName','name'
                ErrorAction = 'Stop'
            }
            if ($domainFqdn) { $params.Server = $domainFqdn }

            $group = Get-ADGroup @params
            if ($group) {
                $Results.Add([pscustomobject]@{
                    UserSAM = $u.SamAccountName
                    Group   = $group.Name  # or $group.samAccountName
                }) | Out-Null
            }
        }
        catch {
            $err = $_
            $Failures.Add([pscustomobject]@{
                Timestamp    = Get-Date
                UserSAM      = $u.SamAccountName
                GroupDN      = $gdn
                DomainTried  = $domainFqdn
                ErrorType    = $err.Exception.GetType().FullName
                ErrorMessage = $err.Exception.Message
            }) | Out-Null

            Write-Warning ("Failed to resolve group for {0}: {1} ({2})" -f $u.SamAccountName, $gdn, $err.Exception.Message)
        }
    }
}

# --- Export results ---
$Results | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
Write-Host "Exported user-group mapping to: $outCsv" -ForegroundColor Green

if ($Failures.Count -gt 0) {
    $Failures | Export-Csv -Path $failCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Exported failure log to: $failCsv" -ForegroundColor Yellow
} else {
    Write-Host "No group resolution failures detected." -ForegroundColor Green
}
