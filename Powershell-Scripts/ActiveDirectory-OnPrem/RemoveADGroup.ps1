Import-Module ActiveDirectory -ErrorAction Stop

Write-Host "Paste or type one username per line (SamAccountName or UPN). Press Enter on a blank line when done." -ForegroundColor Cyan

# Collect users interactively (one per line)
$inputs = New-Object System.Collections.Generic.List[string]
while ($true) {
    $line = Read-Host "User"
    if ([string]::IsNullOrWhiteSpace($line)) { break }
    $inputs.Add($line.Trim())
}

if ($inputs.Count -eq 0) {
    Write-Error "No users provided. Exiting."
    return
}

# Ask for the group
$groupInput = Read-Host "Enter the AD group (Name, sAMAccountName, or DistinguishedName)"

# Resolve the group robustly
try {
    $group = Get-ADGroup -Identity $groupInput -ErrorAction Stop
}
catch {
    # Try by Name if Identity didn't work
    $escaped = $groupInput -replace '([\\()\.\*\+\?\|\{\}\[\]\^\$])','\$1'
    $candidates = Get-ADGroup -LDAPFilter "(name=$escaped)" -ErrorAction SilentlyContinue
    if ($candidates -and $candidates.Count -eq 1) {
        $group = $candidates
    }
    elseif ($candidates -and $candidates.Count -gt 1) {
        Write-Error "Multiple groups matched '$groupInput'. Please specify sAMAccountName or the DN."
        return
    }
    else {
        Write-Error "Group '$groupInput' not found."
        return
    }
}

Write-Host "Removing users from group: $($group.Name) ($($group.DistinguishedName))" -ForegroundColor Yellow

$results = foreach ($u in $inputs) {
    # Resolve user by Identity, then fall back to SAM/UPN filter
    $user = $null
    try { $user = Get-ADUser -Identity $u -ErrorAction Stop } catch {}

    if (-not $user) {
        $user = Get-ADUser -LDAPFilter "(|(sAMAccountName=$u)(userPrincipalName=$u))" -ErrorAction SilentlyContinue
    }

    if (-not $user) {
        [pscustomobject]@{
            User              = $u
            UserDN            = $null
            Group             = $group.Name
            Status            = "User not found"
        }
        continue
    }

    try {
        # Remove direct membership (no per-item confirmation)
        Remove-ADGroupMember -Identity $group.DistinguishedName -Members $user.DistinguishedName -Confirm:$false -ErrorAction Stop
        [pscustomobject]@{
            User              = $user.SamAccountName
            UserDN            = $user.DistinguishedName
            Group             = $group.Name
            Status            = "Removed"
        }
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'is not a member') {
            $status = "Not a direct member"
        } else {
            $status = "Error: $msg"
        }
        [pscustomobject]@{
            User              = $user.SamAccountName
            UserDN            = $user.DistinguishedName
            Group             = $group.Name
            Status            = $status
        }
    }
}

# Show summary
$results | Sort-Object Status, User | Format-Table -AutoSize

# Optional: export a log
# $results | Export-Csv -Path ".\RemovedUsersFrom_$($group.Name -replace '[^\w-]','_').csv" -NoTypeInformation -Encoding UTF8
