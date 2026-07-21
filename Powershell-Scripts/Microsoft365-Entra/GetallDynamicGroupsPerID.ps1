# Requires: Microsoft.Graph & ImportExcel modules
# Run this script directly in PowerShell
#Combine with FormatDynamicGroupListtoComma.ps1 for ease of updating Staff_Roles_and_Permissions_EC.xlsx

# Prompt for Object IDs in terminal
Write-Host "Paste Object IDs (one per line). Press Enter on a blank line to finish:`n"

$groupIds = @()
while ($true) {
    $line = Read-Host
    if ([string]::IsNullOrWhiteSpace($line)) { break }
    if ($line -match '^[a-fA-F0-9\-]{36}$') {
        $groupIds += $line
    } else {
        Write-Warning "Invalid GUID format: $line"
    }
}

if ($groupIds.Count -eq 0) {
    Write-Host "No valid Object IDs entered. Exiting."
    return
}

# Connect to Microsoft Graph
$requiredScopes = @('Directory.Read.All','Group.Read.All')
if (-not (Get-MgContext) -or @((Get-MgContext).Scopes) -notcontains 'Group.Read.All') {
    Connect-MgGraph -Scopes $requiredScopes
}

# Get all dynamic groups
$dynGroups = Get-MgGroup -Filter "groupTypes/any(s:s eq 'DynamicMembership')" -All `
                         -Property "id,displayName,membershipRule"

$results = @()

foreach ($id in $groupIds) {
    try {
        $cloudGroup = Get-MgGroup -GroupId $id -Property displayName
        $displayName = $cloudGroup.DisplayName
    } catch {
        Write-Warning "Could not resolve group $id"
        continue
    }

    $pattern = [Regex]::Escape($id)
    $matches = $dynGroups | Where-Object { $_.membershipRule -and ($_.membershipRule -match $pattern) }

    foreach ($match in $matches) {
        $snippet = $match.MembershipRule -replace ".*?($id).*?", '...$1...'
        $results += [PSCustomObject]@{
            ObjectId     = $id
            DisplayName  = $displayName
            DynamicGroup = $match.DisplayName
            RuleSnippet  = $snippet
            FullRule     = $match.MembershipRule
        }
    }
}

# Export to Excel
if ($results.Count -eq 0) {
    Write-Host "No dynamic group references found."
} else {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel
    $results | Export-Excel -Path ".\DynamicGroupAssignments.xlsx" -AutoSize -WorksheetName "DynamicGroupRefs"
    Write-Host "✅ Exported results to DynamicGroupAssignments.xlsx"
}
