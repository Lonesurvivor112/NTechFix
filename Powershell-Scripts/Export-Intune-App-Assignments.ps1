# Get-IntuneAppAssignments.ps1
# Exports all Intune app assignments (group name, assignment type) to CSV
# Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Devices.CorporateManagement

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "DeviceManagementApps.Read.All", "Group.Read.All"

# Get all mobile apps
$apps = Get-MgDeviceAppManagementMobileApp -All

$results = @()

foreach ($app in $apps) {
    # Get assignments for each app
    $assignments = Get-MgDeviceAppManagementMobileAppAssignment -MobileAppId $app.Id

    if ($assignments) {
        foreach ($assignment in $assignments) {
            $groupName = "N/A"
            $targetType = $assignment.Target.AdditionalProperties.'@odata.type'

            # Resolve group name if assigned to a group
            if ($targetType -eq "#microsoft.graph.groupAssignmentTarget" -or
                $targetType -eq "#microsoft.graph.exclusionGroupAssignmentTarget") {
                $groupId = $assignment.Target.AdditionalProperties.groupId
                try {
                    $group = Get-MgGroup -GroupId $groupId
                    $groupName = $group.DisplayName
                } catch {
                    $groupName = "Group not found ($groupId)"
                }
            } elseif ($targetType -eq "#microsoft.graph.allDevicesAssignmentTarget") {
                $groupName = "All Devices"
            } elseif ($targetType -eq "#microsoft.graph.allLicensedUsersAssignmentTarget") {
                $groupName = "All Users"
            }

            $results += [PSCustomObject]@{
                AppName        = $app.DisplayName
                AppType        = $app.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
                AssignmentType = switch ($assignment.Intent) {
                    "required"       { "Required" }
                    "available"      { "Available" }
                    "uninstall"      { "Uninstall" }
                    default          { $assignment.Intent }
                }
                GroupName      = $groupName
                TargetType     = $targetType -replace '#microsoft.graph.', ''
            }
        }
    } else {
        $results += [PSCustomObject]@{
            AppName        = $app.DisplayName
            AppType        = $app.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
            AssignmentType = "No Assignment"
            GroupName      = "N/A"
            TargetType     = "N/A"
        }
    }
}

# Export to CSV
$outputPath = "X:\IntuneAppAssignments.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "Export complete: $outputPath" -ForegroundColor Green
Write-Host "Total apps: $($apps.Count)" -ForegroundColor Cyan
Write-Host "Total assignment entries: $($results.Count)" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph