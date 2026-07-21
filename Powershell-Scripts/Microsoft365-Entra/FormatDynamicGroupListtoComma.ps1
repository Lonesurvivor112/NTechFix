# Path to your input CSV file
$inputFile = "X:\LifecyclemanagementClone\DynamicGroup_ReferencedGroups.csv"
$outputFile = "X:\LifecyclemanagementClone\CommaSeperatedDynamicGroupMembers.xlsx"

# Import the CSV
$data = Import-Csv -Path $inputFile

# Group by DisplayName and join DynamicGroups
$grouped = $data | Group-Object -Property DynamicGroupName | ForEach-Object {
    [PSCustomObject]@{
        DisplayName = $_.Name
        DynamicGroups = ($_.Group.ReferencedName -join ", ")
    }
}

# Export to Excel (requires ImportExcel module) or CSV
$grouped | Export-Excel -Path $outputFile -AutoSize -AutoFilter
# If you don't have ImportExcel module, use CSV:
# $grouped | Export-Csv -Path "C:\\path\\to\\grouped_output.csv" -NoTypeInformation