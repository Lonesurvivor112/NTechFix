# =============================================================================
# OneDrive / SharePoint Recursive File Counter
# =============================================================================
#
# PREREQUISITES
# -------------
# 1. Install the Microsoft Graph PowerShell module (only needed once):
#       Install-Module -Name Microsoft.Graph -Force
#
# 2. Connect to Microsoft Graph with your work account:
#       Connect-MgGraph -Scopes "Files.Read.All"
#    A browser login window will open. Sign in with your Microsoft/work account.
#
# =============================================================================
# STEP 1 - GET YOUR DRIVE ID
# =============================================================================
# Run this command, replacing the email with your own:
#
#   Get-MgUserDrive -UserId "yourname@yourcompany.com"
#
# Example output:
#   Name     Id                              DriveType
#   ----     --                              ---------
#   OneDrive b!2hx1FzwN_ku1Qs....           business
#
# Copy the full value in the "Id" column -- that is your DRIVE ID.
#
# =============================================================================
# STEP 2 - GET YOUR FOLDER ID
# =============================================================================
# Run this command using the Drive ID you just copied:
#
#   Get-MgDriveRootChild -DriveId "YOUR_DRIVE_ID_HERE"
#
# Example output:
#   Name                  Id                       WebUrl
#   ----                  --                       ------
#   Financial Clearance   016GCNV5IE2OTE2EDOKV...  https://...
#   Documents             016GCNV5PWVHJNSXJHIF...  https://...
#
# Find the folder you want to count and copy its full "Id" -- that is your FOLDER ID.
#
# =============================================================================
# STEP 3 - PASTE YOUR IDs BELOW AND RUN THE SCRIPT
# =============================================================================

# Replace these two values with your own:
$driveId  = "b!s0OPiWLy5k--OTglGCSMxwyJO79OPlFGhA9DVkcdV2o5bj8K7qb6RIwoulz-SN8G"
$folderId = "YOUR_FOLDER_ID_HERE"

$allFiles    = @()
$folderCount = 0
$queue       = [System.Collections.Generic.Queue[string]]::new()
$queue.Enqueue($folderId)

while ($queue.Count -gt 0) {
    $currentId = $queue.Dequeue()
    $folderCount++
    Write-Host "Processing folder $folderCount | Files found so far: $($allFiles.Count) | Folders remaining in queue: $($queue.Count)"

    $url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$currentId/children?`$top=999"

    do {
        $response = Invoke-MgGraphRequest -Uri $url -Method GET
        foreach ($item in $response.value) {
            if ($item.folder) {
                $queue.Enqueue($item.id)
            } else {
                $allFiles += $item
            }
        }
        $url = $response.'@odata.nextLink'
    } while ($url)
}

Write-Host ""
Write-Host "============================================"
Write-Host " DONE!"
Write-Host " Total folders processed : $folderCount"
Write-Host " Total files found       : $($allFiles.Count)"
Write-Host "============================================"