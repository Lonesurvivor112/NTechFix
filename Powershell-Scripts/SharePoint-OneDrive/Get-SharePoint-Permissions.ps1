# Make sure you're connected to THAT site (the one in the expiration email)
 >  # Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/THESITE" -Interactive -ClientId "your-app-id"
 >  
 >  # Show the site's permission groups and their members
 >  Get-PnPGroup | ForEach-Object {
 >      $g = $_
 >      Write-Host "`n=== GROUP: $($g.Title) ===" -ForegroundColor Cyan
 >      Get-PnPGroupMember -Group $g | Select-Object Title, Email, LoginName | Format-Table -AutoSize
 >  }
Get-PnPGroup | Where-Object { $_.Title -like "SharingLinks*" } | Select Title, Id
 >  
 >  # For each, see who's in it — find the one(s) with your two users
 >  Get-PnPGroup | Where-Object { $_.Title -like "SharingLinks*" } | ForEach-Object {
 >      $members = Get-PnPGroupMember -Group $_
 >      if (($members.Email -contains "imascam@gmail.com") -or ($members.Email -contains "Jeffery@mememeh.net")) {
 >          Write-Host "MATCH in group: $($_.Title)" -ForegroundColor Green
 >          $members | Select Title, Email | Format-Table -AutoSize
 >      }
 >  }
 # Find which item each sharing-link group is attached to
 >  $targetGroups = @(194, 195)
 >  
 >  Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } | ForEach-Object {
 >      $list = $_
 >      Get-PnPListItem -List $list -PageSize 500 | ForEach-Object {
 >          $item = $_
 >          $fileRef = $item["FileRef"]
 >          if (-not $fileRef) { return }
 >          try {
 >              $hasUnique = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
 >              if ($hasUnique) {
 >                  $ras = Get-PnPProperty -ClientObject $item -Property RoleAssignments
 >                  foreach ($ra in $ras) {
 >                      $m = Get-PnPProperty -ClientObject $ra -Property Member
 >                      if ($targetGroups -contains $m.Id) {
 >                          Write-Host "GROUP $($m.Id) -> $fileRef" -ForegroundColor Green
 >                      }
 >                  }
 >              }
 >          } catch {}
 >      }
 >  }
$targetGroups = @(194, 195)
 >  
 >  Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } | ForEach-Object {
 >      $list = $_
 >      Get-PnPListItem -List $list -PageSize 500 | ForEach-Object {
 >          $item = $_
 >          $fileRef = $item["FileRef"]
 >          if (-not $fileRef) { return }
 >          try {
 >              $hasUnique = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
 >              if ($hasUnique) {
 >                  $ras = Get-PnPProperty -ClientObject $item -Property RoleAssignments
 >                  foreach ($ra in $ras) {
 >                      $m = Get-PnPProperty -ClientObject $ra -Property Member
 >                      if ($targetGroups -contains $m.Id) {
 >                          Write-Host "GROUP $($m.Id) | Type: $($item.FileSystemObjectType) | $fileRef" -ForegroundColor Green
 >                      }
 >                  }
 >              }
 >          } catch {}
 >      }
 >  }