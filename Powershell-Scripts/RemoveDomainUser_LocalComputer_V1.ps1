#Param takes input "-user user" and relays the info into the script to remove the correct user
param($user)
#Explaining whats going on
Write-Host "The Proper User Removal Tool" -BackgroundColor Green
Write-Host "User Specified will be permenently deleted and potentially CANNOT be recovered" -ForegroundColor Red
Write-Host "!!!PLEASE CTRL C NOW IF MISTAKE!!!" -BackgroundColor Red
start-sleep -Seconds 8
Write-host "Starting Removal of" $user -BackgroundColor Cyan
#Can Be used potentially as input while running the script#
###$user = Read-Host "Enter Username you would like to remove"
Get-CimInstance -Class Win32_UserProfile | Where-Object { $_.LocalPath.split('\')[-1] -eq $user } | Remove-CimInstance