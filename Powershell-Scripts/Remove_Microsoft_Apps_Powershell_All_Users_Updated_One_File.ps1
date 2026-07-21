#Removal Of Microsoft Apps Script#

# Start of Current User Script#
Write-Host "Starting Removal Of Microsoft Apps" -ForegroundColor Green
write-host ""
Write-Host "Applies to Current User" -ForegroundColor Green
Write-host "Removing Windows Apps Current User" -ForegroundColor Green
Write-host "Running"

Write-host "Uninstalling 3D Builder" -ForegroundColor Cyan
Get-AppxPackage *3dbuilder* | Remove-AppxPackage

Write-host "Uninstalling Calendar and Mail" -ForegroundColor Cyan
Get-AppxPackage *windowscommunicationsapps* | Remove-AppxPackage

Write-host "Uninstalling Get Office" -ForegroundColor Cyan
Get-AppxPackage *officehub* | Remove-AppxPackage

Write-host "Uninstalling Get Skype" -ForegroundColor Cyan
Get-AppxPackage *skypeapp* | Remove-AppxPackage

Write-host "Uninstalling Groove Music" -ForegroundColor Cyan
Get-AppxPackage *zunemusic* | Remove-AppxPackage

Write-host "Uninstalling Maps" -ForegroundColor Cyan
Get-AppxPackage *windowsmaps* | Remove-AppxPackage

Write-host "Uninstalling Microsoft Solitaire Collection" -ForegroundColor Cyan
Get-AppxPackage *solitairecollection* | Remove-AppxPackage

Write-host "Uninstalling Bing Finance" -ForegroundColor Cyan
Get-AppxPackage *bingfinance* | Remove-AppxPackage

Write-host "Uninstalling Zune Movies & TV" -ForegroundColor Cyan
Get-AppxPackage *zunevideo* | Remove-AppxPackage

Write-host "Uninstalling People" -ForegroundColor Cyan
Get-AppxPackage *people* | Remove-AppxPackage

Write-host "Uninstalling Phone Companion" -ForegroundColor Cyan
Get-AppxPackage *windowsphone* | Remove-AppxPackage

Write-host "Uninstalling Xbox app" -ForegroundColor Cyan
Get-AppxPackage *xboxapp* | Remove-AppxPackage

Write-Host ""
Write-Host "Transition To Next Script" -ForegroundColor DarkYellow
Write-Host "Transition To Next Script" -ForegroundColor DarkYellow
Write-Host "Transition To Next Script" -ForegroundColor DarkYellow
Write-Host ""

#End Of Current User Removal#
#Start Of Future User Removal#


Write-Host "Starting Removal Of Microsoft Apps" -ForegroundColor Green
Write-Host ""
Write-Host "Applys To Future Users" -ForegroundColor Green
Write-host ""
Write-Host "Starting" -ForegroundColor Blue


Write-host "Your Phone" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.YourPhone*"} | remove-appxprovisionedpackage –online

Write-host "Office" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.Office.Onenote*"} | remove-appxprovisionedpackage –online

Write-host "Mail and Calendar" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.windowscommunicationsapps*"} | remove-appxprovisionedpackage –online

Write-host "Xbox Game Bar" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.XboxGamingOverlay*"} | remove-appxprovisionedpackage –online

Write-host "Skype App" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.SkypeApp*"} | remove-appxprovisionedpackage –online

Write-host "Xbox Identity Provider" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.XboxIdentityProvider*"} | remove-appxprovisionedpackage –online

Write-host "Office Hub" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.MicrosoftOfficeHub*"} | remove-appxprovisionedpackage –online

Write-host "Solitaire" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.MicrosoftSolitaireCollection*"} | remove-appxprovisionedpackage –online

Write-host "VR App" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.MixedReality.Portal*"} | remove-appxprovisionedpackage –online

Write-host "People App" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.People*"} | remove-appxprovisionedpackage –online
 
Write-host "Wallet App" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.Wallet*"} | remove-appxprovisionedpackage –online

Write-host "Feedback Hub" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.WindowsFeedbackHub*"} | remove-appxprovisionedpackage –online

Write-host "Windows Maps" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.WindowsMaps*"} | remove-appxprovisionedpackage -online

Write-host "More Xbox Stuff" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.XboxApp*"} | remove-appxprovisionedpackage –online

Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.XboxGameOverlay*"} | remove-appxprovisionedpackage –online

Write-host "Microsoft Music" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.ZuneMusic*"} | remove-appxprovisionedpackage –online

Write-host "GetStarted" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.getstarted*"} | remove-appxprovisionedpackage –online

Write-host "bing Finance" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.bingfinance*"} | remove-appxprovisionedpackage –online

Write-host "Bing sports" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.bingsports*"} | remove-appxprovisionedpackage –online

Write-host "Sway" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.sway*"} | remove-appxprovisionedpackage –online

Write-host "holographic" -foregroundcolor Green
Get-appxprovisionedpackage –online | where-object {$_.packagename –like
"*microsoft.holographic*"} | remove-appxprovisionedpackage –online

write-host ""
Write-host "Reverting Execution Policy"
Set-ExecutionPolicy Restricted
write-host "Getting the current Execution Policy after change"
Get-ExecutionPolicy


Write-Host "The Script should complete without Errors, If you enounter errors check the script file" -Foregroundcolor Red