# Search by SAM account name (on-prem synced users only)
$UserPrincipalNameReal = read-host "Enter the User Principal Name (UPN) of the user"
Write-Output "You Entered: $UserPrincipalNameReal"
$user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalNameReal'"

# Store Object ID
$userId = $user.Id

#Get TAP Methods and Remove ALL
$tapMethods = Get-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId
foreach ($method in $tapMethods) {
    Write-Output "Removing TAP with ID: $($method.Id)"
    Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId -TemporaryAccessPassAuthenticationMethodId $method.Id
}


# Confirm removal
$remainingTaps = Get-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId
if ($remainingTaps.Count -eq 0) {
    Write-Output "✅ All TAP methods have been successfully removed for $UserPrincipalNameReal."
} else {
    Write-Output "⚠️ Some TAP methods still remain:"
    $remainingTaps | ForEach-Object { Write-Output " - ID: $($_.Id)" }
}
