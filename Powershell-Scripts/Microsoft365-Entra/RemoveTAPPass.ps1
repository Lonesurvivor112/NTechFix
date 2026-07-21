
# Prompt for User ID and TAP ID
$userId = Read-Host "Enter the User ID (UPN or Object ID)"
$tapId = Read-Host "Enter the Temporary Access Pass Method ID"

# Confirm deletion
Write-Host "You are about to delete TAP ID '$tapId' for user '$userId'"
$confirmation = Read-Host "Type 'YES' to confirm"

if ($confirmation -eq "YES") {
try {
Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId -TemporaryAccessPassAuthenticationMethodId $tapId
Write-Host "Temporary Access Pass deleted successfully." -ForegroundColor Green
} catch {
Write-Host "Error deleting TAP: $_" -ForegroundColor Red
}
}