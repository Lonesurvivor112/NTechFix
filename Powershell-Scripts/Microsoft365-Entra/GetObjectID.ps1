
# Connect to Microsoft Graph
#Connect-MgGraph -Scopes "User.Read.All"
# Search by SAM account name (on-prem synced users only)
$UserPrincipalNameReal = read-host "Enter the User Principal Name (UPN) of the user"
Write-Output "You Entered: $UserPrincipalNameReal"
$user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalNameReal'"

# Store Object ID
$objectId = $user.Id

# Output
Write-Output "Object ID $objectId"
