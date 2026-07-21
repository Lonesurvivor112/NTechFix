# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"
# Search by SAM account name (on-prem synced users only)
$Usernameget = read-host "Enter the User Principal Name (UPN) of the user"
$UserPrincipalNameReal = "$Usernameget@contoso.com"
Write-Output "You Entered: $UserPrincipalNameReal"
$user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalNameReal'"

# Store Object ID
$objectId = $user.Id

# Output
Write-Output "Object ID $objectId"

# Connect to Entra Application
$tenantId = ""
$clientId = ""
$clientSecret = ""
$userId = "$objectId"

#Credentials for Entra Application
$body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $clientId
    client_secret = $clientSecret
}

# Get access token for Microsoft Graph API
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
$token = $tokenResponse.access_token

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

# Set Lifetime for the Temporary Access Pass (TAP)
$tapBody = @{
    lifetimeInMinutes = 60
    isUsableOnce = $true
} | ConvertTo-Json


# Create Temporary Access Pass (TAP) for the user
Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$userId/authentication/temporaryAccessPassMethods" -Headers $headers -Body $tapBody