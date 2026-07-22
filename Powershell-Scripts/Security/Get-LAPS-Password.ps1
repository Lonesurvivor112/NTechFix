
$tenantId = ""
$clientId = ""
$clientSecret = ""

$body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $clientId
    client_secret = $clientSecret
}

$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
$accessToken = $tokenResponse.access_token


$deviceName = Read-Host "Enter the computer name"
$headers = @{ Authorization = "Bearer $accessToken" }

# Properly encode the filter query
$filter = "displayName eq '$deviceName'"
$encodedFilter = [System.Uri]::EscapeDataString($filter)

# Correct way to construct the URL (no backslash escaping)

$deviceUrl = 'https://graph.microsoft.com/beta/devices?$filter=' + $encodedFilter


# Make the request
$deviceResponse = Invoke-RestMethod -Uri $deviceUrl -Headers $headers -Method Get

# Extract device ID
if ($deviceResponse.value.Count -gt 0) {
    $deviceId = $deviceResponse.value[0].id
    Write-Host "Device ID: $deviceId"
} else {
    Write-Host "No device found with the name '$deviceName'"
}


$headers = @{
    Authorization = "Bearer $accessToken"
    "Content-Type" = "application/json"
}


# Construct the LAPS password URL
$lapsUrl = "https://graph.microsoft.com/beta/devices/$deviceId/getLocalAdminPassword"

# Make the POST request

$lapsResponse = Invoke-RestMethod -Uri $lapsUrl -Headers $headers -Method Post -Body '{}' -ContentType 'application/json'


# Display the password info
if ($lapsResponse.value) {
    foreach ($cred in $lapsResponse.value) {
        Write-Host "Account Name: $($cred.accountName)"
        Write-Host "Password: $($cred.password)"
        Write-Host "Created: $($cred.createdDateTime)"
        Write-Host "-----------------------------"
    }
} else {
    Write-Host "No LAPS credentials found for this device."
}
