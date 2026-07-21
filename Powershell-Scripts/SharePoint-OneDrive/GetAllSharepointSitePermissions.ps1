#Connect to SharePoint 
Connect-SPOService -Url https://contoso.sharepoint.com 
# Set the user's email address 
$userEmail = "jdoe@contoso.com" 
# Get all site collections in the tenant 
$sites = Get-SPOSite -Limit All 
# Loop through each collection and check if the user has access foreach 
($site in $sites) {     
# Get the user's permissions for the site collection     
$permissions = Get-SPOUser -Site $site.Url -LoginName $userEmail -ErrorAction SilentlyContinue     if ($permissions) {         
    # If the user has permissions, output the site collection URL         
Write-Host "User $($userEmail) has access to $($site.Url)"     } }