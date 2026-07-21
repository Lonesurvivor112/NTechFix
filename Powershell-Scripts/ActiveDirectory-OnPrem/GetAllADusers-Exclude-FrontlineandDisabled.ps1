# Define the OUs to exclude
#$excludedOUs = @(
#    "OU=No-Longer-Employed-By-CC,OU=Users,OU=Contoso Corp,DC=contoso,DC=com",
#    "OU=Frontline-Workers,OU=Users,OU=Contoso Corp,DC=contoso,DC=com"

# Define the OU to exclude
$excludedOU = "OU=333,OU=Users,OU=Contoso Corp,DC=contoso,DC=com"

# Get all enabled users and exclude those in the specified OU
Get-ADUser -Filter 'Enabled -eq $true' -Properties DistinguishedName, description |
Where-Object {
    $_.DistinguishedName -notlike "*$excludedOU*"
} |
Select-Object Name, SamAccountName, Enabled, DistinguishedName
