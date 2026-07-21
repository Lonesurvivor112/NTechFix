# Prompt for user details
$firstname = Read-Host -Prompt "Enter the First Name"
$lastname = Read-Host -Prompt "Enter the Last Name"
$fullname = "$firstname $lastname"
$samacctname = Read-Host -Prompt "Enter the username"
$jobtitle = Read-Host -Prompt "Enter the Job Title"
$manager = Read-Host -Prompt "Enter the Department Manager's username (e.g., jdoe)"

# Prompt for group name with fallback
$groupname = Read-Host -Prompt "Enter the group to add (leave blank for default 'sec-gl-role-dsp')"
if ([string]::IsNullOrWhiteSpace($groupname)) {
    $groupname = "sec-gl-role-dsp"
}
$groupdn = "CN=$groupname,OU=Roles,OU=Contoso Corp,DC=contoso,DC=com"

# Prompt for OU under Users
$subou = Read-Host -Prompt "Enter the sub-OU under Users (leave blank for default 'Frontline-Workers')"
if ([string]::IsNullOrWhiteSpace($subou)) {
    $subou = "Frontline-Workers"
}
$oupath = "OU=$subou,OU=Users,OU=Contoso Corp,DC=contoso,DC=com"

# Email setup
$emaildomain = "contoso.com"
$emailaddy = "$samacctname@$emaildomain"

# Confirm creation
$confirm = Read-Host "Do you want to create user $fullname? (yes/no)"
if ($confirm -eq "yes") {
    New-ADUser -DisplayName $fullname `
               -GivenName $firstname `
               -Surname $lastname `
               -EmailAddress $emailaddy `
               -Name $fullname `
               -Path $oupath `
               -SamAccountName $samacctname `
               -Server "DC01.contoso.com" `
               -Type user `
               -UserPrincipalName $emailaddy

    Set-ADUser -Identity $samacctname `
               -SamAccountName $samacctname `
               -Description $jobtitle `
               -Title $jobtitle `
               -Company "Contoso Corp" `
               -State "Michigan" `
               -Manager $manager
}


#Set Pass

# Set password and enable account
$securepassword = ConvertTo-SecureString -String Welcome1 -AsPlainText -Force
Set-ADAccountPassword -Identity $samacctname -NewPassword:$securepassword -Reset:$true -Server:"DC01.contoso.com"
Enable-ADAccount -Identity $samacctname -Server "DC01.contoso.com"
Set-ADAccountControl -Identity $samacctname `
                     -AccountNotDelegated $false `
                     -AllowReversiblePasswordEncryption $false `
                     -CannotChangePassword $false `
                     -DoesNotRequirePreAuth $false `
                     -PasswordNeverExpires $false `
                     -UseDESKeyOnly $false `
                     -Server "DC01.contoso.com"
Set-ADUser -Identity $samacctname `
           -ChangePasswordAtLogon $true `
           -SmartcardLogonRequired $false `
           -Server "DC01.contoso.com"

# Add to group
Add-ADPrincipalGroupMembership -Identity $samacctname -MemberOf $groupdn
