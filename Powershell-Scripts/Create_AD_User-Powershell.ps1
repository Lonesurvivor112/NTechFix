$firstname = Read-Host -Prompt "Enter The First Name"
$lastname = Read-Host -Prompt "Enter The Last Name"
$fullname = "$firstname $lastname"
$samacctname = Read-Host -Prompt "Enter the username"
#$decription = Read-Host -Prompt "Enter The Description"
#$jobtitle = $description
$emaildomain = "contoso.com"
$emailaddy = "$samacctname@$emaildomain"
$oupath = "OU=Frontline-Workers,OU=Users,OU=Contoso Corp,DC=contoso,dc=com"
$realsamacctname = $emailaddy

# Create user
$confirm = Read-Host "Do you want to create user $fullname"
if ($confirm -eq "yes") {
    New-ADUser -DisplayName $fullname -GivenName $firstname -Surname $lastname -EmailAddress $emailaddy -Name $fullname -Path $oupath -SamAccountName $samacctname -Server "DC01.contoso.com" -Type user -UserPrincipalName $emailaddy
    Set-ADUser -Identity $samacctname -SamAccountName $samacctname -Description "DSP" -Title "DSP" -Company "Contoso Corp" -State "Michigan"
}

$securepassword = ConvertTo-SecureString -String Welcome1 -AsPlainText -Force
#Set Pass
Set-ADAccountPassword -Identity $samacctname -NewPassword:$securepassword -Reset:$true -Server:"DC01.contoso.com"
#Enable AD
Enable-ADAccount -Identity $samacctname -Server:"DC01.contoso.com"
Set-ADAccountControl -AccountNotDelegated:$false -AllowReversiblePasswordEncryption:$false -CannotChangePassword:$false -DoesNotRequirePreAuth:$false -Identity $samacctname -PasswordNeverExpires:$false -Server:"DC01.contoso.com" -UseDESKeyOnly:$false
Set-ADUser -ChangePasswordAtLogon:$true -Identity $samacctname -Server:"DC01.contoso.com" -SmartcardLogonRequired:$false

#Add Groupsl
Add-ADPrincipalGroupMembership -Identity $samacctname -MemberOf:"CN=sec-gl-role-dsp,OU=Roles,OU=Contoso Corp,DC=contoso,DC=com"


