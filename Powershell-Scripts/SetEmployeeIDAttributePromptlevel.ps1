$user = Read-Host -Prompt "Enter the username"

# Check Current EmployeeID
$confirm = Read-Host "Do you want to check the current EmployeeID? (yes/no)"
if ($confirm -eq "yes") {
    Get-ADUser -Identity $user -Properties EmployeeID
}

# Clear EmployeeID so you can enter a new one
$confirm = Read-Host "Do you want to clear the EmployeeID? (yes/no)"
if ($confirm -eq "yes") {
    Set-ADUser -Identity $user -Clear EmployeeID
}

# Set the EmployeeID attribute for the specified user
$confirm = Read-Host "Do you want to set a new EmployeeID for $user? (yes/no)"
if ($confirm -eq "yes") {
    $EmployeeID = Read-Host "Enter the new EmployeeID"
    Set-ADUser -Identity $user -Add @{"EmployeeID"="$EmployeeID"}
}
