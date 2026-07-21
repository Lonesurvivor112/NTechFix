# Simple script to get all users who report to a specified manager
# Usage: .\Get-UsersByManager.ps1 "jsmith"

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$ManagerIdentity
)

# Import the Active Directory module
Import-Module ActiveDirectory

# Try to find the manager
try {
    # First attempt to find by SAM account name (most common scenario)
    $manager = Get-ADUser -Identity $ManagerIdentity -Properties DisplayName -ErrorAction SilentlyContinue
    
    # If not found by SAM, try by display name
    if (-not $manager) {
        $manager = Get-ADUser -Filter "Name -eq '$ManagerIdentity'" -Properties DisplayName
        
        # If still not found, exit with error
        if (-not $manager) {
            Write-Host "Manager '$ManagerIdentity' not found in Active Directory." -ForegroundColor Red
            exit
        }
    }
    
    Write-Host "Finding users who report to: $($manager.DisplayName)" -ForegroundColor Cyan
    
    # Find all users who report to the specified manager
    $users = Get-ADUser -Filter "Manager -eq '$($manager.DistinguishedName)'" -Properties DisplayName, Title, Department, EmailAddress
    
    # Check if any users were found
    if ($users.Count -eq 0) {
        Write-Host "No users found reporting to this manager." -ForegroundColor Yellow
        exit
    }
    
    Write-Host "Found $($users.Count) users:" -ForegroundColor Green
    
    # Display the users in a simple table format
    $users | Select-Object DisplayName, SamAccountName, Title, Department, EmailAddress | Format-Table
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}