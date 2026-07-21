# Microsoft 365 Guest User Management Script
# This script helps administrators view and manage guest users in a Microsoft 365 tenant

# Prerequisites:
# 1. You must have the Microsoft Graph PowerShell SDK installed
# 2. You must have sufficient permissions (Global Admin or User Admin)

#TO GET STARTED RUN AND DEBUG THE SCRIPT

#Main Menu   
function Show-Menu {
    Write-Host "`n===== Microsoft 365 Guest User Management =====" -ForegroundColor Cyan
    Write-Host "Get-AllGuestUsers: View all guest users" -ForegroundColor Green
    Write-Host "Get-GuestUsersByDate: Display guest users with create date" -ForegroundColor Blue
    Write-Host "Export-GuestUsersToCSV: Export guest users to CSV" -ForegroundColor White
    Write-Host "Remove-MultipleGuestUsers: Delete multiple guest users" -ForegroundColor Yellow
    Write-Host "CTRL+C: Exit the script" -ForegroundColor Red
    Write-Host "Show-Menu" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
}
Write-host "A List of your Options:" 
Write-Host "`n===== Microsoft 365 Guest User Management =====" -ForegroundColor Cyan
Write-Host "Get-AllGuestUsers: View all guest users" -ForegroundColor Green
Write-Host "Get-GuestUsersByDate: Display guest users with create date" -ForegroundColor Blue
Write-Host "Export-GuestUsersToCSV: Export guest users to CSV" -ForegroundColor White
Write-Host "Remove-MultipleGuestUsers: Delete multiple guest users" -ForegroundColor Yellow
Write-Host "CTRL+C: Exit the script" -ForegroundColor Red
Write-Host "Show-Menu: Shows the Menu" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# Install required modules if not already installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Host "Installing Microsoft Graph Users module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Force
}

# Connect to Microsoft Graph with appropriate permissions
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -NoWelcome
Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green

# Function to get all guest users
function Get-AllGuestUsers {
    Write-Host "Retrieving all guest users..." -ForegroundColor Cyan
    
    $guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -All -Property Id, DisplayName, Mail, UserPrincipalName, CreatedDateTime, UserType
    
    if ($guestUsers.Count -eq 0) {
        Write-Host "No guest users found in the tenant." -ForegroundColor Yellow
        return $null
    }
    
    Write-Host "Found $($guestUsers.Count) guest users." -ForegroundColor Green
    return $guestUsers
}

## Function to export guest users to CSV
Function Export-GuestUsersToCSV {
    param(
        [parameter(Mandatory = $true)]
        [string]$ExportPath
        )
Get-AllGuestUsers | Select-Object createddatetime, Id, Displayname, Mail, UserPrinciplalName | export-csv -Path $ExportPath -NoTypeInformation
}
# Function to remove a guest user
function Remove-GuestUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    
    try {
        Remove-MgUser -UserId $UserId
        Write-Host "User with ID $UserId has been deleted successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to delete user with ID $UserId. Error: $_" -ForegroundColor Red
        return $false
    }
}
# Function to delete multiple guest users
function Remove-MultipleGuestUsers {
    param (
        [Parameter(Mandatory = $true)]
        [array]$UserIds
    )
    
    $successCount = 0
    $failCount = 0
    
    foreach ($userId in $UserIds) {
        $result = Remove-GuestUser -UserId $userId
        if ($result) {
            $successCount++
        }
        else {
            $failCount++
        }
    }
    
    Write-Host "Deletion complete: $successCount users deleted successfully, $failCount failures." -ForegroundColor Cyan
}

# Function to filter guest users by creation date
function Get-GuestUsersByDate {
    param (
        [Parameter(Mandatory = $true)]
        [array]$GuestUsers,
        
        [Parameter(Mandatory = $true)]
        [DateTime]$OlderThan
    )
    
    $filteredUsers = $GuestUsers | Where-Object { $_.CreatedDateTime -lt $OlderThan }
    
    Write-Host "Found $($filteredUsers.Count) guest users created before $($OlderThan.ToString('yyyy-MM-dd'))." -ForegroundColor Cyan
    
    return $filteredUsers
}

   #!!! # Always disconnect from Microsoft Graph when done #!!!
    #Disconnect-MgGraph | Out-Null
    #Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
