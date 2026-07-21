# create-user.ps1
# script to create a user in AD and license appropriately

function New-ECUser {
    param(
        [Parameter(Mandatory=$true)]
        [string] $GivenName,
        [Parameter(Mandatory=$true)]
        [string] $Surname,
        [Parameter(Mandatory=$true)]
        [string] $userType,
        [string] $Department,
        [string] $ManagerSAMAccountName,
        [string] $EmployeeID,
        [string] $365LicenseOverride
    )
    
    $Name = "$GivenName $Surname"
    $sAMAccountName = New-SAMAccountName $GivenName $Surname
}

function New-SAMAccountName {
    param (
        [string] $Given, 
        [string] $Sur
    )
    # construct initial SAM name
    
    $initialSAM = $Given.ToLower()[0]+$Sur.ToLower()
    
    # does the sAM already exist?
    if ((Test-UniqueSAM $initialSAM)) {
        $initialSAM
    } else {
            # try appending a 2
            $suffixNum = 2
            $uniqueSAM = $false

            while (-not $uniqueSAM) {
                $TestSAM = $initialSAM + $suffixNum
                if ((Test-UniqueSAM $TestSAM)) {
                    $uniqueSAM =  $true
                }
                $suffixNum++
            }
            $TestSAM
        }
}

# Return true if input is unused sAM
function Test-UniqueSAM {
    param(
      [Parameter(Mandatory)]
      [String]
      $sAMAccountName
    )
    $null -eq ([ADSISearcher] "(sAMAccountName=$sAMAccountName)").FindOne()
}