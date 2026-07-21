
Import-Module ActiveDirectory

$computerName = Read-Host "Enter the computer name"

# Get all recovery objects and filter by computer name
$recoveryObjects = Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -Properties * |
    Where-Object { $_.DistinguishedName -like "*$computerName*" }

# Output all properties
#$recoveryObjects | Format-List

$recoveryObjects | Select-Object `
    @{Name='ComputerName';Expression={$computerName}},
    @{Name='RecoveryKeyID';Expression={[System.Guid]::new($_.'msFVE-RecoveryGuid')}},
    @{Name='RecoveryPassword';Expression={$_.'msFVE-RecoveryPassword'}},
    @{Name='DateAdded';Expression={$_.WhenCreated}}
