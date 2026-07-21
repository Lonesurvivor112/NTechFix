
# Ensure Microsoft Graph module is installed
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Connect to Microsoft Graph with required permissions
Connect-MgGraph -Scopes "Device.Read.All", "Directory.Read.All"

# Prompt for device name
$deviceName = Read-Host "Enter the device name"

# Get the device objects by name
$devices = Get-MgDevice -Filter "displayName eq '$deviceName'"

if ($devices) {
    foreach ($device in $devices) {
        $deviceId = $device.Id
        Write-Host "Attempting to retrieve LAPS password for Device ID: $deviceId"

        # Attempt to retrieve the LAPS password
        try {
            $lapsPassword = Get-MgDeviceLocalAdminPassword -DeviceId $deviceId
            Write-Host "LAPS Password for Device ID $deviceId: $($lapsPassword.SecretText)"
        } catch {
            Write-Host "Failed to retrieve LAPS password for Device ID $deviceId. Ensure you have the necessary permissions and the device is LAPS-enabled."
        }
    }
} else {
    Write-Host "Device not found. Please check the name and try again."
}
