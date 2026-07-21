
# Uninstall existing Microsoft.Graph and PowerShellGet modules
Get-InstalledModule -Name Microsoft.Graph -AllVersions | Uninstall-Module -Force
Get-InstalledModule -Name PowerShellGet -AllVersions | Uninstall-Module -Force

# Install the latest version of PowerShell 7
$installerUrl = "https://github.com/PowerShell/PowerShell/releases/latest/download/PowerShell-7.3.4-win-x64.msi"
$installerPath = "$env:TEMP\PowerShell-7.3.4-win-x64.msi"
Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
Start-Process -FilePath msiexec.exe -ArgumentList "/i", $installerPath, "/quiet", "/norestart" -Wait

# Restart PowerShell session
Write-Host "PowerShell 7 installed. Please restart your PowerShell session."

# Install Microsoft.Graph module
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Import Microsoft.Graph module and select beta profile
Import-Module Microsoft.Graph
Select-MgProfile -Name "beta"

# Connect to Microsoft Graph with required scopes
Connect-MgGraph -Scopes "Device.Read.All", "Directory.Read.All", "DeviceLocalCredential.Read.All"

Write-Host "Environment setup complete. You can now use Microsoft Graph SDK with LAPS support."
