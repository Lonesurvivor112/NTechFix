# Ensure NuGet provider is available
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false

# Set execution policy for this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force

# Install the Autopilot info script
Install-Script -Name Get-WindowsAutopilotInfo -Force -Confirm:$false

# Register device hash to Autopilot and assign profile
Get-WindowsAutopilotInfo -Online -Assign