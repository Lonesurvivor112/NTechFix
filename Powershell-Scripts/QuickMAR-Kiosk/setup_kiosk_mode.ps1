
# Create local user 'medpass' with no password
$Username = "medpass"
$Password = ConvertTo-SecureString "" -AsPlainText -Force
New-LocalUser -Name $Username -Password $Password -FullName "MedCart Kiosk User" -Description "Kiosk user for med passing" -PasswordNeverExpires -UserMayNotChangePassword
Add-LocalGroupMember -Group "Users" -Member $Username

# Enable auto-login for 'medpass'
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty -Path $RegPath -Name "AutoAdminLogon" -Value "1"
Set-ItemProperty -Path $RegPath -Name "DefaultUsername" -Value $Username
Set-ItemProperty -Path $RegPath -Name "DefaultPassword" -Value ""
Set-ItemProperty -Path $RegPath -Name "DefaultDomainName" -Value $env:COMPUTERNAME

# Create Startup folder and shortcut for Quickmar
$StartupPath = "C:\Users\$Username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
New-Item -ItemType Directory -Force -Path $StartupPath
$ShortcutPath = "$StartupPath\Quickmar.lnk"
$TargetPath = "C:\Program Files (x86)\Quickmar\Quickmar.exe"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.Save()

# Hide desktop icons and taskbar
$ExplorerPolicies = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
New-Item -Path $ExplorerPolicies -Force | Out-Null
Set-ItemProperty -Path $ExplorerPolicies -Name "NoDesktop" -Value 1
Set-ItemProperty -Path $ExplorerPolicies -Name "NoTaskbar" -Value 1

Write-Host "✅ Kiosk setup complete. Reboot and log in as 'medpass' to test."
