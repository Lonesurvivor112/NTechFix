# ============================================================
# CreateShortcut.ps1
# Creates Safety Zone shortcut on the Public Desktop
# Intended for Intune Win32 App deployment
# ============================================================

$ErrorActionPreference = "Stop"

# App details
$ShortcutName = "Safety Zone.lnk"
$SafetyZoneUrl = "https://events.healthcaresafetyzone.com/EventsModuleWeb/default.aspx?cs=77666c71-e7b6-4126-a0fd-7b2c3ec34552"

# Paths
$PublicDesktop = "C:\Users\Public\Desktop"
$ShortcutPath = Join-Path $PublicDesktop $ShortcutName

$CompanyFolder = "C:\ProgramData\ContosoCorp"
$IconFolder = Join-Path $CompanyFolder "Icons"
$LogFolder = Join-Path $CompanyFolder "Logs"
$LogPath = Join-Path $LogFolder "SafetyZoneShortcut-Install.log"

$PackageIconPath = Join-Path $PSScriptRoot "SafetyZone.ico"
$LocalIconPath = Join-Path $IconFolder "SafetyZone.ico"

# Create required folders
New-Item -ItemType Directory -Path $PublicDesktop -Force | Out-Null
New-Item -ItemType Directory -Path $IconFolder -Force | Out-Null
New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null

Start-Transcript -Path $LogPath -Append

try {
    Write-Output "Starting Safety Zone shortcut creation..."

    # Find Microsoft Edge
    $EdgePathX86 = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    $EdgePathX64 = "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"

    if (Test-Path $EdgePathX86) {
        $EdgePath = $EdgePathX86
    }
    elseif (Test-Path $EdgePathX64) {
        $EdgePath = $EdgePathX64
    }
    else {
        throw "Microsoft Edge executable was not found."
    }

    Write-Output "Edge path detected: $EdgePath"

    # Copy custom icon if included in Intune package
    if (Test-Path $PackageIconPath) {
        Copy-Item -Path $PackageIconPath -Destination $LocalIconPath -Force
        $IconLocation = "$LocalIconPath,0"
        Write-Output "Custom icon copied to: $LocalIconPath"
    }
    else {
        $IconLocation = "$EdgePath,0"
        Write-Output "Custom icon not found in package. Using Edge icon."
    }

    # Remove old shortcut if it already exists
    if (Test-Path $ShortcutPath) {
        Remove-Item -Path $ShortcutPath -Force
        Write-Output "Existing shortcut removed: $ShortcutPath"
    }

    # Create shortcut
    $Shell = New-Object -ComObject WScript.Shell
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $EdgePath
    $Shortcut.Arguments = $SafetyZoneUrl
    $Shortcut.WorkingDirectory = Split-Path $EdgePath
    $Shortcut.IconLocation = $IconLocation
    $Shortcut.Description = "Safety Zone"
    $Shortcut.Save()

    Write-Output "Shortcut created successfully at: $ShortcutPath"

    # Verify shortcut exists
    if (Test-Path $ShortcutPath) {
        Write-Output "Detection check passed. Shortcut exists."
        exit 0
    }
    else {
        throw "Shortcut was not created successfully."
    }
}
catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}
finally {
    Stop-Transcript
}
