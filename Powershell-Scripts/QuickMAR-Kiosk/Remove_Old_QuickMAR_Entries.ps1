
# Remove leftover QuickMAR registry entries
$versionsToRemove = @("5.4.0.2393", "5.4.0.2287")
$uninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($path in $uninstallPaths) {
    Get-ChildItem $path | ForEach-Object {
        $displayName = (Get-ItemProperty $_.PSPath).DisplayName
        $displayVersion = (Get-ItemProperty $_.PSPath).DisplayVersion
        if ($displayName -like "*QuickMAR*" -and $versionsToRemove -contains $displayVersion) {
            Write-Host "Removing $displayName $displayVersion from $path"A
            Remove-Item $_.PSPath -Force
        }
    }
}
