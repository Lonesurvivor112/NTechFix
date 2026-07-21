if (Get-Service -Name "CagService" -ErrorAction SilentlyContinue) {
    Write-Output "Installed"
    exit 0
}

exit 1