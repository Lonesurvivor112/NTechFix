while ($true) {
    $process = Get-Process -Name "QuickMAR" -ErrorAction SilentlyContinue
    if (-not $process) {
        Start-Process "C:\Program Files (x86)\QuickMAR\QuickMAR.exe"
        Start-Sleep -Seconds 10
    }
    Start-Sleep -Seconds 5
}
