# Paths and Task Details
$updatePath = "C:\Program Files (x86)\QuickMAR\Updates\QuickMarUpdate_5_5_0_2552.exe"
$monitorScriptPath = "C:\Scripts\Run-MonitorQuickmar.ps1"
$taskName = "MonitorQuickmar"

# Stop QuickMAR
function Stop-QuickMAR {
    $process = Get-Process -Name "QuickMAR" -ErrorAction SilentlyContinue
    if ($process) {
        Write-Output "Stopping QuickMAR..."
        Stop-Process -Name "QuickMAR" -Force
    } else {
        Write-Output "QuickMAR is not running."
    }
}

# Disable Task Scheduler job
function Disable-QuickMARTask {
    Write-Output "Disabling Task Scheduler job: $taskName"
    schtasks /Change /TN "$taskName" /DISABLE
}
# Stop monitoring PowerShell script
function Stop-MonitorScript {
    $monitorProc = Get-CimInstance Win32_Process | Where-Object {
        $_.Name -eq "powershell.exe" -and $_.CommandLine -match 'monitorquickmar\.ps1'
    }
    if ($monitorProc) {
        Write-Output "Stopping MonitorQuickmar PowerShell script..."
        foreach ($proc in $monitorProc) {
            Stop-Process -Id $proc.ProcessId -Force
        }
    } else {
        Write-Output "MonitorQuickmar script not running."
    }
}

# Main update logic
if (Test-Path $updatePath) {
    Write-Output "QuickMAR update found. Preparing for update..."

    Stop-QuickMAR
    Stop-MonitorScript
    Disable-QuickMARTask

    Write-Output "Starting silent update..."
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $updatePath
    $processInfo.Arguments = "/quiet /norestart"
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true
    $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

    $process = [System.Diagnostics.Process]::Start($processInfo)
    $process.WaitForExit()

    Write-Output "QuickMAR update completed."

# Enable Task Scheduler job
function Enable-QuickMARTask {
    Write-Output "Re-enabling Task Scheduler job: $taskName"
    schtasks /Change /TN "$taskName" /ENABLE
}

    Enable-QuickMARTask


    # Restart QuickMAR
    Start-Process "C:\Program Files (x86)\QuickMAR\QuickMAR.exe"
    Write-Output "QuickMAR restarted."

    # Prompt for restart
    Write-Host "It is recommended you restart the computer for the Task Scheduler monitor to rerun."
    $restart = Read-Host "Restart now? (Yes/No)"
    if ($restart -match '^(Y|Yes)$') {
        Write-Output "Restarting computer..."
        Restart-Computer -Force
    } else {
        Write-Output "Restart skipped. Please restart manually later."
    }
} else {
    Write-Output "Error: QuickMAR update file not found at: $updatePath"
}
