
function StartAdSync {
    Start-Job -ScriptBlock {
        Invoke-Command -ComputerName computerdc01 -ScriptBlock {
            Start-ADSyncSyncCycle -PolicyType Delta
        }
    }
}

Set-Alias -Name deltasync -Value StartAdSync

$text = "Starting The Party, Running AD Sync"
$colors = @("Red", "Yellow", "Green", "Cyan", "Blue", "Magenta")
$startTime = Get-Date
$duration = New-TimeSpan -Seconds 5

# Save cursor position to restore it later
$originalPos = $host.UI.RawUI.CursorPosition

# Start the sync in the background
StartAdSync

# Rainbow animation
while ((Get-Date) -lt ($startTime + $duration)) {
    $firstColor = $colors[0]
    $colors = $colors[1..($colors.Length - 1)] + $firstColor
    $host.UI.RawUI.CursorPosition = $originalPos

    for ($i = 0; $i -lt $text.Length; $i++) {
        $colorIndex = $i % $colors.Count
        Write-Host $text[$i] -ForegroundColor $colors[$colorIndex] -NoNewline
    }

    Start-Sleep -Milliseconds 100
}

Write-Host "Waiting for AD Sync to complete..." -ForegroundColor Cyan

# Spinner animation while waiting for the job
$spinner = @("|", "/", "-", "\")
$spinnerIndex = 0

while ((Get-Job | Where-Object { $_.State -eq 'Running' })) {
    Write-Host -NoNewline "`r$($spinner[$spinnerIndex]) Sync in progress..."
    Start-Sleep -Milliseconds 200
    $spinnerIndex = ($spinnerIndex + 1) % $spinner.Length
}

# Final message
Write-Host "✔ AD Delta Sync completed." -ForegroundColor Green

# Clean up the job
Get-Job | Remove-Job
