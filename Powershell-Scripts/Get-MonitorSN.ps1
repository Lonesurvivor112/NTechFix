function Decode {param ([byte[]]$Data)
    [System.Text.Encoding]::ASCII.GetString($Data)}
 
$Monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi
 
foreach ($Monitor in $Monitors) {$SerialNumber = Decode $Monitor.SerialNumberID
    Write-Output "Monitor Serial Number: $SerialNumber"}
