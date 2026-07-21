$guid = "c4e9a55a-1564-495f-9805-afac5c2b70a9"
$path = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"

if (-not (Test-Path $path)) {
    New-Item -Path $path -Force:$false | Out-Null
}

$current = (Get-ItemProperty -Path $path -Name "AADJMachineDomainGuid" -ErrorAction SilentlyContinue).AADJMachineDomainGuid

if ($current -ne $guid) {
    Set-ItemProperty -Path $path -Name "AADJMachineDomainGuid" -Value $guid -Type String
}