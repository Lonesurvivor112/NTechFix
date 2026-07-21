#Take Input From User Running
$Computer = Read-Host "Please Enter Computer Name: "
$user = Read-Host "Enter User ID: "
#Invoke runs only on REMOTE Computer
Write-Host "This Script will only work REMOTELY!!"
Invoke-Command -ComputerName $computer -ScriptBlock {
    param($user)
    $localpath = 'c:\users\' + $user
#Finds Local Path of user and Selects it 
    Get-WmiObject -Class Win32_UserProfile | Where-Object {$_.LocalPath -eq $localpath} | 
    Remove-WmiObject
} -ArgumentList $user