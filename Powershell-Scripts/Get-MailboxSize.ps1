# Prompt for username
$username = Read-Host "Enter mailbox username (example: jdoe)"

$domain  = "contoso.com"
$mailbox = "$username@$domain"

Write-Host "Reporting folder sizes for $mailbox" -ForegroundColor Cyan

# Disable broken WAM path just in case
$env:EXO_DISABLE_WAM = "1"

Connect-ExchangeOnline -ShowBanner:$false | Out-Null

# ✅ NO -ResultSize here
$stats = Get-EXOMailboxFolderStatistics -Identity $mailbox

$report = $stats | Where-Object FolderSize | ForEach-Object {

    $sizeString = $_.FolderSize.ToString()

    if ($sizeString -match '([\d\.,]+)\s*(GB|MB|KB|B)') {
        $num  = ($matches[1] -replace ',', '') -as [double]
        $unit = $matches[2]

        $sizeGB = switch ($unit) {
            'GB' { $num }
            'MB' { $num / 1024 }
            'KB' { $num / 1024 / 1024 }
            'B'  { $num / 1024 / 1024 / 1024 }
        }
    } else {
        $sizeGB = 0
    }

    [pscustomobject]@{
        FolderPath = $_.FolderPath
        Items      = $_.ItemsInFolder
        SizeGB     = [math]::Round($sizeGB, 3)
    }
}

$report |
    Sort-Object SizeGB -Descending |
    Format-Table -AutoSize