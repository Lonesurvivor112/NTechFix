#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports AutoSubscribeNewMembers status and member/subscriber counts for all Unified Groups.

.EXAMPLE
    .\Get-UnifiedGroupAutoSubscribeStatus.ps1

.EXAMPLE
    .\Get-UnifiedGroupAutoSubscribeStatus.ps1 -ExportCsv "C:\Reports\GroupAudit.csv"
#>

param (
    [string]$ExportCsv
)

# Verify Exchange Online session
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
} catch {
    Write-Error "Not connected to Exchange Online. Run Connect-ExchangeOnline first."
    exit 1
}

Write-Host "`n[*] Retrieving all Unified Groups ..." -ForegroundColor Cyan
$allGroups = Get-UnifiedGroup -ResultSize Unlimited -IncludeAllProperties
Write-Host "    Found $($allGroups.Count) group(s).`n" -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($group in $allGroups) {

    $members     = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members     -ResultSize Unlimited
    $subscribers = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Subscribers -ResultSize Unlimited

    $memCount = ($members    | Measure-Object).Count
    $subCount = ($subscribers | Measure-Object).Count

    $result = [PSCustomObject]@{
        DisplayName             = $group.DisplayName
        PrimarySmtpAddress      = $group.PrimarySmtpAddress
        AutoSubscribeNewMembers = $group.AutoSubscribeNewMembers
        MemberCount             = $memCount
        SubscriberCount         = $subCount
        CountMismatch           = $memCount -ne $subCount
    }

    $results.Add($result)

    $color = if ($group.AutoSubscribeNewMembers) { "Green" } else { "Yellow" }
    $autoSubLabel = if ($group.AutoSubscribeNewMembers) { "AutoSub: YES" } else { "AutoSub: NO " }
    $mismatchLabel = if ($memCount -ne $subCount) { " | ⚠ Count mismatch (Members: $memCount, Subscribers: $subCount)" } else { "" }

    Write-Host "  [$autoSubLabel]  $($group.DisplayName)$mismatchLabel" -ForegroundColor $color
}

# Summary
$autoSubOn  = ($results | Where-Object { $_.AutoSubscribeNewMembers }).Count
$autoSubOff = ($results | Where-Object { -not $_.AutoSubscribeNewMembers }).Count
$mismatches = ($results | Where-Object { $_.CountMismatch }).Count

Write-Host "`n──────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "  Total groups              : $($results.Count)"
Write-Host "  AutoSubscribe ENABLED     : $autoSubOn"  -ForegroundColor Green
Write-Host "  AutoSubscribe DISABLED    : $autoSubOff" -ForegroundColor Yellow
Write-Host "  Member/Subscriber mismatch: $mismatches" -ForegroundColor $(if ($mismatches -gt 0) { "Yellow" } else { "Green" })
Write-Host "──────────────────────────────────────────────`n" -ForegroundColor Cyan

if ($ExportCsv) {
    $results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "[*] Exported to: $ExportCsv`n" -ForegroundColor Cyan
}

# Display as table in console
$results | Sort-Object AutoSubscribeNewMembers | Format-Table -AutoSize