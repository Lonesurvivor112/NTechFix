<#
.SYNOPSIS
    Add or remove yourself as Site Collection Administrator on a list of SharePoint sites.
.DESCRIPTION
    Standalone permissions management script. Run with $Mode = "Add" before
    pulling analytics, then re-run with $Mode = "Remove" when you're done.
.NOTES
    Requires: Microsoft.Online.SharePoint.PowerShell module
    Permissions: SharePoint Administrator role
#>

# Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force

# ============================================================
# CONFIG — edit these
# ============================================================
$Mode           = "Add"                                              # "Add" or "Remove"
$AdminUser      = "jdoe@contoso.com"                   # ← your account
$TenantAdminUrl = "https://contoso-admin.sharepoint.com"

# Log file
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir    = "X:\SharePointSCA"
$logFile   = Join-Path $logDir "SCA_${Mode}_$timestamp.csv"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

# ============================================================
# SITE LIST
# ============================================================
$siteUrls = @(
    "https://contoso.sharepoint.com/sites/admin"
    "https://contoso.sharepoint.com/sites/admissions"
    "https://contoso.sharepoint.com/sites/emerson-mgmt"
    "https://contoso.sharepoint.com/sites/appcatalog"
    "https://contoso.sharepoint.com/sites/BA"
    "https://contoso.sharepoint.com/sites/BoardFiles"
    "https://contoso.sharepoint.com/sites/trustees"
    "https://contoso.sharepoint.com/sites/ccm"
    "https://contoso.sharepoint.com/sites/DayProgram"
    "https://contoso.sharepoint.com/sites/DCWECStaffSchedulingAccess"
    "https://contoso.sharepoint.com/sites/dcw"
    "https://contoso.sharepoint.com/sites/DrGutterman'sClinic"
    "https://contoso.sharepoint.com/sites/ECBulletin"
    "https://contoso.sharepoint.com/sites/ECCare"
    "https://contoso.sharepoint.com/sites/knowledge.workers"
    "https://contoso.sharepoint.com/sites/ContosoStaffScheduling3200"
    "https://contoso.sharepoint.com/sites/ContosoUnitStaffSchedulesEditAccess"
    "https://contoso.sharepoint.com/sites/ECUnitStaffSchedulesReadONLY"
    "https://contoso.sharepoint.com/sites/EC-Work-Instructions"
    "https://contoso.sharepoint.com"
    "https://contoso.sharepoint.com/sites/ContosoParkingPolicy"
    "https://contoso.sharepoint.com/sites/EncryptedEmail2"
    "https://contoso.sharepoint.com/sites/FallWinterHolidayCommittee"
    "https://contoso.sharepoint.com/sites/Finance"
    "https://contoso.sharepoint.com/sites/fct2"
    "https://contoso.sharepoint.com/sites/HR-archive"
    "https://contoso.sharepoint.com/sites/HR"
    "https://contoso.sharepoint.com/sites/IT"
    "https://contoso.sharepoint.com/sites/ITHelpGuidesandCybersecurityTraining"
    "https://contoso.sharepoint.com/sites/Kitchen"
    "https://contoso.sharepoint.com/sites/Leadership"
    "https://contoso.sharepoint.com/sites/Maintenance"
    "https://contoso.sharepoint.com/sites/manchester.communication"
    "https://contoso.sharepoint.com/sites/MichiganClinicalDepartment"
    "https://contoso.sharepoint.com/sites/MichiganECDynamic"
    "https://contoso.sharepoint.com/sites/mi-eoc"
    "https://contoso.sharepoint.com/sites/MI-LevelOfCare2"
    "https://contoso.sharepoint.com/sites/mi-moves"
    "https://contoso.sharepoint.com/sites/mi-soc2"
    "https://contoso.sharepoint.com/sites/Nursing"
    "https://contoso.sharepoint.com/sites/o365-enforced-classification"
    "https://contoso.sharepoint.com/sites/OccupationalTherapy802"
    "https://contoso.sharepoint.com/sites/ProjectAnura-PhoneSystemConsolidation"
    "https://contoso.sharepoint.com/sites/pwa"
    "https://contoso.sharepoint.com/sites/qualitydepartment"
    "https://contoso.sharepoint.com/sites/RecTherapy"
    "https://contoso.sharepoint.com/sites/receipts"
    "https://contoso.sharepoint.com/sites/residentialmanagement"
    "https://contoso.sharepoint.com/sites/rcm"
    "https://contoso.sharepoint.com/sites/ECVirtualTraining"
    "https://contoso.sharepoint.com/sites/Security-Testing"
    "https://contoso.sharepoint.com/sites/VOC"
)

# ============================================================
# VALIDATE MODE
# ============================================================
if ($Mode -notin @("Add","Remove")) {
    throw "`$Mode must be 'Add' or 'Remove'. Currently: $Mode"
}

$isAdding = $Mode -eq "Add"
$action   = if ($isAdding) { "Adding" } else { "Removing" }
$flagVal  = $isAdding   # $true to add, $false to remove

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " MODE  : $Mode" -ForegroundColor Cyan
Write-Host " USER  : $AdminUser" -ForegroundColor Cyan
Write-Host " SITES : $($siteUrls.Count)" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Safety prompt for Remove
if (-not $isAdding) {
    $confirm = Read-Host "`nRemove SCA rights from $($siteUrls.Count) sites? Type YES to continue"
    if ($confirm -ne "YES") {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ============================================================
# CONNECT
# ============================================================
Write-Host "`nConnecting to SharePoint Online..." -ForegroundColor Cyan
try {
    Connect-SPOService -Url $TenantAdminUrl -ErrorAction Stop
}
catch {
    Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ============================================================
# RUN
# ============================================================
$counter = 0
$results = foreach ($url in $siteUrls) {
    $counter++
    Write-Progress -Activity "$action Site Collection Administrator" `
                   -Status "$counter of $($siteUrls.Count): $url" `
                   -PercentComplete (($counter / $siteUrls.Count) * 100)

    $status  = "OK"
    $message = ""

    try {
        Set-SPOUser -Site $url -LoginName $AdminUser -IsSiteCollectionAdmin $flagVal -ErrorAction Stop | Out-Null
        if ($isAdding) {
            Write-Host "[+] $url" -ForegroundColor Green
        } else {
            Write-Host "[-] $url" -ForegroundColor Yellow
        }
    }
    catch {
        $status  = "Failed"
        $message = $_.Exception.Message
        Write-Host "[X] $url  --  $message" -ForegroundColor Red
    }

    [PSCustomObject]@{
        SiteUrl   = $url
        Action    = $Mode
        User      = $AdminUser
        Status    = $status
        Message   = $message
        Timestamp = (Get-Date).ToString("u")
    }
}
Write-Progress -Activity "$action Site Collection Administrator" -Completed

# ============================================================
# EXPORT LOG
# ============================================================
$results | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8

$okCount   = ($results | Where-Object Status -eq "OK").Count
$failCount = ($results | Where-Object Status -eq "Failed").Count

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " RESULTS" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Succeeded : $okCount"   -ForegroundColor Green
Write-Host " Failed    : $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Gray" })
Write-Host " Log file  : $logFile"   -ForegroundColor Gray
Write-Host ""

Disconnect-SPOService