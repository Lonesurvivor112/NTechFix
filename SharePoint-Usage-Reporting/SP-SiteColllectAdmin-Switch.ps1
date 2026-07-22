<#
.SYNOPSIS
    Add or remove a user as Site Collection Administrator on a list of SharePoint sites.

.DESCRIPTION
    Standalone SharePoint Online permissions management script.

    Run with $Mode = "Add" before collecting analytics or performing administrative
    reporting tasks, then re-run with $Mode = "Remove" when finished.

.NOTES
    Requires:
        Microsoft.Online.SharePoint.PowerShell module

    Permissions:
        SharePoint Administrator role

    Install module if needed:
        Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
#>

# ============================================================
# CONFIGURATION
# ============================================================

# Choose whether to add or remove Site Collection Administrator rights.
# Valid values: "Add" or "Remove"
$Mode = "Add"

# Account to add/remove as Site Collection Administrator.
# Replace with your admin account.
$AdminUser = "admin@contoso.com"

# SharePoint Online tenant admin URL.
# Replace "contoso" with your tenant name.
$TenantAdminUrl = "https://contoso-admin.sharepoint.com"

# Log directory.
# Update this path as needed.
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir    = "C:\Logs\SharePointSCA"
$logFile   = Join-Path $logDir "SCA_${Mode}_$timestamp.csv"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# ============================================================
# SITE LIST
# ============================================================

# Replace these example URLs with your SharePoint site URLs.
$siteUrls = @(
    "https://contoso.sharepoint.com"
    "https://contoso.sharepoint.com/sites/Administration"
    "https://contoso.sharepoint.com/sites/Admissions"
    "https://contoso.sharepoint.com/sites/Management"
    "https://contoso.sharepoint.com/sites/AppCatalog"
    "https://contoso.sharepoint.com/sites/BusinessAnalytics"
    "https://contoso.sharepoint.com/sites/BoardDocuments"
    "https://contoso.sharepoint.com/sites/BoardOfDirectors"
    "https://contoso.sharepoint.com/sites/CareManagement"
    "https://contoso.sharepoint.com/sites/DayPrograms"
    "https://contoso.sharepoint.com/sites/StaffScheduling"
    "https://contoso.sharepoint.com/sites/Operations"
    "https://contoso.sharepoint.com/sites/SpecialtyClinic"
    "https://contoso.sharepoint.com/sites/CompanyNews"
    "https://contoso.sharepoint.com/sites/PatientCare"
    "https://contoso.sharepoint.com/sites/KnowledgeBase"
    "https://contoso.sharepoint.com/sites/WorkSchedules"
    "https://contoso.sharepoint.com/sites/SchedulingEditors"
    "https://contoso.sharepoint.com/sites/SchedulingReadOnly"
    "https://contoso.sharepoint.com/sites/WorkInstructions"
    "https://contoso.sharepoint.com/sites/ParkingPolicies"
    "https://contoso.sharepoint.com/sites/SecureEmail"
    "https://contoso.sharepoint.com/sites/EmployeeCommittee"
    "https://contoso.sharepoint.com/sites/Finance"
    "https://contoso.sharepoint.com/sites/Compliance"
    "https://contoso.sharepoint.com/sites/HRArchive"
    "https://contoso.sharepoint.com/sites/HumanResources"
    "https://contoso.sharepoint.com/sites/InformationTechnology"
    "https://contoso.sharepoint.com/sites/ITTraining"
    "https://contoso.sharepoint.com/sites/FoodServices"
    "https://contoso.sharepoint.com/sites/Leadership"
    "https://contoso.sharepoint.com/sites/Facilities"
    "https://contoso.sharepoint.com/sites/Communications"
    "https://contoso.sharepoint.com/sites/ClinicalServices"
    "https://contoso.sharepoint.com/sites/DynamicForms"
    "https://contoso.sharepoint.com/sites/EmergencyOperations"
    "https://contoso.sharepoint.com/sites/LevelOfCare"
    "https://contoso.sharepoint.com/sites/CaseManagement"
    "https://contoso.sharepoint.com/sites/SocialServices"
    "https://contoso.sharepoint.com/sites/Nursing"
    "https://contoso.sharepoint.com/sites/DataClassification"
    "https://contoso.sharepoint.com/sites/OccupationalTherapy"
    "https://contoso.sharepoint.com/sites/PhoneSystemMigration"
    "https://contoso.sharepoint.com/sites/ProjectManagement"
    "https://contoso.sharepoint.com/sites/Quality"
    "https://contoso.sharepoint.com/sites/RecreationalTherapy"
    "https://contoso.sharepoint.com/sites/Receipts"
    "https://contoso.sharepoint.com/sites/ResidentialPrograms"
    "https://contoso.sharepoint.com/sites/RevenueCycle"
    "https://contoso.sharepoint.com/sites/VirtualTraining"
    "https://contoso.sharepoint.com/sites/SecurityTesting"
    "https://contoso.sharepoint.com/sites/VocationalServices"
)

# ============================================================
# VALIDATE MODE
# ============================================================

if ($Mode -notin @("Add", "Remove")) {
    throw "`$Mode must be 'Add' or 'Remove'. Current value: $Mode"
}

$isAdding = $Mode -eq "Add"
$action   = if ($isAdding) { "Adding" } else { "Removing" }
$flagVal  = $isAdding

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " SHAREPOINT SITE COLLECTION ADMIN TOOL" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " MODE  : $Mode" -ForegroundColor Cyan
Write-Host " USER  : $AdminUser" -ForegroundColor Cyan
Write-Host " SITES : $($siteUrls.Count)" -ForegroundColor Cyan
Write-Host " LOG   : $logFile" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# ============================================================
# SAFETY PROMPT FOR REMOVE MODE
# ============================================================

if (-not $isAdding) {
    $confirm = Read-Host "`nRemove Site Collection Administrator rights from $($siteUrls.Count) sites? Type YES to continue"

    if ($confirm -ne "YES") {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ============================================================
# CONNECT TO SHAREPOINT ONLINE
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
# PROCESS SITES
# ============================================================

$counter = 0

$results = foreach ($url in $siteUrls) {
    $counter++

    Write-Progress `
        -Activity "$action Site Collection Administrator" `
        -Status "$counter of $($siteUrls.Count): $url" `
        -PercentComplete (($counter / $siteUrls.Count) * 100)

    $status  = "OK"
    $message = ""

    try {
        Set-SPOUser `
            -Site $url `
            -LoginName $AdminUser `
            -IsSiteCollectionAdmin $flagVal `
            -ErrorAction Stop | Out-Null

        if ($isAdding) {
            Write-Host "[+] $url" -ForegroundColor Green
        }
        else {
            Write-Host "[-] $url" -ForegroundColor Yellow
        }
    }
    catch {
        $status  = "Failed"
        $message = $_.Exception.Message

        Write-Host "[X] $url -- $message" -ForegroundColor Red
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

$okCount   = ($results | Where-Object { $_.Status -eq "OK" }).Count
$failCount = ($results | Where-Object { $_.Status -eq "Failed" }).Count

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " RESULTS" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Succeeded : $okCount" -ForegroundColor Green

if ($failCount -gt 0) {
    Write-Host " Failed    : $failCount" -ForegroundColor Red
}
else {
    Write-Host " Failed    : $failCount" -ForegroundColor Gray
}

Write-Host " Log file  : $logFile" -ForegroundColor Gray
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# DISCONNECT
# ============================================================

Disconnect-SPOService