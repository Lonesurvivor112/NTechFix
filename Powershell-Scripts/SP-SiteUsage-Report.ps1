<#
.SYNOPSIS
    Pulls SharePoint site usage analytics for a specific list of sites.
.NOTES
    Requires: Microsoft.Graph PowerShell module
    Permissions: Sites.Read.All
#>

# Install-Module Microsoft.Graph -Scope CurrentUser -Force

Import-Module Microsoft.Graph.Sites

Connect-MgGraph -Scopes "Sites.Read.All"

# --- CONFIG ---
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

$daysBack    = 30
$timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputDir   = "X:\SharePointUsage"
$detailCsv   = Join-Path $outputDir "SiteUsage_Detail_$timestamp.csv"
$summaryCsv  = Join-Path $outputDir "SiteUsage_Summary_$timestamp.csv"

if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$startDate = (Get-Date).AddDays(-$daysBack).ToString("yyyy-MM-dd")
$endDate   = (Get-Date).ToString("yyyy-MM-dd")

Write-Host "Pulling analytics for $($siteUrls.Count) sites ($startDate to $endDate)..." -ForegroundColor Cyan

$counter = 0
$results = foreach ($siteUrl in $siteUrls) {
    $counter++
    Write-Progress -Activity "Pulling site analytics" `
                   -Status "$counter of $($siteUrls.Count): $siteUrl" `
                   -PercentComplete (($counter / $siteUrls.Count) * 100)

    $uniqueViewers = $null
    $totalVisits   = $null
    $totalHits     = $null
    $avgTimeSpent  = $null
    $avgSecondsRaw = $null
    $status        = "OK"
    $siteId        = $null

    try {
        # Convert URL to Graph site ID lookup format
        $uri      = [uri]$siteUrl
        $hostname = $uri.Host
        $path     = $uri.AbsolutePath.TrimEnd('/')

        $lookupUri = if ([string]::IsNullOrWhiteSpace($path)) {
            "https://graph.microsoft.com/v1.0/sites/$hostname"
        } else {
            "https://graph.microsoft.com/v1.0/sites/${hostname}:${path}"
        }

        $site   = Invoke-MgGraphRequest -Method GET -Uri $lookupUri -ErrorAction Stop
        $siteId = $site.id
    }
    catch {
        $status = "Site lookup failed: $($_.Exception.Message -replace '\s+',' ' | Select-Object -First 200)"
    }

    if ($siteId) {
        try {
            $statsUri = "https://graph.microsoft.com/beta/sites/$siteId/getActivityStats(startDateTime=$startDate,endDateTime=$endDate)"
            $stats = Invoke-MgGraphRequest -Method GET -Uri $statsUri -ErrorAction Stop

            if ($stats.value -and $stats.value.Count -gt 0) {
                $uniqueViewers = ($stats.value | Measure-Object -Property actorCount -Sum).Sum
                $totalVisits   = ($stats.value | Measure-Object -Property visitCount -Sum).Sum
                $totalHits     = ($stats.value | Measure-Object -Property hits       -Sum).Sum
                $avgSecondsRaw = ($stats.value | Measure-Object -Property timeSpentInSeconds -Average).Average
                if ($avgSecondsRaw) {
                    $avgTimeSpent = [TimeSpan]::FromSeconds($avgSecondsRaw).ToString("mm\m\ ss\s")
                }
            } else {
                $status = "No activity in period"
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            # Pull out the HTTP status code if present
            if ($errMsg -match '(\d{3})\s*\(([^)]+)\)') {
                $status = "Analytics: $($matches[1]) $($matches[2])"
            } else {
                $status = "Analytics: $($errMsg.Substring(0, [Math]::Min(150, $errMsg.Length)))"
            }
        }
    }

    [PSCustomObject]@{
        SiteUrl              = $siteUrl
        SiteName             = $site.displayName
        UniqueViewers        = $uniqueViewers
        SiteVisits           = $totalVisits
        PageViews            = $totalHits
        AvgTimeSpentPerUser  = $avgTimeSpent
        AvgTimeSpentSeconds  = if ($avgSecondsRaw) { [math]::Round($avgSecondsRaw, 0) } else { $null }
        Status               = $status
        ReportPeriod         = "Last $daysBack days"
        ReportGeneratedUtc   = (Get-Date).ToUniversalTime().ToString("u")
    }
}
Write-Progress -Activity "Pulling site analytics" -Completed

# --- Export per-site detail CSV ---
$results |
    Sort-Object SiteVisits -Descending |
    Export-Csv -Path $detailCsv -NoTypeInformation -Encoding UTF8

# --- Tenant summary ---
$totalUniqueViewers = ($results | Measure-Object -Property UniqueViewers -Sum).Sum
$totalSiteVisits    = ($results | Measure-Object -Property SiteVisits    -Sum).Sum
$totalPageViews     = ($results | Measure-Object -Property PageViews     -Sum).Sum
$avgTimeAcross      = ($results | Where-Object AvgTimeSpentSeconds | Measure-Object -Property AvgTimeSpentSeconds -Average).Average
$avgTimeFormatted   = if ($avgTimeAcross) { [TimeSpan]::FromSeconds($avgTimeAcross).ToString("mm\m\ ss\s") } else { "N/A" }

$summary = [PSCustomObject]@{
    ReportPeriod         = "Last $daysBack days"
    ReportGeneratedUtc   = (Get-Date).ToUniversalTime().ToString("u")
    TotalSitesQueried    = $results.Count
    SitesWithData        = ($results | Where-Object { $_.Status -eq "OK" }).Count
    TotalUniqueViewers   = $totalUniqueViewers
    TotalSiteVisits      = $totalSiteVisits
    TotalPageViews       = $totalPageViews
    AvgTimeSpentPerUser  = $avgTimeFormatted
}

$summary | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8

# --- Confirm ---
Write-Host "`n========== EXPORT COMPLETE ==========" -ForegroundColor Green
Write-Host "Detail CSV  : $detailCsv  ($($results.Count) rows)" -ForegroundColor Green
Write-Host "Summary CSV : $summaryCsv" -ForegroundColor Green

Write-Host "`n========== TENANT-WIDE TOTALS ==========" -ForegroundColor Yellow
$summary | Format-List

Disconnect-MgGraph | Out-Null