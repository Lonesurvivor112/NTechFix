<#
.SYNOPSIS
    Reports M365 Group members/owners and optionally SharePoint site groups.
    Uses ExchangeOnlineManagement + Microsoft.Online.SharePoint.PowerShell.

.PARAMETER GroupFilter
    Partial name to search (e.g. "Nursing"). Prompts if not supplied.

.PARAMETER TenantAdminUrl
    SharePoint admin URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER SiteDetails
    Also reports each group's SharePoint site Owners / Members / Visitors.

.PARAMETER OutputPath
    Folder to write CSV reports (default: current directory).

.EXAMPLE
    .\Get-GroupReport.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"
    .\Get-GroupReport.ps1 -GroupFilter "Nursing" -TenantAdminUrl "https://contoso-admin.sharepoint.com" -SiteDetails
#>

#Requires -Modules ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell

[CmdletBinding()]
param(
    [string]$GroupFilter,

    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [switch]$SiteDetails,

    [string]$OutputPath = (Get-Location).Path
)

$ErrorActionPreference = "Stop"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ── Title lookup cache (avoids duplicate Get-User calls) ──────────────────────

$titleCache = @{}

function Get-CachedUserTitle {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return "" }
    if (-not $titleCache.ContainsKey($Email)) {
        $u = Get-User -Identity $Email -ErrorAction SilentlyContinue
        $titleCache[$Email] = if ($u) { $u.Title } else { "" }
    }
    return $titleCache[$Email]
}

# ── Prompt for filter if not supplied ─────────────────────────────────────────

if (-not $GroupFilter) {
    $GroupFilter = Read-Host "Enter group name filter (e.g. 'Nursing')"
    if ([string]::IsNullOrWhiteSpace($GroupFilter)) { throw "A group name filter is required." }
}

# ── Connect ───────────────────────────────────────────────────────────────────

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

if ($SiteDetails) {
    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
    try {
        Connect-SPOService -Url $TenantAdminUrl -Interactive
    }
    catch {
        # Older module versions use -UseWebLogin instead of -Interactive
        Connect-SPOService -Url $TenantAdminUrl -UseWebLogin
    }
}

# ── Find matching M365 groups ─────────────────────────────────────────────────

Write-Host "Searching for M365 groups matching '$GroupFilter'..." -ForegroundColor Cyan
$allGroups = Get-UnifiedGroup -Filter "DisplayName -like '*$GroupFilter*'" -ResultSize Unlimited

if (-not $allGroups) {
    Write-Warning "No M365 groups found matching '$GroupFilter'."
    return
}

Write-Host "  Found $($allGroups.Count) group(s)." -ForegroundColor Green

# ── Report 1: Owners & Members ────────────────────────────────────────────────

$memberReport = [System.Collections.Generic.List[PSObject]]::new()

foreach ($group in $allGroups) {
    Write-Host "  Processing: $($group.DisplayName)" -ForegroundColor Gray

    $owners  = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Owners  -ResultSize Unlimited -ErrorAction SilentlyContinue
    $members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue

    foreach ($owner in $owners) {
        $memberReport.Add([PSCustomObject]@{
            GroupName  = $group.DisplayName
            GroupEmail = $group.PrimarySmtpAddress
            Role       = "Owner"
            Name       = $owner.DisplayName
            Title      = Get-CachedUserTitle $owner.PrimarySmtpAddress
            Email      = $owner.PrimarySmtpAddress
        })
    }

    $ownerEmails = $owners.PrimarySmtpAddress

    foreach ($member in $members) {
        if ($ownerEmails -contains $member.PrimarySmtpAddress) { continue }  # skip duplicate owners
        $memberReport.Add([PSCustomObject]@{
            GroupName  = $group.DisplayName
            GroupEmail = $group.PrimarySmtpAddress
            Role       = "Member"
            Name       = $member.DisplayName
            Title      = Get-CachedUserTitle $member.PrimarySmtpAddress
            Email      = $member.PrimarySmtpAddress
        })
    }
}

$report1Path = Join-Path $OutputPath "${GroupFilter}_M365Groups_${timestamp}.csv"
$memberReport | Export-Csv -Path $report1Path -NoTypeInformation -Encoding UTF8
Write-Host "`nReport 1 saved: $report1Path" -ForegroundColor Green

# ── Report 2: SharePoint Site Groups (--SiteDetails) ─────────────────────────

if ($SiteDetails) {
    $siteReport = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($group in $allGroups) {
        $siteUrl = $group.SharePointSiteUrl
        if ([string]::IsNullOrWhiteSpace($siteUrl)) {
            Write-Warning "  No site URL for '$($group.DisplayName)' — skipping."
            continue
        }

        Write-Host "  Getting site groups: $siteUrl" -ForegroundColor Gray

        try {
            $spGroups = Get-SPOSiteGroup -Site $siteUrl -ErrorAction Stop

            # Focus on the three default groups; adjust the filter if your naming differs
            $defaultGroups = $spGroups | Where-Object {
                $_.Title -match "Owners|Members|Visitors"
            }

            foreach ($spGroup in $defaultGroups) {
                $users = Get-SPOUser -Site $siteUrl -Group $spGroup.Title -ErrorAction SilentlyContinue

                foreach ($user in $users) {
                    $siteReport.Add([PSCustomObject]@{
                        GroupName   = $group.DisplayName
                        SiteUrl     = $siteUrl
                        SPGroup     = $spGroup.Title
                        DisplayName = $user.DisplayName
                        Title       = Get-CachedUserTitle $user.Email
                        LoginName   = $user.LoginName
                        Email       = $user.Email
                        UserType    = if ($user.IsSiteAdmin) { "SiteAdmin" } else { "User" }
                    })
                }
            }
        }
        catch {
            Write-Warning "  Could not retrieve site groups for '$siteUrl': $_"
        }
    }

    $report2Path = Join-Path $OutputPath "${GroupFilter}_SiteDetails_${timestamp}.csv"
    $siteReport | Export-Csv -Path $report2Path -NoTypeInformation -Encoding UTF8
    Write-Host "Report 2 saved: $report2Path" -ForegroundColor Green
}

# ── Summary ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "  Filter       : $GroupFilter"
Write-Host "  Groups found : $($allGroups.Count)"
Write-Host "  Rows (R1)    : $($memberReport.Count)"
if ($SiteDetails) { Write-Host "  Rows (R2)    : $($siteReport.Count)" }
