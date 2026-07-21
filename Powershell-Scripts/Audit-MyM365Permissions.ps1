<#
.SYNOPSIS
    Audits and removes your own delegated permissions across all M365 mailboxes and OneDrives.

.DESCRIPTION
    Connects to Exchange Online and SharePoint Online, then finds everywhere the
    running admin account has been granted:
      - Mailbox Full Access
      - Mailbox Send As
      - Mailbox Send on Behalf
      - OneDrive Site Collection Admin

    Results are displayed interactively. You can remove individual entries or all at once.

.PARAMETER AdminUPN
    Your admin UPN (e.g. admin@contoso.com). If not supplied, the script will
    detect it after connecting to Exchange Online.

.PARAMETER SharePointAdminUrl
    Your SharePoint admin URL (e.g. https://contoso-admin.sharepoint.com).
    Required for OneDrive checks.

.PARAMETER ReportOnly
    If set, findings are shown but no removal prompts appear.

.EXAMPLE
    .\Audit-MyM365Permissions.ps1 -SharePointAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    .\Audit-MyM365Permissions.ps1 -AdminUPN "admin@contoso.com" `
        -SharePointAdminUrl "https://contoso-admin.sharepoint.com" -ReportOnly
#>

[CmdletBinding()]
param(
    [string]$AdminUPN,
    [string]$SharePointAdminUrl,
    [switch]$ReportOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Colour helpers ---------------------------------------------------------

function Write-Header  { param($t) Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Write-Success { param($t) Write-Host $t -ForegroundColor Green }
function Write-Warn    { param($t) Write-Host $t -ForegroundColor Yellow }
function Write-Info    { param($t) Write-Host $t -ForegroundColor Gray }

# --- Module check -----------------------------------------------------------

foreach ($mod in @("ExchangeOnlineManagement", "Microsoft.Online.SharePoint.PowerShell")) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Warn "Module '$mod' not found. Install it with: Install-Module $mod -Scope CurrentUser"
        exit 1
    }
}

# SPO module was built for Windows PowerShell 5.1 and needs a compatibility shim in PS7
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -WarningAction SilentlyContinue
} else {
    Import-Module Microsoft.Online.SharePoint.PowerShell
}

# --- Connect ----------------------------------------------------------------

Write-Header "Connecting to Exchange Online"
try {
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Success "Connected to Exchange Online."
} catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}

# Resolve running admin's UPN if not supplied
if (-not $AdminUPN) {
    $AdminUPN = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
    if (-not $AdminUPN) {
        $AdminUPN = Read-Host "Could not auto-detect UPN. Enter your UPN"
    }
}
Write-Info "Auditing permissions for: $AdminUPN"

# SharePoint / OneDrive
if (-not $SharePointAdminUrl) {
    $SharePointAdminUrl = Read-Host "`nEnter your SharePoint Admin URL (e.g. https://contoso-admin.sharepoint.com)"
}

Write-Header "Connecting to SharePoint Online"
try {
    Connect-SPOService -Url $SharePointAdminUrl
    Write-Success "Connected to SharePoint Online."
} catch {
    Write-Error "Failed to connect to SharePoint Online: $_"
}

# --- Helpers ----------------------------------------------------------------

# Normalise identity strings so we can match $AdminUPN reliably
function Matches-Identity {
    param([string]$Identity, [string]$UPN)
    return ($Identity -like "*$UPN*") -or ($Identity -like "*$($UPN.Split('@')[0])*")
}

# --- 1. MAILBOX PERMISSIONS -------------------------------------------------

Write-Header "Scanning mailboxes (this may take a few minutes)..."

$allMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object -Property DisplayName, PrimarySmtpAddress, GrantSendOnBehalfTo
$total        = $allMailboxes.Count
$counter      = 0

$findings = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($mbx in $allMailboxes) {
    $counter++
    $pct = [int](($counter / $total) * 100)
    Write-Progress -Activity "Checking mailboxes" -Status "$counter / $total  ($($mbx.PrimarySmtpAddress))" -PercentComplete $pct

    $smtp = $mbx.PrimarySmtpAddress

    # Full Access
    try {
        Get-MailboxPermission -Identity $smtp -User $AdminUPN -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -like "*FullAccess*" -and $_.Deny -eq $false } |
        ForEach-Object {
            $findings.Add([PSCustomObject]@{
                Type     = "Mailbox - Full Access"
                Target   = "$($mbx.DisplayName) <$smtp>"
                Identity = $smtp
                Detail   = "FullAccess"
            })
        }
    } catch { Write-Warn "  Skipped Full Access check on $smtp`: $_" }

    # Send As
    try {
        Get-RecipientPermission -Identity $smtp -Trustee $AdminUPN -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -like "*SendAs*" } |
        ForEach-Object {
            $findings.Add([PSCustomObject]@{
                Type     = "Mailbox - Send As"
                Target   = "$($mbx.DisplayName) <$smtp>"
                Identity = $smtp
                Detail   = "SendAs"
            })
        }
    } catch { Write-Warn "  Skipped Send As check on $smtp`: $_" }

    # Send on Behalf
    if ($mbx.GrantSendOnBehalfTo) {
        $mbx.GrantSendOnBehalfTo | ForEach-Object {
            if (Matches-Identity -Identity $_ -UPN $AdminUPN) {
                $findings.Add([PSCustomObject]@{
                    Type     = "Mailbox - Send on Behalf"
                    Target   = "$($mbx.DisplayName) <$smtp>"
                    Identity = $smtp
                    Detail   = "SendOnBehalf"
                })
            }
        }
    }
}

Write-Progress -Activity "Checking mailboxes" -Completed

# --- 2. ONEDRIVE SITE COLLECTION ADMIN -------------------------------------

Write-Header "Scanning OneDrive sites for Site Collection Admin access..."

try {
    $oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'" |
                     Select-Object -Property Url, Owner, Title
    $odTotal   = $oneDriveSites.Count
    $odCounter = 0

    foreach ($site in $oneDriveSites) {
        $odCounter++
        $pct = [int](($odCounter / $odTotal) * 100)
        Write-Progress -Activity "Checking OneDrive sites" -Status "$odCounter / $odTotal  ($($site.Url))" -PercentComplete $pct

        try {
            $admins = Get-SPOUser -Site $site.Url -Limit All -ErrorAction SilentlyContinue |
                      Where-Object { $_.IsSiteAdmin -eq $true }

            $admins | ForEach-Object {
                if (Matches-Identity -Identity $_.LoginName -UPN $AdminUPN) {
                    $findings.Add([PSCustomObject]@{
                        Type     = "OneDrive - Site Collection Admin"
                        Target   = "$($site.Title) ($($site.Url))"
                        Identity = $site.Url
                        Detail   = "SiteCollectionAdmin"
                    })
                }
            }
        } catch {
            Write-Warn "  Skipped $($site.Url): $_"
        }
    }

    Write-Progress -Activity "Checking OneDrive sites" -Completed

} catch {
    Write-Warn "OneDrive scan failed: $_"
}

# --- 3. REPORT --------------------------------------------------------------

Write-Header "Results"

if ($findings.Count -eq 0) {
    Write-Success "No delegated permissions found for $AdminUPN."
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit 0
}

Write-Warn "$($findings.Count) permission(s) found for $AdminUPN :`n"

$i = 1
$findings | ForEach-Object {
    Write-Host ("  [{0,2}]  {1,-36}  {2}" -f $i, $_.Type, $_.Target) -ForegroundColor Yellow
    $i++
}

if ($ReportOnly) {
    Write-Info "`n-ReportOnly specified - no changes made."
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit 0
}

# --- 4. REMOVAL -------------------------------------------------------------

Write-Header "Removal"

$choice = Read-Host "`nEnter numbers to remove (comma-separated), 'ALL' to remove everything, or ENTER to skip"

$toRemove = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($choice -eq "ALL") {
    $findings | ForEach-Object { $toRemove.Add($_) }
} elseif ($choice -match "\d") {
    $choice -split "," | ForEach-Object {
        $token = $_.Trim() -replace '[^0-9]', ''
        $idx   = 0
        if ($token -ne '' -and [int]::TryParse($token, [ref]$idx)) {
            $idx--
            if ($idx -ge 0 -and $idx -lt $findings.Count) {
                $toRemove.Add($findings[$idx])
            } else {
                Write-Warn "  Number out of range: $($idx + 1) - skipped."
            }
        } else {
            Write-Warn "  Skipped non-numeric token: '$($_.Trim())'"
        }
    }
} else {
    Write-Info "No changes made."
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    exit 0
}

Write-Host ""

foreach ($item in $toRemove) {
    Write-Info "Removing: $($item.Type) on $($item.Target)..."

    try {
        switch ($item.Detail) {

            "FullAccess" {
                Remove-MailboxPermission -Identity $item.Identity `
                    -User $AdminUPN -AccessRights FullAccess -Confirm:$false
                Write-Success "  Removed Full Access on $($item.Identity)"
            }

            "SendAs" {
                Remove-RecipientPermission -Identity $item.Identity `
                    -Trustee $AdminUPN -AccessRights SendAs -Confirm:$false
                Write-Success "  Removed Send As on $($item.Identity)"
            }

            "SendOnBehalf" {
                $mbx = Get-Mailbox -Identity $item.Identity
                $updated = $mbx.GrantSendOnBehalfTo |
                    Where-Object { -not (Matches-Identity -Identity $_ -UPN $AdminUPN) }
                Set-Mailbox -Identity $item.Identity -GrantSendOnBehalfTo $updated
                Write-Success "  Removed Send on Behalf on $($item.Identity)"
            }

            "SiteCollectionAdmin" {
                # Check current lock state - locked sites (disabled/deleted accounts) reject permission changes
                $siteInfo    = Get-SPOSite -Identity $item.Identity -ErrorAction Stop
                $priorLock   = $siteInfo.LockState
                $wasUnlocked = $false

                if ($priorLock -ne "Unlock") {
                    Write-Info "  Site is locked ($priorLock) - temporarily unlocking..."
                    Set-SPOSite -Identity $item.Identity -LockState "Unlock" -ErrorAction Stop
                    $wasUnlocked = $true
                    # Brief pause for the lock change to propagate
                    Start-Sleep -Seconds 3
                }

                Set-SPOUser -Site $item.Identity -LoginName $AdminUPN -IsSiteCollectionAdmin $false -ErrorAction Stop

                if ($wasUnlocked) {
                    Write-Info "  Re-applying lock state ($priorLock)..."
                    Set-SPOSite -Identity $item.Identity -LockState $priorLock -ErrorAction SilentlyContinue
                }

                Write-Success "  Removed Site Collection Admin on $($item.Identity)"
            }
        }
    } catch {
        Write-Warn "  FAILED to remove $($item.Type) on $($item.Identity): $_"
    }
}

# --- Disconnect -------------------------------------------------------------

Write-Header "Done"
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Success "Disconnected from Exchange Online."
Write-Info "Note: To disconnect SharePoint run: Disconnect-SPOService"
