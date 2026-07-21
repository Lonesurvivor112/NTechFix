<#
.SYNOPSIS
  Inventories Microsoft 365 Groups to identify Planner “sprawl”:
  - Groups with Planner plans
  - Whether they’re Teams-enabled
  - Their connected SharePoint site URL
  - Exports results to CSV

.REQUIREMENTS
  Microsoft Graph PowerShell SDK

  Delegated scopes recommended:
    Group.Read.All
    Tasks.Read
    Sites.Read.All

  Key APIs:
    - List group-owned Planner plans: GET /groups/{id}/planner/plans [2](https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans?view=graph-rest-1.0)
    - Detect Teams-enabled groups via resourceProvisioningOptions 'Team' [1](https://learn.microsoft.com/en-us/graph/teams-list-all-teams)
    - Get group SharePoint site: GET /groups/{id}/sites/root [3](https://learn.microsoft.com/en-us/graph/api/site-get?view=graph-rest-1.0)
#>

[CmdletBinding()]
param(
  [string]$OutputPath = ".\PlannerSprawlReport.csv",
  [int]$ThrottleMs = 150,
  [switch]$IncludeNoPlannerGroups,  # if set, outputs all groups; otherwise outputs only groups with Planner plans (or unknown access)
  [switch]$VerboseErrors
)

$ErrorActionPreference = "Stop"

function Ensure-Graph {
  if (-not (Get-Module Microsoft.Graph -ListAvailable)) {
    throw "Microsoft.Graph PowerShell SDK not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
  }

  if (-not (Get-MgContext)) {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    # Delegated scopes for interactive use:
    Connect-MgGraph -Scopes "Group.Read.All","Tasks.Read","Sites.Read.All" | Out-Null
  }
}

function Get-GroupPlans {
  param([Parameter(Mandatory)][string]$GroupId)

  try {
    # Graph PowerShell cmdlet maps to /groups/{id}/planner/plans [2](https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans?view=graph-rest-1.0)
    $plans = @(Get-MgGroupPlannerPlan -GroupId $GroupId -ErrorAction Stop)
    return ,$plans
  } catch {
    if ($VerboseErrors) {
      Write-Warning "Planner plans query failed for GroupId=$GroupId : $($_.Exception.Message)"
    }
    return $null  # distinguish "no access / error" from "empty list"
  }
}

function Get-GroupSiteUrl {
  param([Parameter(Mandatory)][string]$GroupId)

  try {
    # /groups/{id}/sites/root returns the group’s SharePoint site [3](https://learn.microsoft.com/en-us/graph/api/site-get?view=graph-rest-1.0)
    $site = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/sites/root" -f $GroupId)
    return $site.webUrl
  } catch {
    if ($VerboseErrors) {
      Write-Warning "Site query failed for GroupId=$GroupId : $($_.Exception.Message)"
    }
    return $null
  }
}

# -------------------- Main --------------------
Ensure-Graph

Write-Host "Fetching Microsoft 365 (Unified) groups..." -ForegroundColor Cyan
$groups = Get-MgGroup -All `
  -Filter "groupTypes/any(c:c eq 'Unified')" `
  -Property "id,displayName,createdDateTime,visibility,resourceProvisioningOptions,mailNickname"

Write-Host "Scanning $($groups.Count) groups..." -ForegroundColor Cyan

$results = New-Object System.Collections.Generic.List[object]
$counter = 0

foreach ($g in $groups) {
  $counter++
  Write-Progress -Activity "Planner Sprawl Scan" `
    -Status "$counter / $($groups.Count): $($g.DisplayName)" `
    -PercentComplete (($counter / $groups.Count) * 100)

  # Teams-enabled detection: resourceProvisioningOptions contains 'Team' [1](https://learn.microsoft.com/en-us/graph/teams-list-all-teams)
  $hasTeam = $false
  if ($null -ne $g.ResourceProvisioningOptions) {
    $hasTeam = $g.ResourceProvisioningOptions -contains "Team"
  }

  # Planner plans: /groups/{id}/planner/plans [2](https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans?view=graph-rest-1.0)
  $plans = Get-GroupPlans -GroupId $g.Id
  $hasPlanner = $false
  $plannerCount = 0
  $plannerTitles = $null
  $plannerAccessState = "Ok"

  if ($plans -eq $null) {
    # Could be permissions or transient error
    $plannerAccessState = "UnknownAccess"
  } else {
    $plannerCount = $plans.Count
    if ($plannerCount -gt 0) {
      $hasPlanner = $true
      $plannerTitles = ($plans | Select-Object -ExpandProperty Title) -join " | "
    }
  }

  # SharePoint site
  $siteUrl = Get-GroupSiteUrl -GroupId $g.Id

  # Classification
  $sprawlType =
    if ($plannerAccessState -eq "UnknownAccess") { "UnknownAccess" }
    elseif ($hasPlanner -and -not $hasTeam)      { "PlannerOnly" }
    elseif ($hasPlanner -and $hasTeam)           { "Planner+Team" }
    else                                         { "NoPlanner" }

  $row = [pscustomobject]@{
    GroupDisplayName     = $g.DisplayName
    GroupId              = $g.Id
    CreatedDateTime      = $g.CreatedDateTime
    Visibility           = $g.Visibility
    HasTeam              = $hasTeam
    HasPlanner           = $hasPlanner
    PlannerPlanCount     = $plannerCount
    PlannerPlanTitles    = $plannerTitles
    PlannerAccessState   = $plannerAccessState
    SharePointSiteUrl    = $siteUrl
    SprawlType           = $sprawlType
  }

  if ($IncludeNoPlannerGroups) {
    $results.Add($row)
  } else {
    # Default output: only sprawl-relevant rows (Planner exists or unknown access)
    if ($row.HasPlanner -or $row.PlannerAccessState -eq "UnknownAccess") {
      $results.Add($row)
    }
  }

  Start-Sleep -Milliseconds $ThrottleMs
}

Write-Host "Exporting CSV to: $OutputPath" -ForegroundColor Green
$results |
  Sort-Object -Property `
    SprawlType,
    @{ Expression = 'PlannerPlanCount'; Descending = $true },
    GroupDisplayName |
  Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Done. Exported rows: $($results.Count)" -ForegroundColor Green
Write-Host "Tip: Filter CSV where SprawlType = 'PlannerOnly' for classic Planner sprawl." -ForegroundColor Yellow