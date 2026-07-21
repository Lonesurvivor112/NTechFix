<#
Requires:
  - Microsoft.Graph (Directory.Read.All, Group.Read.All)
  - ImportExcel

What it does:
  For every Dynamic Membership group, extract all GUIDs found in the membershipRule,
  resolve those GUIDs to group display names, and export results to Excel in both
  flat and aggregated forms.
#>

[CmdletBinding()]
param(
  [string[]]$DynamicGroupIds = @()  # Optional: limit processing to these DG IDs
)

# ---------- Connect to Graph ----------
$requiredScopes = @('Directory.Read.All','Group.Read.All')
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if (!$ctx -or !($ctx.Scopes -contains 'Group.Read.All')) {
  Connect-MgGraph -Scopes $requiredScopes
}

# ---------- Pull dynamic groups ----------
$prop = 'id,displayName,membershipRule'
$filter = "groupTypes/any(s:s eq 'DynamicMembership')"

$dynGroups = Get-MgGroup -Filter $filter -All -Property $prop |
  Where-Object { $_.MembershipRule }  # ensure rule exists

if ($DynamicGroupIds.Count) {
  $dynGroups = $dynGroups | Where-Object { $DynamicGroupIds -contains $_.Id }
}

if (-not $dynGroups) {
  Write-Host "No dynamic groups found (filter may be too narrow)." -ForegroundColor Yellow
  return
}

# ---------- Extract every GUID from rules ----------
$guidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# All unique GUIDs across all dynamic groups (so we can resolve names once)
$allGuids =
  $dynGroups |
  ForEach-Object {
    [regex]::Matches($_.MembershipRule, $guidPattern) |
      ForEach-Object { $_.Value.ToLower() }
  } |
  Where-Object { $_ } |
  Sort-Object -Unique

# ---------- Resolve GUIDs → Group display names ----------
$guidToName = @{}
if ($allGuids.Count) {
  Write-Host "Resolving $($allGuids.Count) referenced group IDs..." -ForegroundColor Cyan

  # Graph doesn't support a giant "id in (...)" filter for Groups reliably, so resolve one-by-one.
  # If you have thousands, consider batching with Invoke-MgGraphRequest $batch.
  foreach ($gid in $allGuids) {
    try {
      $g = Get-MgGroup -GroupId $gid -Property displayName -ErrorAction Stop
      $guidToName[$gid] = $g.DisplayName
    } catch {
      # Could also be an app/service principal or a deleted group referenced historically
      $guidToName[$gid] = $null
    }
  }
}

# ---------- Build flat rows ----------
$flat = New-Object System.Collections.Generic.List[object]

foreach ($dg in $dynGroups) {
  $ruleGuids =
    [regex]::Matches($dg.MembershipRule, $guidPattern) |
    ForEach-Object { $_.Value.ToLower() } |
    Sort-Object -Unique

  if (-not $ruleGuids) {
    # Dynamic group with no explicit GUIDs (may use attributes only)
    $flat.Add([pscustomobject]@{
      DynamicGroupId    = $dg.Id
      DynamicGroupName  = $dg.DisplayName
      ReferencedGroupId = $null
      ReferencedName    = $null
      RuleSnippet       = $dg.MembershipRule.Substring(0, [Math]::Min(250, $dg.MembershipRule.Length))
      FullRule          = $dg.MembershipRule
    })
    continue
  }

  foreach ($gid in $ruleGuids) {
    $name = $guidToName[$gid]
    # Create a short snippet around the GUID for quick visual
    $snippet = ($dg.MembershipRule -replace ".*?($gid).*?", '...$1...')
    $flat.Add([pscustomobject]@{
      DynamicGroupId    = $dg.Id
      DynamicGroupName  = $dg.DisplayName
      ReferencedGroupId = $gid
      ReferencedName    = $name
      RuleSnippet       = $snippet
      FullRule          = $dg.MembershipRule
    })
  }
}

# ---------- Aggregated (per DG → CSV of referenced groups) ----------
$aggregated =
  $flat |
  Group-Object DynamicGroupId |
  ForEach-Object {
    $dgId   = $_.Name
    $first  = $_.Group | Select-Object -First 1
    $pairs  = $_.Group |
              Where-Object { $_.ReferencedGroupId } |
              ForEach-Object {
                if ($_.ReferencedName) { "$($_.ReferencedName) <$($_.ReferencedGroupId)>" }
                else { $_.ReferencedGroupId }
              }

    [pscustomobject]@{
      DynamicGroupId    = $dgId
      DynamicGroupName  = $first.DynamicGroupName
      ReferencedCount   = ($pairs | Measure-Object).Count
      ReferencedGroups  = ($pairs -join ', ')
      FullRule          = $first.FullRule
    }
  } |
  Sort-Object DynamicGroupName

# ---------- Export to Excel ----------
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
  try { Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop } catch {}
}
Import-Module ImportExcel -ErrorAction SilentlyContinue

$xlPath = ".\DynamicGroup_ReferencedGroups.xlsx"
Remove-Item $xlPath -ErrorAction SilentlyContinue

# Flat sheet
$flat | Export-Excel -Path $xlPath -WorksheetName 'Flat' -AutoSize -FreezeTopRow

# Aggregated sheet
$aggregated | Export-Excel -Path $xlPath -WorksheetName 'Aggregated' -AutoSize -FreezeTopRow -Append

Write-Host "✅ Exported to $xlPath" -ForegroundColor Green