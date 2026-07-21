function Find-EntraDynamicParentsOfOnPremGroup {
<#
.SYNOPSIS
  Interactive: paste one or more GUIDs (on-prem AD objectGUIDs or cloud group Ids) and list Entra dynamic groups
  whose dynamic membership rules reference those groups via memberOf. Shows details and a comma-separated names line.

.PARAMETER ExportCsv
  Optional CSV export path.

.EXAMPLE
  Find-EntraDynamicParentsOfOnPremGroup
  # You will be prompted to paste GUIDs (comma/space/newline separated).
#>
  [CmdletBinding()]
  param(
    [string]$ExportCsv
  )

  # --- Prompt for GUIDs (paste list) ---
  $raw = Read-Host "Paste one or more GUIDs (on-prem objectGUIDs or cloud group Ids), separated by comma, space, or newlines"
  $candidates = $raw -split '[,\s]+' | Where-Object { $_ } | Select-Object -Unique
  if (-not $candidates) { throw "No GUIDs provided." }

  # Helper to validate GUIDs
  function _IsGuid([string]$s) {
    $tmp = [Guid]::Empty
    return [Guid]::TryParse($s, [ref]$tmp)
  }

  # --- Load AD module if present (for on-prem resolution) ---
  $adAvailable = $false
  if (Get-Module -ListAvailable -Name ActiveDirectory) {
    try { Import-Module ActiveDirectory -ErrorAction Stop; $adAvailable = $true } catch { }
  }

  # --- Connect to Graph (ensure scopes) ---
  $requiredScopes = @('Directory.Read.All','Group.Read.All')
  $ctx = Get-MgContext -ErrorAction SilentlyContinue
  if (-not $ctx -or -not $ctx.Scopes -or @($ctx.Scopes) -notcontains 'Group.Read.All') {
    Connect-MgGraph -Scopes $requiredScopes
  }

  # Resolve inputs to cloud group Ids
  $resolvedMap = @{}  # CloudGroupId -> PSCustomObject(Source='OnPrem'|'Cloud'; Input; ADName; CloudName)
  $unresolved  = @()

  foreach ($g in $candidates) {
    if (-not (_IsGuid $g)) { Write-Warning "Skipping invalid GUID: $g"; continue }
    $handled = $false

    # Try as on-prem AD group first (objectGUID -> SID -> cloud)
    if ($adAvailable) {
      try {
        $adGroup = Get-ADGroup -Identity $g -Properties SID,SamAccountName -ErrorAction Stop
        $sid = $adGroup.SID.Value
        $headers = @{ ConsistencyLevel = 'eventual' }
        $cloudMatches = Get-MgGroup -Filter "onPremisesSecurityIdentifier eq '$sid'" `
                                    -ConsistencyLevel eventual -All `
                                    -Property "id,displayName,onPremisesSyncEnabled" -Headers $headers
        if ($cloudMatches) {
          if ($cloudMatches.Count -gt 1) {
            Write-Warning "Multiple Entra groups share SID $sid. Using first: $($cloudMatches[0].DisplayName)."
          }
          $cid = $cloudMatches[0].Id
          $resolvedMap[$cid] = [pscustomobject]@{
            Source   = 'OnPrem'
            Input    = $g
            ADName   = $adGroup.SamAccountName
            CloudName= $cloudMatches[0].DisplayName
          }
          $handled = $true
        }
      } catch { }
    }

    # If not handled, try directly as a cloud group ObjectId
    if (-not $handled) {
      try {
        $grp = Get-MgGroup -GroupId $g -Property "id,displayName"
        if ($grp) {
          $resolvedMap[$grp.Id] = [pscustomobject]@{
            Source   = 'Cloud'
            Input    = $g
            ADName   = $null
            CloudName= $grp.DisplayName
          }
          $handled = $true
        }
      } catch { }
    }

    if (-not $handled) { $unresolved += $g }
  }

  if ($unresolved.Count -gt 0) {
    Write-Warning ("Unresolved GUIDs (not AD objectGUIDs or cloud group Ids): " + ($unresolved -join ', '))
  }

  $cloudIds = @($resolvedMap.Keys | Select-Object -Unique)
  if (-not $cloudIds -or $cloudIds.Count -eq 0) {
    throw "No valid/resolvable cloud group Ids were provided."
  }

  # --- Enumerate dynamic groups once ---
  $dynGroups = Get-MgGroup -Filter "groupTypes/any(s:s eq 'DynamicMembership')" -All `
                           -Property "id,displayName,membershipRule,membershipRuleProcessingState,description"
  if (-not $dynGroups) {
    Write-Output "No dynamic groups found in tenant."
    return
  }

  # Helper for small snippet around a matched Id
  function Get-RuleSnippet([string]$rule,[string]$id) {
    if (-not $rule -or -not $id) { return $rule }
    $idx = $rule.IndexOf($id, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) { return $rule }
    $start = [Math]::Max(0, $idx - 50)
    $len   = [Math]::Min(120, $rule.Length - $start)
    $rule.Substring($start, $len)
  }

  # Match rules against any of provided cloud Ids
  $results = New-Object System.Collections.Generic.List[object]
  foreach ($dg in $dynGroups) {
    if (-not $dg.MembershipRule) { continue }
    $matchedIds = @()
    foreach ($cid in $cloudIds) {
      $pattern = [Regex]::Escape($cid.ToString())
      if ($dg.MembershipRule -match $pattern) { $matchedIds += $cid }
    }
    $matchedIds = $matchedIds | Select-Object -Unique
    if ($matchedIds.Count -gt 0) {
      $matchedNames = foreach ($cid in $matchedIds) {
        if ($resolvedMap.ContainsKey($cid) -and $resolvedMap[$cid].CloudName) { $resolvedMap[$cid].CloudName } else { $cid }
      }
      $snippet = Get-RuleSnippet $dg.MembershipRule $matchedIds[0]
      $results.Add([pscustomobject]@{
        DynamicGroup           = $dg.DisplayName
        DynamicGroupId         = $dg.Id
        ProcessingState        = $dg.MembershipRuleProcessingState
        MatchedChildGroupNames = ($matchedNames -join '; ')
        MatchedChildGroupIds   = ($matchedIds -join '; ')
        RuleSnippet            = $snippet
        FullRule               = $dg.MembershipRule
        Description            = $dg.Description
      })
    }
  }

  if ($results.Count -eq 0) {
    Write-Output "No dynamic group rules reference the provided group Id(s)."
    return
  }

  # --- Output: detailed table, then comma-separated names ---
  $sorted = $results | Sort-Object DynamicGroup

  if ($ExportCsv) {
    $sorted | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Output "Exported to $ExportCsv"
  }

  $sorted | Format-Table DynamicGroup, DynamicGroupId, ProcessingState, MatchedChildGroupNames, RuleSnippet -Wrap

  ""
  "Comma-separated dynamic group names:"
  ($sorted | Select-Object -ExpandProperty DynamicGroup -Unique | Sort-Object -Unique | ForEach-Object { $_ }) -join ', '
}
