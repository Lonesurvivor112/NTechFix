function Find-EntraDynamicParentsOfOnPremGroup {
<#
.SYNOPSIS
  Given an on-prem AD group's objectGUID (or a cloud group Id), list Entra dynamic groups
  whose dynamic membership rules reference that group via memberOf.

.PARAMETER OnPremGroupGuid
  The on-prem AD group's objectGUID (Guid). If provided, the function resolves it to the cloud group.

.PARAMETER CloudGroupId
  The Microsoft Entra group ObjectId (Guid) of the synced group. Skips on-prem lookup if supplied.

.PARAMETER ExportCsv
  Optional CSV export path.

.EXAMPLE
  THIS DOES NOT WORK JUST USE OTHER - Find-EntraDynamicParentsOfOnPremGroup -OnPremGroupGuid '9a6b9c9f-5a4b-4b1e-9a7f-123456789abc'

.EXAMPLE
  Find-EntraDynamicParentsOfOnPremGroup -CloudGroupId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' -ExportCsv .\parents.csv
#>
  [CmdletBinding()]
  param(
    [Guid]$OnPremGroupGuid,
    [Guid]$CloudGroupId,
    [string]$ExportCsv
  )

  if (-not $OnPremGroupGuid -and -not $CloudGroupId) {
    throw "Provide either -OnPremGroupGuid or -CloudGroupId."
  }

  # --- Load modules and connect to Graph ---
  if ($OnPremGroupGuid) {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
      throw "ActiveDirectory module not found. Install RSAT or run on a management host."
    }
    Import-Module ActiveDirectory
  }

  $requiredScopes = @('Directory.Read.All','Group.Read.All')
  $ctx = Get-MgContext -ErrorAction SilentlyContinue
  if (-not $ctx -or @($ctx.Scopes) -notcontains 'Group.Read.All') {
    Connect-MgGraph -Scopes $requiredScopes
  }

  # --- Resolve on-prem GUID -> cloud group Id (via SID) ---
  if ($OnPremGroupGuid) {
    try {
      $adGroup = Get-ADGroup -Identity $OnPremGroupGuid -Properties SID, SamAccountName
    } catch {
      throw "Could not find an AD group with objectGUID $OnPremGroupGuid. $_"
    }
    $sid = $adGroup.SID.Value
    Write-Verbose "Resolved AD group '$($adGroup.SamAccountName)' SID: $sid"

    # Advanced query header for certain directory filters
    $headers = @{ ConsistencyLevel = 'eventual' }
    $cloudMatches = Get-MgGroup -Filter "onPremisesSecurityIdentifier eq '$sid'" `
                                -ConsistencyLevel eventual -CountVariable c -All `
                                -Property "id,displayName,onPremisesSyncEnabled" -Headers $headers
    if (-not $cloudMatches) {
      throw "No Entra group found with onPremisesSecurityIdentifier '$sid'. Has the group synced yet?"
    }
    if ($cloudMatches.Count -gt 1) {
      Write-Warning "Multiple Entra groups share SID $sid. Using the first match: $($cloudMatches[0].DisplayName)."
    }
    $CloudGroupId = $cloudMatches[0].Id
    Write-Verbose "Cloud group Id: $CloudGroupId"
  }

  # --- Enumerate dynamic groups and find rules that reference the cloud group Id ---
  $dynGroups = Get-MgGroup -Filter "groupTypes/any(s:s eq 'DynamicMembership')" -All `
                           -Property "id,displayName,membershipRule,membershipRuleProcessingState,description"
  if (-not $dynGroups) {
    Write-Output "No dynamic groups found in tenant."
    return
  }

  $pattern = [Regex]::Escape($CloudGroupId.ToString())
  $hits = $dynGroups | Where-Object { $_.membershipRule -and ($_.membershipRule -match $pattern) }

  if (-not $hits) {
    Write-Output "No dynamic group rules reference cloud group Id $CloudGroupId."
    return
  }

  # Optional: helpful snippet around the matched Id
  function Get-RuleSnippet([string]$rule,[string]$id) {
    $idx = $rule.IndexOf($id, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) { return $rule }
    $start = [Math]::Max(0, $idx - 50)
    $len   = [Math]::Min(120, $rule.Length - $start)
    $rule.Substring($start, $len)
  }

  $result = $hits | Select-Object `
    @{n='DynamicGroup';e={$_.DisplayName}},
    @{n='DynamicGroupId';e={$_.Id}},
    @{n='ProcessingState';e={$_.MembershipRuleProcessingState}},
    @{n='RuleSnippet';e={ Get-RuleSnippet $_.MembershipRule $CloudGroupId }},
    @{n='FullRule';e={$_.MembershipRule}},
    @{n='Description';e={$_.Description}}

  if ($ExportCsv) {
    $result | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Output "Exported to $ExportCsv"
  } else {
    $result | Sort-Object DynamicGroup | Format-Table -Wrap
  }

  Write-Verbose "Note: memberOf rules include only direct members and cannot be combined with other operators."
}
