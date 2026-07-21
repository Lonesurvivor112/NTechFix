# DRY RUN: check current state for sec-gl-role-receptionistrelief (no changes)
$ErrorActionPreference = 'Stop'
$GroupName = 'sec-gl-role-receptionistrelief'

# Connect to Graph
$scopes = @('Group.ReadWrite.All','Directory.ReadWrite.All')
Connect-MgGraph -Scopes $scopes

# Resolve the AIP SKU (RIGHTSMANAGEMENT_CE)
$sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'RIGHTSMANAGEMENT_CE' }
if (-not $sku) { throw "RIGHTSMANAGEMENT_CE SKU not found in this tenant." }
$skuId = $sku.SkuId

# Resolve group robustly
$g = Resolve-Group -Name $GroupName
if (-not $g) { throw "Group '$GroupName' not found by displayName/mailNickname/search." }

Write-Host ("Group found: {0}  ({1})  SecurityEnabled={2}  GroupTypes={3}" -f `
    $g.displayName, $g.id, $g.securityEnabled, ($g.groupTypes -join ','))

# Check if group already has the SKU
$gFull  = Get-MgGroup -GroupId $g.Id -Property AssignedLicenses,AssignedPlans
$hasSku = $false
if ($gFull.AssignedLicenses) { $hasSku = $gFull.AssignedLicenses.SkuId -contains $skuId }

if ($hasSku) {
  Write-Host "[DRY RUN] Would SKIP: $($g.displayName) already has RIGHTSMANAGEMENT_CE" -ForegroundColor Yellow
} else {
  Write-Host "[DRY RUN] Would ASSIGN RIGHTSMANAGEMENT_CE to $($g.displayName) (no plans disabled)" -ForegroundColor Cyan
}