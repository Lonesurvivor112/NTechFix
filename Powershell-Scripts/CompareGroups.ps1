<#
Prereqs:
Install-Module Microsoft.Graph -Scope CurrentUser
#>

# 0) Connect (once per session)
Connect-MgGraph -Scopes "Group.Read.All","User.Read.All" | Out-Null

# 1) Helper: get USERS from a group by GUID (direct or transitive)
function Get-UsersFromGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [switch]$Transitive  # use to include nested groups
    )

    # Get direct or transitive members; DO NOT $select '@odata.type'
    $raw = if ($Transitive) {
        Get-MgGroupTransitiveMember -GroupId $GroupId -All -Property "id,displayName,userPrincipalName,mail"
    } else {
        Get-MgGroupMember -GroupId $GroupId -All -Property "id,displayName,userPrincipalName,mail"
    }

    # Keep only user objects (works whether @odata.type is top-level or in AdditionalProperties)
    $userMembers = $raw | Where-Object {
        ($_.'@odata.type' -eq '#microsoft.graph.user') -or
        ($_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') -or
        ($_.PSObject.TypeNames -match 'User') -or
        ($_.PSObject.Properties.Name -contains 'UserPrincipalName') -or
        ($_.AdditionalProperties.Keys -contains 'userPrincipalName')
    }

    # For each user object ID, get full user so we always have UPN reliably
    $users = foreach ($m in $userMembers) {
        try {
            $u = Get-MgUser -UserId $m.Id -Property "id,displayName,userPrincipalName,mail"
            [pscustomobject]@{
                Id                = $u.Id
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
                Mail              = $u.Mail
            }
        } catch {
            # If a member somehow isn't a user or access denied, skip or note it
            [pscustomobject]@{
                Id                = $m.Id
                DisplayName       = $m.DisplayName
                UserPrincipalName = $null
                Mail              = $null
                Note              = "Lookup failed (not a user or access denied)"
            }
        }
    }

    # Return unique users by Id (some tenants return dups on transitive queries)
    $users | Group-Object Id | ForEach-Object { $_.Group[0] }
}

# === Example usage ===
$groupAId = '3318843e-2e32-458d-a174-ea11dd2f4f78'  # replace with your GUID Nicks Test Ktichen
$groupBId = '53f245aa-0284-4f2a-8966-7b8727f9b5ed'  # replace with your GUID Michigan EC

# 2) Pull users from each group (direct members). Add -Transitive if you want nested included.
$A = Get-UsersFromGroup -GroupId $groupAId        # -Transitive
$B = Get-UsersFromGroup -GroupId $groupBId        # -Transitive

# 3) Compare by stable key (Id). You can switch to UPN if desired.
$idxA = $A | Group-Object -Property Id -AsHashTable -AsString
$idxB = $B | Group-Object -Property Id -AsHashTable -AsString

$InBoth = $A | Where-Object { $idxB.ContainsKey($_.Id) }
$OnlyA  = $A | Where-Object { -not $idxB.ContainsKey($_.Id) }
$OnlyB  = $B | Where-Object { -not $idxA.ContainsKey($_.Id) }

# 4) Display a quick summary + tables
"Counts => A: $($A.Count)  B: $($B.Count)  InBoth: $($InBoth.Count)  OnlyA: $($OnlyA.Count)  OnlyB: $($OnlyB.Count)"

"`n===== In BOTH ====="
$InBoth | Sort-Object DisplayName, UserPrincipalName | Format-Table -AutoSize

"`n===== Only in A ====="
$OnlyA  | Sort-Object DisplayName, UserPrincipalName | Format-Table -AutoSize

"`n===== Only in B ====="
$OnlyB  | Sort-Object DisplayName, UserPrincipalName | Format-Table -AutoSize
