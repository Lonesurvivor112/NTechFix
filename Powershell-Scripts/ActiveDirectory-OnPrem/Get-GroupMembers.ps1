# Get-GroupMembers.ps1
# Reports members of specified Active Directory groups

$groups = @(
    "sec-gl-role-rcmstaff",
    "sec-gl-role-qualityStaff",
    "sec-gl-role-accounting",
    "sec-gl-role-hr",
    "sec-gl-role-payroll",
    "sec-gl-role-qualityDirector",
    "sec-gl-role-ClinicalDirector",
    "sec-gl-role-ccmDirector",
    "sec-gl-role-ccm",
    "sec-gl-role-admissionsDirector",
    "sec-gl-role-recruiter",
    "sec-gl-role-admissionsStaff"
)

foreach ($group in $groups) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Group: $group" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    try {
        $members = Get-ADGroupMember -Identity $group -ErrorAction Stop
        
        if ($members) {
            $members | ForEach-Object {
                Write-Host "  - $($_.Name) ($($_.SamAccountName))" -ForegroundColor Green
            }
            Write-Host "`nTotal members: $($members.Count)" -ForegroundColor Yellow
        }
        else {
            Write-Host "  No members found" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
