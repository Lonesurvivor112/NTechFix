Get-MgGroup -Filter "DisplayName eq 'Day Program'"

$users = Get-MgGroupMember -GroupId 1a1ae78c-59fd-4b1b-9cb0-94c2d8195175

$users | foreach  {Add-UnifiedGroupLinks -Identity 'Day Program' -Links $_.id -LinkType Subscribers}