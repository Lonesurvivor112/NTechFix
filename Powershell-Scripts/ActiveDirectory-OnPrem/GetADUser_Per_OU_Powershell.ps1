$OUpath = 'ou=Administration,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Administration.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Business Operations,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_BO.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Contract Clinical Staff,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Contract Clinical Staff.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Dining Hall,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Dining Hall.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=First Aid Trainers,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_FirstAidTrainers.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=HR,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_HR.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Information Technology,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_IT.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Jacksonville,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Jacksonville.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Maintenance,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Maint.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Manchester,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Manchester.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=MDT,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_MDT.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Nurses,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Nurses.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Programs,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Programs.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Phychologist,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_Phychologist.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Recreation Therapy,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_PT.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Residential Coordinators,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_RC.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Selfridge,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_selfridge.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath

$OUpath = 'ou=Supported Employement,ou=Users,ou=Contoso Corp,dc=contoso,dc=com'
$ExportPath = 'c:\data\users_in_RC.csv'
Get-ADUser -Filter * -SearchBase $OUpath | Select-object DistinguishedName, Name,UserPrincipalName | Export-Csv -NoType $ExportPath
