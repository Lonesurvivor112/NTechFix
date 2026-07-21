set-aduser jsmith -Add @{proxyAddresses="SMTP:jsmith@contoso.com","smtp:jsmith@contoso.com"

Get-ADUser jsmith -Properties displayName,givenName,sn,userPrincipalName,mail,proxyAddresses,sAMAccountName,objectGUID,ms-DS-ConsistencyGuid |
 >    Select displayName,givenName,sn,userPrincipalName,mail,proxyAddresses,sAMAccountName,objectGUID,ms-DS-ConsistencyGuid |
 >    Format-List

displayName           : Jane Smith
givenName             : Jane
sn                    : Smith
userPrincipalName     : jsmith@contoso.com
mail                  : jsmith@contoso.com
proxyAddresses        : {}
sAMAccountName        : jsmith
objectGUID            : 00000000-0000-0000-0000-000000000000
ms-DS-ConsistencyGuid : {0, 0, 0, 0…}

 jdoe   SnipeIT PY
 SnipeIT PY  Get-ADUser jsmith -Properties * |
 >    Select Name,displayName,userPrincipalName,mail,proxyAddresses,sAMAccountName,
 >           objectGUID,'ms-DS-ConsistencyGuid',legacyExchangeDN,targetAddress,
 >           mailNickname,'msExchRecipientTypeDetails','msExchRemoteRecipientType' |
 >    Format-List

Name                       : Jane Doe
displayName                : Jane Doe
userPrincipalName          : jsmith@contoso.com
mail                       : jsmith@contoso.com
proxyAddresses             : {smtp:jdoe2@contoso.com}
sAMAccountName             : jsmith
objectGUID                 : 00000000-0000-0000-0000-000000000000
ms-DS-ConsistencyGuid      : {0, 0, 0, 0…}
legacyExchangeDN           :
targetAddress              :
mailNickname               :
msExchRecipientTypeDetails :
msExchRemoteRecipientType  :
