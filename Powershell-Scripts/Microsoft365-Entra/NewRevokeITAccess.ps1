foreach ($method in $authMethods) {
    Write-Host "Processing method: $($method.AdditionalProperties['@odata.type'])"
    switch ($method.AdditionalProperties['@odata.type']) {
        '#microsoft.graph.fido2AuthenticationMethod' {
            Write-Host "Removing FIDO2 method"
            Remove-MgUserAuthenticationFido2Method -UserId $userId -Fido2AuthenticationMethodId $method.Id
        }
        '#microsoft.graph.emailAuthenticationMethod' {
            Write-Host "Removing email method"
            Remove-MgUserAuthenticationEmailMethod -UserId $userId -EmailAuthenticationMethodId $method.Id
        }
        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
            Write-Host "Removing Microsoft Authenticator method"
            Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $userId -MicrosoftAuthenticatorAuthenticationMethodId $method.Id
        }
        '#microsoft.graph.phoneAuthenticationMethod' {
            Write-Host "Removing phone method"
            Remove-MgUserAuthenticationPhoneMethod -UserId $userId -PhoneAuthenticationMethodId $method.Id
        }
        '#microsoft.graph.softwareOathAuthenticationMethod' {
            Write-Host "Removing software OATH method"
            Remove-MgUserAuthenticationSoftwareOathMethod -UserId $userId -SoftwareOathAuthenticationMethodId $method.Id
        }
        '#microsoft.graph.temporaryAccessPassAuthenticationMethod' {
            Write-Host "Removing temporary access pass method"
            Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $userId -TemporaryAccessPassAuthenticationMethodId $method.Id
        }
        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
            Write-Host "Removing Windows Hello for Business method"
            Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $userId -WindowsHelloForBusinessAuthenticationMethodId $method.Id
        }
        default {
            Write-Host "This script does not handle removing this auth method type: $($method.AdditionalProperties['@odata.type'])"
        }
    }
}