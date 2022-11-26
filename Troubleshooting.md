# Troubleshooting

## Microsoft Tokens have expired

1. Make sure you have the latest version of the code from the main repository. If you have forked your own version choose to sync with the main fork:

![image](https://user-images.githubusercontent.com/79275328/204080193-4983da88-6a25-4801-b756-63c6900bf4d5.png)

2. Once your repository is up to date make sure your Function app is running the latest version of your code. Go to https://portal.azure.com/#view/HubsExtension/BrowseResource/resourceType/Microsoft.Web%2Fsites/kind/functionapp and click on the name of your Hudu M365 function app:

![image](https://user-images.githubusercontent.com/79275328/204080687-16bd5dfe-167e-4a27-b517-2be9ec9da577.png)

3. Click on Deployment Centre on the left and then the sync button at the top:

![image](https://user-images.githubusercontent.com/79275328/204080657-0ff0960d-7636-4da9-82c0-e52f250e79e1.png)

4. Next go to Key Vault management and select your Hudu M365 Key Vault: https://portal.azure.com/#view/HubsExtension/BrowseResource/resourceType/Microsoft.KeyVault%2Fvaults

![image](https://user-images.githubusercontent.com/79275328/204080719-4aa52a66-977d-4109-ad49-b5e30acfc3cc.png)

5. Add your user account and give it all permissions on Secrets. For your Hudu M365 function app make sure it has Get, Set and List permissions on Secrets:

![image](https://user-images.githubusercontent.com/79275328/204080842-33f27fb2-deb7-48a2-bd2f-f8bcd8449d1f.png)

6. In the Key Vault under secrets obtain your Tenant ID, ApplicationID and ApplicationSecret. Put them into this script and run the script. Follow the instructions and make sure you login with an account with Microsoft MFA enabled.

```PowerShell
### User Input Variables ###

### Enter the details of your Secure Access Model Application below ###

$ApplicationId           = '<YOUR APPLICATION ID>'
$ApplicationSecret       = '<YOUR APPLICATION SECRET>'
$TenantId                = '<YOUR TENANT ID>'

### STOP EDITING HERE ###

### Create credential object using UserEntered(ApplicationID) and UserEntered(ApplicationSecret) ###

$Credential = New-Object System.Management.Automation.PSCredential($ApplicationId, ($ApplicationSecret | ConvertTo-SecureString -AsPlainText -Force))

### Splat Params required for Updating Refresh Token ###

$UpdateRefreshTokenParamaters = @{
    ApplicationID        = $ApplicationId
    Tenant               = $TenantId
    Scopes               = 'https://api.partnercenter.microsoft.com/user_impersonation'
    Credential           = $Credential
    UseAuthorizationCode = $true
    ServicePrincipal     = $true
}

### Splat Params required for Updating Exchange Refresh Token ###

$UpdateExchangeTokenParamaters = @{
    ApplicationID           = 'a0c73c16-a7e3-4564-9a95-2bdf47383716'
    Scopes                  = 'https://outlook.office365.com/.default'
    Tenant                  = $TenantId
    UseDeviceAuthentication = $true
}

### Create new Refresh Token using previously splatted paramaters ###

$Token = New-PartnerAccessToken @UpdateRefreshTokenParamaters

### Create new Exchange Refresh Token using previously splatted paramaters ###

$Exchangetoken = New-PartnerAccessToken @UpdateExchangeTokenParamaters

### Output Refresh Tokens and Exchange Refresh Tokens ###

Write-Host "================ Secrets ================"
Write-Host "`$ApplicationId         = $($ApplicationId)"
Write-Host "`$ApplicationSecret     = $($ApplicationSecret)"
Write-Host "`$TenantID              = $($TenantId)"
Write-Host "`$RefreshToken          = $($Token.refreshtoken)" -ForegroundColor Blue
Write-Host "`$ExchangeRefreshToken  = $($ExchangeToken.Refreshtoken)" -ForegroundColor Green
Write-Host "================ Secrets ================"
Write-Host "     SAVE THESE IN A SECURE LOCATION     "
```

6. Take the RefreshToken and ExchangeRefreshToken provided by the script and add new versions for each into your key vault.
7. Go back to the function app and choose to stop it.
8. Select Configuration on the left of the function app settings.
9. Rename the ExchangeRefreshToken setting and RefreshToken setting adding a 1 at the end of the name.
10. Start the function app, wait 3 minutes and then stop it again.
11. Rename the application settings back to their original name.
12. Wait 5 minutes.
13. Start the function app again and the tokens should be updated. In the future they should update themselves every week.

