# Input bindings are passed in via param block.
param($Timer)

Write-Host "Starting Token Refresh"
# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

$Refreshtoken = (Get-GraphToken -ReturnRefresh $true).Refresh_token
$ExchangeRefreshtoken = (Get-GraphToken -AppID 'a0c73c16-a7e3-4564-9a95-2bdf47383716' -refreshtoken $ENV:ExchangeRefreshtoken -ReturnRefresh $true).Refresh_token

if ($env:MSI_SECRET) {
    Disable-AzContextAutosave -Scope Process | Out-Null
    Write-Host "Connecting to Azure AD"
    $AzSession = Connect-AzAccount -Identity
}

Write-Host "Getting Key Vault"
$KV = $ENV:WEBSITE_DEPLOYMENT_ID

if ($Refreshtoken -and $KV) { 
    Set-AzKeyVaultSecret -VaultName $kv -Name 'RefreshToken' -SecretValue (ConvertTo-SecureString -String $Refreshtoken -AsPlainText -Force)
}
else { Write-Host "Could not update refresh token. Will try again in 7 days."}
if ($ExchangeRefreshtoken -and $KV) {
    Set-AzKeyVaultSecret -VaultName $kv -Name 'ExchangeRefreshToken' -SecretValue (ConvertTo-SecureString -String $ExchangeRefreshtoken -AsPlainText -Force)
    Write-Host "System API: Updated Exchange Refresh token."
}
else {
    Write-Host "Could not update Exchange refresh token. Will try again in 7 days."
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
