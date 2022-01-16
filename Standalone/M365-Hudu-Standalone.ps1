# M365 to Hudu Syncronisation Script
# For configuration details and an Azure Function version of the script please visit here: https://github.com/lwhitelock/HuduM365Automation
# This project utilizes some of the helper functions written by Kelvin Tegelaar from the CIPP project https://github.com/KelvinTegelaar/CIPP and is licensed under the same terms.

#### Settings ####
# Your Azure Keyvault name
$VaultName = "Your Key Vault"
# Microsoft Secure Application Model Info
$customerExclude = (Get-AzKeyVaultSecret -vaultName $VaultName -name "customerExclude" -AsPlainText) -split ',' 
$script:ApplicationId = Get-AzKeyVaultSecret -vaultName $VaultName -name "ApplicationID" -AsPlainText
$script:ApplicationSecret = Get-AzKeyVaultSecret -vaultName $VaultName -name "ApplicationSecret" -AsPlainText
$script:TenantID = Get-AzKeyVaultSecret -vaultName $VaultName -name "TenantID" -AsPlainText
$script:RefreshToken = Get-AzKeyVaultSecret -vaultName $VaultName -name "RefreshToken"-AsPlainText
$script:ExchangeRefreshToken = Get-AzKeyVaultSecret -vaultName $VaultName -name "ExchangeRefreshToken"-AsPlainText
$script:UPN = Get-AzKeyVaultSecret -vaultName $VaultName -name "UPN" -AsPlainText

#### Hudu Settings ####
$HuduAPIKey = Get-AzKeyVaultSecret -vaultName $VaultName -name "HuduAPIKey" -AsPlainText
# Set the base domain of your Hudu instance without a trailing /
$HuduBaseDomain = Get-AzKeyVaultSecret -vaultName $VaultName -name "HuduBaseDomain" -AsPlainText

##########################          Settings         ############################
# Enable Verbose Mode Uncomment the below line
#$VerbosePreference = "continue"

$PeopleLayoutName = "People"
# If this is enabled users not in Hudu will be created if they don't exist
$CreateUsers = $True

# Set the asset layout names Asset layout names
$DesktopsName = "Desktops / Laptops"
$MobilesName = "Mobile Devices"
# If this is enabled devices will be created in Hudu if they don't exist
$CreateDevices = $True
$CreateMobileDevices = $True

$CreateInOverview = $true
$OverviewCompany = 'Overview - M365'

#This will toggle on and off importing domains from M365 to Hudu
$importDomains = $true

#For imported domains this will set if monitoring is enabled or disabled
$monitorDomains = $true

# Document your own tenant settings
$DocumentPartnerTenant = $True
$PartnerTenant = [PSCustomObject]@{
    customerId        = $script:TenantID
    defaultDomainName = 'yourdomain.onmicrosoft.com'
    displayName       = 'Your Partner Name'
}

# Intune Desktop / Laptop Device Types
$IntuneDesktopDeviceTypes = @('windowsRT', 'macMDM')

# Serial numbers that should be excluded
$ExcludeSerials = @("0", "SystemSerialNumber", "To Be Filled By O.E.M.", "System Serial Number", "0123456789", "123456789", "............")

# PSA User URL (Only supports Halo PSA at the moment)
$PSAUserURL = 'https://your.halopsadomain.com/customers?mainview=user&inactive=false&userid='

# RMM Device URL (Only Supports Datto RMM at the moment)
$RMMDeviceURL = 'https://merlotrmm.centrastage.net/device/'

# RMM Remote Access URL (Only Supports Datto RMM at the moment)
$RMMRemoteURL = 'https://merlot.centrastage.net/csm/remote/rto/'

##########################      End Settings       ############################

####### License Lookup Hash #########
$LicenseLookup = @{
    'SPZA_IW'                                 = 'App Connect Iw'
    'AAD_BASIC'                               = 'Azure Active Directory Basic'
    'AAD_PREMIUM'                             = 'Azure Active Directory Premium P1'
    'AAD_PREMIUM_P2'                          = 'Azure Active Directory Premium P2'
    'RIGHTSMANAGEMENT'                        = 'Azure Information Protection Plan 1'
    'MCOCAP'                                  = 'Common Area Phone'
    'MCOPSTNC'                                = 'Communications Credits'
    'DYN365_ENTERPRISE_PLAN1'                 = 'Dynamics 365 Customer Engagement Plan Enterprise Edition'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE'      = 'Dynamics 365 For Customer Service Enterprise Edition'
    'DYN365_FINANCIALS_BUSINESS_SKU'          = 'Dynamics 365 For Financials Business Edition'
    'DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE' = 'Dynamics 365 For Sales And Customer Service Enterprise Edition'
    'DYN365_ENTERPRISE_SALES'                 = 'Dynamics 365 For Sales Enterprise Edition'
    'DYN365_ENTERPRISE_TEAM_MEMBERS'          = 'Dynamics 365 For Team Members Enterprise Edition'
    'DYN365_TEAM_MEMBERS'                     = 'Dynamics 365 Team Members'
    'Dynamics_365_for_Operations'             = 'Dynamics 365 Unf Ops Plan Ent Edition'
    'EMS'                                     = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                              = 'Enterprise Mobility + Security E5'
    'EXCHANGESTANDARD'                        = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                      = 'Exchange Online (Plan 2)'
    'EXCHANGEARCHIVE_ADDON'                   = 'Exchange Online Archiving For Exchange Online'
    'EXCHANGEARCHIVE'                         = 'Exchange Online Archiving For Exchange Server'
    'EXCHANGEESSENTIALS'                      = 'Exchange Online Essentials'
    'EXCHANGE_S_ESSENTIALS'                   = 'Exchange Online Essentials'
    'EXCHANGEDESKLESS'                        = 'Exchange Online Kiosk'
    'EXCHANGETELCO'                           = 'Exchange Online Pop'
    'INTUNE_A'                                = 'Intune'
    'M365EDU_A1'                              = 'Microsoft 365 A1'
    'M365EDU_A3_FACULTY'                      = 'Microsoft 365 A3 For Faculty'
    'M365EDU_A3_STUDENT'                      = 'Microsoft 365 A3 For Students'
    'M365EDU_A5_FACULTY'                      = 'Microsoft 365 A5 For Faculty'
    'M365EDU_A5_STUDENT'                      = 'Microsoft 365 A5 For Students'
    'O365_BUSINESS'                           = 'Microsoft 365 Apps For Business'
    'SMB_BUSINESS'                            = 'Microsoft 365 Apps For Business'
    'OFFICESUBSCRIPTION'                      = 'Microsoft 365 Apps For Enterprise'
    'MCOMEETADV'                              = 'Microsoft 365 Audio Conferencing'
    'MCOMEETADV_GOC'                          = 'Microsoft 365 Audio Conferencing For Gcc'
    'O365_BUSINESS_ESSENTIALS'                = 'Microsoft 365 Business Basic'
    'SMB_BUSINESS_ESSENTIALS'                 = 'Microsoft 365 Business Basic'
    'SPB'                                     = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'                   = 'Microsoft 365 Business Standard'
    'SMB_BUSINESS_PREMIUM'                    = 'Microsoft 365 Business Standard'
    'MCOPSTN_5'                               = 'Microsoft 365 Domestic Calling Plan (120 Minutes)'
    'SPE_E3'                                  = 'Microsoft 365 E3'
    'SPE_E3_USGOV_DOD'                        = 'Microsoft 365 E3_Usgov_Dod'
    'SPE_E3_USGOV_GCCHIGH'                    = 'Microsoft 365 E3_Usgov_Gcchigh'
    'SPE_E5'                                  = 'Microsoft 365 E5'
    'INFORMATION_PROTECTION_COMPLIANCE'       = 'Microsoft 365 E5 Compliance'
    'IDENTITY_THREAT_PROTECTION'              = 'Microsoft 365 E5 Security'
    'IDENTITY_THREAT_PROTECTION_FOR_EMS_E5'   = 'Microsoft 365 E5 Security For Ems E5'
    'M365_F1'                                 = 'Microsoft 365 F1'
    'SPE_F1'                                  = 'Microsoft 365 F3'
    'M365_G3_GOV'                             = 'Microsoft 365 Gcc G3'
    'MCOEV'                                   = 'Microsoft 365 Phone System'
    'PHONESYSTEM_VIRTUALUSER'                 = 'Microsoft 365 Phone System - Virtual User'
    'MCOEV_DOD'                               = 'Microsoft 365 Phone System For Dod'
    'MCOEV_FACULTY'                           = 'Microsoft 365 Phone System For Faculty'
    'MCOEV_GOV'                               = 'Microsoft 365 Phone System For Gcc'
    'MCOEV_GCCHIGH'                           = 'Microsoft 365 Phone System For Gcchigh'
    'MCOEVSMB_1'                              = 'Microsoft 365 Phone System For Small And Medium Business'
    'MCOEV_STUDENT'                           = 'Microsoft 365 Phone System For Students'
    'MCOEV_TELSTRA'                           = 'Microsoft 365 Phone System For Telstra'
    'MCOEV_USGOV_DOD'                         = 'Microsoft 365 Phone System_Usgov_Dod'
    'MCOEV_USGOV_GCCHIGH'                     = 'Microsoft 365 Phone System_Usgov_Gcchigh'
    'WIN_DEF_ATP'                             = 'Microsoft Defender Advanced Threat Protection'
    'CRMSTANDARD'                             = 'Microsoft Dynamics Crm Online'
    'CRMPLAN2'                                = 'Microsoft Dynamics Crm Online Basic'
    'FLOW_FREE'                               = 'Microsoft Flow Free'
    'INTUNE_A_D_GOV'                          = 'Microsoft Intune Device For Government'
    'POWERAPPS_VIRAL'                         = 'Microsoft Power Apps Plan 2 Trial'
    'TEAMS_FREE'                              = 'Microsoft Team (Free)'
    'TEAMS_EXPLORATORY'                       = 'Microsoft Teams Exploratory'
    'IT_ACADEMY_AD'                           = 'Ms Imagine Academy'
    'ENTERPRISEPREMIUM_FACULTY'               = 'Office 365 A5 For Faculty'
    'ENTERPRISEPREMIUM_STUDENT'               = 'Office 365 A5 For Students'
    'EQUIVIO_ANALYTICS'                       = 'Office 365 Advanced Compliance'
    'ATP_ENTERPRISE'                          = 'Microsoft Defender for Office 365 (Plan 1)'
    'STANDARDPACK'                            = 'Office 365 E1'
    'STANDARDWOFFPACK'                        = 'Office 365 E2'
    'ENTERPRISEPACK'                          = 'Office 365 E3'
    'DEVELOPERPACK'                           = 'Office 365 E3 Developer'
    'ENTERPRISEPACK_USGOV_DOD'                = 'Office 365 E3_Usgov_Dod'
    'ENTERPRISEPACK_USGOV_GCCHIGH'            = 'Office 365 E3_Usgov_Gcchigh'
    'ENTERPRISEWITHSCAL'                      = 'Office 365 E4'
    'ENTERPRISEPREMIUM'                       = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'            = 'Office 365 E5 Without Audio Conferencing'
    'DESKLESSPACK'                            = 'Office 365 F3'
    'ENTERPRISEPACK_GOV'                      = 'Office 365 Gcc G3'
    'MIDSIZEPACK'                             = 'Office 365 Midsize Business'
    'LITEPACK'                                = 'Office 365 Small Business'
    'LITEPACK_P2'                             = 'Office 365 Small Business Premium'
    'WACONEDRIVESTANDARD'                     = 'Onedrive For Business (Plan 1)'
    'WACONEDRIVEENTERPRISE'                   = 'Onedrive For Business (Plan 2)'
    'POWER_BI_STANDARD'                       = 'Power Bi (Free)'
    'POWER_BI_ADDON'                          = 'Power Bi For Office 365 Add-On'
    'POWER_BI_PRO'                            = 'Power Bi Pro'
    'PROJECTCLIENT'                           = 'Project For Office 365'
    'PROJECTESSENTIALS'                       = 'Project Online Essentials'
    'PROJECTPREMIUM'                          = 'Project Online Premium'
    'PROJECTONLINE_PLAN_1'                    = 'Project Online Premium Without Project Client'
    'PROJECTPROFESSIONAL'                     = 'Microsoft Project Plan 3'
    'PROJECTONLINE_PLAN_2'                    = 'Project Online With Project For Office 365'
    'SHAREPOINTSTANDARD'                      = 'Sharepoint Online (Plan 1)'
    'SHAREPOINTENTERPRISE'                    = 'Sharepoint Online (Plan 2)'
    'MCOIMP'                                  = 'Skype For Business Online (Plan 1)'
    'MCOSTANDARD'                             = 'Skype For Business Online (Plan 2)'
    'MCOPSTN2'                                = 'Skype For Business Pstn Domestic And International Calling'
    'MCOPSTN1'                                = 'Skype For Business Pstn Domestic Calling'
    'MCOPSTN5'                                = 'Skype For Business Pstn Domestic Calling (120 Minutes)'
    'MCOPSTNEAU2'                             = 'Telstra Calling For O365'
    'TOPIC_EXPERIENCES'                       = 'Topic Experiences'
    'VISIOONLINE_PLAN1'                       = 'Visio Online Plan 1'
    'VISIOCLIENT'                             = 'Visio Online Plan 2'
    'VISIOCLIENT_GOV'                         = 'Visio Plan 2 For Gov'
    'WIN10_PRO_ENT_SUB'                       = 'Windows 10 Enterprise E3'
    'WIN10_VDA_E3'                            = 'Windows 10 Enterprise E3'
    'WIN10_VDA_E5'                            = 'Windows 10 Enterprise E5'
    'WINDOWS_STORE'                           = 'Windows Store For Business'
    'RMSBASIC'                                = 'Azure Information Protection Basic'
    'UNIVERSAL_PRINT_M365'                    = 'Universal Print'
    'RIGHTSMANAGEMENT_ADHOC'                  = 'Rights Management Service Basic Content Protection'
    'SKU_Dynamics_365_for_HCM_Trial'          = 'Dynamics 365 for Talent'
    'PROJECT_P1'                              = 'Project Plan 1'
    'PROJECT_PLAN1_DEPT'                      = 'Project Plan  1 (Self Service)'
    'SHAREPOINTSTORAGE'                       = 'Microsoft Office 365 Extra File Storage'
    'NONPROFIT_PORTAL'                        = 'Non Profit Portal'
    'MDE_SMB'                                 = 'Microsoft Defender for Endpoint (Business Premium)'
}

# Assigned Licenses Map
$AssignedMap = [pscustomobject]@{
    'AADPremiumService'             = 'o-skypeforbusiness'
    'MultiFactorService'            = 'o-skypeforbusiness'
    'RMSOnline'                     = 'o-skypeforbusiness'
    'MicrosoftPrint'                = 'o-yammer'
    'WindowsDefenderATP'            = 'o-skypeforbusiness'
    'exchange'                      = 'o-exchange'
    'ProcessSimple'                 = 'o-onedrive'
    'OfficeForms'                   = 'o-yammer'
    'SCO'                           = 'o-skypeforbusiness'
    'MicrosoftKaizala'              = 'o-yammer'
    'Adallom'                       = 'o-skypeforbusiness'
    'ProjectWorkManagement'         = 'o-yammer'
    'TeamspaceAPI'                  = 'o-teams'
    'MicrosoftOffice'               = 'o-yammer'
    'PowerAppsService'              = 'o-onedrive'
    'SharePoint'                    = 'o-sharepoint'
    'MicrosoftCommunicationsOnline' = 'o-teams'
    'Deskless'                      = 'o-yammer'
    'MicrosoftStream'               = 'o-yammer'
    'Sway'                          = 'o-yammer'
    'To-Do'                         = 'o-yammer'
    'WhiteboardServices'            = 'o-yammer'
    'Windows'                       = 'o-skypeforbusiness'
    'YammerEnterprise'              = 'o-yammer'
}

$AssignedNameMap = @{
    'AADPremiumService'             = 'Azure Active Directory Premium'
    'MultiFactorService'            = 'Azure Multi-Factor Authentication'
    'RMSOnline'                     = 'Azure Rights Management'
    'MicrosoftPrint'                = 'Cloud Print'
    'WindowsDefenderATP'            = 'Defender for Endpoint'
    'exchange'                      = 'Exchange Online'
    'ProcessSimple'                 = 'Flow'
    'OfficeForms'                   = 'Forms'
    'SCO'                           = 'Intune'
    'MicrosoftKaizala'              = 'Kaizala'
    'Adallom'                       = 'Microsoft Cloud App Security'
    'ProjectWorkManagement'         = 'Microsoft Planner'
    'TeamspaceAPI'                  = 'Microsoft Teams'
    'MicrosoftOffice'               = 'Office 365'
    'PowerAppsService'              = 'PowerApps'
    'SharePoint'                    = 'SharePoint Online'
    'MicrosoftCommunicationsOnline' = 'Skype for Business'
    'Deskless'                      = 'Staff Hub'
    'MicrosoftStream'               = 'Stream'
    'Sway'                          = 'Sway'
    'To-Do'                         = 'To-Do'
    'WhiteboardServices'            = 'Whiteboard'
    'Windows'                       = 'Windows'
    'YammerEnterprise'              = 'Yammer'
}

### Functions ###
### These are some helper functions borrowed from Kelvin Tegelaar's CIPP project https://github.com/KelvinTegelaar/CIPP
function New-GraphGetRequest ($uri, $tenantid, $scope, $AsApp, $noPagination) {

    if ($scope -eq "ExchangeOnline") { 
        $Headers = $Script:ExchangeAuthheaders
    } else {
        $headers = $Script:Authheaders
    }
    Write-Verbose "Using $($uri) as url"
    $nextURL = $uri
    $ReturnedData = do {
        try {
            $Data = (Invoke-RestMethod -Uri $nextURL -Method GET -Headers $headers -ContentType "application/json; charset=utf-8")
            if ($data.value) { $data.value } else { ($Data) }
            if ($noPagination) { $nextURL = $null } else { $nextURL = $data.'@odata.nextLink' }                
        } catch {
            $Message = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message
            if ($null -eq $Message) { $Message = $($_.Exception.Message) }
            throw $Message
        }
    } until ($null -eq $NextURL)
   
    return $ReturnedData   

}

function Get-GraphToken($tenantid, $scope, $AsApp, $AppID, $refreshToken, $ReturnRefresh) {
    if (!$scope) { $scope = 'https://graph.microsoft.com/.default' }

    $AuthBody = @{
        client_id     = $script:ApplicationId
        client_secret = $script:ApplicationSecret
        scope         = $Scope
        refresh_token = $script:RefreshToken
        grant_type    = "refresh_token"
                    
    }
    if ($asApp -eq $true) {
        $AuthBody = @{
            client_id     = $script:ApplicationId
            client_secret = $script:ApplicationSecret
            scope         = $Scope
            grant_type    = "client_credentials"
        }
    }

    if ($null -ne $AppID -and $null -ne $refreshToken) {
        $AuthBody = @{
            client_id     = $appid
            refresh_token = $RefreshToken
            scope         = $Scope
            grant_type    = "refresh_token"
        }
    }

    if (!$tenantid) { $tenantid = $script:tenantid }
    $AccessToken = (Invoke-RestMethod -Method post -Uri "https://login.microsoftonline.com/$($tenantid)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop)
    if ($ReturnRefresh) { $header = $AccessToken } else { $header = @{ Authorization = "Bearer $($AccessToken.access_token)" } }

    return $header
}

function New-ClassicAPIPostRequest($TenantID, $Uri, $Method = 'POST', $Resource = 'https://admin.microsoft.com', $Body) {
    $token = Get-ClassicAPIToken -Tenant $tenantID -Resource $Resource
    try {
        $ReturnedData = Invoke-RestMethod -ContentType "application/json;charset=UTF-8" -Uri $Uri -Method $Method -Body $Body -Headers @{
            Authorization            = "Bearer $($token.access_token)";
            "x-ms-client-request-id" = [guid]::NewGuid().ToString();
            "x-ms-client-session-id" = [guid]::NewGuid().ToString()
            'x-ms-correlation-id'    = [guid]::NewGuid()
            'X-Requested-With'       = 'XMLHttpRequest' 
        } 
                       
    } catch {
        throw "Failed to make Classic Get Request $_"
    }
    return $ReturnedData
}

function Get-ClassicAPIToken($tenantID, $Resource) {
    $uri = "https://login.microsoftonline.com/$($TenantID)/oauth2/token"
    $body = "resource=$Resource&grant_type=refresh_token&refresh_token=$($script:ExchangeRefreshToken)"
    try {
        $token = Invoke-RestMethod $uri -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction SilentlyContinue -Method post
        return $token
    } catch {
        Write-Error "Failed to obtain Classic API Token for $Tenant - $_"        
    }
}

function New-ExoRequest ($tenantid, $cmdlet, $cmdParams) {
    $Headers = Get-GraphToken -AppID 'a0c73c16-a7e3-4564-9a95-2bdf47383716' -RefreshToken $script:ExchangeRefreshToken -Scope 'https://outlook.office365.com/.default' -Tenantid $tenantid 
    $tenant = $tenantid
    if ($cmdParams) {
        $Params = $cmdParams
    } else {
        $Params = @{}
    }
    $ExoBody = @{
        CmdletInput = @{
            CmdletName = $cmdlet
            Parameters = $Params
        }
    } | ConvertTo-Json
    $ReturnedData = Invoke-RestMethod "https://outlook.office365.com/adminapi/beta/$($tenant)/InvokeCommand" -Method POST -Body $ExoBody -Headers $Headers -ContentType "application/json; charset=utf-8"
    return $ReturnedData.value   
    
}

function New-GraphBulkRequest ($Requests, $tenantid) {

    $headers = Get-GraphToken -tenantid $tenantid
    $URL = 'https://graph.microsoft.com/beta/$batch'

    $ReturnedData = for ($i = 0; $i -lt $Requests.count; $i += 20) {                                                                                                                                              
        $req = @{}                
        # Use select to create hashtables of id, method and url for each call                                     
        $req['requests'] = ($Requests[$i..($i + 19)])
        Invoke-RestMethod -Uri $URL -Method POST -Headers $headers -ContentType "application/json; charset=utf-8" -Body ($req | convertto-json)                                                                                                                                
    }
    
    foreach ($MoreData in $ReturnedData.Responses | Where-Object { $_.body.'@odata.nextLink' }) {
        $AdditionalValues = New-GraphGetRequest -uri $MoreData.body.'@odata.nextLink' -tenantid $TenantFilter
        $NewValues = [System.Collections.Generic.List[PSCustomObject]]$MoreData.body.value
        $AdditionalValues | foreach-object { $NewValues.add($_) }
        $MoreData.body.value = $NewValues
    }

    return $ReturnedData.responses
}

function Get-BulkResultByID ($Results, $ID) {
    ($Results | Where-Object { $_.id -eq $ID }).body.value
}

function Get-FormatedField ($Title, $Value) {
    return @"
<div class="card__item">
    <div class="card__item-slot">
        $Title
    </div>
    <div class="card__item-slot">
        $Value
    </div>
</div>
"@
}

function Get-FormattedBlock ($Heading, $Body) {
    return @"
<div class="nasa__block" style="margin-bottom: 20px;">
    <header class='nasa__block-header' style="padding-top: 15px;">
            <h1>$Heading</h1>
        </header>
        <div style="padding-left: 15px; padding-right: 15px; padding-bottom: 15px;">
        $Body
    </div>
</div>
"@
}

function Get-LinkBlock($URL, $Icon, $Title) {
    return "<div class='o365__app' style='text-align:center'><a href=$URL target=_blank><h3><i class=`"$Icon`">&nbsp;&nbsp;&nbsp;</i>$Title</h3></a></div>"
}

### Start ###

import-module HuduAPI

#Login to Hudu
New-HuduAPIKey $HuduAPIKey
New-HuduBaseUrl $HuduBaseDomain

$Script:Authheaders = Get-GraphToken -tenantid $script:Tenantid

# Get Customers
[System.Collections.Generic.List[PSCustomObject]]$Customers = (New-GraphGetRequest -uri "https://graph.microsoft.com/beta/contracts?`$top=999" -tenantid $script:Tenantid) | Select-Object CustomerID, DefaultdomainName, DisplayName | Where-Object -Property DisplayName -NotIn $customerExclude

if ($DocumentPartnerTenant -eq $true) {
    $Customers.add($PartnerTenant)
}

foreach ($Customer in $Customers) {
    $TenantFilter = $Customer.CustomerId
    write-host "$(Get-Date) - #############################################"
    write-host "$(Get-Date) - Starting $($customer.DisplayName)"
    #Check if they are in Hudu before doing any unnessisary work
    $defaultdomain = $customer.DefaultDomainName
    $hududomain = Get-HuduWebsites -name "https://$defaultdomain"
    if ($($hududomain.id.count) -gt 0) {
        $company_name = $hududomain[0].company_name
        $company_id = $hududomain[0].company_id

        # Get Auth Headers
        $Script:ExchangeAuthHeaders = Get-GraphToken -AppID 'a0c73c16-a7e3-4564-9a95-2bdf47383716' -RefreshToken $script:ExchangeRefreshToken -Scope 'https://outlook.office365.com/.default' -Tenantid $TenantFilter
        $Script:Authheaders = Get-GraphToken -tenantid $TenantFilter

        write-verbose "$(Get-Date) - Fetching Hudu Details"
        $PeopleLayout = Get-HuduAssetLayouts -name $PeopleLayoutName
        $People = Get-HuduAssets -companyid $company_id -assetlayoutid $PeopleLayout.id

        $DesktopsLayout = Get-HuduAssetLayouts -name $DesktopsName
        $HuduDesktopDevices = Get-HuduAssets -companyid $company_id -assetlayoutid $DesktopsLayout.id

        $MobilesLayout = Get-HuduAssetLayouts -name $MobilesName
        $HuduMobileDevices = Get-HuduAssets -companyid $company_id -assetlayoutid $MobilesLayout.id

        $HuduDevices = $HuduDesktopDevices + $HuduMobileDevices
		
        #Create a table to send into Hudu
        $CustomerLinks = "<div class=`"nasa__content`"> 
        <div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://portal.office.com/Partner/BeginClientSession.aspx?CTID=$($customer.CustomerId)&CSDEST=o365admincenter')`"><h3><i class=`"fas fa-cogs`">&nbsp;&nbsp;&nbsp;</i>M365 Admin Portal</h3></button></div>
        <div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://outlook.office365.com/ecp/?rfr=Admin_o365&exsvurl=1&delegatedOrg=$($Customer.DefaultDomainName)')`"><h3><i class=`"fas fa-mail-bulk`">&nbsp;&nbsp;&nbsp;</i>Exchange Admin Portal</h3></button></div>
        <div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://aad.portal.azure.com/$($Customer.DefaultDomainName)')`" ><h3><i class=`"fas fa-users-cog`">&nbsp;&nbsp;&nbsp;</i>Azure Active Directory</h3></button></div>
		<div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://endpoint.microsoft.com/$($customer.DefaultDomainName)/')`"><h3><i class=`"fas fa-laptop`">&nbsp;&nbsp;&nbsp;</i>Endpoint Management</h3></button></td></div>
									
		<div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://portal.office.com/Partner/BeginClientSession.aspx?CTID=$($Customer.CustomerId)&CSDEST=MicrosoftCommunicationsOnline')`"><h3><i class=`"fab fa-skype`">&nbsp;&nbsp;&nbsp;</i>Sfb Portal</h3></button></div>
        <div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://admin.teams.microsoft.com/?delegatedOrg=$($Customer.DefaultDomainName)')`"><h3><i class=`"fas fa-users`">&nbsp;&nbsp;&nbsp;</i>Teams Portal</h3></button></div>
        <div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://portal.azure.com/$($customer.DefaultDomainName)')`"><h3><i class=`"fas fa-server`">&nbsp;&nbsp;&nbsp;</i>Azure Portal</h3></button></div>
        <div class=`"nasa__block`"><button class=`"button`" onclick=`"window.open('https://account.activedirectory.windowsazure.com/usermanagement/multifactorverification.aspx?tenantId=$($Customer.CustomerId)&culture=en-us&requestInitiatedContext=users')`" ><h3><i class=`"fas fa-key`">&nbsp;&nbsp;&nbsp;</i>MFA Portal (Read Only)</h3></button></div>
		
		</div>"
        
        # Build bulk requests array.
        [System.Collections.Generic.List[PSCustomObject]]$TenantRequests = @(
            @{
                id     = 'Users'
                method = 'GET'
                url    = '/users'
            },
            @{
                id     = 'AllRoles'
                method = 'GET'
                url    = '/directoryRoles'
            },
            @{
                id     = 'RawDomains'
                method = 'GET'
                url    = '/domains'
            },
            @{
                id     = 'Licenses'
                method = 'GET'
                url    = '/subscribedSkus'
            },
            @{
                id     = 'Devices'
                method = 'GET'
                url    = '/deviceManagement/managedDevices'
            },
            @{
                id     = 'DeviceCompliancePolicies'
                method = 'GET'
                url    = '/deviceManagement/deviceCompliancePolicies/'
            },
            @{
                id     = 'DeviceApps'
                method = 'GET'
                url    = '/deviceAppManagement/mobileApps'
            },
            @{
                id     = 'Groups'
                method = 'GET'
                url    = '/groups'
            },
            @{
                id     = 'ConditionalAccess'
                method = 'GET'
                url    = '/identity/conditionalAccess/policies'
            }
                
        )

        write-verbose "$(Get-Date) - Fetching Bulk Data"
        try {
            $TenantResults = New-GraphBulkRequest -Requests $TenantRequests -tenantid $TenantFilter
        } catch {
            Write-Error "Failed to fetch bulk company data"
        }

        $Users = Get-BulkResultByID -Results $TenantResults -ID 'Users'

        write-verbose "$(Get-Date) - Parsing Users"
        # Grab licensed users	
        $licensedUsers = $Users | where-object { $null -ne $_.AssignedLicenses.SkuId } | Sort-Object UserPrincipalName			
            
        write-verbose "$(Get-Date) - Parsing Roles"    
        # Get All Roles
        $AllRoles = Get-BulkResultByID -Results $TenantResults -ID 'AllRoles'
            
        $SelectList = 'id', 'displayName', 'userPrincipalName'

        [System.Collections.Generic.List[PSCustomObject]]$RolesRequestArray = @()
        foreach ($Role in $AllRoles) {
            $RolesRequestArray.add(@{
                    id     = $Role.id
                    method = 'GET'
                    url    = "/directoryRoles/$($Role.id)/members?`$select=$($selectlist -join ',')"
                })
        }

        try {
            $MemberReturn = New-GraphBulkRequest -Requests $RolesRequestArray -tenantid $TenantFilter
        } catch {
            $MemberReturn = $null
        }

        $Roles = foreach ($Result in $MemberReturn) {
            [PSCustomObject]@{
                ID            = $Result.id
                DisplayName   = ($AllRoles | where-object { $_.id -eq $Result.id }).displayName
                Description   = ($AllRoles | where-object { $_.id -eq $Result.id }).description
                Members       = $Result.body.value
                ParsedMembers = $Result.body.value.Displayname -join ', '
            }
        }
         

        write-verbose "$(Get-Date) - Building Magic Dash"

        $pre = "<div class=`"nasa__block`"><header class='nasa__block-header'>
			<h1><i class='fas fa-users icon'></i>Assigned Roles</h1>
			 </header>"

        $post = "</div>"
        $RolesHtml = $Roles | Select-Object DisplayName, Description, ParsedMembers | ConvertTo-Html -PreContent $pre -PostContent $post -Fragment | ForEach-Object { $tmp = $_ -replace "&lt;", "<"; $tmp -replace "&gt;", ">"; } | Out-String

        $AdminUsers = (($Roles | Where-Object { $_.Displayname -match "Administrator" }).Members | where-object { $null -ne $_.displayName } | Select-Object @{N = 'Name'; E = { "<a target='_blank' href='https://aad.portal.azure.com/$($Customer.DefaultDomainName)/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/$($_.Id)'>$($_.DisplayName) - $($_.UserPrincipalName)</a>" } } -unique).name -join "<br/>"
            
        write-verbose "$(Get-Date) - Fetching Domains"
        try {
            $RawDomains = Get-BulkResultByID -Results $TenantResults -ID 'RawDomains'
        } catch {
            $RawDomains = $null
        }
        $customerDomains = ($RawDomains | Where-Object { $_.IsVerified -eq $True }).id -join ', ' | Out-String
        
        $detailstable = "<div class='nasa__block'>
							<header class='nasa__block-header'>
							<h1><i class='fas fa-info-circle icon'></i>Basic Info</h1>
							 </header>
								<main>
								<article>
								<div class='basic_info__section'>
								<h2>Tenant Name</h2>
								<p>
									$($customer.DisplayName)
								</p>
								</div>
								<div class='basic_info__section'>
								<h2>Tenant ID</h2>
								<p>
									$($customer.customerId)
								</p>
								</div>
								<div class='basic_info__section'>
								<h2>Default Domain</h2>
								<p>
									$defaultdomain
								</p>
								</div>
								<div class='basic_info__section'>
								<h2>Customer Domains</h2>
								<p>
									$customerDomains
								</p>
								</div>
								<div class='basic_info__section'>
								<h2>Admin Users</h2>
								<p>
									$AdminUsers
								</p>
								</div>
						</article>
						</main>
						</div>
"       
        write-verbose "$(Get-Date) - Parsing Licenses"
        # Get Licenses
        $Licenses = Get-BulkResultByID -Results $TenantResults -ID 'Licenses'

        # Get the license overview for the tenant
        if ($Licenses) {
            $pre = "<div class=`"nasa__block`"><header class='nasa__block-header'>
			<h1><i class='fas fa-info-circle icon'></i>Current Licenses</h1>
			 </header>"
			
            $post = "</div>"

            $licenseOut = $Licenses | where-object { $_.PrepaidUnits.Enabled -gt 0 } | Select-Object @{N = 'License Name'; E = { $($LicenseLookup.$($_.SkuPartNumber)) } }, @{N = 'Active'; E = { $_.PrepaidUnits.Enabled } }, @{N = 'Consumed'; E = { $_.ConsumedUnits } }, @{N = 'Unused'; E = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits } }
            $licenseHTML = $licenseOut | ConvertTo-Html -PreContent $pre -PostContent $post -Fragment | Out-String
        }
        
        write-verbose "$(Get-Date) - Parsing Devices"
        # Get all devices from Intune
        $devices = Get-BulkResultByID -Results $TenantResults -ID 'Devices'

        write-verbose "$(Get-Date) - Parsing Device Compliance Polcies"
        # Fetch Compliance Policy Status
        $DeviceCompliancePolicies = Get-BulkResultByID -Results $TenantResults -ID 'DeviceCompliancePolicies'
           
        # Get the status of each device for each policy
        [System.Collections.Generic.List[PSCustomObject]]$PolicyRequestArray = @()
        foreach ($CompliancePolicy in $DeviceCompliancePolicies) {
            $PolicyRequestArray.add(@{
                    id     = $CompliancePolicy.id
                    method = 'GET'
                    url    = "/deviceManagement/deviceCompliancePolicies/$($CompliancePolicy.id)/deviceStatuses"
                })
        }

        try {
            $PolicyReturn = New-GraphBulkRequest -Requests $PolicyRequestArray -tenantid $TenantFilter
        } catch {
            $PolicyReturn = $null
        }

        $DeviceComplianceDetails = foreach ($Result in $PolicyReturn) {
            [pscustomobject]@{
                ID             = ($DeviceCompliancePolicies | where-object { $_.id -eq $Result.id }).id
                DisplayName    = ($DeviceCompliancePolicies | where-object { $_.id -eq $Result.id }).DisplayName
                DeviceStatuses = $Result.body.value
            }
        }
            
        write-verbose "$(Get-Date) - Parsing Apps"
        # Fetch Apps  
        $DeviceApps = Get-BulkResultByID -Results $TenantResults -ID 'DeviceApps'

        # Fetch the App status for each device
        [System.Collections.Generic.List[PSCustomObject]]$RequestArray = @()
        foreach ($InstalledApp in $DeviceApps | where-object { $_.isAssigned -eq $True }) {
            $RequestArray.add(@{
                    id     = $InstalledApp.id
                    method = 'GET'
                    url    = "/deviceAppManagement/mobileApps/$($InstalledApp.id)/deviceStatuses"
                })
        }

        try {
            $InstalledAppDetailsReturn = New-GraphBulkRequest -Requests $RequestArray -tenantid $TenantFilter
        } catch {
            $InstalledAppDetailsReturn = $null
        }
        $DeviceAppInstallDetails = foreach ($Result in $InstalledAppDetailsReturn) {
            [pscustomobject]@{
                ID                  = $Result.id
                DisplayName         = ($DeviceApps | where-object { $_.id -eq $Result.id }).DisplayName 
                InstalledAppDetails = $result.body.value
            }
        }

        write-verbose "$(Get-Date) - Parsing Groups"
        # Fetch Groups  
        $AllGroups = Get-BulkResultByID -Results $TenantResults -ID 'Groups'

        # Fetch the App status for each device
        [System.Collections.Generic.List[PSCustomObject]]$GroupRequestArray = @()
        foreach ($Group in $AllGroups) {
            $GroupRequestArray.add(@{
                    id     = $Group.id
                    method = 'GET'
                    url    = "/groups/$($Group.id)/members"
                })
        }

        try {
            $GroupMembersReturn = New-GraphBulkRequest -Requests $GroupRequestArray -tenantid $TenantFilter
        } catch {
            $GroupMembersReturn = $null
        }
        $Groups = foreach ($Result in $GroupMembersReturn) {
            [pscustomobject]@{
                ID          = $Result.id
                DisplayName = ($AllGroups | where-object { $_.id -eq $Result.id }).DisplayName 
                Members     = $result.body.value
            }
        }

        write-verbose "$(Get-Date) - Parsing Conditional Access Polcies"
        # Fetch and parse conditional access polcies
        $AllConditionalAccessPolcies = Get-BulkResultByID -Results $TenantResults -ID 'ConditionalAccess'

        $ConditionalAccessMembers = foreach ($CAPolicy in $AllConditionalAccessPolcies) {
            #Setup User Array
            [System.Collections.Generic.List[PSCustomObject]]$CAMembers = @()

            # Check for All Include
            if ($CAPolicy.conditions.users.includeUsers -contains 'All') {
                $Users | foreach-object { $null = $CAMembers.add($_.id) }
            } else {
                # Add any specific all users to the array
                $CAPolicy.conditions.users.includeUsers | foreach-object { $null = $CAMembers.add($_) }
            }

            # Now all members of groups
            foreach ($CAIGroup in $CAPolicy.conditions.users.includeGroups) {
                foreach ($Member in ($Groups | where-object { $_.id -eq $CAIGroup }).Members) {
                    $null = $CAMembers.add($Member.id)
                }
            }

            # Now all members of roles
            foreach ($CAIRole in $CAPolicy.conditions.users.includeRoles) {
                foreach ($Member in ($Roles | where-object { $_.id -eq $CAIRole }).Members) {
                    $null = $CAMembers.add($Member.id)
                }
            }

            # Parse to Unique members
            $CAMembers = $CAMembers | select-object -unique

            if ($CAMembers) {
                # Now remove excluded users
                $CAPolicy.conditions.users.excludeUsers | foreach-object { $null = $CAMembers.remove($_) }

                # Excluded Groups
                foreach ($CAEGroup in $CAPolicy.conditions.users.excludeGroups) {
                    foreach ($Member in ($Groups | where-object { $_.id -eq $CAEGroup }).Members) {
                        $null = $CAMembers.remove($Member.id)
                    }
                }

                # Excluded Roles
                foreach ($CAIRole in $CAPolicy.conditions.users.excludeRoles) {
                    foreach ($Member in ($Roles | where-object { $_.id -eq $CAERole }).Members) {
                        $null = $CAMembers.remove($Member.id)
                    }
                }
            }

            [pscustomobject]@{
                ID          = $CAPolicy.id
                DisplayName = $CAPolicy.DisplayName
                Members     = $CAMembers
            }
        }
            
        write-verbose "$(Get-Date) - Fetching One Drive Details"
        try {
            $OneDriveDetails = New-GraphGetRequest -uri "https://graph.microsoft.com/beta/reports/getOneDriveUsageAccountDetail(period='D7')" -tenantid $TenantFilter | convertfrom-csv 
        } catch {
            $OneDriveDetails = $null
        }

        write-verbose "$(Get-Date) - Fetching CAS Mailbox Details"
        try {
            $CASFull = New-GraphGetRequest -uri "https://outlook.office365.com/adminapi/beta/$($tenantfilter)/CasMailbox" -Tenantid $tenantfilter -scope ExchangeOnline -noPagination $true
        } catch {
            $CASFull = $null
        }
            
        write-verbose "$(Get-Date) - Fetching Mailbox Details"
        try {
            $MailboxDetailedFull = New-ExoRequest -TenantID $TenantFilter -cmdlet 'Get-Mailbox'
        } catch {
            $MailboxDetailedFull = $null
        }

        write-verbose "$(Get-Date) - Fetching Mailbox Stats"
        try {
            $MailboxStatsFull = New-GraphGetRequest -uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')" -tenantid $TenantFilter | convertfrom-csv 
        } catch {
            $MailboxStatsFull = $null
        }

        # Get the details of each licensed user in the tenant
        if ($licensedUsers) {
            $pre = "<div class=`"nasa__block`"><header class='nasa__block-header'>
			<h1><i class='fas fa-users icon'></i>Licensed Users</h1>
			 </header>"

            $post = "</div>"

            $UserCount = 1
                
            $OutputUsers = foreach ($user in $licensedUsers) {
				
                write-verbose "$(Get-Date) - Processing $($User.displayName)"

                # User Groups
                $UserGroups = foreach ($Group in $Groups) {
                    if ($User.id -in $Group.Members.id) {
                        $FoundGroup = $AllGroups | Where-Object { $_.id -eq $Group.id }
                        [PSCustomObject]@{
                            'Display Name'   = $FoundGroup.displayName
                            'Mail Enabled'   = $FoundGroup.mailEnabled
                            'Mail'           = $FoundGroup.mail
                            'Security Group' = $FoundGroup.securityEnabled
                            'Group Types'    = $FoundGroup.groupTypes -join ','
                        }
                    }
                }

                # Fetch Applied Conditional Access
                $UserPolicies = foreach ($cap in $ConditionalAccessMembers) {
                    if ($User.id -in $Cap.Members) {
                        $temp = [PSCustomObject]@{
                            displayName = $cap.displayName
                        }
                        $temp
                    }
                }


                $PermsRequest = ''
                $StatsRequest = ''
                $MailboxDetailedRequest = ''
                $CASRequest = ''

                write-verbose "$(Get-Date) - Fetching Mail Details"
                # Fetch Mail details
                try {
                    $CASRequest = $CASFull | where-object { $_.ExternalDirectoryObjectId -eq $User.iD }
                    $MailboxDetailedRequest = $MailboxDetailedFull | where-object { $_.ExternalDirectoryObjectId -eq $User.iD }
                    $StatsRequest = $MailboxStatsFull | where-object { $_.'User Principal Name' -eq $User.UserPrincipalName }
                    Write-Host "$UserCount Getting $($User.UserPrincipalName)"
                    $UserCount++
                    $PermsRequest = New-GraphGetRequest -uri "https://outlook.office365.com/adminapi/beta/$($tenantfilter)/Mailbox('$($User.ID)')/MailboxPermission" -Tenantid $tenantfilter -scope ExchangeOnline -noPagination $true
                } catch {
                    Write-Error "Failed Fetching Data $_"
                }

                $ParsedPerms = foreach ($Perm in $PermsRequest) {
                    if ($Perm.User -ne 'NT AUTHORITY\SELF') {
                        [pscustomobject]@{
                            User         = $Perm.User
                            AccessRights = $Perm.PermissionList.AccessRights -join ', '
                        }
                    }
                }

                $UserMailSettings = [pscustomobject]@{
                    ForwardAndDeliver        = $MailboxDetailedRequest.DeliverToMailboxAndForward
                    ForwardingAddress        = $MailboxDetailedRequest.ForwardingAddress + ' ' + $MailboxDetailedRequest.ForwardingSmtpAddress
                    LitiationHold            = $MailboxDetailedRequest.LitigationHoldEnabled
                    HiddenFromAddressLists   = $MailboxDetailedRequest.HiddenFromAddressListsEnabled
                    EWSEnabled               = $CASRequest.EwsEnabled
                    MailboxMAPIEnabled       = $CASRequest.MAPIEnabled
                    MailboxOWAEnabled        = $CASRequest.OWAEnabled
                    MailboxImapEnabled       = $CASRequest.ImapEnabled
                    MailboxPopEnabled        = $CASRequest.PopEnabled
                    MailboxActiveSyncEnabled = $CASRequest.ActiveSyncEnabled
                    Permissions              = $ParsedPerms
                    ProhibitSendQuota        = [math]::Round([float]($MailboxDetailedRequest.ProhibitSendQuota -split ' GB')[0], 2)
                    ProhibitSendReceiveQuota = [math]::Round([float]($MailboxDetailedRequest.ProhibitSendReceiveQuota -split ' GB')[0], 2)
                    ItemCount                = [math]::Round($StatsRequest.'Item Count', 2)
                    TotalItemSize            = [math]::Round($StatsRequest.'Storage Used (Byte)' / 1Gb, 2)
                }

                write-verbose "$(Get-Date) - Parsing Data"

                $userDevices = ($devices | Where-Object { $_.userPrincipalName -eq $user.UserPrincipalName } | Select-Object @{N = 'Name'; E = { "<a target='_blank' href=https://endpoint.microsoft.com/$($customer.DefaultDomainName)/#blade/Microsoft_Intune_Devices/DeviceSettingsBlade/overview/mdmDeviceId/$($_.id)>$($_.deviceName) ($($_.operatingSystem))" } }).name -join "<br/>"
				
                $UserDevicesDetailsRaw = $devices | Where-Object { $_.userPrincipalName -eq $user.UserPrincipalName } | Select-Object @{N = 'Name'; E = { "<a target='_blank' href=https://endpoint.microsoft.com/$($customer.DefaultDomainName)/#blade/Microsoft_Intune_Devices/DeviceSettingsBlade/overview/mdmDeviceId/$($_.id)>$($_.deviceName)</a>" } }, @{n = 'Owner'; e = { $_.managedDeviceOwnerType } }, `
                @{n = 'Enrolled'; e = { $_.enrolledDateTime } }, `
                @{n = 'Last Sync'; e = { $_.lastSyncDateTime } }, `
                @{n = 'OS'; e = { $_.operatingSystem } }, `
                @{n = 'OS Version'; e = { $_.osVersion } }, `
                @{n = 'State'; e = { $_.complianceState } }, `
                @{n = 'Model'; e = { $_.model } }, `
                @{n = 'Manufacturer'; e = { $_.manufacturer } },
                deviceName,
                @{n = 'url'; e = { "https://endpoint.microsoft.com/$($customer.DefaultDomainName)/#blade/Microsoft_Intune_Devices/DeviceSettingsBlade/overview/mdmDeviceId/$($_.id)" } }
                    				
                # Format Aliases
                $aliases = (($user.ProxyAddresses | Where-Object { $_ -cnotmatch "SMTP" -and $_ -notmatch ".onmicrosoft.com" }) -replace "SMTP:", " ") -join ", "
                        
                # Formay User Licenses
                $userLicenses = ($user.AssignedLicenses.SkuID | ForEach-Object {
                        $UserLic = $_
                        $SkuPartNumber = ($Licenses | Where-Object { $_.SkuId -eq $UserLic }).SkuPartNumber
                        try {
                            "$($LicenseLookup.$SkuPartNumber)"
                        } catch {
                            "$SkuPartNumber"
                        }
                    }) -join ', '


                $UserOneDriveDetails = $OneDriveDetails | where-object { $_.'Owner Principal Name' -eq $user.UserPrincipalName } 
                # Parse One Drive Settings
                [System.Collections.Generic.List[PSCustomObject]]$OneDriveFormatted = @()
                if ($UserOneDriveDetails) {
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'Owner Principal Name'  -Value "$($UserOneDriveDetails.'Owner Principal Name')"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'One Drive URL'  -Value "<a href=$($UserOneDriveDetails.'Site URL')>$($UserOneDriveDetails.'Site URL')</a>"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'Is Deleted'  -Value "$($UserOneDriveDetails.'Is Deleted')"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'Last Activity Date'  -Value "$($UserOneDriveDetails.'Last Activity Date')"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'File Count'  -Value "$($UserOneDriveDetails.'File Count')"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'Active File Count'  -Value "$($UserOneDriveDetails.'Active File Count')"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'Storage Used (Byte)'  -Value "$($UserOneDriveDetails.'Storage Used (Byte)')"))
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'Storage Allocated (Byte)'  -Value "$($UserOneDriveDetails.'Storage Allocated (Byte)')"))
                    $OneDriveUsePercent = [math]::Round([float](($UserOneDriveDetails.'Storage Used (Byte)' / $UserOneDriveDetails.'Storage Allocated (Byte)') * 100), 2)
                    $OneDriveUserUsage = @"
                        <div class="o365-usage">
                        <div class="o365-mailbox">
                            <div class="o365-used" style="width: $OneDriveUsePercent%;"></div>
                        </div>
                        <div><b>$([math]::Round($UserOneDriveDetails.'Storage Used (Byte)' /1024 /1024 /1024, 2)) GB</b> used, <b>$OneDriveUsePercent%</b> of <b>$([math]::Round($UserOneDriveDetails.'Storage Allocated (Byte)' /1024 /1024 /1024, 2)) GB</b></div>
                    </div>
"@
                        
                    $OneDriveFormatted.add($(Get-FormatedField -Title 'One Drive Usage'  -Value $OneDriveUserUsage))
                }
                        
                        
                # Parse user mailbox details
                [System.Collections.Generic.List[PSCustomObject]]$UserMailSettingsFormatted = @()
                [System.Collections.Generic.List[PSCustomObject]]$UserMailboxDetailsFormatted = @()
                if ($UserMailSettings) {
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'Forward and Deliver'  -Value "$($UserMailSettings.ForwardAndDeliver)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'Forwarding Address'  -Value "$($UserMailSettings.ForwardingAddress)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'Litiation Hold'  -Value "$($UserMailSettings.LitiationHold)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'Hidden From Address Lists'  -Value "$($UserMailSettings.HiddenFromAddressLists)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'EWS Enabled'  -Value "$($UserMailSettings.EWSEnabled)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'MAPI Enabled'  -Value "$($UserMailSettings.MailboxMAPIEnabled)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'OWA Enabled'  -Value "$($UserMailSettings.MailboxOWAEnabled)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'IMAP Enabled'  -Value "$($UserMailSettings.MailboxImapEnabled)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'POP Enabled'  -Value "$($UserMailSettings.MailboxPopEnabled)"))
                    $UserMailSettingsFormatted.add($(Get-FormatedField -Title 'Active Sync Enabled'  -Value "$($UserMailSettings.MailboxActiveSyncEnabled)"))


                    $UserMailboxDetailsFormatted.add($(Get-FormatedField -Title 'Permissions'  -Value "$($UserMailSettings.Permissions | ConvertTo-HTML -Fragment | Out-String)"))
                    $UserMailboxDetailsFormatted.add($(Get-FormatedField -Title 'Prohibit Send Quota'  -Value "$($UserMailSettings.ProhibitSendQuota)"))
                    $UserMailboxDetailsFormatted.add($(Get-FormatedField -Title 'Prohibit Send Receive Quota'  -Value "$($UserMailSettings.ProhibitSendReceiveQuota)"))
                    $UserMailboxDetailsFormatted.add($(Get-FormatedField -Title 'Item Count'  -Value "$($UserMailSettings.ItemCount)"))
                    $UserMailboxDetailsFormatted.add($(Get-FormatedField -Title 'Total Mailbox Size'  -Value "$($UserMailSettings.TotalItemSize)"))
                    try {
                        $UserMailboxUsePercent = [math]::Round([float](($UserMailSettings.TotalItemSize / $UserMailSettings.ProhibitSendReceiveQuota) * 100), 2)
                    } catch {
                        $UserMailboxUsePercent = 100
                    }
                    $UserMailboxUsage = @"
                            <div class="o365-usage">
                        <div class="o365-mailbox">
                            <div class="o365-used" style="width: $UserMailboxUsePercent%;"></div>
                        </div>
                        <div><b>$([math]::Round($UserMailSettings.TotalItemSize,2)) GB</b> used, <b>$UserMailboxUsePercent%</b> of <b>$([math]::Round($UserMailSettings.ProhibitSendReceiveQuota, 2)) GB</b></div>
                    </div>
"@
                    $UserMailboxDetailsFormatted.add($(Get-FormatedField -Title 'Mailbox Usage'  -Value $UserMailboxUsage))

                }

                # Create conditional access policy list
                $UserPoliciesFormatted = '<ul>'
                foreach ($Policy in $UserPolicies) {
                    $UserPoliciesFormatted = $UserPoliciesFormatted + "<li>$($Policy.displayName)</li>"
                }
                $UserPoliciesFormatted = $UserPoliciesFormatted + '</ul>'

                # Create user overview details
                [System.Collections.Generic.List[PSCustomObject]]$UserOverviewFormatted = @()
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'User Name'  -Value "$($User.displayName)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'User Principal Name'  -Value "$($User.userPrincipalName)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'User ID'  -Value "$($User.ID)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'User Enabled'  -Value "$($User.accountEnabled)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'Job Title'  -Value "$($User.jobTitle)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'Mobile Phone'  -Value "$($User.mobilePhone)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'Business Phones'  -Value "$($User.businessPhones -join ', ')"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'Office Location'  -Value "$($User.officeLocation)"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'Aliases'  -Value "$aliases"))
                $UserOverviewFormatted.add($(Get-FormatedField -Title 'Licenses'  -Value "$($userLicenses)"))


                # Assigned plan information
                $AssignedPlans = $User.assignedplans | Where-Object { $_.capabilityStatus -eq 'Enabled' } | Select-Object @{n = 'Assigned'; e = { $_.assignedDateTime } }, @{n = 'Service'; e = { $_.service } } -unique
                [System.Collections.Generic.List[PSCustomObject]]$AssignedPlansFormatted = @()
                foreach ($AssignedPlan in $AssignedPlans) {
                    if ($AssignedPlan.service -in ($AssignedMap | Get-Member -MemberType NoteProperty).name) {
                        $CSSClass = $AssignedMap."$($AssignedPlan.service)"
                        $PlanDisplayName = $AssignedNameMap."$($AssignedPlan.service)"
                        $ParsedDate = get-date($AssignedPlan.Assigned) -format 'yyyy-MM-dd HH:mm:ss'
                        $AssignedPlansFormatted.add("<div class='o365__app $CSSClass' style='text-align:center'><strong>$PlanDisplayName</strong><font style='font-size: .72rem;'>Assigned $($ParsedDate)</font></div>")
                    }
                }
                $AssignedPlansBlock = "<div class='o365'>$($AssignedPlansFormatted -join '')</div>"


                # Format into blocks
                if ($UserMailSettingsFormatted) {
                    $UserMailSettingsBlock = Get-FormattedBlock -Heading "Mailbox Settings" -Body ($UserMailSettingsFormatted -join '')
                } else {
                    $UserMailSettingsBlock = $null
                }
                
                if ($UserMailboxDetailsFormatted) {
                    $UserMailDetailsBlock = Get-FormattedBlock -Heading "Mailbox Details" -Body ($UserMailboxDetailsFormatted -join '')
                } else {
                    $UserMailDetailsBlock = $null
                }
                
                if ($UserGroups) {
                    $UserGroupsBlock = Get-FormattedBlock -Heading "User Groups" -Body $($UserGroups | ConvertTo-Html -Fragment -As Table | out-string)                    
                } else {
                    $UserGroupsBlock = $null
                }

                if ($UserPoliciesFormatted) {
                    $UserPoliciesBlock = Get-FormattedBlock -Heading "Assigned Conditional Access Polcies" -Body $UserPoliciesFormatted
                } else {
                    $UserPoliciesBlock = $null
                }

                if ($OneDriveFormatted) {
                    $OneDriveBlock = Get-FormattedBlock -Heading "One Drive Details" -Body ($OneDriveFormatted -join '')
                } else {
                    $OneDriveBlock = $null
                }

                if ($UserOverviewFormatted) {
                    $UserOverviewBlock = Get-FormattedBlock -Heading "User Details" -Body $UserOverviewFormatted
                } else {
                    $UserOverviewBlock = $null
                }

                if ($UserDevicesDetailsRaw) {
                    $UserDevicesDetailsBlock = Get-FormattedBlock -Heading "Intune Devices" -Body $($UserDevicesDetailsRaw | Select-Object -ExcludeProperty deviceName, url | convertto-html -fragment | ForEach-Object { $tmp = $_ -replace "&lt;", "<"; $tmp -replace "&gt;", ">"; } | out-string)
                } else {
                    $UserDevicesDetailsBlock = $null
                }


                $HuduUser = $People | where-object { $_.primary_mail -eq $user.UserPrincipalName }

                # Build User Link Buttons.
                [System.Collections.Generic.List[PSCustomObject]]$UserLinksFormatted = @()
                $UserLinksFormatted.add((Get-LinkBlock -URL "https://aad.portal.azure.com/$($Customer.defaultDomainName)/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/$($User.id)" -Icon "fas fa-users-cog" -Title "Azure AD"))
                $UserLinksFormatted.add((Get-LinkBlock -URL "https://aad.portal.azure.com/$($Customer.defaultDomainName)/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/SignIns/userId/$($User.id)" -Icon "fas fa-history" -Title "Sign Ins"))
                $UserLinksFormatted.add((Get-LinkBlock -URL "https://admin.teams.microsoft.com/users/$($User.id)/account?delegatedOrg=$($Customer.defaultDomainName)" -Icon "fas fa-users" -Title "Teams Admin"))
                $UserLinksFormatted.add((Get-LinkBlock -URL "https://endpoint.microsoft.com/$($Customer.defaultDomainName)/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/$($User.ID)" -Icon "fas fa-laptop" -Title "EPM (User)"))
                $UserLinksFormatted.add((Get-LinkBlock -URL "https://endpoint.microsoft.com/$($Customer.defaultDomainName)/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Devices/userId/$($User.ID)" -Icon "fas fa-laptop" -Title "EPM (Devices)"))

                # Check for Halo PSA
                if ($HuduUser) {
                    $HaloCard = $HuduUser.cards | where-object { $_.integrator_name -eq 'halo' }
                    if ($HaloCard) {
                        $UserLinksFormatted.add((Get-LinkBlock -URL "$($PSAUserUrl)$($HaloCard.sync_id)" -Icon "far fa-id-card" -Title "Halo PSA"))
                    }
                        
                }
                                        
                $UserLinksBlock = "<div>Management Links</div><div class='o365'>$($UserLinksFormatted -join '')</div>"


                # Assemble the full body of the user details
                $UserBody = "<div>$AssignedPlansBlock<br />$UserLinksBlock<br /><div class=`"nasa__content`">$($UserOverviewBlock)$($UserMailDetailsBlock)$($OneDriveBlock)$($UserMailSettingsBlock)$($UserPoliciesBlock)</div><div class=`"nasa__content`">$($UserDevicesDetailsBlock)</div><div class=`"nasa__content`">$($UserGroupsBlock)</div></div>"

                $UserAssetFields = @{
                    microsoft_365 = $UserBody
                }

                # Update / Create the User Asset
                $HuduUser = $People | where-object { $_.primary_mail -eq $user.UserPrincipalName }
                if ($HuduUser) {
                    $null = Set-HuduAsset -asset_id $HuduUser.id -name $HuduUser.name -company_id $company_id -asset_layout_id $PeopleLayout.id -fields $UserAssetFields
                    write-verbose "$(Get-Date) - User Updated"  

                } else {
                    if ($CreateUsers -eq $True) {
                        $HuduUser = (New-HuduAsset -name $User.DisplayName -company_id $company_id -asset_layout_id $PeopleLayout.id -fields $UserAssetFields -primary_mail $user.UserPrincipalName).asset
                        write-verbose "$(Get-Date) - User Created"
                    }
                }

                # Create relations to devices.
                foreach ($Device in $UserDevicesDetailsRaw) {
                    $HuduDevice = $HuduDevices | Where-Object { $_.name -eq $device.deviceName }
                    if ($HuduDevice) {                            
                        try {
                            $null = New-HuduRelation -FromableType "Asset" -FromableId $HuduUser.id -ToableType "Asset" -ToableId $HuduDevice.id -ea stop
                        } catch {
                            Write-Verbose "Relationship already exists, or creation failed"
                        }
                    }
                }

                $UserLink = "<a target=_blank href=$($HuduUser.url)>$($user.DisplayName)</a>"
               
                [PSCustomObject]@{
                    "Display Name"      = $UserLink
                    "Addresses"         = "<strong>$($user.UserPrincipalName)</strong><br/>$aliases"
                    "EPM Devices"       = $userDevices
                    "Assigned Licenses" = $userLicenses
                    "Options"           = "<a target=`"_blank`" href=https://aad.portal.azure.com/$($Customer.DefaultDomainName)/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/$($user.ObjectId)>Azure AD</a> | <a <a target=`"_blank`" href=https://portal.office.com/Partner/BeginClientSession.aspx?CTID=$($customer.CustomerContextId)&CSDEST=o365admincenter/Adminportal/Home#/users/:/UserDetails/$($user.ObjectId)>M365 Admin</a>"
                }
            }

            $licensedUserHTML = $OutputUsers | ConvertTo-Html -PreContent $pre -PostContent $post -Fragment | ForEach-Object { $tmp = $_ -replace "&lt;", "<"; $tmp -replace "&gt;", ">"; } | Out-String

        }

        # Loop all Devices
        foreach ($Device in $Devices) {
            # Create user overview details
            [System.Collections.Generic.List[PSCustomObject]]$DeviceOverviewFormatted = @()
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'Device Name'  -Value "$($Device.deviceName)"))
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'User'  -Value "$($Device.userDisplayName)"))
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'User Email'  -Value "$($Device.userPrincipalName)"))
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'Owner'  -Value "$($Device.ownerType)"))
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'Enrolled'  -Value "$($Device.enrolledDateTime)"))
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'Last Checkin'  -Value "$($Device.lastSyncDateTime)"))
            if ($Device.complianceState -eq 'compliant') {
                $CompliantSymbol = '<font color=green><em class="fas fa-check-circle">&nbsp;&nbsp;&nbsp;</em></font>'
            } else {
                $CompliantSymbol = '<font color=red><em class="fas fa-times-circle">&nbsp;&nbsp;&nbsp;</em></font>'
            }
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'Compliant'  -Value "$($CompliantSymbol)$($Device.complianceState)"))
            $DeviceOverviewFormatted.add($(Get-FormatedField -Title 'Management Type'  -Value "$($Device.managementAgent)"))

            # Create hardware details
            [System.Collections.Generic.List[PSCustomObject]]$DeviceHardwareFormatted = @()
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'Serial Number'  -Value "$($Device.serialNumber)"))
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'OS'  -Value "$($Device.operatingSystem)"))
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'OS Versions'  -Value "$($Device.osVersion)"))                
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'Chassis'  -Value "$($Device.chassisType)"))  
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'Model'  -Value "$($Device.model)"))
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'Manufacturer'  -Value "$($Device.manufacturer)"))
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'Total Storage'  -Value "$([math]::Round($Device.totalStorageSpaceInBytes /1024 /1024 /1024, 2))"))
            $DeviceHardwareFormatted.add($(Get-FormatedField -Title 'Free Storage'  -Value "$([math]::Round($Device.freeStorageSpaceInBytes /1024 /1024 /1024, 2))"))

            # Device guard details
            [System.Collections.Generic.List[PSCustomObject]]$DeviceEnrollmentFormatted = @()
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Enrollment Type'  -Value "$($Device.deviceEnrollmentType)"))
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Join Type'  -Value "$($Device.joinType)"))
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Registration State'  -Value "$($Device.deviceRegistrationState)"))
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Autopilot Enrolled'  -Value "$($Device.autopilotEnrolled)"))  
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Device Guard Requirements'  -Value "$($Device.hardwareinformation.deviceGuardVirtualizationBasedSecurityHardwareRequirementState)"))
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Virtualistation Based Security'  -Value "$($Device.hardwareinformation.deviceGuardVirtualizationBasedSecurityState)"))
            $DeviceEnrollmentFormatted.add($(Get-FormatedField -Title 'Credential Guard'  -Value "$($Device.hardwareinformation.deviceGuardLocalSystemAuthorityCredentialGuardState)"))
                

            # Compliance polcies
            $DevicePolciesTable = foreach ($Policy in $DeviceComplianceDetails) {
                if ($device.deviceName -in $Policy.DeviceStatuses.deviceDisplayName) {
                    $Status = $Policy.DeviceStatuses | where-object { $_.deviceDisplayName -eq $device.deviceName }
                    if ($Status.status -ne 'unknown') {
                        [PSCustomObject]@{
                            Name           = $Policy.DisplayName
                            Status         = ($Status.status | select-object -unique) -join ', '
                            'Last Report'  = "$(get-date($Status.lastReportedDateTime[0]) -format 'yyyy-MM-dd HH:mm:ss')"
                            'Grace Expiry' = "$(get-date($Status.complianceGracePeriodExpirationDateTime[0]) -format 'yyyy-MM-dd HH:mm:ss')"
                        }
                    }
                }
            }
            $DevicePolciesFormatted = $DevicePolciesTable | ConvertTo-Html -fragment | Out-String

            # Device Groups
            $DeviceGroupsTable = foreach ($Group in $Groups) {
                if ($device.azureADDeviceId -in $Group.members.deviceId) {
                    [PSCustomObject]@{
                        Name = $Group.displayName
                    }
                }
            }
            $DeviceGroupsFormatted = $DeviceGroupsTable | ConvertTo-Html -fragment | Out-String

            # Device Apps
            $DeviceAppsTable = foreach ($App in $DeviceAppInstallDetails) {
                if ($device.id -in $App.InstalledAppDetails.deviceId) {
                    $Status = $App.InstalledAppDetails | where-object { $_.deviceId -eq $device.id }
                    [PSCustomObject]@{
                        Name             = $App.DisplayName
                        'Install Status' = ($Status.installState | select-object -unique ) -join ','
                    }
                }
            }
            $DeviceAppsFormatted = $DeviceAppsTable | ConvertTo-Html -fragment | Out-String
                
                
            # Build Blocks
            $DeviceOverviewBlock = Get-FormattedBlock -Heading "Device Details" -Body ($DeviceOverviewFormatted -join '')
            $DeviceHardwareBlock = Get-FormattedBlock -Heading "Hardware Details" -Body ($DeviceHardwareFormatted -join '')
            $DeviceEnrollmentBlock = Get-FormattedBlock -Heading "Device Enrollment Details" -Body ($DeviceEnrollmentFormatted -join '')
            $DevicePolicyBlock = Get-FormattedBlock -Heading "Compliance Polcies" -Body ($DevicePolciesFormatted -join '')
            $DeviceAppsBlock = Get-FormattedBlock -Heading "App Details" -Body ($DeviceAppsFormatted -join '')
            $DeviceGroupsBlock = Get-FormattedBlock -Heading "Device Groups" -Body ($DeviceGroupsFormatted -join '')

            # Match on name or serial
            if ("$($device.serialNumber)" -in $ExcludeSerials) {
                $HuduDevice = $HuduDevices | Where-Object { $_.name -eq $device.deviceName }
            } else {
                $HuduDevice = $HuduDevices | Where-Object { $_.primary_serial -eq $device.serialNumber }
            }
                
            # Build User Link Buttons.
            [System.Collections.Generic.List[PSCustomObject]]$DeviceLinksFormatted = @()
            $DeviceLinksFormatted.add((Get-LinkBlock -URL "https://endpoint.microsoft.com/$($Customer.defaultDomainName)/#blade/Microsoft_Intune_Devices/DeviceSettingsBlade/overview/mdmDeviceId/$($Device.id)" -Icon "fas fa-laptop" -Title "Endpoint Manager"))

            # Check for Datto RMM
            if ($HuduDevice) {
                $DRMMCard = $HuduDevice.cards | where-object { $_.integrator_name -eq 'dattormm' }
                if ($DRMMCard) {
                    $DeviceLinksFormatted.add((Get-LinkBlock -URL "$($RMMDeviceURL)$($DRMMCard.data.id)" -Icon "fas fa-laptop-code" -Title "Datto RMM"))
                    $DeviceLinksFormatted.add((Get-LinkBlock -URL "$($RMMRemoteURL)$($DRMMCard.data.id)" -Icon "fas fa-desktop" -Title "Datto RMM Remote"))
                }
                    
            }
                                    
            $DeviceLinksBlock = "<div>Management Links</div><div class='o365'>$($DeviceLinksFormatted -join '')</div>"

            $DeviceIntuneDetailshtml = "<div>$DeviceLinksBlock<br /><div class=`"nasa__content`">$($DeviceOverviewBlock)$($DeviceHardwareBlock)$($DeviceEnrollmentBlock)$($DevicePolicyBlock)$($DeviceAppsBlock)$($DeviceGroupsBlock)</div></div>"

            $DeviceAssetFields = @{
                microsoft_365 = $DeviceIntuneDetailshtml
            }

            if ($HuduDevice) {
                if (($HuduDevice | measure-object).count -eq 1) {                                  
                    $null = Set-HuduAsset -asset_id $HuduDevice.id -name $HuduDevice.name -company_id $company_id -asset_layout_id $HuduDevice.asset_layout_id -fields $DeviceAssetFields -PrimarySerial $Device.serialNumber
                    Write-Verbose "Device Updated"
                } else {
                    Write-Verbose "Multiple Devices Matched. $($HuduDevice.name)"
                }                    
            } else {
                if ($device.deviceType -in $IntuneDesktopDeviceTypes) {
                    $DeviceLayoutID = $DesktopsLayout.id
                    $DeviceCreation = $CreateDevices
                } else {
                    $DeviceLayoutID = $MobilesLayout.id
                    $DeviceCreation = $CreateMobileDevices
                }
                if ($DeviceCreation -eq $true) {
                    # Create New Device
                    $HuduDevice = (New-HuduAsset -name $device.deviceName -company_id $company_id -asset_layout_id $DeviceLayoutID -fields $DeviceAssetFields -PrimarySerial $Device.serialNumber).asset
                    write-verbose "$(Get-Date) - Device Created"

                    $HuduUser = $People | where-object { $_.primary_mail -eq $Device.userPrincipalName }
                    if ($HuduUser) {
                        try {
                            $null = New-HuduRelation -FromableType "Asset" -FromableId $HuduUser.id -ToableType "Asset" -ToableId $HuduDevice.id -ea stop
                        } catch {
                            Write-Verbose "Relationship already exists, or creation failed"
                        }
                    }
                }
            }

        }
		 

        #Build the output
        $body = "<div class='nasa__block'>
			<header class='nasa__block-header'>
			<h1><i class='fas fa-cogs icon'></i>Administrative Portals</h1>
	 		</header>
			<div>$CustomerLinks</div> 
			<br/>
			</div>
			<br/>
			<div class=`"nasa__content`">
			 $detailstable
			 $licenseHTML
			 </div>
             <br/>
			 <div class=`"nasa__content`">
			 $RolesHtml
			 </div>
			 <br/>
			 <div class=`"nasa__content`">
			 $licensedUserHTML
			 </div>"
      
   	
        $null = Set-HuduMagicDash -title "Microsoft 365 - $($customer.DisplayName)" -company_name $company_name -message "$($licensedUsers.count) Licensed Users" -icon "fab fa-microsoft" -content $body -shade "success"	
            
        if ($CreateInOverview -eq $true) {
            $null = Set-HuduMagicDash -title "$($customer.DisplayName)" -company_name $OverviewCompany -message "$($licensedUsers.count) Licensed Users" -icon "fab fa-microsoft" -content $body -shade "success"	
        }
            
        write-verbose "$(Get-Date) - https://$defaultdomain Found in Hudu and MagicDash updated for $($customer.DisplayName)"
		
        #Import Domains if enabled
        if ($importDomains) {
            write-verbose "$(Get-Date) - Processing Domains"
            $domainstoimport = $RawDomains
            foreach ($imp in $domainstoimport) {
                $impdomain = $imp.id
                $huduimpdomain = Get-HuduWebsites -name "https://$impdomain"
                if ($($huduimpdomain.id.count) -gt 0) {
                    write-verbose "$(Get-Date) - https://$impdomain Found in Hudu"
                } else {
                    if ($monitorDomains) {
                        $null = New-HuduWebsite -name "https://$impdomain" -notes $HuduNotes -paused "false" -companyid $company_id -disabledns "false" -disablessl "false" -disablewhois "false"
                        write-verbose "$(Get-Date) - https://$impdomain Created in Hudu with Monitoring"
                    } else {
                        $null = New-HuduWebsite -name "https://$impdomain" -notes $HuduNotes -paused "true" -companyid $company_id -disabledns "true" -disablessl "true" -disablewhois "true"
                        write-verbose "$(Get-Date) - https://$impdomain Created in Hudu with Monitoring"
                    }

                }		
            }
      
        }

        Write-Host "Time Taken: $((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalMinutes)"
    } else {
        write-verbose "$(Get-Date) - https://$defaultdomain Not found in Hudu please add it to the correct client"
    }
    
}

