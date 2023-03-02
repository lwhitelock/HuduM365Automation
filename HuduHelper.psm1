
### Functions ###
### The graph helper functions are based on the ones from Kelvin Tegelaar's CIPP project https://github.com/KelvinTegelaar/CIPP
function New-GraphGetRequest {
    Param(
        $uri,
        $tenantid,
        $scope,
        $AsApp,
        $noPagination,
        $Headers
    )

    Write-Verbose "Using $($uri) as url"
    $nextURL = $uri
    $ReturnedData = do {
        try {
            $Data = (Invoke-RestMethod -Uri $nextURL -Method GET -Headers $headers -ContentType 'application/json; charset=utf-8')
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
        client_id     = $env:ApplicationId
        client_secret = $env:ApplicationSecret
        scope         = $Scope
        refresh_token = $env:RefreshToken
        grant_type    = 'refresh_token'

    }
    if ($asApp -eq $true) {
        $AuthBody = @{
            client_id     = $env:ApplicationId
            client_secret = $env:ApplicationSecret
            scope         = $Scope
            grant_type    = 'client_credentials'
        }
    }

    if ($null -ne $AppID -and $null -ne $refreshToken) {
        $AuthBody = @{
            client_id     = $appid
            refresh_token = $RefreshToken
            scope         = $Scope
            grant_type    = 'refresh_token'
        }
    }

    if (!$tenantid) { $tenantid = $env:tenantid }
    $AccessToken = (Invoke-RestMethod -Method post -Uri "https://login.microsoftonline.com/$($tenantid)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop)
    if ($ReturnRefresh) { $header = $AccessToken } else { $header = @{ Authorization = "Bearer $($AccessToken.access_token)" } }

    return $header
}

function New-ExoRequest ($tenantid, $cmdlet, $cmdParams, $useSystemMailbox, $Anchor) {
    $token = Get-ClassicAPIToken -resource 'https://outlook.office365.com' -Tenantid $tenantid

    if ($cmdParams) {
        $Params = $cmdParams
    } else {
        $Params = @{}
    }
    $ExoBody = ConvertTo-Json -Depth 5 -InputObject @{
        CmdletInput = @{
            CmdletName = $cmdlet
            Parameters = $Params
        }
    }
    if (!$Anchor) {
        if ($cmdparams.Identity) { $Anchor = $cmdparams.Identity }
        if ($cmdparams.anr) { $Anchor = $cmdparams.anr }
        if ($cmdparams.User) { $Anchor = $cmdparams.User }

        if (!$Anchor -or $useSystemMailbox) {
            $OnMicrosoft = (New-GraphGetRequest -uri 'https://graph.microsoft.com/beta/domains?$top=999' -tenantid $tenantid | Where-Object -Property isInitial -EQ $true).id
            $anchor = "UPN:SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@$($OnMicrosoft)"

        }
    }
    Write-Host "Using $Anchor"
    $Headers = @{
        Authorization     = "Bearer $($token.access_token)"
        Prefer            = 'odata.maxpagesize = 1000'
        'X-AnchorMailbox' = $anchor

    }
    try {
        $ReturnedData = Invoke-RestMethod "https://outlook.office365.com/adminapi/beta/$($tenantid)/InvokeCommand" -Method POST -Body $ExoBody -Headers $Headers -ContentType 'application/json; charset=utf-8'
    } catch {
        $ReportedError = ($_.ErrorDetails | ConvertFrom-Json -ErrorAction SilentlyContinue)
        $Message = if ($ReportedError.error.details.message) { $ReportedError.error.details.message } else { $ReportedError.error.innererror.internalException.message }
        if ($Message -eq $null) { $Message = $($_.Exception.Message) }
        throw $Message
    }
    return $ReturnedData.value
}

function Get-ClassicAPIToken($tenantID, $Resource) {
    Write-Host 'Using classic'
    $uri = "https://login.microsoftonline.com/$($TenantID)/oauth2/token"
    $Body = @{
        client_id     = $env:ApplicationID
        client_secret = $env:ApplicationSecret
        resource      = $Resource
        refresh_token = $env:RefreshToken
        grant_type    = 'refresh_token'

    }

    $token = Invoke-RestMethod $uri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction SilentlyContinue -Method post
    return $token
}

function New-GraphBulkRequest ($Requests, $tenantid, $Headers) {
    $URL = 'https://graph.microsoft.com/beta/$batch'
    $ReturnedData = for ($i = 0; $i -lt $Requests.count; $i += 20) {
        $req = @{}
        # Use select to create hashtables of id, method and url for each call
        $req['requests'] = ($Requests[$i..($i + 19)])
        Invoke-RestMethod -Uri $URL -Method POST -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ($req | ConvertTo-Json -Depth 10)
    }

    $Headers['ConsistencyLevel'] = 'eventual'
    foreach ($MoreData in $ReturnedData.Responses | Where-Object { $_.body.'@odata.nextLink' }) {
        $AdditionalValues = New-GraphGetRequest -Headers $Headers -uri $MoreData.body.'@odata.nextLink' -tenantid $TenantFilter
        $NewValues = [System.Collections.Generic.List[PSCustomObject]]$MoreData.body.value
        $AdditionalValues | ForEach-Object { $NewValues.add($_) }
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

function Get-AdaptiveColumn($Strings, $Title) {
    [System.Collections.Generic.List[PSCustomObject]]$Items = @()
    $Items.Add([PSCustomObject]@{
            type   = 'TextBlock'
            weight = 'Bolder'
            text   = $Title
        })

    foreach ($String in $Strings) {
        $Items.Add([PSCustomObject]@{
                type      = 'TextBlock'

                separator = $true
                text      = $String
            })
    }

    return [PSCustomObject]@{
        type  = 'Column'
        items = $Items
        width = 'stretch'
    }
}

function Get-LicenseLookup {
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

    return $LicenseLookup

}

function Get-AssignedMap {
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

    return $AssignedMap

}


function Get-AssignedNameMap {

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

    return $AssignedNameMap

}
