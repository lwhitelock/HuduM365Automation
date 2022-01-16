
### Hudu M365 Automation
This is an Azure Function which will syncronise between Microsoft 365 and Hudu.

### Copyright
This project utilises some of the helper fuctions from the CIPP project https://github.com/KelvinTegelaar/CIPP and it licensed under the same license.

### Installation
If you wish to customise the code you can fork this repository and then deploy to Azure. If you would like to use the default version you can deploy it with this button

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Flwhitelock%2FHuduM365Automation%2Fmaster%2FDeployment%2FAzureDeployment.json)

### Settings
#### Core Settings
| Setting Name | Description |
|--|--|
|baseName| Name use as base-template to named the resources deployed in Azure.|
|PeopleLayoutName|The name of the Asset Layout in Hudu used to track People / Users. (Must exist already with a richtext Microsoft 365 field)|
|DesktopsName|The name of the Asset Layout in Hudu used to track Desktops / Laptops. (Must exist already with a richtext Microsoft 365 field)|
|MobilesName|The name of the Asset Layout in Hudu used to track Mobile Devices. (Must exist already with a richtext Microsoft 365 field)|
|customerExclude|A list of customer M365 display names to skip from the sync.|
|ApplicationId|The application ID of your M365 SAM application.|
|ApplicationSecret|The application secret for your M365 SAM application.|
|TenantID|Your Partner Tenant ID.|
|RefreshToken|The long refresh token for your M365 SAM application.|
|ExchangeRefreshToken|The long Exchange refresh token for your SAM application.|
|HuduAPIKey|Your Hudu API Key.|
|HuduBaseDomain|Your Hudu Base Domain.|
|WebhookURL|An incoming Teams Webhook URL to recieve Sync Errors and Reports|
|PSAUserURL|The URL used to link a user in Hudu to your PSA (Only Halo supported at present)|
|RMMDeviceURL|The URL used to link a device to your RMM (Only Datto RMM support at present)|
|RMMRemoteURL|The URL used to launch a remote session on a device in your RMM (Only Datto RMM support at present)|
|CreateInOverview|Set to true to create all Magic Dashes in an overview company as well.|
|OverviewCompany|The name of the overview company to use for Magic Dashes (Must exist already).|
|DocumentPartnerTenant|Set to true to also document your partner tenant to Hudu.|
|PartnerDefaultDomain|Your Partner Tenant's default domain name.|
|PartnerDisplayName|Your Partner Tenant's display name.|
#### Script Settings
| Setting Name | Description |
|--|--|
|CreateUsers|Set to true if you wish to create People / Users in Hudu from M365 if they do not exist.|
|CreateDevices|Set to true if you wish to create Desktops / Laptops in Hudu from M365 if they do not exist.|
|CreateMobileDevices|Set to true if you wish to create Mobile Devices in Hudu from M365 if they do not exist.|
|importDomains|Set to true to import domains in M365 to Hudu Websites.|
|monitorDomains|Set to enable monitoring on imported domains.|
|IntuneDesktopDeviceTypes|Endpoint Manager / Intune device types to identify desktops. All others will be treated as mobile devices|
|ExcludeSerials|Serial numbers to ignore and attempt to match on device name instead. Add any generic serials that might apply to multiple devices|
