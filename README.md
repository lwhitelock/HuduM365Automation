### Hudu M365 Automation
This is an Azure Function which will syncronise between Microsoft 365 and Hudu. For more details please see my blog post here: https://mspp.io/hudu-microsoft-365-integration-and-updated-magic-dash-v3/

### Copyright
This project utilizes some of the helper functions and approaches written by Kelvin Tegelaar from the CIPP project https://github.com/KelvinTegelaar/CIPP and is licensed under the same terms.


### Updates
#### Version 1.2
```
Simplified and Fixed Refresh Token Update
```
#### Version 1.1
```
Added error handling for tokens
Added automatic token refresh
```

### Requirements
For this script you will need
1. A configured M365 Partner SAM Application. To set this up I recommend you follow this guide https://www.gavsto.com/secure-application-model-for-the-layman-and-step-by-step/
2. The following permissions added to your SAM application and granted admin consent:

| Permission | Type |
|--|--|
|Application.Read.All|Application + Delegated|
|AuditLog.Read.All|Delegated|
|DeviceManagementApps.Read.All|Application + Delegated|
|DeviceManagementConfiguration.Read.All|Application + Delegated|
|DeviceManagementManagedDevices.Read.All|Application + Delegated|
|Directory.Read.All|Application|
|Group.Read.All|Application + Delegated|
|Organization.Read.All|Application + Delegated|
|Policy.Read.All|Application + Delegated|
|Reports.Read.All|Application + Delegated|
|SecurityEvents.Read.All|Application + Delegated|

3. Hudu API Keys.
4. M365 to Hudu Tenant mapping completed (Each default domain needs to be setup as a customer under the relevant tenant this script will help with that https://github.com/lwhitelock/HuduAutomation/blob/main/Hudu-M365-Links.ps1)
5. The names of the Hudu asset types for People, Devices and Mobile devices in Hudu.
6. A field called 'Microsoft 365' of type 'RichText' added to each of the above asset layouts
7. An overview company created and the name of it in Hudu if you wish to use that feature.
8. Links to your PSA / RMM if using a supported tool (Currently Halo PSA and Datto RMM).(If you would like me to add support to your PSA / RMM please let me know the URL format to open a link to a user / device and the location in the .card data from the Hudu API where the relevant ID that needs to be passed into the URL is stored)

### Installation
If you wish to customise the code you can fork this repository and then deploy to Azure. If you would like to use the default version which will run the syncronisation daily at midnight, you can deploy it with this button:

#### Custom CSS
Go to Admin -> Design -> Custom CSS and Add in:
```
.card__item table{
	border-collapse: collapse;
	margin: 5px 0;
	font-size: 0.8em;
	font-family: sans-serif;
	min-width: 400px;
	box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}
.card__item h2, p{
	font-size: 0.8em;
	font-family: sans-serif;
}
.card__item th, td {
	padding: 5px 5px;
	width:auto;
}
.card__item thead tr {
	text-align: left;
}
.card__item tr {
	border-bottom: 1px solid #dddddd;
}

.custom-fast-fact.custom-fast-fact--warning {
    background: #f5c086;
}
 .custom-fast-fact.custom-fast-fact--datto-low {
     background: #2C81C8;
}
 .custom-fast-fact.custom-fast-fact--datto-moderate {
     background: #F7C210;
}

 .custom-fast-fact.custom-fast-fact--datto-high {
     background: #F68218;
}

 .custom-fast-fact.custom-fast-fact--datto-critical {
     background: #EC422E;
}

.nasa__block {
   height:auto;
}
```

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Flwhitelock%2FHuduM365Automation%2Fmaster%2FDeployment%2FAzureDeployment.json)

### Settings
#### Core Settings
| Setting Name | Description |
|--|--|
|baseName| Name use as base-template to name the resources deployed in Azure.|
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

#### Running the function
##### Automatically
By default the function will run every day at midnight. It will send a report to your Teams Webhook URL once it completes.

##### Manually
To run the function manually you can use this process:
1. Browse to the Azure Function app in the Azure Portal. It will be named something like m365huduahg45e depending on what you set the base name to
2. Select functions on the left
3. Click on the HuduM365Trigger function
4. Click on the Get Function URL button
5. Paste the URL into a browser. You should recieve a JSON response. 
6. To check the status you can copy and paste the statusQueryGetUri into a browser.

##### Debugging
At the end of the function running a report will be sent to Teams. This should give you details on issues that need to be fixed.
If you have issues the easiest way to debug is to use VSCode with the Azure Functions extension. If you click on theAzure logo in the left and login, you can then find your function app from the list. Right click on the function and choose start streaming logs.
