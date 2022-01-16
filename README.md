
### Hudu M365 Automation
This is an Azure Function which will syncronise between Microsoft 365 and Hudu.

### Copyright
This project utilises some of the helper fuctions from the CIPP project https://github.com/KelvinTegelaar/CIPP and it licensed under the same license.

### Installation
If you wish to customise the code you can fork this repository and then deploy to Azure. If you would like to use the default version you can deploy it with this button

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Flwhitelock%2FHuduM365Automation%2Fmaster%2FDeployment%2FAzureDeployment.json)

### Settings
| Setting Name | Description |
|--|--|
|baseName| Name use as base-template to named the resources deployed in Azure.|
|PeopleLayoutName|The name of the Asset Layout in Hudu used to track People / Users. (Must exist already with a richtext Microsoft 365 field)|
