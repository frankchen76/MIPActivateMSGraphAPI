# Introduction
This repository contains Flight Team sample

# Instructure
* Create local.settings.json file under project folder and include the following contents: 

```JSON
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "AzureWebJobsDashboard": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet",
    "TenantId": "[tenant-id]",
    "TenantName": "[tenant-name]",
    "AppId": "[client-id]",
    "AppSecret": "[client-secret]",
    "TeamAppToInstall": "1542629c-01b3-4a6d-8f76-1938b779e48d",
    "FlightAdminSite": "[sitecollection-name]",
    "FlightLogFile": "Flight Log.docx",
    "NotificationAppId": "YOUR CROSS DEVICE EXPERIENCE HOST NAME"
  }
}
```

