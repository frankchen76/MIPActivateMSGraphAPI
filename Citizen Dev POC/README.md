# Introduction
This repository contains the following samples for citizen developer. 
* Groups: the PowerShell samples to retrieve Groups informaiton
* Mail: the PowerShell samples for Mail
* Reporting: the PowerShelll samples for reporting
* Sites: the PowerShell samples for SPO sites
* Teams: the PowerShell samples for MS Teams
* Users: the PowerShell samples for Usres
* Power Automate (FLow): The Power Automate flow packages. 

# Instruction: 
All PowerShell scripts are reading AAD Application Client Id and Secret information from `clientconfiguration.json` file. Please create this file with below content under `config` folder. 
```JSON
{
    "TenantId": "[tenant-id]",
    "ClientId": "[client-id]]",
    "Thumbprint": "[certificate-thumbprint]"
}
```
