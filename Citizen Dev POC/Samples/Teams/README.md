# Microsoft Teams Samples
---------------------------------------
## NewTeamChannel.ps1
Creates a channel on an existing MS Team 

> **_NOTES:_**  
Minimum Application Permission: Group.ReadWrite.All  
https://docs.microsoft.com/en-us/graph/api/channel-post

```powershell
# Variables
$teamDisplayName = "<Team Display Name>"
$channelDisplayName = "<Channel Display Name>"
$channelDescription = "<Channel Description>"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get Team
$uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$teamDisplayName'"
$team = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Query MS Graph - Create channel
$uri = "https://graph.microsoft.com/v1.0/teams/"+$team.ID+"/channels"
$body = @{
    displayName = $ChannelDisplayName
    description = $ChannelDescription
}
$jsonBody = $body | ConvertTo-Json
$createChannel = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri -Method Post -Body $jsonBody
$createChannel 

```

**Sample Output**
```
https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Mar Even Cooler Team' | Found: 1
https://graph.microsoft.com/v1.0/teams/592fa9a5-d54f-4c82-a5f0-3f30c81c4814/channels | Found: 1

@odata.context : https://graph.microsoft.com/v1.0/$metadata#teams('592fa9a5-d54f-4c82-a5f0-3f30c81c4814')/channels/$entity
id             : 19:c0f5b168e1cc4caa80e3679ab467233f@thread.skype
displayName    : New Graph Channel
description    : Graph Description
email          :
webUrl         : https://teams.microsoft.com/l/channel/19%3ac0f5b168e1cc4caa80e3679ab467233f%40thread.skype/New+Graph+Channel?groupId=592fa9a5-d54f-4c82-a5f
                 0-3f30c81c4814&tenantId=73dbc03f-e5a3-435f-b5f9-798f31c9e140

```