# Creates a channel on an existing MS Team
# Minimum Application Permission: Group.ReadWrite.All
# https://docs.microsoft.com/en-us/graph/api/channel-post

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

