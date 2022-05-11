# Get a message header for a given message
# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages

# Variables
$user = 'user@contoso.com'
$searchString = "Weekly digest: Office 365 changes"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - get messages
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages?`$search=`"$searchString`""
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Query MS Graph - get headers from first message
$messageId = $messages[0].id
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages/$messageId/?`$select=internetMessageHeaders"
$m = Invoke-MSGraphQuery -AccessToken $graphAPIAccessToken -Uri $uri

# Output
$rawHeader = $m.internetMessageHeaders
$rawHeader