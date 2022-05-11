# Deletes a user
# Minimum Application Permission: User.ReadWrite.All
# https://docs.microsoft.com/en-us/graph/api/user-delete

# Variables
$user = "user@contoso.com"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get guest users
$uri = "https://graph.microsoft.com/v1.0/users/$user"
$userToDelete = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri -Method Delete
$userToDelete




