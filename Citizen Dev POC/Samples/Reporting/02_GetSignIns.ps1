# Retrieves the Azure AD user sign-ins for your tenant
# Minimum Application Permission: AuditLog.Read.All and Directory.Read.All
# https://docs.microsoft.com/en-us/graph/api/signin-list

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Sign Ins
$uri = "https://graph.microsoft.com/beta/auditLogs/signIns"
$signIns = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri
$signIns | Select-Object -First 5