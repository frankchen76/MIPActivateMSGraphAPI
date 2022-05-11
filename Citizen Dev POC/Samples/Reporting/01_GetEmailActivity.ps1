# Gets details about email activity users have performed.
# Minimum Application Permission: Reports.Read.All
# https://docs.microsoft.com/en-us/graph/api/reportroot-getemailactivityuserdetail

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get email activity
$uri = "https://graph.microsoft.com/beta/reports/getEmailActivityUserDetail(period='D180')?`$format=application/json"
$activity = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri
$activity

