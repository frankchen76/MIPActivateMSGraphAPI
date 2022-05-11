# Gets the parent folder of email messages that meet search criteria for a given user.
# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages
# https://docs.microsoft.com/en-us/graph/api/mailfolder-get


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
$results = @()
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages?`$search=`"$searchString`""
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Query MS Graph - get parent folder
$messages | % {
    $uri = "https://graph.microsoft.com/v1.0/users/$user/MailFolders/$($_.parentFolderId)"
    $m = Invoke-MSGraphQuery -AccessToken $GraphAPIAccessToken -Uri $uri;
    $results += $m 
}

# Output
$results