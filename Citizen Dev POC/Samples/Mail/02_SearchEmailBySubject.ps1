# Gets email for a given user that meet search criteria
# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages

function Format-EmailMessages {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$messages
    )
    
    $messages | Select-Object `
    Id, `
    Subject, `
    ConversationId, `
    ParentFolderId, `
    CreatedDateTime, `
    @{Label="From";Expression={$_.From.EmailAddress.Address}}, `
    IsRead `
    | Sort-Object CreatedDateTime -Descending
}

# Variables
$user = 'user@contoso.com'
$searchString = "Weekly digest: Office 365 changes"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - search for messages
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages?`$search=`"$searchString`""
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Output - Format messages
Format-EmailMessages $messages