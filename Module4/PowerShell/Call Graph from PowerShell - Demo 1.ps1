# Application (client) ID, tenant ID and secret
$root = $PSScriptRoot
$config = Get-Content "$root\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# ---------------- GET ACCESS TOKEN --------------------------------------------
# Construct URI
$uri = "https://login.microsoftonline.com/$($config.TenantId)/oauth2/v2.0/token"

# Construct Body
$body = @{
    client_id     = $config.ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $config.Secret
    grant_type    = "client_credentials"
}

# Get OAuth 2.0 Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# ---------------- QUERY MS GRAPH --------------------------------------------
# Specify the URI to call, method and headers.
$uri = "https://graph.microsoft.com/v1.0/groups"
$originalUri = $Uri
$method = "GET"
$contentType = 'application/json'
$queryResults = @()

# Header
$reqHeader = @{
    'Content-Type'  = $contentType
    'Authorization' = 'Bearer ' + $token
}
$results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $method
if ($null -ne ($results.value)) { $queryResults += $results.value }

write-host "`n$originalUri (PAGINATION)| Found: $($queryResults.Count) items" -ForegroundColor Cyan

$queryResults | Out-GridView