# Call Graph from PowerShell - Demo1.ps1 - Get all groups (paging vs no paging)
# API Permissions: Group.Read.All
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
$pages = 1

# Header
$reqHeader = @{
    'Content-Type'  = $contentType
    'Authorization' = 'Bearer ' + $token
}

# Take care of pagination to ensure ALL results are returned
# Keep querying MS Graph until there is no @odata.nextLink
do {
    $results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $method
    
    if ($null -ne ($results.value)) { $queryResults += $results.value }
    else { $queryResults += $results }          
    $uri = $results.'@odata.nextlink'
    $pages++
}
until ([String]::IsNullOrEmpty($uri))

# All results
write-host "`n$originalUri (PAGINATION)| Found: $($queryResults.Count) items" -ForegroundColor Cyan
$queryResults | Out-GridView
#$queryResults




