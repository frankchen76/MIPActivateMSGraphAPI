# Call Graph from PowerShell - Demo3.ps1 - Get all guest users and they last logon time

# API Permissions: User.Read.All AND AuditLog.Read.All

# Azure AD OAuth Application Token for Graph API
# Get OAuth token for a AAD Application (returned as $token)

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
# Specify the URI to call and method
$uri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'"
$method = "GET"
$contentType = 'application/json'
$queryResults = @()
$pages = 1
$reqHeader = @{
    'Content-Type'  = $contentType
    'Authorization' = 'Bearer ' + $token
}

# Run Graph API query 
$query = Invoke-RestMethod  -Method $method -Uri $uri -Headers $reqHeader -ErrorAction Stop


# Get all items and repeat for every page (100 per page)
do {
    $results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $method

    if ($null -ne ($results.value)) { $queryResults += $results.value }
    else { $queryResults += $results }
    $uri = $results.'@odata.nextlink'
    $pages++
}
until ([String]::IsNullOrEmpty($uri))


# Number of total items
write-host "`n$originalUri | Found: $($queryResults.Count) items" -ForegroundColor Cyan

# Get last logon datetime
$GuestUsers = @()
foreach ($user in $queryResults) {
    $object = New-Object -TypeName PSObject
        
    $url = "https://graph.microsoft.com/beta/auditLogs/signIns?$('$')filter=userID eq '$($user.id)' and status/errorCode eq 0&$('$')top=1&$('$')orderBy=createdDateTime desc"
    $topSignin = Invoke-RestMethod -Headers $reqHeader  $url -Method Get
    if ($topSignin.Value) {
        $lastLogon = [DateTime]::Parse($topSignin.Value.createdDateTime)  
    }
    else {
        $lastLogon = [DateTime]::MinValue
    }
    $object | Add-Member -Name 'UserPrincipalName' -MemberType Noteproperty -Value $user.UserPrincipalName
    $object | Add-Member -Name 'LastLogon' -MemberType Noteproperty -Value $lastLogon
    $GuestUsers += $object
}

#Show in GridView
$GuestUsers  | Out-GridView

# Save to CSV file
#$GuestUsers | Export-Csv .\GuestLastLogon.csv
