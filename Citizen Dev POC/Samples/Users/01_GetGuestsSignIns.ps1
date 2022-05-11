# Gets the sign in information of guest users
# Minimum Application Permission: 
#   - To list users: User.Read.All
#   - To get sign in info: AuditLog.Read.All and Directory.Read.All (see note below in code on using beta endpoint)
# https://docs.microsoft.com/en-us/graph/api/user-list  
# https://docs.microsoft.com/en-us/graph/api/signin-list


# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get guest users
$uri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'"
$guestUsers = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# NOTE: For NON-PRODUCTION environments, we can get the signInActivity information using the Beta endpoint of the /users API (see below commented query).
# This eliminates the need to have Audit Log permissions.
# $uri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Guest'&`$select=displayName,userPrincipalName, mail, id, CreatedDateTime, signInActivity, UserType&`$top=999"
# $guestUsers = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

$guests= @()
foreach ($guestUser in $guestUsers)
{
    # Query MS Graph - Get guest user sign ins
    $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=userID eq '"+$guestUser.id+"' and status/errorCode eq 0&`$top=1&`$orderBy=createdDateTime desc"
    $signIns = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri
    if ($signIns){
        $row = @{ 
            UPN = $guestUser.userPrincipalName
            Email = $signIns.userPrincipalName
            LastSigInDate = $signIns.createdDateTime            
        }
        $guests += $(new-object psobject -Property $row)
    }else{
        $row = @{ 
            UPN = $guestUser.userPrincipalName
            Email = $guestUser.mail
            LastSigInDate = "Longer than the AuditLog period"
        }
        $guests += $(new-object psobject -Property $row)
    }  
}

# Output
$guests | Out-GridView