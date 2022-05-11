function Import-MSALAuthenticaionHelper {
    <#
    .SYNOPSIS
    This function will load the PowerShell Module for Microsoft Authentication Library (MSAL.PS). This is used to handle MSAL Authentication and Access Token management.

    .OUTPUTS
    Returns a boolean value based on success. 

    .EXAMPLE
    Import-MSALAuthenticaionHelper
    #>

    [CmdletBinding()]
    [OutputType([boolean])]     
    param()
  
    try {

        if (Get-Module -Name MSAL.PS -ListAvailable -ErrorAction:SilentlyContinue) {
            return $true;
        } 
        else { 
            Install-Module MSAL.PS 
        }
        Import-Module MSAL.PS
        return $true;
    }        
    catch {     
        Write-Error "An error occurred while loading the MSAL.PS .\r\nEnsure you've performed the setup";
        return $false;
    }
}

function Get-AccessToken {
    <#
	.SYNOPSIS
    Uses the MSAL Library to authenticate against Azure AD using OAUTH using client credential flow with either a client secret or a certificate    
	
    .PARAMETER TenantName
    The name of the tenant. i.e., contoso.onmicrosoft.com or contoso.com

    .PARAMETER ClientId
    The ClientId of the app registered in Azure AD.

    .PARAMETER ClientSecret
    The client secret for the app registered in Azure AD.

    .PARAMETER CertificateThumbprint
    The client certificate thumbprint of the app registered in Azure AD.
    
    .OUTPUTS
    Returns an access token.

    .EXAMPLE
    $authResult = Get-AccessToken -Tenant <tenant ID> -ClientId <client ID> -Certificate '<certificate thumbprint>'
    $authResult = Get-AccessToken -Tenant <tenant ID> -ClientId <client ID> -ClientSecret '<client Secret>'
    #>
    
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = "CertificateThumbprint")]
        [Parameter(Mandatory = $true, ParameterSetName = "ClientSecret")]
        [string]$Tenant,
        
        [Parameter(Mandatory = $true, ParameterSetName = "CertificateThumbprint")]
        [Parameter(Mandatory = $true, ParameterSetName = "ClientSecret")]
        [System.Guid]$ClientID,

        [Parameter(Mandatory = $true, ParameterSetName = "CertificateThumbprint")]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory = $true, ParameterSetName = "ClientSecret")]
        [string]$ClientSecret
    )

    begin {
        $scopes = [string[]]("https://graph.microsoft.com/.default")

        if (!(Import-MSALAuthenticaionHelper)) {
            return;
        }
    }
    process {
        try {

            switch($PSCmdlet.ParameterSetName){
                "CertificateThumbprint" {
                    $certificate = Get-ChildItem Cert:\CurrentUser\my | ? { $_.Thumbprint -eq $CertificateThumbprint };
                    $accessToken = Get-MsalToken -TenantId $Tenant -ClientId $ClientID -ClientCertificate $certificate -Scopes $scopes -ForceRefresh
                }
                "ClientSecret"{
                    $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
                    $accessToken = Get-MsalToken -TenantId $Tenant -ClientId $ClientID -ClientSecret $secureSecret -Scopes $scopes -ForceRefresh
                }
            }
            
            Write-Host "Success!" -ForegroundColor Green;
            return $accessToken
        }
        catch {
            Write-Host
            Write-Error $_.Exception.Message
        }
    }
}

function Get-RawTokenDetails {
    
    param
    (
        [Parameter(Mandatory = $true)]
        $AccessToken
    )

    $tokenData = $AccessToken.AccessToken.Split('.')[1]

    if ($tokendata.Length % 4 -gt 0) {
        $tokendata = $tokendata.PadRight($tokendata.Length + 4 - $tokendata.Length % 4, '=');
    }
    $results = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($tokendata)) | ConvertFrom-Json;

    return $results
}

function Get-TokenExpiry {
    param
    (
        [Parameter(Mandatory = $true)]
        $AccessToken
    )
    try {
        
        $tokenExpiration = $AccessToken.ExpiresOn.LocalDateTime;
        $diff = $tokenExpiration - (Get-Date)
        
        $TokenHealth = New-Object PSObject -prop @{
            TimeToExpiry = $diff.Minutes.ToString() + " min " + $diff.Seconds.ToString() + " sec"
            IsExpired    = (Get-Date) -gt $tokenExpiration   
        }      
        return $TokenHealth
    }
    catch {
    }
}

function Invoke-MSGraphQuery {
    <#
	.SYNOPSIS
    This function will help with making API calls to Microsoft Graph API and assist with pagination.

    .PARAMETER AccessToken
    Valid OAuth token scoped with the correct permissions.

    .PARAMETER Uri
    The API Uri you wish to access.

    .PARAMETER Method
    The HTTP method type. The default is 'GET'. 
    GET, PATCH, POST and DELETE are the only supported methods at this time.

    .PARAMETER Body
    When issuing POST or PATCH methods, you must include a json formatted body.
    
    .EXAMPLE
    $user = 'user@contoso.com'
    $uri = "https://graph.microsoft.com/v1.0/users/$user"
    Invoke-GraphQuery -Uri $uri
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNull()]
        [Microsoft.Identity.Client.AuthenticationResult]$AccessToken,
        
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNull()]
        [System.String]$Uri,
      
        [Parameter(Mandatory = $false)]
        [ValidateSet('Get', 'Post', 'Patch', 'Delete')]
        [Microsoft.PowerShell.Commands.WebRequestMethod]$Method = 'Get',

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        $Body
    )
    begin {
        $contentType = 'application/json'
        $queryResults = @()
        $originalUri = $Uri
        $pages = 1

        Write-Progress -Id 1 -Activity "Executing query: $Uri" -CurrentOperation "Invoking MS Graph API"      
    }
    process {
        $reqHeader = @{
            'Content-Type'  = $contentType
            'Authorization' = 'Bearer ' + $AccessToken.AccessToken
        }

        Write-Progress -Id 1 -Activity "Querying Microsoft Graph API."
        switch ($Method) {
            "Get"{
                do{
                    $results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $Method
                    if ($null -ne ($results.value)){ 
                        $queryResults += $results.value 
                    }
                    else {
                        $queryResults += $results
                    }          
                    $uri = $results.'@odata.nextlink'
                    Write-Progress -Id 1 -Activity "Querying Microsoft Graph API." -Status "$($queryResults.Count) results from $pages page(s)."
                    $pages++
                }
                until ([String]::IsNullOrEmpty($uri))
            }
            "Post"{
                if ($null -eq $Body) {
                    throw "When issuing a PATCH or POST, you must include a json formatted body."                
                }
                $results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $Method -Body $Body
                $queryResults += $results
            }
            "Patch"{
                if ($null -eq $Body) {
                    throw "When issuing a PATCH or POST, you must include a json formatted body."               
                }    
                $results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $Method -Body $Body
                $queryResults += $results
            }
            "Delete"{
                $results = Invoke-RestMethod -Headers $reqHeader -Uri $uri -Method $Method
                $queryResults += $results
            }
        }

        Write-Progress -Id 1 -Activity "Querying Microsoft Graph API." -Completed           
        write-host "$originalUri | Found:" ($queryResults).Count
    }
    end {
        Return $queryResults
    }    
}

Export-ModuleMember -Function * -Cmdlet * -Variable * -Alias *