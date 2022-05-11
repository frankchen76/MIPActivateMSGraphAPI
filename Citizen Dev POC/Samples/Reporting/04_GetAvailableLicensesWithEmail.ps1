# Get Available Licenses
# Minimum Application Permission: 
#  - For Licensing: Organization.Read.All
#  - For email: Mail.send
# https://docs.microsoft.com/en-us/graph/api/subscribedsku-list

# Variables
$mailFrom      = "sender@contoso.com"
$mailTo        = "receiver@contoso.com"
$mailsubject   = "Azure/O365 Licensing Report"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - send email
$uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
$subscribedSkus = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Build licensing info array
$licensingInfo = @()
foreach ($subscribedSku in $subscribedSkus){
    $licenseHash = [ordered]@{
        SKUName        = $subscribedSku.skuPartNumber
        SKUId          = $subscribedSku.skuId
        Status         = $subscribedSku.capabilityStatus
        AllocatedUnits = $subscribedSku.prepaidUnits.enabled
        ConsumedUnits  = $subscribedSku.consumedUnits
        AvailableUnits = $subscribedSku.prepaidUnits.enabled - $subscribedSku.consumedUnits
    }
    $licensingInfo += New-Object PSObject -Property $licenseHash
}

# Convert licensing info to HTML table
$licensingHTMLBody = [PSCustomobject] $licensingInfo | ConvertTo-Html

# email css style
$style  ="<style>
body { font-family: Segoe UI; font-size: 13px}
h1, h5, th { text-align: center; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #7da3d8; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #efeff5; }
tr:nth-child(even) { background: #7da3d8; }
tr:nth-child(odd) { background: #b8d1f3; }
a { color: #7da3d8; }
</style>
"
# email content
$mailContent = "$style `
Hello, </br></br>
Please, find below the current licensing status in your tenant: </br></br> `
$licensingHTMLBody </br></br> Thank you."

# Create mail body
$emailBody = [ordered]@{
  message = @{
    subject = $mailsubject
    body = @{
      contentType = "HTML"
      content = $mailContent
    }
    toRecipients = @(
      @{
        emailAddress = @{
          address = $mailTo
        }
      }
    )
  }
}
$jsonBody = $emailBody | ConvertTo-Json -Depth 4

# Query MS Graph - send email
$uri = "https://graph.microsoft.com/v1.0/users/$MailFrom/sendMail"
Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri -Method Post -Body $jsonBody
