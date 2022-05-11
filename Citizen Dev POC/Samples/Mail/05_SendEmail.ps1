# Sends an email
# Minimum Application Permission: Mail.Send
# https://docs.microsoft.com/en-us/graph/api/user-sendmail

# Variables
$mailFrom    = "user@contoso.com"
$mailTo      = "recipient@contoso.com"
$mailsubject = "Hello MS Graph!"
$mailContent = "This is a sample mail sent via MS Graph. How cool is this?"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Create mail body
$emailBody = [ordered]@{
  message = @{
    subject = $mailsubject
    body = @{
      contentType = "Text"
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

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - send email
$uri = "https://graph.microsoft.com/v1.0/users/$MailFrom/sendMail"
Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri -Method Post -Body $jsonBody