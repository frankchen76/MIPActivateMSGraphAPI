# Creates a user and assigns licenses with specific service plans enabled.
# Minimum Application Permission (to create users and assign license) : User.ReadWrite.All
# https://docs.microsoft.com/en-us/graph/api/user-post-users
# https://docs.microsoft.com/en-us/graph/api/user-assignlicense

# Variables
$displayName       = "Graph User"
$mailNickName      = "graphUser"
$givenName         = "Graph"
$surname           = "User"
$jobTitle          = "PowerShell Pro"
$userPrincipalName = "graphUser@contoso"
$password          = "P4ssw0rd"
$usageLocation     = "US"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Create new user body
$userBody = [ordered]@{
    accountEnabled = $true
    displayName = $displayName
    mailNickName = $mailNickName 
    givenName = $givenName
    surname = $surname
    jobTitle = $jobTitle
    userPrincipalName = $userPrincipalName
    passwordProfile =  @{
        forceChangePasswordNextSignIn = $false
        password = $password 
    }
    passwordPolicies = "DisablePasswordExpiration, DisableStrongPassword"
    usageLocation = $usageLocation
}
$jsonBody = $userBody | ConvertTo-Json -Depth 4

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - create user
$uri = "https://graph.microsoft.com/v1.0/users"
$newUser = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri -Method Post -Body $jsonBody

<#
    Office 365 E3 (ENTERPRISEPREMIUM) Sku ID: 6fd2c87f-b296-42f0-b197-1e91e994b900
    Service Plans under E3:
        ServicePlanName       ServicePlanId                       
        ---------------       -------------                       
        CDS_O365_P2           95b76021-6a53-4741-ab8b-1d1f3d66a95a
        PROJECT_O365_P2       31b4e2fc-4cd6-4e7d-9c1b-41407303bd66
        DYN365_CDS_O365_P2    4ff01e01-1ba7-4d71-8cf8-ce96c3bbcf14
        MICROSOFTBOOKINGS     199a5c09-e0ca-4e37-8f7c-b05d533e1ea2
        KAIZALA_O365_P3       aebd3021-9f8f-4bf8-bbe3-0ed2f4f047a1
        MICROSOFT_SEARCH      94065c59-bc8e-4e8b-89e5-5138d471eaff
        WHITEBOARD_PLAN2      94a54592-cd8b-425e-87c6-97868b000b91
        MIP_S_CLP1            5136a095-5cf0-4aff-bec3-e84448b38ea5
        MYANALYTICS_P2        33c4f319-9bdd-48d6-9c4d-410b750a4a5a
        BPOS_S_TODO_2         c87f142c-d1e9-4363-8630-aaea9c4d9ae5
        FORMS_PLAN_E3         2789c901-c14e-48ab-a76a-be334d9d793a
        STREAM_O365_E3        9e700747-8b1d-45e5-ab8d-ef187ceec156
        Deskless              8c7d2df8-86f0-4902-b2ed-a0458298f3b3
        FLOW_O365_P2          76846ad7-7776-4c40-a281-a386362dd1b9
        POWERAPPS_O365_P2     c68f8d98-5534-41c8-bf36-22fa496fa792
        TEAMS1                57ff2da0-773e-42df-b2af-ffb7a2317929
        PROJECTWORKMANAGEMENT b737dad2-2f6c-4c65-90e3-ca563267e8b9
        SWAY                  a23b959c-7ce8-4e57-9140-b90eb88a9e97
        INTUNE_O365           882e1d05-acd1-4ccb-8708-6ee03664b117
        YAMMER_ENTERPRISE     7547a3fe-08ee-4ccb-b430-5077c5041653
        RMS_S_ENTERPRISE      bea4c11e-220a-4e6d-8eb8-8ea15d019f90
        OFFICESUBSCRIPTION    43de0ff5-c92c-492b-9116-175376d08c38
        MCOSTANDARD           0feaeb32-d00e-4d66-bd5a-43b5b83db82c
        SHAREPOINTWAC         e95bec33-7c88-4a70-8e19-b10bd9d0c014
        SHAREPOINTENTERPRISE  5dbe027f-2339-4123-9542-606e4d348a72
        EXCHANGE_S_ENTERPRISE efb87545-963c-4e0d-99df-69c6916d9eb0
#>

# O365 E3 SKU ID
$entPremiumSkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"

# Enable SHAREPOINTWAC, SHAREPOINTENTERPRISE, EXCHANGE_S_ENTERPRISE
$servicePlanstoEnable = @("efb87545-963c-4e0d-99df-69c6916d9eb0","e95bec33-7c88-4a70-8e19-b10bd9d0c014","5dbe027f-2339-4123-9542-606e4d348a72")

# All Service Plans included in O365 E3 Sku
$servicePlans = @("95b76021-6a53-4741-ab8b-1d1f3d66a95a","31b4e2fc-4cd6-4e7d-9c1b-41407303bd66","4ff01e01-1ba7-4d71-8cf8-ce96c3bbcf14", `
        "199a5c09-e0ca-4e37-8f7c-b05d533e1ea2","aebd3021-9f8f-4bf8-bbe3-0ed2f4f047a1","94065c59-bc8e-4e8b-89e5-5138d471eaff","94a54592-cd8b-425e-87c6-97868b000b91", `
        "5136a095-5cf0-4aff-bec3-e84448b38ea5","33c4f319-9bdd-48d6-9c4d-410b750a4a5a","c87f142c-d1e9-4363-8630-aaea9c4d9ae5","2789c901-c14e-48ab-a76a-be334d9d793a", `
        "9e700747-8b1d-45e5-ab8d-ef187ceec156","8c7d2df8-86f0-4902-b2ed-a0458298f3b3","76846ad7-7776-4c40-a281-a386362dd1b9","c68f8d98-5534-41c8-bf36-22fa496fa792", `
        "57ff2da0-773e-42df-b2af-ffb7a2317929","b737dad2-2f6c-4c65-90e3-ca563267e8b9","a23b959c-7ce8-4e57-9140-b90eb88a9e97","882e1d05-acd1-4ccb-8708-6ee03664b117", `
        "7547a3fe-08ee-4ccb-b430-5077c5041653","bea4c11e-220a-4e6d-8eb8-8ea15d019f90","43de0ff5-c92c-492b-9116-175376d08c38","0feaeb32-d00e-4d66-bd5a-43b5b83db82c", `
        "e95bec33-7c88-4a70-8e19-b10bd9d0c014","5dbe027f-2339-4123-9542-606e4d348a72","efb87545-963c-4e0d-99df-69c6916d9eb0")

$disabledPlans = {$servicePlans}.Invoke()

# Remove $servicePlansToEnable from $disabledPlans
$servicePlanstoEnable | % {$disabledPlans.Remove($_)} | Out-Null

# Build license body
$licenseBody = [ordered] @{
    "addLicenses" = @(
        @{
        disabledPlans = $disabledPlans
        skuId = $entPremiumSkuId
        }
    )
    removeLicenses = @()
}
$jsonBody = $licenseBody | ConvertTo-Json -Depth 4

# Query MS Graph - assign licenses
$uri = "https://graph.microsoft.com/v1.0/users/graphUser@e3jaimes.onmicrosoft.com/assignLicense"
$newLicenses = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri -Method Post -Body $jsonBody

