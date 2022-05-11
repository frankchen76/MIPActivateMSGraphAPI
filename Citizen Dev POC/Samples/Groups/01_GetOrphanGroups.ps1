# Gets groups that are either orphan (no owners) or only have one owner.
# Minimum Application Permission: Group.Read.All
# https://docs.microsoft.com/en-us/graph/api/group-get
# This script requires at least PowerShell 6


# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken = Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph
$uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c+eq+'Unified')"
$query = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

$noOwner = @()
$oneOwner = @()

foreach ($group in $query)
{

    # get owners
    $uri = "https://graph.microsoft.com/v1.0/groups/"+$group.id+"/owners"
    $query = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri
    $owners = $query
    $OwnerCount = $owners.Count

    # get members 
    $uri = "https://graph.microsoft.com/v1.0/groups/"+$group.id+"/members"
    $query = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri
    $members = $query
    $membersCount = $members.count
    $team=""
    if($group.resourceProvisioningOptions.Contains("Team")){$team="Team"}
    # No owner
    if ($OwnerCount -eq 0){
        $row = @{ 
                    Displayname= $group.displayName
                    GroupID = $group.id
                    ResourceProvisioningOptions = $team
                    MembersCount= $membersCount
                }
         $noOwner += $(new-object psobject -Property $row)  
    }
    # One Owner
    if ($OwnerCount -eq 1){
        $row = @{ 
                    Displayname= $group.displayName
                    GroupID = $group.id
                    ResourceProvisioningOptions =$team
                    MembersCount= $membersCount
                }
         $OneOwner += $(new-object psobject -Property $row)
    }
}

# Output - No Owner
$noOwner | Out-GridView

# Output - One Owner
$oneOwner | Out-GridView