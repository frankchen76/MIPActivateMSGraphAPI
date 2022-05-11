#requires -version 5
#requires -Module Microsoft.Graph.Authentication
#requires -Module Microsoft.Graph.Users
#requires -Module Microsoft.Graph.Reports

# Call Graph SDK from PowerShell - Demo.ps1 - Get all guest users and their last logon time

# API Permissions: User.Read.All, AuditLog.Read.All

# Setting up for use with application credentials
# $clientId = "<guid>"
# $tenantId = "<guid>"
# $certPath = "c:\path\to\certificate.pfx"

# switch to PS 7 environment
# connect to m365x725618 
# $clientId = "e2b23b83-4856-4029-9284-54b08a285564"
# $tenantId = "8a5ee357-7de0-4836-ab20-9173b12cdce9"
# $certPath = "C:\AzureDevOps\PFEProjects-Private\PS-Samples\SPO\Authentication\SPOFullTrustCert\SPOFullTrust.pfx"
$clientId = "37ef28ef-bfeb-40a5-92d9-0c3fe99889c8"
$tenantId = "5ce1ecac-a23b-4b82-81d2-a5ac987fa83b"
$certPath = "C:\Projects\MIPs\Activate MS Graph for M365\Cert\ActivateMSGraphAPICert.pfx"


$certificate = Get-PfxCertificate -FilePath $certPath
Connect-Graph -ClientId $clientId -TenantId $tenantId -Certificate $certificate

# If you want to use delegated credentials, replace the above lines with:
# Connect-Graph -Scopes 'User.Read.All','AuditLog.Read.All'
# This requires global admin, or whole-organization consent previously granted by a local admin

# Connect-Graph -Scopes 'User.Read.All', 'AuditLog.Read.All'

# Disconnect-Graph

# Run Graph API query
$allUsers = Get-MgUser -Filter "userType eq 'Guest'"

# Number of total items
Write-Host "Number of items :" $allUsers.count

$guestLogins = [System.Collections.ArrayList]::new()

foreach ($user in $allUsers) {

    $signIn = Get-MgAuditLogSignIn -Filter "userId eq '$($user.Id)' and status/errorCode eq 0" -Top 1 -Sort 'createdDateTime desc'

    [void]$guestLogins.Add([PSCustomObject]@{
            Mail      = $user.Mail
            # Reports "1 January, 0001 00:00:00" if the guest user has never signed in
            LastLogin = if ($signIn) { $signIn.CreatedDateTime }
            else { [datetime]::MinValue }
        })

}

#Show in GridView
$guestLogins | Out-GridView

# Save to CSV file
#$guestLogins | Export-Csv .\GuestLastLogon.csv