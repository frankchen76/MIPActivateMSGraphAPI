# Reporting and Audits
---------------------------------------
## GetEmailActivity.ps1
Gets details about email activity users have performed. 

> **_NOTES:_**   
Minimum Application Permission: Reports.Read.All  
https://docs.microsoft.com/en-us/graph/api/reportroot-getemailactivityuserdetail

```powershell
# Activity
$uri = "https://graph.microsoft.com/beta/reports/getEmailActivityUserDetail(period='D7')?`$format=application/json"
$activity =Invoke-MSGraphQuery -AccessToken $GraphAPIAccessToken -Uri $uri;
$activity[0]
```

**Sample Output**
```
https://graph.microsoft.com/beta/reports/getEmailActivityUserDetail(period='D7')?$format=application/json | Found: 9


@odata.type       : #microsoft.graph.emailActivityUserDetail
reportRefreshDate : 2019-12-21
userPrincipalName : CharlotteRm1@contoso.com
displayName       : #Charlotte Room 1
isDeleted         : False
deletedDate       :
lastActivityDate  : 2019-09-09
sendCount         : 0
receiveCount      : 0
readCount         : 0
assignedProducts  : {OFFICE 365 E5}
reportPeriod      : 7

```
---------------------------------------

## GetSignIns.ps1
Retrieves the Azure AD user sign-ins for your tenant

> **_NOTES:_**  
Minimum Application Permission: AuditLog.Read.All and Directory.Read.All  
https://docs.microsoft.com/en-us/graph/api/signin-list

```powershell
# Audit Logs
$uri = "https://graph.microsoft.com/beta/auditLogs/signIns"
$signIns = Invoke-MSGraphQuery -AccessToken $GraphAPIAccessToken -Uri $uri;
$signIns[0]

```

**Sample Output**
```
https://graph.microsoft.com/beta/auditLogs/signIns | Found: 3125


id                                : 98cdfe54-d310-4759-861a-edfd4a469600
createdDateTime                   : 2019-12-23T19:26:34.1263588Z
userDisplayName                   : On-Premises Directory Synchronization Service Account
userPrincipalName                 : sync_mail01_73d7ca34ee96@brndv.onmicrosoft.com
userId                            : 4809c180-1585-4d91-8c40-80ec7e879a80
appId                             : cb1056e2-e479-49de-ae31-7812af012ed8
appDisplayName                    : Microsoft Azure Active Directory Connect
ipAddress                         : 8.37.44.13
clientAppUsed                     : Mobile Apps and Desktop clients
correlationId                     : 9a2c4a74-81ea-4bdc-9711-905938753d80
conditionalAccessStatus           : success
originalRequestId                 : 98cdfe54-d310-4759-861a-edfd4a469600
isInteractive                     : False
tokenIssuerName                   :
tokenIssuerType                   : AzureAD
processingTimeInMilliseconds      : 99
riskDetail                        : none
riskLevelAggregated               : none
riskLevelDuringSignIn             : none
riskState                         : none
riskEventTypes                    : {}
resourceDisplayName               : Windows Azure Active Directory
resourceId                        : 00000002-0000-0000-c000-000000000000
authenticationMethodsUsed         : {}
alternateSignInName               : Sync_MAIL01_73d7ca34ee96@brndv.onmicrosoft.com
servicePrincipalName              :
servicePrincipalId                :
mfaDetail                         :
status                            : @{errorCode=0; failureReason=; additionalDetails=}
deviceDetail                      : @{deviceId=; displayName=; operatingSystem=Windows 8; browser=Rich Client 5.0.5.0; isCompliant=; isManaged=; trustType=}
location                          : @{city=Charlotte; state=North Carolina; countryOrRegion=US; geoCoordinates=}
appliedConditionalAccessPolicies  : {@{id=8d9d1ec7-6a67-452d-be02-ee8fb4a8c42c; displayName=Combined Security Registration; enforcedGrantControls=System.Object[];
                                    enforcedSessionControls=System.Object[]; result=notApplied; conditionsSatisfied=none; conditionsNotSatisfied=application},
                                    @{id=RequireMfaForAzureResourceManager; displayName=Baseline policy: Require MFA for Service Management;
                                    enforcedGrantControls=System.Object[]; enforcedSessionControls=System.Object[]; result=notApplied; conditionsSatisfied=none;
                                    conditionsNotSatisfied=application}, @{id=RequireMfaForAdmins; displayName=Baseline policy: Require MFA for admins;
                                    enforcedGrantControls=System.Object[]; enforcedSessionControls=System.Object[]; result=notApplied; conditionsSatisfied=none;
                                    conditionsNotSatisfied=users}, @{id=BlockLegacyAuthentication; displayName=Baseline policy: Block legacy authentication;
                                    enforcedGrantControls=System.Object[]; enforcedSessionControls=System.Object[]; result=notApplied; conditionsSatisfied=users;
                                    conditionsNotSatisfied=clientType}}
authenticationProcessingDetails   : {}
networkLocationDetails            : {}
authenticationDetails             : {@{authenticationStepDateTime=2019-12-23T19:26:34.1263588Z; authenticationMethod=PHS; authenticationMethodDetail=; succeeded=True;
                                    authenticationStepResultDetail=; authenticationStepRequirement=Primary Authentication}}
authenticationRequirementPolicies : {}
```


---------------------------------------

## GetUserActivityReport.ps1
This is a sample script that shows how to query the 'reports' endpoint in MS Graph in order to gather user activity information from various O365 workloads.
All the gathered information is presented in one report.
This script is an adapted version of the one found here: https://github.com/12Knocksinna/Office365itpros/blob/master/GetGraphUserStatisticsReportV2.PS1

> **_NOTES:_**  
Minimum Application Permission: Reports.Read.All  
https://docs.microsoft.com/en-us/graph/api/resources/report?view=graph-rest-1.0


```powershell
# Variables
$outFile = "c:\temp\Office365TenantUsage.csv"
$period = "D180"
$StartTime1 = Get-Date

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# NOTE: substring(3) in the queries below is to get rid of initial characters in the response (ï»¿) we get from the reports API
# Query MS Graph - Get Teams Usage Data  
Write-Host "`nFetching Teams user activity data from Microsoft Graph" -ForegroundColor Cyan 
$TeamsUserReportsURI = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='$period')"
$TeamsUserData = (Invoke-MSGraphQuery -Uri $TeamsUserReportsURI -AccessToken $graphApiAccessToken).substring(3) | ConvertFrom-Csv 

# Query MS Graph - Get OneDrive for Business data
Write-Host "`nFetching OneDrive for Business user activity data from Microsoft Graph" -ForegroundColor Cyan
$OneDriveUsageURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$period')"
$OneDriveData = (Invoke-MSGraphQuery -Uri $OneDriveUsageURI -AccessToken $graphApiAccessToken).substring(3) | ConvertFrom-Csv 

# Query MS Graph - Get Exchange Activity Data
Write-Host "`nFetching Exchange Online user activity data from Microsoft Graph" -ForegroundColor Cyan
$EmailReportsURI = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='$period')"
$EmailData = (Invoke-MSGraphQuery -Uri $EmailReportsURI -AccessToken $graphApiAccessToken).substring(3) | ConvertFrom-Csv 
# Query MS Graph - Get Exchange Storage Data    
$MailboxUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$period')"
$MailboxUsage = (Invoke-MSGraphQuery -Uri $MailboxUsageReportsURI -AccessToken $graphApiAccessToken).substring(3) | ConvertFrom-Csv 

# Query MS Graph - Get SharePoint usage data
Write-Host "`nFetching SharePoint Online user activity data from Microsoft Graph" -ForegroundColor Cyan
$SPOUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='$period')"
$SPOUsage = (Invoke-MSGraphQuery -Uri $SPOUsageReportsURI -AccessToken $graphApiAccessToken).substring(3) | ConvertFrom-Csv 

# Query MS Graph - Get Yammer usage data
Write-Host "`nFetching Yammer user activity data from Microsoft Graph" -ForegroundColor Cyan
$YammerUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getYammerActivityUserDetail(period='$period')"
$YammerUsage = (Invoke-MSGraphQuery -Uri $YammerUsageReportsURI -AccessToken $graphApiAccessToken).substring(3) | ConvertFrom-Csv 

# Create hash table for user sign in data
$UserSignIns = @{}
# And hash table for the output data
$DataTable = @{}

# Query MS Graph - Get User sign in data
Write-Host "`nFetching user sign-in data from Microsoft Graph" -ForegroundColor Cyan
$URI = "https://graph.microsoft.com/beta/users?`$select=displayName,userPrincipalName, mail, id, CreatedDateTime, signInActivity, UserType&`$top=999"
$SignInData = (Invoke-MSGraphQuery -Uri $URI -AccessToken $graphApiAccessToken) 
# Update the user sign in hash table
ForEach ($U in $SignInData) {
   If ($U.UserType -eq "Member") {
     $DataTable.Add([String]$U.UserPrincipalName,$Null)
     If ($U.SignInActivity.LastSignInDateTime) {
          $LastSignInDate = Get-Date($U.SignInActivity.LastSignInDateTime) -format g
          $UserSignIns.Add([String]$U.UserPrincipalName, $LastSignInDate) }
}}

Write-Host "`n*** Processing activity data fetched from Microsoft Graph" -ForegroundColor Cyan

$StartTime2 = Get-Date
# Process Teams Data
ForEach ($T in $TeamsUserData) {
   If ([string]::IsNullOrEmpty($T."Last Activity Date")) { 
      $TeamsLastActivity = "No activity"
      $TeamsDaysSinceActive = "N/A" }
   Else {
      $TeamsLastActivity = Get-Date($T."Last Activity Date") -format "dd-MMM-yyyy" 
      $TeamsDaysSinceActive = (New-TimeSpan($TeamsLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     TeamsUPN               = $T."User Principal Name"
     TeamsLastActive        = $TeamsLastActivity  
     TeamsDaysSinceActive   = $TeamsDaysSinceActive      
     TeamsReportDate        = Get-Date($T."Report Refresh Date") -format "dd-MMM-yyyy"  
     TeamsLicense           = $T."Assigned Products"
     TeamsChannelChats      = $T."Team Chat Message Count"
     TeamsPrivateChats      = $T."Private Chat Message Count"
     TeamsCalls             = $T."Call Count"
     TeamsMeetings          = $T."Meeting Count"
     TeamsRecordType        = "Teams"}
   $DataTable[$T."User Principal Name"] = $ReportLine} 

# Process Exchange Data
ForEach ($E in $EmailData) {
   $ExoDaysSinceActive = $Null
   If ([string]::IsNullOrEmpty($E."Last Activity Date")) { 
      $ExoLastActivity = "No activity"
      $ExoDaysSinceActive = "N/A" }
   Else {
      $ExoLastActivity = Get-Date($E."Last Activity Date") -format "dd-MMM-yyyy"
      $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days }
  $ReportLine  = [PSCustomObject] @{          
     ExoUPN                = $E."User Principal Name"
     ExoDisplayName        = $E."Display Name"
     ExoLastActive         = $ExoLastActivity   
     ExoDaysSinceActive    = $ExoDaysSinceActive    
     ExoReportDate         = Get-Date($E."Report Refresh Date") -format "dd-MMM-yyyy"  
     ExoSendCount          = [int]$E."Send Count"
     ExoReadCount          = [int]$E."Read Count"
     ExoReceiveCount       = [int]$E."Receive Count"
     ExoIsDeleted          = $E."Is Deleted"
     ExoRecordType         = "Exchange Activity"}
   [Array]$ExistingData = $DataTable[$E."User Principal Name"] 
   [Array]$NewData = $ExistingData + $ReportLine
   $DataTable[$E."User Principal Name"] = $NewData } 
  
ForEach ($M in $MailboxUsage) {
   If ([string]::IsNullOrEmpty($M."Last Activity Date")) { 
      $ExoLastActivity = "No activity" }
   Else {
      $ExoLastActivity = Get-Date($M."Last Activity Date") -format "dd-MMM-yyyy"
      $ExoDaysSinceActive = (New-TimeSpan($ExoLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     MbxUPN                = $M."User Principal Name"
     MbxDisplayName        = $M."Display Name"
     MbxLastActive         = $ExoLastActivity 
     MbxDaysSinceActive    = $ExoDaysSinceActive          
     MbxReportDate         = Get-Date($M."Report Refresh Date") -format "dd-MMM-yyyy"  
     MbxQuotaUsed          = [Math]::Round($M."Storage Used (Byte)"/1GB,2) 
     MbxItems              = [int]$M."Item Count"
     MbxRecordType         = "Exchange Storage"}
   [Array]$ExistingData = $DataTable[$M."User Principal Name"] 
   [Array]$NewData = $ExistingData + $ReportLine
   $DataTable[$M."User Principal Name"] = $NewData } 

# SharePoint data
ForEach ($S in $SPOUsage) {
   If ([string]::IsNullOrEmpty($S."Last Activity Date")) { 
      $SPOLastActivity = "No activity"
      $SPODaysSinceActive = "N/A" }
   Else {
      $SPOLastActivity = Get-Date($S."Last Activity Date") -format "dd-MMM-yyyy"
      $SPODaysSinceActive = (New-TimeSpan ($SPOLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     SPOUPN              = $S."User Principal Name"
     SPOLastActive       = $SPOLastActivity    
     SPODaysSinceActive  = $SPODaysSinceActive 
     SPOViewedEdited     = [int]$S."Viewed or Edited File Count"     
     SPOSyncedFileCount  = [int]$S."Synced File Count"
     SPOSharedExt        = [int]$S."Shared Externally File Count"
     SPOSharedInt        = [int]$S."Shared Internally File Count"
     SPOVisitedPages     = [int]$S."Visited Page Count" 
     SPORecordType       = "SharePoint Usage"}
   [Array]$ExistingData = $DataTable[$S."User Principal Name"] 
   [Array]$NewData = $ExistingData + $ReportLine
   $DataTable[$S."User Principal Name"] = $NewData }  

# OneDrive for Business data
ForEach ($O in $OneDriveData) {
   $OneDriveLastActivity = $Null
   If ([string]::IsNullOrEmpty($O."Last Activity Date")) { 
      $OneDriveLastActivity = "No activity"
      $OneDriveDaysSinceActive = "N/A" }
   Else {
      $OneDriveLastActivity = Get-Date($O."Last Activity Date") -format "dd-MMM-yyyy" 
      $OneDriveDaysSinceActive = (New-TimeSpan($OneDriveLastActivity)).Days }
   $ReportLine  = [PSCustomObject] @{          
     ODUPN               = $O."Owner Principal Name"
     ODDisplayName       = $O."Owner Display Name"
     ODLastActive        = $OneDriveLastActivity    
     ODDaysSinceActive   = $OneDriveDaysSinceActive    
     ODSite              = $O."Site URL"
     ODFileCount         = [int]$O."File Count"
     ODStorageUsed       = [Math]::Round($O."Storage Used (Byte)"/1GB,4) 
     ODQuota             = [Math]::Round($O."Storage Allocated (Byte)"/1GB,2) 
     ODRecordType        = "OneDrive Storage"}
   [Array]$ExistingData = $DataTable[$O."Owner Principal Name"] 
   [Array]$NewData = $ExistingData + $ReportLine
   $DataTable[$O."Owner Principal Name"] = $NewData }  

# Yammer Data
ForEach ($Y in $YammerUsage) {  
  If ([string]::IsNullOrEmpty($Y."Last Activity Date")) { 
      $YammerLastActivity = "No activity" 
      $YammerDaysSinceActive = "N/A" }
   Else {
      $YammerLastActivity = Get-Date($Y."Last Activity Date") -format "dd-MMM-yyyy" 
      $YammerDaysSinceActive = (New-TimeSpan ($YammerLastActivity)).Days }
  $ReportLine  = [PSCustomObject] @{          
     YUPN             = $Y."User Principal Name"
     YDisplayName     = $Y."Display Name"
     YLastActive      = $YammerLastActivity      
     YDaysSinceActive = $YammerDaysSinceActive   
     YPostedCount     = [int]$Y."Posted Count"
     YReadCount       = [int]$Y."Read Count"
     YLikedCount      = [int]$Y."Liked Count"
     YRecordType      = "Yammer Usage"}
   [Array]$ExistingData = $DataTable[$Y."User Principal Name"] 
   [Array]$NewData = $ExistingData + $ReportLine
   $DataTable[$Y."User Principal Name"] = $NewData }

#CLS
# Create set of users that we've collected data for - each of these users will be in the $DataTable with some information.
[System.Collections.ArrayList]$Users = @()
ForEach ($UserPrincipalName in $DataTable.Keys) { 
   If ($DataTable[$UserPrincipalName]) { #Info exists in datatable
   $obj = [PSCustomObject]@{ 
      UPN  = $UserPrincipalName}
   $Users.add($obj) | Out-Null }
}
$StartTime3 = Get-Date
# Set up progress bar
$ProgressDelta = 100/($Users.Count); $PercentComplete = 0; $UserNumber = 0
$OutData = [System.Collections.Generic.List[Object]]::new() # Create merged output file

# Process each user to extract Exchange, Teams, OneDrive, SharePoint, and Yammer statistics for their activity
ForEach ($UserPrincipalName in $Users) {
  $U = $UserPrincipalName.UPN
  $UserNumber++
  $CurrentStatus = $U + " ["+ $UserNumber +"/" + $Users.Count + "]"
  Write-Progress -Activity "Extracting information for user" -Status $CurrentStatus -PercentComplete $PercentComplete
  $PercentComplete += $ProgressDelta
  $ExoData = $Null; $ExoActiveData = $Null; $TeamsData = $Null; $ODData = $Null; $SPOData = $Null; $YammerData = $Null
  
  $UserData = $DataTable[$U]  # Extract data for the user - everything is in a single keyed access to the hash table

# Process Exchange Data
  [string]$ExoUPN = (Out-String -InputObject $UserData.ExoUPN).Trim()
  [string]$ExoLastActive = (Out-String -InputObject $UserData.ExoLastActive).Trim()
  If ([string]::IsNullOrEmpty($ExoUPN) -or $ExoLastActive -eq "No Activity") {
     $ExoDaysSinceActive  = "N/A"
     $EXoLastActive = "No Activity" }
  Else {
     [string]$ExoLastActive = (Out-String -InputObject $UserData.ExoLastActive).Trim()
     [string]$ExoDaysSinceActive = (Out-String -InputObject $UserData.ExoDaysSinceActive).Trim() }
 
# Parse OneDrive for Business usage data 
  [string]$ODUPN = (Out-String -InputObject $UserData.ODUPN).Trim()
  [string]$ODLastActive = (Out-String -InputObject $UserData.ODLastActive).Trim()  # Possibility of a second OneDrive account for some users.
  If (($ODLastActive -Like "*No Activity*") -or ([string]::IsNullOrEmpty($ODLastActive))) {$ODLastActive = "No Activity"} # this is a hack until I figure out a better way to handle the situation
  If ([string]::IsNullOrEmpty($ODUPN)-eq $Null -or $ODLastActive -eq "No Activity") {
     [string]$ODDaysSinceActive  = "N/A"
     [string]$ODLastActive = "No Activity"
     $ODFiles            = 0
     $ODStorage          = 0
     $ODQuota            = 1024 }
 Else {
     [string]$ODDaysSinceActive = (Out-String -InputObject $UserData.ODDaysSinceActive).Trim()
     [string]$ODLastActive = (Out-String -InputObject $UserData.ODLastActive).Trim()
     [string]$ODFiles = (Out-String -InputObject $UserData.ODFileCount).Trim()
     [string]$ODStorage = (Out-String -InputObject $UserData.ODStorageUsed).Trim()
     [string]$ODQuota = (Out-String -InputObject $UserData.ODQuota).Trim()  }

# Parse Yammer usage data; Yammer isn't used everywhere, so make sure that we record zero data 
  [string]$YUPN = (Out-String -InputObject $UserData.YUPN).Trim()
  [string]$YammerLastActive = (Out-String -InputObject $UserData.YLastActive).Trim()
  If (([string]::IsNullOrEmpty($YUPN) -or ($YammerLastActive -eq "No Activity"))) { 
     $YammerLastActive = "No Activity"  
     $YammerDaysSinceActive  = "N/A" 
     $YammerPosts             = 0
     $YammerReads             = 0
     $YammerLikes             = 0 }
 Else {
     [string]$YammerDaysSinceActive = (Out-String -InputObject $UserData.YDaysSinceActive).Trim()
     [string]$YammerPosts = (Out-String -InputObject $UserData.YPostedCount).Trim()
     [string]$YammerReads = (Out-String -InputObject $UserData.YReadCount).Trim()
     [string]$YammerLikes = (Out-String -InputObject $UserData.YLikedCount).Trim() }
  
 If ($UserData.TeamsDaysSinceActive -gt 0) {
     [string]$TeamsDaysSinceActive = (Out-String -InputObject $UserData.TeamsDaysSinceActive).Trim()
     [string]$TeamsLastActive = (Out-String -InputObject $UserData.TeamsLastActive).Trim() }
 Else { 
     [string]$TeamsDaysSinceActive = "N/A"
     [string]$TeamsLastActive = "No Activity" }
 
 If ($UserData.SPODaysSinceActive -gt 0) {
     [string]$SPODaysSinceActive = (Out-String -InputObject $UserData.SPODaysSinceActive).Trim()
     [string]$SPOLastActive = (Out-String -InputObject $UserData.SPOLastActive).Trim() }
 Else { 
     [string]$SPODaysSinceActive = "N/A"
     [string]$SPOLastActive = "No Activity" }
 
# Fetch the sign in data if available
$LastAccountSignIn = $Null; $DaysSinceSignIn = 0
$LastAccountSignIn = $UserSignIns.Item($U)
If ($LastAccountSignIn -eq $Null) { $LastAccountSignIn = "No sign in data found"; $DaysSinceSignIn = "N/A"}
  Else { $DaysSinceSignIn = (New-TimeSpan($LastAccountSignIn)).Days }
   
# Figure out if the account is used
[int]$ExoDays = 365; [int]$TeamsDays = 365; [int]$SPODays = 365; [int]$ODDays = 365; [int]$YammerDays = 365

# Base is 2 if someuse uses the five workloads because the Graph is usually 2 days behind, but we have some N/A values for days used
  If ($ExoDaysSinceActive -ne "N/A") {$ExoDays = $ExoDaysSinceActive -as [int]}
  If ($TeamsDaysSinceActive -eq "N/A") {$TeamsDays = 365} Else {$TeamsDays = $TeamsDaysSinceActive -as [int]}
  If ($SPODaysSinceActive -eq "N/A") {$SPODays = 365} Else {$SPODays = $SPODaysSinceActive -as [int]}  
  If ($ODDaysSinceActive -eq "N/A") {$ODDays = 365} Else {$ODDays = $ODDaysSinceActive -as [int]} 
  If ($YammerDaysSinceActive -eq "N/A") {$YammerDays = 365} Else {$YammerDays = $YammerDaysSinceActive -as [int]}
   
# Average days per workload used...
  $AverageDaysSinceUse = [Math]::Round((($ExoDays + $TeamsDays + $SPODays + $ODDays + $YammerDays)/5),2)

  Switch ($AverageDaysSinceUse) { # Figure out if account is used
   ({$PSItem -le 8})                          { $AccountStatus = "Heavy usage" }
   ({$PSItem -ge 9 -and $PSItem -le 50} )     { $AccountStatus = "Moderate usage" }   
   ({$PSItem -ge 51 -and $PSItem -le 120} )   { $AccountStatus = "Poor usage" }
   ({$PSItem -ge 121 -and $PSItem -le 300 } ) { $AccountStatus = "Review account"  }
   default                                    { $AccountStatus = "Account unused" }
  } # End Switch

# And an override if someone has been active in just one workload in the last 14 days
  [int]$DaysCheck = 14 # Set this to your chosen value if you want to use a different period.
  If (($ExoDays -le $DaysCheck) -or ($TeamsDays -le $DaysCheck) -or ($SPODays -le $DaysCheck) -or ($ODDays -le $DaysCheck) -or ($YammerDays -le $DaysCheck)) {
     $AccountStatus = "Account in use"}

If ((![string]::IsNullOrEmpty($ExoUPN))) {
# Build a line for the report file with the collected data for all workloads and write it to the list
  $OutLine  = [PSCustomObject] @{          
     UPN                     = $U
     DisplayName             = (Out-String -InputObject $UserData.ExoDisplayName).Trim()
     Status                  = $AccountStatus
     LastSignIn              = $LastAccountSignIn
     DaysSinceSignIn         = $DaysSinceSignIn 
     EXOLastActive           = $ExoLastActive  
     EXODaysSinceActive      = $ExoDaysSinceActive  
     EXOQuotaUsed            = (Out-String -InputObject $UserData.MbxQuotaUsed).Trim()
     EXOItems                = (Out-String -InputObject $UserData.MbxItems).Trim()
     EXOSendCount            = (Out-String -InputObject $UserData.ExoSendCount).Trim()
     EXOReadCount            = (Out-String -InputObject $UserData.ExoReadCount).Trim()
     EXOReceiveCount         = (Out-String -InputObject $UserData.ExoReceiveCount).Trim()
     TeamsLastActive         = $TeamsLastActive
     TeamsDaysSinceActive    = $TeamsDays 
     TeamsChannelChat        = (Out-String -InputObject $UserData.TeamsChannelChats).Trim()
     TeamsPrivateChat        = (Out-String -InputObject $UserData.TeamsPrivateChats).Trim()
     TeamsMeetings           = (Out-String -InputObject $UserData.TeamsMeetings).Trim()
     TeamsCalls              = (Out-String -InputObject $UserData.TeamsCalls).Trim()
     SPOLastActive           = $SPOLastActive
     SPODaysSinceActive      = $SPODays 
     SPOViewedEditedFiles    = (Out-String -InputObject $UserData.SPOViewedEdited).Trim()
     SPOSyncedFiles          = (Out-String -InputObject $UserData.SPOSyncedFileCount).Trim()
     SPOSharedExtFiles       = (Out-String -InputObject $UserData.SPOSharedExt).Trim()
     SPOSharedIntFiles       = (Out-String -InputObject $UserData.SPOSharedInt).Trim()
     SPOVisitedPages         = (Out-String -InputObject $UserData.SPOVisitedPages).Trim()
     OneDriveLastActive      = $ODLastActive
     OneDriveDaysSinceActive = $ODDaysSinceActive
     OneDriveFiles           = $ODFiles
     OneDriveStorage         = $ODStorage
     OneDriveQuota           = $ODQuota
     YammerLastActive        = $YammerLastActive  
     YammerDaysSinceActive   = $YammerDaysSinceActive
     YammerPosts             = $YammerPosts
     YammerReads             = $YammerReads
     YammerLikes             = $YammerLikes
     License                 = (Out-String -InputObject $UserData.TeamsLicense).Trim()
     OneDriveSite            = (Out-String -InputObject $UserData.ODSite).Trim()
     IsDeleted               = (Out-String -InputObject $UserData.ExoIsDeleted).Trim()
     EXOReportDate           = (Out-String -InputObject $UserData.ExoReportDate).Trim()
     TeamsReportDate         = (Out-String -InputObject $UserData.TeamsReportDate).Trim()
     UsageFigure             = $AverageDaysSinceUse }
   $OutData.Add($OutLine)   } 
 } #End processing user data

#CLS
$StartTime4 = Get-Date
$GraphTime = $StartTime2 - $StartTime1
$PrepTime = $StartTime3 - $StartTime2
$ReportTime = $StartTime4 - $StartTime3
$ScriptTime = $StartTime4 - $StartTime1
$AccountsPerMinute = [math]::Round(($Outdata.count/($ScriptTime.TotalSeconds/60)),2)
$GraphElapsed = $GraphTime.Minutes.ToString() + ":" + $GraphTime.Seconds.ToString()
$PrepElapsed = $PrepTime.Minutes.ToString() + ":" + $PrepTime.Seconds.ToString()
$ReportElapsed = $ReportTime.Minutes.ToString() + ":" + $ReportTime.Seconds.ToString()
$ScriptElapsed = $ScriptTime.Minutes.ToString() + ":" + $ScriptTime.Seconds.ToString()

Write-Host "`nStatistics for Graph Report Script V2.0"
Write-Host "---------------------------------------"
Write-Host "Time to fetch data from Microsoft Graph:" $GraphElapsed
Write-Host "Time to prepare date for processing:    " $PrepElapsed
Write-Host "Time to create report from data:        " $ReportElapsed
Write-Host "Total time for script:                  " $ScriptElapsed
Write-Host "Total accounts processed:               " $Outdata.count
Write-Host "Accounts processsed per minute:         " $AccountsPerMinute
Write-Host "`nOutput CSV file available in " $outFile

$OutData | Sort {$_.ExoLastActive -as [DateTime]} -Descending | Out-GridView  
$OutData | Sort $AccountStatus | Export-CSV $outFile -NoTypeInformation

```

**Sample Output**
```
(output below represents one user in the collection)


UPN                     : user@contoso.onmicrosoft.com
DisplayName             : User
Status                  : Account in use
LastSignIn              : 9/10/2020 16:24
DaysSinceSignIn         : 4
EXOLastActive           : 09-Sep-2020
EXODaysSinceActive      : 5
EXOQuotaUsed            : 0.03
EXOItems                : 938
EXOSendCount            : 94
EXOReadCount            : 51
EXOReceiveCount         : 135
TeamsLastActive         : 03-Sep-2020
TeamsDaysSinceActive    : 11
TeamsChannelChat        : 0
TeamsPrivateChat        : 1
TeamsMeetings           : 0
TeamsCalls              : 0
SPOLastActive           : 09-Sep-2020
SPODaysSinceActive      : 5
SPOViewedEditedFiles    : 45
SPOSyncedFiles          : 0
SPOSharedExtFiles       : 0
SPOSharedIntFiles       : 0
SPOVisitedPages         : 48
OneDriveLastActive      : 05-Sep-2019
OneDriveDaysSinceActive : 375
OneDriveFiles           : 519
OneDriveStorage         : 0.0644
OneDriveQuota           : 5120
YammerLastActive        : No Activity
YammerDaysSinceActive   : N/A
YammerPosts             : 0
YammerReads             : 0
YammerLikes             : 0
License                 : MICROSOFT POWER AUTOMATE FREE+POWER BI (FREE)+OFFICE 365 E3+MICROSOFT POWER APPS PLAN 2 TRIAL+POWER APPS PER USER PLAN+POWER BI   
                          PRO
OneDriveSite            : https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com
IsDeleted               : False
EXOReportDate           : 13-Sep-2020
TeamsReportDate         : 13-Sep-2020
UsageFigure             : 152.2

```


---------------------------------------

## GetAvailableLicensesWithEmail.ps1
Gets information on subscribed skus within a tenant and sends a summary report through email.

> **_NOTES:_**  
Minimum Application Permission:  
For Licensing: Organization.Read.All  
For email: Mail.send  
https://docs.microsoft.com/en-us/graph/api/subscribedsku-list  
https://docs.microsoft.com/en-us/graph/api/user-sendmail


```powershell

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

```

**Sample Output**
```
A formatted email with an HTML table containing the following information will be sent:


SKUName        : STREAM
SKUId          : 1f2f344a-700d-42c9-9427-5cea1d5d7ba6
Status         : Enabled
AllocatedUnits : 1000000
ConsumedUnits  : 2
AvailableUnits : 999998

SKUName        : ENTERPRISEPACK
SKUId          : 6fd2c87f-b296-42f0-b197-1e91e994b900
Status         : Enabled
AllocatedUnits : 25
ConsumedUnits  : 21
AvailableUnits : 4

SKUName        : FLOW_FREE
SKUId          : f30db892-07e9-47e9-837c-80727f46fd3d
Status         : Enabled
AllocatedUnits : 10000
ConsumedUnits  : 8
AvailableUnits : 9992

SKUName        : POWERAPPS_VIRAL
SKUId          : dcb1a3ae-b33f-4487-846a-a640262fadf4
Status         : Enabled
AllocatedUnits : 10000
ConsumedUnits  : 7
AvailableUnits : 9993

SKUName        : POWER_BI_STANDARD
SKUId          : a403ebcc-fae0-4ca2-8c8c-7a907fd6c235
Status         : Enabled
AllocatedUnits : 1000000
ConsumedUnits  : 7
AvailableUnits : 999993

```