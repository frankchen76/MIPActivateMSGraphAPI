# Email and Calendar Samples
---------------------------------------
## GetMessages.ps1
Gets All email messages for a given user (including deleted ones)

> **_NOTES:_**  
Minimum Application Permission: Mail.Read  
https://docs.microsoft.com/en-us/graph/api/user-list-messages

```powershell
function Format-EmailMessages {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$messages
    )
    
    $messages | Select-Object `
    Id, `
    Subject, `
    ConversationId, `
    ParentFolderId, `
    CreatedDateTime, `
    @{Label="From";Expression={$_.From.EmailAddress.Address}}, `
    IsRead `
    | Sort-Object CreatedDateTime -Descending
}

# Variables
$user = 'user@contoso.com'

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages"
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Output - Format messages
Format-EmailMessages $messages
```

## SearchEmailBySubject.ps1
Gets email for a given user that meet search criteria

> **_NOTES:_**  
Minimum Application Permission: Mail.Read  
https://docs.microsoft.com/en-us/graph/api/user-list-messages


```powershell
function Format-EmailMessages {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [object]$messages
    )
    
    $messages | Select-Object `
    Id, `
    Subject, `
    ConversationId, `
    ParentFolderId, `
    CreatedDateTime, `
    @{Label="From";Expression={$_.From.EmailAddress.Address}}, `
    IsRead `
    | Sort-Object CreatedDateTime -Descending
}

# Variables
$user = 'user@contoso.com'
$searchString = "Weekly digest: Office 365 changes"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - search for messages
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages?`$search=`"$searchString`""
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Output - Format messages
Format-EmailMessages $messages
```

## MessageFolderName.ps1
Gets the parent folder of email messages that meet search criteria for a given user.

> **_NOTES:_**  
Minimum Application Permission: Mail.Read  
https://docs.microsoft.com/en-us/graph/api/user-list-messages  
https://docs.microsoft.com/en-us/graph/api/mailfolder-get

```powershell
# Variables
$user = 'user@contoso.com'
$searchString = "Weekly digest: Office 365 changes"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - get messages
$results = @()
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages?`$search=`"$searchString`""
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Query MS Graph - get parent folder
$messages | % {
    $uri = "https://graph.microsoft.com/v1.0/users/$user/MailFolders/$($_.parentFolderId)"
    $m = Invoke-MSGraphQuery -AccessToken $GraphAPIAccessToken -Uri $uri;
    $results += $m 
}

# Output
$results
```

## GetMessageHeader.ps1
Get a message header for a given message

> **_NOTES:_**  
Minimum Application Permission: Mail.Read  
https://docs.microsoft.com/en-us/graph/api/user-list-messages

```powershell
# Variables
$user = 'user@contoso.com'
$searchString = "Weekly digest: Office 365 changes"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - get messages
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages?`$search=`"$searchString`""
$messages = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri

# Query MS Graph - get headers from first message
$messageId = $messages[0].id
$uri = "https://graph.microsoft.com/v1.0/users/$user/messages/$messageId/?`$select=internetMessageHeaders"
$m = Invoke-MSGraphQuery -AccessToken $graphAPIAccessToken -Uri $uri

# Output
$rawHeader = $m.internetMessageHeaders
$rawHeader
```

## SendEmail.ps1
Sends an email

> **_NOTES:_**  
Minimum Application Permission: Mail.Send  
https://docs.microsoft.com/en-us/graph/api/user-sendmail

```powershell    
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
```


---------------------------------------

## GetRoomsUtilization.ps1
Collect appointments from a list of resource rooms then calculate utilization for each.
Use Case: Customer is paying for rented meeting spaces throughout several states and wants to determine actual usage so they can decide whether or not to continue to pay for them

> **_NOTES:_**  
Minimum Application Permission: Calendars.Read  
https://docs.microsoft.com/en-us/graph/api/user-list-calendarview


```powershell
# Variables
$startTime = "2020-06-25"
$endTime = "2020-08-03"
$calendars = @("CharlotteRm1@contoso.com","DublinRm1@contoso.com","SeattleRm1@contoso.com")

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# Get Access Token
$graphApiAccessToken =  Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -CertificateThumbprint $config.Thumbprint

$results = $null
$results = @()
[System.DateTime]$formattedStartTime = $startTime
[System.DateTime]$formattedEndTime = $endTime
$reportDays = ((New-TimeSpan).Add($formattedEndTime.subtract($formattedStartTime))).Days

foreach($calendar in $calendars){

    # Query MS Graph - Get appointments
    $uri = "https://graph.microsoft.com/v1.0/users/$calendar/calendar/calendarView?startDateTime=$startTime&endDateTime=$endTime"
    $appointments = Invoke-MSGraphQuery -AccessToken $graphApiAccessToken -Uri $uri
    $bookableTime = New-TimeSpan
    $calEvent = [ordered]@{
        Room = $appointments[0].location.displayName
        ReportPeriod = $reportDays       
        TotalAppointments = "0"
        TotalHoursBooked = "0"
        Utilization = "0"
        Events = @()
    }
  
    foreach ($appointment in $appointments) {
        $TotalDuration = New-timespan
        if($appointment.isAllDay){
            $TotalDuration = New-TimeSpan -Hours 8
        }
        else{
      
        [System.DateTime]$start = $appointment.Start.dateTime        
        [System.DateTime]$end = $appointment.End.dateTime
        $TotalDuration = (New-TimeSpan).Add($end.Subtract($start))
      }
      
      $BookableTime += $TotalDuration;

        $event = $null
        $event += New-Object psobject -Property @{
            Room     = $appointment.location.displayName
            Subject  = $appointment.Subject
            Date     = $start.ToShortDateString()
            Duration =  $TotalDuration
        }

        $calEvent.Events += $event
    }
    $calEvent.TotalAppointments = $appointments.Count
    $calEvent.TotalHoursBooked = $BookableTime.TotalHours
    $calEvent.Utilization = ($calEvent.TotalHoursBooked / $([int]$reportDays * 8)).tostring("P")
    $results += New-Object -TypeName PSObject -Property $calEvent
}

# Output
$results


```
</br>

**Sample Output**
```
Sample Results for 3 rooms over 10 days.


Room              : #Charlotte Room 1
ReportPeriod      : 10
TotalAppointments : 6
TotalHoursBooked  : 5
Utilization       : 6.25%
Events            : {@{Subject=CLT 3; Duration=02:30:00; Room=#Charlotte Room >1; Date=9/10/2019}, @{Subject=CLT 2; Duration=00:30:00;
                    Room=#Charlotte Room 1; Date=9/11/2019}, @{Subject=CLT 1; >Duration=00:30:00; Room=#Charlotte Room 1; Date=9/10/2019}>,
                    @{Subject=CLT 1; Duration=00:30:00; Room=#Charlotte Room >1; Date=9/12/2019}, @{Subject=CLT 1; Duration=00:30:00;
                    Room=#Charlotte Room 1; Date=9/14/2019}, @{Subject=CLT 1; >Duration=00:30:00; Room=#Charlotte Room 1; Date=9/16/2019}>}

Room              : #Dublin Room 1
ReportPeriod      : 10
TotalAppointments : 4
TotalHoursBooked  : 2
Utilization       : 2.50%
Events            : {@{Subject=brandon ; Duration=00:30:00; Room=#Dublin Room >1; Date=9/11/2019}, @{Subject=brandon ; Duration=00:30:00;
                    Room=#Dublin Room 1; Date=9/12/2019}, @{Subject=brandon ; >Duration=00:30:00; Room=#Dublin Room 1; Date=9/13/2019},
                    @{Subject=brandon ; Duration=00:30:00; Room=#Dublin Room >1; Date=9/14/2019}}

Room              : #Seattle Room 1
ReportPeriod      : 10
TotalAppointments : 10
TotalHoursBooked  : 80
Utilization       : 100.00%
Events            : {@{Subject=brandon ; Duration=08:00:00; Room=#Seattle >Room 1; Date=9/14/2019}, @{Subject=brandon ; Duration=08:00:00;
                    Room=#Seattle Room 1; Date=9/14/2019}, @{Subject=brandon ;> Duration=08:00:00; Room=#Seattle Room 1; Date=9/14/2019},
                    @{Subject=brandon ; Duration=08:00:00; Room=#Seattle Room >1; Date=9/14/2019}, @{Subject=brandon ; Duration=08:00:00;
                    Room=#Seattle Room 1; Date=9/14/2019}, @{Subject=brandon ;> Duration=08:00:00; Room=#Seattle Room 1; Date=9/14/2019},
                    @{Subject=brandon ; Duration=08:00:00; Room=#Seattle Room >1; Date=9/14/2019}, @{Subject=brandon ; Duration=08:00:00;
                    Room=#Seattle Room 1; Date=9/14/2019}, @{Subject=brandon ;> Duration=08:00:00; Room=#Seattle Room 1; Date=9/14/2019},
                    @{Subject=brandon ; Duration=08:00:00; Room=#Seattle Room >1; Date=9/14/2019}}
```