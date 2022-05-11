# Gets room utilization based on their calendars
# Minimum Application Permission: Calendars.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-calendarview

# Variables
$startTime = "2020-06-25"
$endTime = "2020-08-03"
$calendars = @("CharlotteRm1@contoso.com","DublinRm1@contoso.com","SeattleRm1@contoso.com")

$token = 
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