$results = @()
$calendars = @('Adams@M365x725618.onmicrosoft.com')
$startDate = ([DateTime]'2022-05-08')
$reportDays = 7

# Get config and helper
#$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent)
$root = $PSScriptRoot
$config = Get-Content "$root\clientconfiguration.json" -Raw | ConvertFrom-Json
Import-Module "$root\GraphAPIHelper.psm1" -Force

# $client = '<replace me with a guid>'
# $tenant = '<replace me with a guid or routing domain>'
# $secret = '<replace me with a secret>'

# Get-AccessToken is a custom function, code provided in this module
#$token = Get-AccessToken -Tenant $tenant -ClientID $client -ClientSecret $secret
$token = Get-AccessToken -ClientID $config.ClientId -Tenant $config.TenantId -ClientSecret $config.Secret

foreach ($calendar in $calendars) {

    # Graph wants dates and times formatted as ISO 8601/round trip
    $start = $startDate.ToString('o')                       # 2021-05-24T00:00:00.0000000
    $end = $startDate.AddDays($reportDays).ToString('o')  # 2021-05-28T00:00:00.0000000
    $uri = "https://graph.microsoft.com/v1.0/users/$calendar/calendar/calendarView?startDateTime=$start&endDateTime=$end"
  
    # Invoke-MSGraphQuery is a custom function, code provided in this module
    $appointments = Invoke-MSGraphQuery -AccessToken $token -Uri $uri

    # Create a new measurement of time to track overall booked time for the room
    $bookableTime = New-TimeSpan

    # Create an object to represent our room metrics
    $roomMetrics = [ordered]@{
        Room              = $appointments[0].location.displayName;
        ReportPeriod      = $reportDays
        TotalAppointments = $appointments.Count
        TotalHoursBooked  = 0
        Utilization       = 0
    }
  
    # Loop through all appointments
    foreach ($appointment in $appointments) {

        # Establish duration of the appointment. Consider "all day" appointments to be 8 hours
        if ($appointment.isAllDay) {
            $totalDuration = New-TimeSpan -Hours 8
        }
        else {
            # If not all day, calculate the time between the start and end times of the appointment
            $totalDuration = New-TimeSpan -Start ([DateTime]$appointment.Start.dateTime) -End ([DateTime]$appointment.End.dateTime)
        }
    
        # Add the duration of the appointment to the total time for this room
        $bookableTime += $totalDuration
    }

    # Fill in some blanks on our room metrics object
    $roomMetrics.TotalHoursBooked = $bookableTime.TotalHours
    # ToString('P') formats as a percentage
    $roomMetrics.Utilization = ($roomMetrics.TotalHoursBooked / $([int]$reportDays * 8)).ToString("P")

    # Add our room metrics to the overall report
    $results += New-Object -TypeName PSObject -Property $roomMetrics;
}

# Output the report
$results