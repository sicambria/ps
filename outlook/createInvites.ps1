# PowerShell script to create Outlook events based on CSV input
# Usage: https://github.com/sicambria/ps/blob/main/outlook/createInvites.MD

param(
    [Parameter(Mandatory=$true)]
    [string]$csvPath, # CSV input file path

    [Parameter(Mandatory=$false)]
    [ValidateSet('HU', 'UK', 'US')] # Supported date formats: Hungarian (HU), UK (UK), US (US)
    [string]$dateFormat = 'HU' # Default to HU format if not provided
)

# Ensure the Outlook COM object is available
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"

function Get-FileEncoding {
    param ([string]$filePath)
    $byte = [System.IO.File]::ReadAllBytes($filePath)[0]
    switch -Regex ($byte) {
        "EF" { return "UTF8" }
        "FF" { return "Unicode" }
        "FE" { return "BigEndianUnicode" }
        default { return "Default" }
    }
}

# Function to parse different date formats
function ParseDate ($dateString) {
    try {
        switch ($dateFormat) {
            'HU' { # Hungarian format: yyyy.MM.dd.
                return [DateTime]::ParseExact($dateString.Trim(), "yyyy.MM.dd", $null)
            }
            'UK' { # UK format: dd/MM/yyyy
                return [DateTime]::ParseExact($dateString.Trim(), "dd/MM/yyyy", $null)
            }
            'US' { # US format: MM/dd/yyyy
                return [DateTime]::ParseExact($dateString.Trim(), "MM/dd/yyyy", $null)
            }
        }
    } catch {
        throw "Unable to parse the date '$dateString' using format '$dateFormat'."
    }
}

# Read CSV file

# Guess the encoding
$fileEncoding = Get-FileEncoding -filePath $csvPath

# Import the CSV file with detected encoding
$events = Import-Csv -Path $csvPath -Delimiter ";" -Encoding $fileEncoding

Write-Host "First event's subject: $($events[0].'meeting name')"
Write-Host "First event's body: $($events[0].'meeting body text')"

$eventId = 0

# Loop through each row in the CSV and create an event
foreach ($event in $events) {
    Write-Host "Processing event: $($event.'meeting name')..."

    try {
        # Logging CSV inputs for debugging
        Write-Host "Date: $($event.date)"
        Write-Host "Time of Start: $($event.'time of start')"
        Write-Host "Time of Ending: $($event.'time of ending')"

        # Create a new Outlook application
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")

        # Create a new appointment item
        $appointment = $outlook.CreateItem(1)

        # Set the properties for the appointment from the CSV data

        $tempSubject = $event.'meeting name'
        $tempBody = $event.'meeting body text'

            #Write-Host "tempSubject: $tempSubject"
            #Write-Host "tempBody: $tempBody"

        $appointment.Subject = "$tempSubject"
        $appointment.Body = "$tempBody"


        Write-Host "Set appointment subject to: $($appointment.Subject)"
        Write-Host "Set appointment body to: $($appointment.Body)"


        $appointment.Start = (ParseDate $event.date).Add([TimeSpan]::Parse($event.'time of start'))
        $appointment.End = (ParseDate $event.date).Add([TimeSpan]::Parse($event.'time of ending'))
        $appointment.RequiredAttendees = $event.'required invitees'
        $appointment.OptionalAttendees = $event.'optional invitees'

        # Save the appointment
        $appointment.Save()

        Write-Host "Successfully created event: $($event.'meeting name')"

        $eventId++

    } catch {
        Write-Host "Error while processing $($event.'meeting name'). Error details: $_" -ForegroundColor Red
    }
}

Write-Host "All events processed!" 
