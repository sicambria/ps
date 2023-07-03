# Set the start and end dates
$startDate = ""
$endDate = ""

# comment the next 2 lines out with # sign to read dates from console
#$startDate = Get-Date -Year 2023 -Month 8 -Day 1
#$endDate = Get-Date -Year 2023 -Month 10 -Day 1


# Function to read date from console
function ReadDate($prompt) {
    Write-Host "$prompt Year:"
    $year = Read-Host
    Write-Host "$prompt Month:"
    $month = Read-Host
    Write-Host "$prompt Day:"
    $day = Read-Host

    return Get-Date -Year $year -Month $month -Day $day
}

# Set the start and end dates, if not already set
if (!$startDate) {
    $startDate = ReadDate "Enter the start date"
}
if (!$endDate) {
    $endDate = ReadDate "Enter the end date"
}



Write-Host "Date range is set from $startDate to $endDate"

# Create an Outlook.Application object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get the default calendar and all available calendar subfolders
$defaultCalendar = $namespace.GetDefaultFolder(9) # 9 = olFolderCalendar
$calendarFolders = $defaultCalendar.Folders

# List the default calendar and all calendar folders
Write-Host "Available calendars:"
Write-Host "1: $($defaultCalendar.Name)"
for ($i=0; $i -lt $calendarFolders.Count; $i++) {
    Write-Host "$($i+2): $($calendarFolders.Item($i+1).Name)"
}

# Prompt user to select a calendar
$selectedCalendarIndex = Read-Host -Prompt 'Enter the number of the calendar you want to select'

if ($selectedCalendarIndex -eq 1) {
    $calendar = $defaultCalendar
} else {
    $calendar = $calendarFolders.Item($selectedCalendarIndex-1)
}

Write-Host "You've selected the $($calendar.Name) calendar"

# Get the items in the date range
$items = $calendar.Items
$items.IncludeRecurrences = $true
$items.Sort("[Start]", $true)

# Get system regional settings and format the date accordingly
$region = (Get-Culture).Name
if ($region -eq "hu-HU") {
    $startDateString = $startDate.ToString("yyyy.MM.dd. HH:mm")
    $endDateString = $endDate.ToString("yyyy.MM.dd. HH:mm")
} else {
    $startDateString = $startDate.ToString("MM/dd/yyyy HH:mm")
    $endDateString = $endDate.ToString("MM/dd/yyyy HH:mm")
}

$restriction = "[Start] >= `"$startDateString`" AND [Start] < `"$endDateString`""
$items = $items.Restrict($restriction)

# List all meetings found in the calendar
Write-Host "Meetings found in the $($calendar.Name) calendar:"
foreach ($item in $items) {
    Write-Host "Meeting: $($item.Subject), Start: $($item.Start), End: $($item.End), Duration: $($item.Duration) minutes, All day event: $($item.AllDayEvent), Status: $($item.BusyStatus)"
}


Write-Host ""
Write-Host "---- DRY RUN - LIST 20+ hour long meetings -----"
Write-Host ""

# Loop through each item
foreach ($item in $items) {
    # Check if the event is an "All day event" or if the duration is more than 20 hours
    if ($item.AllDayEvent -eq $true -or $item.Duration -gt 1200) { # 1200 minutes = 20 hours
        Write-Host "Identified item: "$($item.Subject)" - which is an all-day event or longer than 20 hours"

    }
}

Write-Host ""
Read-Host "Continue to set them FREE? Press CTRL-C to abort."
Write-Host ""


# Loop through each item
foreach ($item in $items) {
    # Check if the event is an "All day event" or if the duration is more than 20 hours
    if ($item.AllDayEvent -eq $true -or $item.Duration -gt 1200) { # 1200 minutes = 20 hours
        Write-Host "Identified item: $($item.Subject) which is an all-day event or longer than 20 hours"
        $item.BusyStatus = 0  # 0 = olFree
        $item.Save()
        Write-Host "BusyStatus updated to FREE for item: $($item.Subject)"
    }
}



# Clean up
Write-Host "Initiating cleanup..."
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Cleanup completed and script execution completed!"
