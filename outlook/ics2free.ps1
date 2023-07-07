param(
    [Parameter(Mandatory=$true)]
    [string]$inputIcsFile,

    [Parameter(Mandatory=$false)]
    [string]$outputDirectory = (Get-Location).Path,

    [switch]$detailedLog
)

# Check if the input ICS file exists
if(!(Test-Path $inputIcsFile)){
    Write-Error "The input file $inputIcsFile does not exist."
    return
}

# Check if the output directory exists
if(!(Test-Path $outputDirectory)){
    Write-Error "The output directory $outputDirectory does not exist."
    return
}

try{
    # Construct the new file name as "filename.free.ics"
    $outputIcsFile = Join-Path $outputDirectory ((Split-Path $inputIcsFile -Leaf).Replace(".ics", ".free.ics"))
    $inEvent = $false
    $addTransparent = $true
    $addBusyStatus = $true
    $fileContent = Get-Content $inputIcsFile
    $newContent = @()

    # Iterate through each line of the file content
    foreach($line in $fileContent){
        if($detailedLog){Write-Host "Processing line: $line"}

        # Check if we are inside a VEVENT block
        if($line -match "BEGIN:VEVENT"){
            $inEvent = $true
        }
        elseif($line -match "END:VEVENT"){
            $inEvent = $false
            # Append the additional lines before the end of the VEVENT if they are not already there
            if($addTransparent){$newContent += "TRANSP:TRANSPARENT"}
            if($addBusyStatus){$newContent += "X-MICROSOFT-CDO-BUSYSTATUS:FREE"}
            # Reset the flags for the next VEVENT
            $addTransparent = $true
            $addBusyStatus = $true
        }

        # Check if the TRANSP or BUSYSTATUS properties already exist
        if($line -match "TRANSP:TRANSPARENT"){
            $addTransparent = $false
        }
        elseif($line -match "X-MICROSOFT-CDO-BUSYSTATUS:FREE"){
            $addBusyStatus = $false
        }

        # Append the current line to the new content
        $newContent += $line
    }

    # Write the new content to the output file
    $newContent | Out-File -FilePath $outputIcsFile -Encoding utf8

    if($detailedLog){Write-Host "File written to $outputIcsFile"}
}
catch{
    Write-Error "An error occurred: $_"
}
