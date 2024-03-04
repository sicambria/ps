# Get the current directory
$CurrentDirectory = Get-Location

# Define the path for the output file
$OutputFile = Join-Path -Path $CurrentDirectory -ChildPath "output.txt"

# Create or clear the output file
if (Test-Path $OutputFile) {
    Clear-Content $OutputFile
} else {
    New-Item -Path $OutputFile -ItemType File
}

# Regex to match YouTube URLs
$YouTubeRegex = 'https?://(?:www\.)?youtube\.com/watch\?v=[^ \r\n]+|https?://youtu\.be/[^ \r\n]+'

# Get all .txt files in the current directory
$Files = Get-ChildItem -Path $CurrentDirectory -Filter *.txt

# Initialize counter
$LinkCount = 0

foreach ($File in $Files) {
    try {
        Write-Verbose "Processing file: $($File.Name)"
        # Read the content of the file
        $Content = Get-Content -Path $File.FullName -ErrorAction Stop
        
        # Find all matches
        $Matches = Select-String -InputObject $Content -Pattern $YouTubeRegex -AllMatches
        
        foreach ($Match in $Matches.Matches) {
            # Remove a trailing quotation mark if present
            $CleanedLink = $Match.Value -replace '"$', ''
            
            # Append the cleaned URL to the output file
            Add-Content -Path $OutputFile -Value $CleanedLink
            $LinkCount++
        }
    } catch {
        Write-Warning "An error occurred processing file: $($File.Name). Error: $_"
    }
}

Write-Host "YouTube links have been extracted to $OutputFile. Total links found: $LinkCount"
