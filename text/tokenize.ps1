<#
.SYNOPSIS
This script reads an input text file, tokenizes its content, and creates smaller chunk files in a "chunks" subfolder.

.DESCRIPTION
Each of the generated chunk files contains a maximum number of tokens, which can be configured via the -MaxTokens parameter.

.PARAMETER FilePath
The path to the input text file.

.PARAMETER MaxTokens
The maximum number of tokens each chunk file should contain. Default is 1000.

.EXAMPLE
.\tokenize.ps1 -FilePath "C:\path\to\your\file.txt" -MaxTokens 500
#>

# Define parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,

    [int]$MaxTokens = 1000,

    [switch]$Clipboard
)

# Ensure the input file exists
if (-not (Test-Path $FilePath)) {
    Write-Error "File $FilePath does not exist."
    exit
}

# Create 'chunks' directory if it doesn't exist and Clipboard is not selected
$chunksDir = Join-Path (Get-Item $FilePath).DirectoryName 'chunks'
if (-not (Test-Path $chunksDir) -and -not $Clipboard) {
    New-Item -ItemType Directory -Path $chunksDir | Out-Null
}

# Read the file and split into lines
try {
    $lines = Get-Content -Path $FilePath -Encoding UTF8
} catch {
    Write-Error "Error reading file ${FilePath}: $_"
    exit
}

# Process lines and count tokens while preserving structure
$chunk = @()
$tokenCount = 0
$chunkNumber = 0

foreach ($line in $lines) {
    $lineTokens = $line -split '\s+'
    $tokenCount += $lineTokens.Length

    if ($tokenCount -le $MaxTokens) {
        $chunk += $line
    } else {
        if ($Clipboard) {
            $chunk -join "`r`n" | Set-Clipboard
            Write-Host "Chunk $chunkNumber is now on the clipboard. Press enter to continue with the next chunk..."
            [void](Read-Host)
            $chunkNumber++
        } else {
            $chunkFilePath = Join-Path $chunksDir ("chunk_" + $chunkNumber + ".txt")
            $chunkNumber++

            try {
                $chunk | Out-File $chunkFilePath -Encoding UTF8
            } catch {
                Write-Error "Error writing to ${chunkFilePath}: $_"
                exit
            }
        }

        # Reset chunk and token count for next iteration
        $chunk = @($line)
        $tokenCount = $lineTokens.Length
    }
}

# Handle any remaining content
if ($chunk.Count -gt 0) {
    if ($Clipboard) {
        $chunk -join "`r`n" | Set-Clipboard
        Write-Host "Chunk $chunkNumber is now on the clipboard. All chunks have been processed."
    } else {
        $chunkFilePath = Join-Path $chunksDir ("chunk_" + $chunkNumber + ".txt")
        try {
            $chunk | Out-File $chunkFilePath -Encoding UTF8
        } catch {
            Write-Error "Error writing to ${chunkFilePath}: $_"
            exit
        }
    }
}

if (-not $Clipboard) {
    Write-Output "File splitting completed!"
}
