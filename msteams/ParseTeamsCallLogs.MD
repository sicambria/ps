# MS Teams Call Logs Attendance Summary Script

This PowerShell script parses CSV files exported from MS Teams call logs and summarizes the overall attendance time for each participant. It handles multiple CSV files and offers additional functionalities through optional parameters.

## Features

- **Aggregates Attendance Time**: Calculates total attendance time in hours and minutes for each participant across multiple CSV files.
- **Handles Multiple Files**: Processes all CSV files within a specified directory or the current directory if none is specified.
- **Date Extraction**: Extracts dates from filenames for enhanced reporting.
- **Optional Parameters**: Additional functionalities like listing dates and outputting results in CSV format.

## Parameters

- `-csvDirectory <String>`: (Optional) The directory path where the CSV files are located. If not specified, the script uses the current working directory.
- `-listDates`: (Optional) Lists the dates of the meetings considered in the calculation.
- `-CSVToConsole`: (Optional) Outputs the data in CSV format to the console, displaying the date, name, and time for each participant. Suppresses verbose logging and additional processing.

## Usage

1. **Default Mode**: Run the script in default mode to display the total attendance time for each participant. If no directory is specified, the script processes CSV files in the current directory.

    ```powershell
    .\ParseTeamsCallLogs.ps1 -csvDirectory "Path\To\CSVFiles"
    ```

    or

    ```powershell
    .\ParseTeamsCallLogs.ps1
    ```

2. **List Dates**: List the dates of the meetings considered in the calculation:

    ```powershell
    .\ParseTeamsCallLogs.ps1 -csvDirectory "Path\To\CSVFiles" -listDates
    ```

    or

    ```powershell
    .\ParseTeamsCallLogs.ps1 -listDates
    ```

3. **CSV to Console**: Output the data in CSV format to the console:

    ```powershell
    .\ParseTeamsCallLogs.ps1 -csvDirectory "Path\To\CSVFiles" -CSVToConsole
    ```

    or

    ```powershell
    .\ParseTeamsCallLogs.ps1 -CSVToConsole
    ```

## Output

- **Default Mode**: Displays the total attendance time in hours and minutes for each participant, sorted from the least to the most time attended.
- **List Dates**: Shows the dates of the meetings considered in the calculation, sorted from earliest to latest.
- **CSV to Console**: Outputs the date, name, and time in CSV format for each participant in each file, without additional logging or processing.
