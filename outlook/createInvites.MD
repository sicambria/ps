# PowerShell Script to Create Outlook Events from CSV

This script reads from a CSV file and creates corresponding Outlook events based on the entries.

## Script Overview

The main operations of this script are:
1. Ensure the availability of the Outlook COM object.
2. Define a function to parse various date formats.
3. Read the CSV file and process each entry.
4. Create an Outlook event for each CSV entry.

## Prerequisites

- Microsoft Office with Outlook must be installed.

## Parameters

- **csvPath**: Path to the input CSV file. This parameter is mandatory.
  
- **dateFormat**: Defines the date format used in the CSV. The script currently supports Hungarian (HU), UK (UK), and US (US) date formats. If this parameter is not provided, it defaults to the Hungarian (HU) format.

## CSV Input Format

The CSV file should be formatted as follows:

```scss
meeting name;meeting body text;required invitees;optional invitees;date;time of start;time of ending
```

**Example CSV Entry:**

```scss
Team Meeting;Discuss Q3 Goals;john.doe@example.com;alice.smith@example.com;2023.08.03;11:30:00;12:30:00
```

**Date Format Examples:**

- Hungarian (HU): `2023.08.03`
- UK (UK): `03/08/2023`
- US (US): `08/03/2023`

**Running the Script:**

To execute the script, navigate to the directory containing the script file and run:

```powershell
.\ScriptName.ps1 -csvPath "path_to_csv.csv" -dateFormat "HU"
```

Replace `ScriptName.ps1` with the actual script filename and `path_to_csv.csv` with the path to your CSV file. The `-dateFormat` parameter is optional; if not provided, it will default to the Hungarian format.

**Error Handling:**

The script contains error handling to catch issues that might arise while processing each CSV entry. Any errors encountered will be printed to the console with details.
