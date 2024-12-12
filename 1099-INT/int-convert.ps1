# Define the input directory containing the files
$inputDirectory = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\1099-INT" # Change this to your directory path
$outputDirectory = "$inputDirectory\Output" # Directory for output files
$archiveDirectory = "$inputDirectory\Archive" # Directory for archived files

# Ensure the output and archive directories exist
if (-not (Test-Path -Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory
}
if (-not (Test-Path -Path $archiveDirectory)) {
    New-Item -ItemType Directory -Path $archiveDirectory
}

# Import the Excel module if it's not already imported
if (-not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module ImportExcel
}

# Get all files in the directory with .csv or .xlsx extensions
$files = Get-ChildItem -Path $inputDirectory -Recurse | Where-Object { $_.Extension -match '\.csv|\.xlsx' }

# Define columns for money formatting
$moneyColumns = @(
    "Recipient TIN"
    "TIN Type",
    "Recipient Acct",
    "Group",
    "Box 1 Rents",
    "Check Date"
    "Transaction Description",
    "Print Indicator"
)

# Process each file
foreach ($file in $files) {
    Write-Host "Processing file: $($file.FullName)" -ForegroundColor Gray

    $extensionTag = if ($file.Extension -eq ".csv") { "-csv" } elseif ($file.Extension -eq ".xlsx") { "-xlsx" } else { "" }
    $outputFileName = "$outputDirectory\$($file.BaseName)$extensionTag-Output.txt"

    switch ($file.Extension) {
        ".csv" {
            # Handle CSV files
            $csvData = Import-Csv -Path $file.FullName

            # Create a list to hold the output lines
            $outputLines = @()

            # Add the headers as the first line in the output file
            $headers = ($csvData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
            $outputLines += $headers
            
            # Process each row of the CSV data
            foreach ($row in $csvData) {
                foreach ($moneyColumn in $moneyColumns) {
                    if ($row.PSObject.Properties[$moneyColumn] -and $row.$moneyColumn -match '^(?:\d+|\.\d+|\d+\.\d{1})$') {
                        $row.$moneyColumn = [string]::Format("{0:F2}", [decimal]$row.$moneyColumn)
                    }
                }

                # Join the values of each row using a tab separator
                $outputLine = ($row.PSObject.Properties.Value -join "`t")
                $outputLines += $outputLine
            }

            # Write the output to the text file
            $outputLines | Set-Content -Path $outputFileName
        }

        ".xlsx" {
            # Handle XLSX files
            $excelData = Import-Excel -Path $file.FullName

            # Create a list to hold the output lines
            $outputLines = @()

            # Add the headers as the first line in the output file
            $headers = ($excelData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
            $outputLines += $headers

            # Process each row of the Excel data
            foreach ($row in $excelData) {
                foreach ($moneyColumn in $moneyColumns) {
                    if ($row.PSObject.Properties[$moneyColumn] -and $row.$moneyColumn -match '^(?:\d+|\.\d+|\d+\.\d{1})$') {
                        $row.$moneyColumn = [string]::Format("{0:F2}", [decimal]$row.$moneyColumn)
                    }
                }
    
                # Join the values of each row using a tab separator
                $outputLine = ($row.PSObject.Properties.Value -join "`t")
                $outputLines += $outputLine
            }

            # Write the output to the text file
            $outputLines | Set-Content -Path $outputFileName
        }

        default {
            Write-Host "-- Unsupported file format: $($file.Extension)" -ForegroundColor Red
        }
    }

    # Move processed file to archive directory
    Move-Item -Path $file.FullName -Destination "$archiveDirectory\$($file.Name)"

    Write-Host " -- File processed and archived: $outputFileName" -ForegroundColor Magenta
}

Write-Host "All files have been processed, respective output files saved to $outputDirectory, and files archived to $archiveDirectory" -ForegroundColor Green