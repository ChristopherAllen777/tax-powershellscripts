# Define the input directory containing the files
$inputDirectory = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\1099-MISC" # Change this to your directory path
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

# Define column letters and their numeric positions (1 = A, 2 = B, etc.)
$moneyColumnsMap = @{
    # "A" = 1  # Column A
    # "B" = 2  # Column B
    # "C" = 3  # Column C
    # "D" = 4  # Column D
    # "E" = 5  # Column E
    # "F" = 6  # Column F
    # "G" = 7   # Column G
    # "H" = 8   # Column H
    # "I" = 9   # Column I
    # "J" = 10  # Column J
    # "K" = 11  # Column K
    # "L" = 12  # Column L
    # "M" = 13  # Column M
    # "N" = 14  # Column N
    # "O" = 15  # Column O
    # "P" = 16  # Column P
    # "Q" = 17  # Column Q
    # "R" = 18  # Column R
    # "S" = 19  # Column S
    "T" = 20  # Column T
    "U" = 21  # Column U
    "V" = 22  # Column V
    "W" = 23  # Column W
    "X" = 24  # Column X
    "Y" = 25  # Column Y
    "Z" = 26  # Column Z
    "AA" = 27  # Column AA
    "AB" = 28  # Column AB
    "AC" = 29  # Column AC
    "AD" = 30  # Column AD
    "AE" = 31  # Column AE
    "AF" = 32  # Column AF
    "AG" = 33  # Column AG
    # "AH" = 34  # Column AH
    "AI" = 35  # Column AI
    # "AJ" = 36  # Column AJ
    # "AK" = 37  # Column AK
    # "AL" = 38  # Column AL
    # "AM" = 39  # Column AM
    # "AN" = 40  # Column AN
    # "AO" = 41  # Column AO
    # "AP" = 42  # Column AP
    # "AQ" = 43  # Column AQ
    # "AR" = 44  # Column AR
    # "AS" = 45  # Column AS
    # "AT" = 46  # Column AT
    # "AU" = 47  # Column AU
    # "AV" = 48  # Column AV
    # "AW" = 49  # Column AW
    # "AX" = 50  # Column AX
    # "AY" = 51  # Column AY
    # "AZ" = 52  # Column AZ
}

# Get all files in the directory with .csv or .xlsx extensions
$files = Get-ChildItem -Path $inputDirectory | Where-Object { $_.Extension -match '\.csv|\.xlsx' }

# Process each file
foreach ($file in $files) {
    Write-Host "Processing file: $($file.FullName)" -ForegroundColor Gray

    $extensionTag = if ($file.Extension -eq ".csv") { "-csv" } elseif ($file.Extension -eq ".xlsx") { "-xlsx" } else { "" }
    $outputFileName = "$outputDirectory\$($file.BaseName)$extensionTag-Output-(tab-delimited).txt"

    switch ($file.Extension) {
        ".csv" {
            # Handle CSV files
            $csvData = Import-Csv -Path $file.FullName
            $outputLines = @()

            # Add the headers as the first line in the output file
            $headers = ($csvData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
            $outputLines += $headers
            
            # Process each row of the CSV data
            foreach ($row in $csvData) {
                foreach ($colLetter in $moneyColumnsMap.Keys) {
                    $columnIndex = $moneyColumnsMap[$colLetter] - 1
                    $value = $row.PSObject.Properties.Value[$columnIndex]

                    if ($value -match '^(?:\d+|\.\d+|\d+\.\d{1})$') {
                        $row.PSObject.Properties.Value[$columnIndex] = [string]::Format("{0:F2}", [decimal]$value)
                    }
                }

                # Combine row values with tabs
                $outputLine = ($row.PSObject.Properties.Value -join "`t")
                $outputLines += $outputLine
            }

            # Write the output to the text file
            $outputLines | Set-Content -Path $outputFileName
        }

        ".xlsx" {
            # Handle Excel files
            $excelData = Import-Excel -Path $file.FullName
            $outputLines = @()

            # Add the headers as the first line
            $headers = ($excelData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
            $outputLines += $headers

            foreach ($row in $excelData) {
                foreach ($colLetter in $moneyColumnsMap.Keys) {
                    $columnIndex = $moneyColumnsMap[$colLetter]
                    $propertyName = $row.PSObject.Properties.Name[$columnIndex - 1]
                    $value = $row.$propertyName

                    if ($value -match '^(?:\d+|\.\d+|\d+\.\d{1})$') {
                        $row.$propertyName = [string]::Format("{0:F2}", [decimal]$value)
                    }
                }

                # Combine row values with tabs
                $outputLine = ($row.PSObject.Properties.Value -join "`t")
                $outputLines += $outputLine
            }

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

