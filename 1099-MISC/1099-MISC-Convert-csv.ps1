# Define the input CSV file and output text file
$inputCsv = "1099_MISC_Test_file_1.csv"  # Change this to your input CSV file path
$outputText = "1099-MISC-OutputFile(noformula)-TxtTabDeliniated.txt"  # Change this to your desired output text file path

# Read the data from the CSV file
$csvData = Import-Csv -Path $inputCsv

# Create a list to hold the output lines
$outputLines = @()

# Add the headers as the first line in the output file
$headers = ($csvData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
$outputLines += $headers

# Process each row of the CSV data
foreach ($row in $csvData) {
    # Join the values of each row using a tab separator
    $outputLine = ($row.PSObject.Properties.Value -join "`t")
    # Add the line to the output array
    $outputLines += $outputLine
}

# Write the output to the text file
$outputLines | Set-Content -Path $outputText

Write-Host "CSV data has been successfully transformed to a tab-delimited format with headers."
