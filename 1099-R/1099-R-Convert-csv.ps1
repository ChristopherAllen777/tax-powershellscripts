# Define the input CSV file and output text file
$inputCsv = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\1099-R\1099-R-Testfile.csv"  # Change this to your input CSV file path
$outputText = "1099-R-OutputFile(csv)-TxtTabDeliniated.txt"  # Change this to your desired output text file path

# Read the data from the CSV file
$csvData = Import-Csv -Path $inputCsv

# Create a list to hold the output lines
$outputLines = @()

# Add the headers as the first line in the output file
$headers = ($csvData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
$outputLines += $headers

# Define the list of columns that are numeric and should be formatted as money with 2 decimal places (e.g., 1.00)
# Add additional columns as needed
$moneyColumns = @(
    "Box 1 - Gross Distribution"
)

# Process each row of the CSV data
foreach ($row in $csvData) {
    # Loop through each column in the money column list
    foreach ($moneyColumn in $moneyColumns) {
        # Check if the column exists in the row and the value matches the numeric format
        if ($row.PSObject.Properties[$moneyColumn] -and $row.$moneyColumn -match '^(?:\d+|\.\d+|\d+\.\d{1})$') {
            # Format the value to two decimal places
            $row.$moneyColumn = [string]::Format("{0:F2}", [decimal]$row.$moneyColumn)
        }
    }

    # Join the values of each row using a tab separator
    $outputLine = ($row.PSObject.Properties.Value -join "`t")
    # Add the line to the output array
    $outputLines += $outputLine
}

# Write the output to the text file
$outputLines | Set-Content -Path $outputText

Write-Host "CSV data has been successfully transformed to a tab-delimited format with headers and formatted money values."