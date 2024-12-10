# Define the input XLSX file and output text file
$inputXlsx = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\1099-MISC\1099 MISC Test file 1.xlsx"  # Change this to your input XLSX file path
$outputText = "1099-MISC-OutputFile(standard)-TxtTabDeliniated.txt"  # Change this to your desired output text file path

# Import the Excel module if it's not already imported
if (-not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module ImportExcel
}

# Read the data from the XLSX file
$excelData = Import-Excel -Path $inputXlsx

# Create a list to hold the output lines
$outputLines = @()

# Add the headers as the first line in the output file
$headers = ($excelData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
$outputLines += $headers

# Process each row of the Excel data
foreach ($row in $excelData) {
    # Join the values of each row using a tab separator
    $outputLine = ($row.PSObject.Properties.Value -join "`t")
    # Add the line to the output array
    $outputLines += $outputLine
}

# Write the output to the text file
$outputLines | Set-Content -Path $outputText

Write-Host "XLSX data has been successfully transformed to a tab-delimited format with headers."

#------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------
# Define the input XLSX file and output text file
$inputXlsx = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\1099-MISC\1099 MISC Test file 1.xlsx"  # Change this to your input XLSX file path
$outputText = "1099-MISC-OutputFile(moneyvalues)-TxtTabDeliniated.txt"  # Change this to your desired output text file path

# Import the Excel module if it's not already imported
if (-not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module ImportExcel
}

# Read the data from the XLSX file
$excelData = Import-Excel -Path $inputXlsx

# Create a list to hold the output lines
$outputLines = @()

# Add the headers as the first line in the output file
$headers = ($excelData | Select-Object -First 1).PSObject.Properties.Name -join "`t"
$outputLines += $headers

# Define the list of columns that are numeric and should be formatted as money
$moneyColumns = @(
    "Onesource",
    "Box 1 Rents",
    "Box 10 Gross proceeds paid to an attorney", 
    "Box 5 Fishing boat proceeds", 
    "Box 7 Nonemployee compensation"  # Add additional columns as needed
)

# Process each row of the Excel data
foreach ($row in $excelData) {
    # Loop through each column in the money column list
    foreach ($moneyColumn in $moneyColumns) {
        # Check if the column exists in the row and the value is numeric
        if ($row.PSObject.Properties[$moneyColumn] -and $row.$moneyColumn -match '^\d+(\.\d+)?$') {
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

Write-Host "XLSX data has been successfully transformed to a tab-delimited format with headers and formatted money values."


