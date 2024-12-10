# Define the input XLSX file and output text file
$inputXlsx = "1099-INT TEST FILE.xlsx"  # Change this to your input XLSX file path
$outputText = "1099-INT-OutputFile-TxtTabDeliniated.txt"  # Change this to your desired output text file path

# Import the Excel module if it's not already imported
if (-not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module ImportExcel
}

# Read the data from the XLSX file
$excelData = Import-Excel -Path $inputXlsx

# Create a list to hold the output lines
$outputLines = @()

# Process each row of the Excel data
foreach ($row in $excelData) {
    # Join the values of each row using a tab separator
    $outputLine = ($row.PSObject.Properties.Value -join "`t")
    # Add the line to the output array
    $outputLines += $outputLine
}

# Write the output to the text file
$outputLines | Set-Content -Path $outputText

Write-Host "XLSX data has been successfully transformed to a tab-delimited format."

