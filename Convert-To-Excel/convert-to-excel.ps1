# Check if the ImportExcel module is installed; if not, install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Input and output file paths
$InputFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\file.txt"  # Tab-delimited TXT file
$OutputFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\converted-file.xlsx"  # Desired Excel file path

# Import the tab-delimited TXT file
Write-Host "Importing data from the tab-delimited file..."
$Data = Import-Csv -Path $InputFile -Delimiter "`t"

# Export the data to an Excel file
Write-Host "Exporting data to an Excel file..."
$Data | Export-Excel -Path $OutputFile -AutoSize

Write-Host "Conversion complete. Excel file saved at $OutputFile"
