# Input and output file paths
$InputFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\file.txt"  # Tab-delimited TXT file
$OutputFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\converted-file.csv"  # Desired CSV file path

# Read the input file as plain text to process headers
Write-Host "Processing the input file to consolidate headers..."
$Lines = Get-Content -Path $InputFile

# Combine header lines and clean them up
$Headers = $Lines[0..2] -join " "  # Assuming headers span the first 3 lines
$Headers = $Headers -replace "``t+", " " -replace ",,", "," -replace "  ", " "

# Get the remaining data lines
$DataLines = $Lines[3..($Lines.Count - 1)]

# Create a temporary file with the corrected content
$TempFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\temp-file.txt"
$Headers | Out-File -FilePath $TempFile
$DataLines | Out-File -FilePath $TempFile -Append

# Import the corrected tab-delimited file
Write-Host "Importing data from the processed file..."
$Data = Import-Csv -Path $TempFile -Delimiter "`t"

# Export the data to a CSV file
Write-Host "Exporting data to a CSV file..."
$Data | Export-Csv -Path $OutputFile -NoTypeInformation -Force

# Cleanup temporary file
Remove-Item -Path $TempFile -Force

Write-Host "Conversion complete. CSV file saved at $OutputFile"


