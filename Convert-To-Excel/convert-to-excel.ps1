### Convert a tab-delimited .txt file to an Excel spreadsheet

# Input and output file paths
$InputFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\file.txt"  # Tab-delimited TXT file
$OutputFile = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-To-Excel\converted-file.xlsx"  # Desired Excel file path

# Check if Excel COM object is available
Write-Host "Checking for Excel installation..."
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel is installed. Proceeding with conversion."
} catch {
    Write-Error "Excel is not installed on this system. Please install Microsoft Excel to proceed."
    exit
}

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

# Create an Excel application instance
Write-Host "Creating Excel file..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Add a workbook and worksheet
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)

    # Write headers to the worksheet
    $colIndex = 1
    foreach ($header in $Headers.Split(" ")) {
        $worksheet.Cells.Item(1, $colIndex).Value2 = $header
        $colIndex++
    }

    # Write data to the worksheet
    $rowIndex = 2
    foreach ($row in $Data) {
        $colIndex = 1
        foreach ($property in $row.PSObject.Properties) {
            $worksheet.Cells.Item($rowIndex, $colIndex).Value2 = $property.Value
            $colIndex++
        }
        $rowIndex++
    }

    # Save the workbook
    $workbook.SaveAs($OutputFile, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
    Write-Host "Excel file saved at $OutputFile"
} catch {
    Write-Error "An error occurred: $_"
} finally {
    # Close the workbook and quit Excel
    $workbook.Close($false)
    $excel.Quit()

    # Release COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Cleanup temporary file
Remove-Item -Path $TempFile -Force

Write-Host "Conversion complete. Excel file saved at $OutputFile"
