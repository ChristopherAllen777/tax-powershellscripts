# Load the Excel file and modify column N data
$excelPath = "C:\Users\5CA\Documents\Scripts\FormatCSVScript\Convert-zip\Copy of Final-TR25OSF-1099-INT--TEST-Brian with Headers-decimal.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$worksheet = $workbook.Sheets.Item(1)

# Column T = 20, Column X = 24
$columnsToModify = @(20, 24)

# Loop through rows and modify columns T and X data
$rowCount = $worksheet.UsedRange.Rows.Count
foreach ($col in $columnsToModify) {
    for ($row = 2; $row -le $rowCount; $row++) {
        # Fetch the cell value
        $cellValue = $worksheet.Cells.Item($row, $col).Value2
        if ($cellValue -ne $null) {
            # Convert to a double and format with two decimal places
            try {
                $decimalValue = [double]$cellValue
                # Apply the formatting to ensure two decimal places, even if one exists
                $worksheet.Cells.Item($row, $col).Value2 = $decimalValue.ToString("F2")
            } catch {
                # If conversion fails, skip the value
                Write-Host "Skipping non-numeric value in Row $row, Column $col"
            }
        }
    }
}

# Save and close the workbook
$workbook.Save()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)



