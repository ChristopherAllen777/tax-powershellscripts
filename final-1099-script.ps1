# Define the base directory containing all 1099 folders
$baseDirectory = "C:\Users\5CA\Documents\Scripts\FormatCSVScript"

# List of subdirectories
$subDirectories = @(
    "1099-INT",
    "1099-MISC",
    "1099-NEC",
    "1099-R"
)

# Loop through each subdirectory and execute its script
foreach ($subDir in $subDirectories) {
    $scriptPath = Join-Path -Path $baseDirectory -ChildPath $subDir

    # Define the specific script to run in each subdirectory
    $scriptFile = Join-Path -Path $scriptPath -ChildPath "convert.ps1"  # Change "ProcessScript.ps1" to your script name

    # Check if the script exists
    if (Test-Path -Path $scriptFile) {
        Write-Host "Running script in ${subDir}: $scriptFile" -ForegroundColor Green

        # Execute the script
        & $scriptFile

        Write-Host "Completed script for $subDir" -ForegroundColor Cyan
    } else {
        Write-Host "No script found in $subDir. Skipping..." -ForegroundColor Yellow
    }
}

Write-Host "All scripts executed." -ForegroundColor Green
