#1099 - Runs script that converts all excel/csv files to tab-delimited txt files.
try {
	Write-Output "$("[{0:MM/dd/yy} {0:HH:mm:ss:fff}]" -f (Get-Date)) Start final-1099-script.ps1" | Tee-Object -FilePath $LogFileName -Append
	& final-1099-script.ps1
	Write-Output "$("[{0:MM/dd/yy} {0:HH:mm:ss:fff}]" -f (Get-Date)) End final-1099-script.ps1" | Tee-Object -FilePath $LogFileName -Append
	Send-MailMessage -To "chris.allen@citizensinc.com","cti@citizensinc.com" -From "<cti@citizensinc.com>" -Subject "1099 Files Scripts Ran" -Body "Stack Trace: $_" -SmtpServer "smtp.domain01.local"
} catch {
	Write-Output "$("[{0:MM/dd/yy} {0:HH:mm:ss:fff}]" -f (Get-Date)) *****-> Univista_PolicyChangesOutbound FAILED: $_" | Tee-Object -FilePath $ScriptDir"\"$LogFileName -Append
	Send-MailMessage -To "chris.allen@citizensinc.com","cti@citizensinc.com" -From "<cti@citizensinc.com>" -Subject "Error in 1099 Script" -Body "Stack Trace: $_" -SmtpServer "smtp.domain01.local"
}