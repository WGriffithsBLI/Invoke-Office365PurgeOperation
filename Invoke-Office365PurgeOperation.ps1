# Script to remove all traces of Office 365 and the Click to Run client.
# Periodically, the Office 365 C2R gets stuck at 0-2% complete, this script will resolve this issue.
# Script by Will Griffiths - wgriffiths@quartech.com
Write-Host "Welcome! This script will purge all Office 365 Components. It is normal that this script will produce errors as some components may be already uninstalled."
Read-Host -Prompt "This script must be run with Administrative Rights. Press Enter to begin operation or Ctrl+C to cancel."

Write-Host "Step 1: Removing Office 365 Binaries"
Remove-Item "C:\Program Files\Microsoft Office 16" -Recurse

Write-Host "Step 2: Removing Scheduled Tasks"
C:\Windows\System32\schtasks.exe /delete /tn "\Microsoft\Office\Office Automatic Updates"
C:\Windows\System32\schtasks.exe /delete /tn "\Microsoft\Office\Office Subscription Maintenance"
C:\Windows\System32\schtasks.exe /delete /tn "\Microsoft\Office\Office ClickToRun Service Monitor"
C:\Windows\System32\schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentLogOn2016"
C:\Windows\System32\schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentFallBack2016"

Write-Host "Step 3: Ending Click-to-Run Tasks"
Stop-Process -Name "OfficeClickToRun.exe" -Force
Stop-Process -Name "OfficeC2RClient.exe" -Force
Stop-Process -Name "AppVshNotify.exe" -Force
Stop-Process -Name "msiexec.exe" -Force

Write-Host "Step 4: De-registering Office 365 Services"
C:\Windows\System32\sc.exe delete "ClickToRunSvc"

Write-Host "Step 5: Deleting Office files"
Remove-Item "C:\Program Files\Microsoft Office" -Recurse
Remove-Item "C:\Program Files (x86)\Microsoft Office" -Recurse
Remove-Item "C:\Program Files\Common Files\microsoft shared\ClickToRun" -Recurse
Remove-Item "C:\ProgramData\Microsoft\ClickToRun" -Recurse
Remove-Item "C:\ProgramData\Microsoft\Office\ClickToRunPackageLocker"

Write-Host "Step 6: Deleting Office 365 Registry Entries"
Remove-Item "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun" -Recurse
Remove-Item "HKLM:\SOFTWARE\Microsoft\AppVISV" -Recurse
Remove-Item "HKCU:\Software\Microsoft\Office" -Recurse

Write-Host "Office 365 and Click-to-Run Purge Complete."