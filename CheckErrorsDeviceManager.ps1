#PowerShell Script to detect Errors in device manager

$devices = Get-PnpDevice -PresentOnly -Status Error, Unknown #,DEGRADED

if ($devices.Count -eq 0) 
{
Write-Host -ForegroundColor Green "Ignore the above error."
Write-Host -ForegroundColor Green "All devices are working properly. No errors."
}

else 
{
Write-Host -ForegroundColor Red "Errors Present"
Write-Host -ForegroundColor Red "The following devices have Errors"
$devices
return $false
}