# PowerShell script to register daily 9:00 AM Task Scheduler task
$TaskName = "RestaurantRatingsDailyMailer"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = (Get-Command python.exe -ErrorAction SilentlyContinue).Source

if (-not $PythonExe) {
    Write-Host "Python not found in PATH. Defaulting to 'python'" -ForegroundColor Yellow
    $PythonExe = "python.exe"
}

$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument "`"$ScriptDir\main.py`" --run-now" -WorkingDirectory $ScriptDir
$Trigger = New-ScheduledTaskTrigger -Daily -At 9:00AM
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Write-Host "Registering Scheduled Task '$TaskName' to run daily at 9:00 AM..." -ForegroundColor Cyan

try {
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Force | Out-Null
    Write-Host "[SUCCESS] Task '$TaskName' registered successfully!" -ForegroundColor Green
    Write-Host "The automation will run every morning at 09:00 AM." -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to register task: $_" -ForegroundColor Red
    Write-Host "Please make sure to run PowerShell as Administrator." -ForegroundColor Yellow
}
