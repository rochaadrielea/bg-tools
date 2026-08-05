# ============================================================
# install_service.ps1
# Registers the Quality BRM dashboard as a Windows scheduled task.
#
# What it does:
#   - Runs automatically when you log in
#   - Restarts if it crashes (up to 999 times)
#   - Runs hidden in the background
#
# How to run:
#   1. Right-click this file -> Run with PowerShell
#      (if blocked, open PowerShell and run:
#       Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
#       then run: .\install_service.ps1)
# ============================================================

$TaskName = "QualityBRMDashboard"
$ScriptFolder = $PSScriptRoot
$RunBat = Join-Path $ScriptFolder "run.bat"

if (-not (Test-Path $RunBat)) {
    Write-Host "ERROR: run.bat not found in $ScriptFolder" -ForegroundColor Red
    exit 1
}

# Remove existing task if it exists
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# Action: run the bat file
$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$RunBat`"" `
    -WorkingDirectory $ScriptFolder

# Trigger: at logon of current user
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"

# Settings: restart on failure, hidden, no time limit
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit (New-TimeSpan -Days 365) `
    -Hidden

# Register the task
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Quality BRM Streamlit dashboard - auto-start on logon"

Write-Host ""
Write-Host "SUCCESS: Task '$TaskName' registered." -ForegroundColor Green
Write-Host ""
Write-Host "The dashboard will start automatically at next logon."
Write-Host "To start it right now, run:"
Write-Host "    Start-ScheduledTask -TaskName $TaskName" -ForegroundColor Cyan
Write-Host ""
Write-Host "To check status:"
Write-Host "    Get-ScheduledTask -TaskName $TaskName" -ForegroundColor Cyan
Write-Host ""
Write-Host "To stop it:"
Write-Host "    Stop-ScheduledTask -TaskName $TaskName" -ForegroundColor Cyan
Write-Host ""
Write-Host "To remove it:"
Write-Host "    Unregister-ScheduledTask -TaskName $TaskName -Confirm:`$false" -ForegroundColor Cyan
