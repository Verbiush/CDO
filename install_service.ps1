$ErrorActionPreference = "Stop"

# Get current directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# Target the AGENT, not the Client, for background service
$ExePath = Join-Path $ScriptDir "CDO_Agente.exe"
$TaskName = "CDO_Agente_AutoStart"

if (-not (Test-Path $ExePath)) {
    Write-Host "Agent Executable not found at $ExePath"
    exit 1
}

# Unregister if exists
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

# Create Scheduled Task for Current User
$Action = New-ScheduledTaskAction -Execute $ExePath -WorkingDirectory $ScriptDir
$Trigger = New-ScheduledTaskTrigger -AtLogon
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0 -Priority 7

# Register
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings

# Start it now
Start-ScheduledTask -TaskName $TaskName

Write-Host "Agent Service installed and started successfully."
