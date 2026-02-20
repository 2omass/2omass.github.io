param(
  [Parameter(Mandatory=$true)][string]$SourceDir,
  [string]$Department = "HR",
  [string]$PythonExe = "python",
  [int]$IntervalSeconds = 30,
  [string]$TaskName = "HRFileOrganizerARM"
)

$ScriptPath = Join-Path $PSScriptRoot "hr_file_organizer.py"
$OutputDir = Join-Path $SourceDir "organized"
$ReportPath = Join-Path $OutputDir "hr_organizer_report.csv"

$argument = "`"$ScriptPath`" `"$SourceDir`" --department `"$Department`" --output `"$OutputDir`" --report `"$ReportPath`" --watch --interval $IntervalSeconds"

$action = New-ScheduledTaskAction -Execute $PythonExe -Argument $argument -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "Task '$TaskName' registered. It will run continuously at startup on this Windows ARM device."
