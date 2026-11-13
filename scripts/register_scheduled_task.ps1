param(
    [Parameter(Mandatory = $true)]
    [string]$TaskName,
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
    )]
    [string]$DayOfWeek,
    [Parameter(Mandatory = $true)]
    [string]$AtTime,
    [Parameter(Mandatory = $true)]
    [string]$Password
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Cmd = Join-Path $PSScriptRoot "refresh_all.cmd"

if (-not (Test-Path $Cmd)) {
    Write-Error "Missing $Cmd"
}

$Action = New-ScheduledTaskAction -Execute $Cmd -WorkingDirectory $Root
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $AtTime
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

$User = "$env:USERDOMAIN\$env:USERNAME"

$Existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($Existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -User $User `
    -Password $Password `
    -RunLevel Limited | Out-Null

Write-Host "Registered: $TaskName ($DayOfWeek $AtTime)"
