# Daily 11:00 AM refresh (daily + weekly workbooks to OneDrive).
$ErrorActionPreference = "Stop"
$Helper = Join-Path $PSScriptRoot "register_scheduled_task.ps1"

Write-Host ""
Write-Host "RAIL dashboard — daily automatic refresh"
Write-Host "  Every day at 11:00 AM (local)"
Write-Host ""
Write-Host "Enter your Windows sign-in password (Task Scheduler on this PC only)."
$Secure = Read-Host "Windows password" -AsSecureString
$Ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
try {
    $Plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($Ptr)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Ptr)
}

$OldNames = @(
    "RAIL dashboard refresh",
    "RAIL dashboard refresh Thursday"
)
foreach ($Name in $OldNames) {
    $Existing = Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue
    if ($Existing) {
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
        Write-Host "Removed old task: $Name"
    }
}

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Cmd = Join-Path $PSScriptRoot "refresh_all_scheduled.cmd"
$Action = New-ScheduledTaskAction -Execute $Cmd -WorkingDirectory $Root
$Trigger = New-ScheduledTaskTrigger -Daily -At "11:00AM"
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

$TaskName = "RAIL dashboard refresh daily"
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
    -Password $Plain `
    -RunLevel Limited | Out-Null

Write-Host ""
Write-Host "Registered: $TaskName"
Write-Host "Close Excel before 11:00. Use VPN for Jenkins. Logs: $Root\reports\"
Write-Host "Test: schtasks /Run /TN `"$TaskName`""
