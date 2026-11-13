<#
    Registers the daily OTIF build with Windows Task Scheduler.

    Run this yourself when you are ready:
        powershell -ExecutionPolicy Bypass -File .\register_task.ps1

    Defaults to 09:45 every day, running as the current user. The task only
    fires while this account is logged on, and StartWhenAvailable catches up on
    a missed start.

    Interactive logon is the correct choice here rather than a limitation:
    PARQUET_DIR lives under the user profile in OneDrive, and the OneDrive sync
    client only runs inside a logged-on user session. Running the job while
    logged off would write files that never reach SharePoint until the user
    signs in anyway.

    To remove:  Unregister-ScheduledTask -TaskName "OTIF Daily Build"
#>
[CmdletBinding()]
param(
    [string] $TaskName = "OTIF Daily Build",
    [string] $At       = "09:45",
    [string] $Arguments = ""
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$runner = Join-Path $projectRoot "run_otif.cmd"

if (-not (Test-Path $runner)) {
    throw "Cannot find run_otif.cmd next to this script (looked in $projectRoot)."
}

# -Argument rejects an empty string, so only pass it when there is something to pass.
if ([string]::IsNullOrWhiteSpace($Arguments)) {
    $action = New-ScheduledTaskAction -Execute $runner -WorkingDirectory $projectRoot
} else {
    $action = New-ScheduledTaskAction -Execute $runner -Argument $Arguments -WorkingDirectory $projectRoot
}
$trigger = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $existing) {
    Write-Host "Task '$TaskName' already exists - updating it."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Builds the OTIF report from Business Central OData and writes Parquet + Excel." | Out-Null

Write-Host "Registered '$TaskName' to run daily at $At."
Write-Host "Runner : $runner"
Write-Host ""
Write-Host "Test it now with:"
Write-Host "    Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Then check logs\otif_<today>.log"
