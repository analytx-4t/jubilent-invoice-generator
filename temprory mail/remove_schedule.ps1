$ErrorActionPreference = "Stop"

$taskName = "Jubilant Temporary Invoice Debit Note Daily Mail"
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($null -ne $task) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Output "Removed scheduled task: $taskName"
} else {
    Write-Output "Scheduled task was not installed."
}
