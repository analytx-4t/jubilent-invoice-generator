$ErrorActionPreference = "Stop"

$taskName = "Jubilant Temporary Invoice Debit Note Daily Mail"
$python = (Get-Command python.exe -ErrorAction Stop).Source
$scriptPath = Join-Path $PSScriptRoot "send_scheduled_report.py"
$projectRoot = Split-Path -Parent $PSScriptRoot
$workingDirectory = [System.Security.SecurityElement]::Escape($projectRoot)
$escapedPython = [System.Security.SecurityElement]::Escape($python)
$escapedScript = [System.Security.SecurityElement]::Escape(('"' + $scriptPath + '"'))
$userSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

$indiaNow = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(
    [DateTimeOffset]::Now,
    "India Standard Time"
)
$firstRun = $indiaNow.Date.AddHours(9)
if ($firstRun -le $indiaNow.DateTime) {
    $firstRun = $firstRun.AddDays(1)
}
$startBoundary = $firstRun.ToString("yyyy-MM-dd'T'HH:mm:ss") + "+05:30"

$taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$startBoundary</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$userSid</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <StartWhenAvailable>true</StartWhenAvailable>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
    <Enabled>true</Enabled>
    <AllowStartOnDemand>true</AllowStartOnDemand>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$escapedPython</Command>
      <Arguments>$escapedScript</Arguments>
      <WorkingDirectory>$workingDirectory</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

Register-ScheduledTask -TaskName $taskName -Xml $taskXml -Force | Out-Null
Write-Output "Daily email scheduled for 9:00 AM Asia/Kolkata."
Write-Output "The task runs as the current Windows user and requires that user to be signed in."
Write-Output "To remove it, run .\remove_schedule.ps1 before deleting this folder."
