# Temporary daily report email

This optional feature is isolated in this folder and does not change the core mailer.
It sends the existing `reports/Monthly_Management_Summary_Report.pdf` every day at
9:00 AM Asia/Kolkata, using the existing email subject, message, sender, and recipient
settings. The saved PDF is reused as-is; it is not regenerated each day. The outgoing
attachment is named `Invoice & Debit Note Generation Report.pdf`.

## Install

From this folder, run `.\install_schedule.ps1` in PowerShell. The task runs as the
current Windows user, so that user must be signed in at the scheduled time. If the
computer is asleep or unavailable, Windows starts the missed run when it becomes
available.

The existing project `.env` must contain the SMTP settings and recipient list.
To test the sender without waiting for the schedule, run:

```powershell
python .\send_scheduled_report.py
```

## Remove

Run `.\remove_schedule.ps1` before deleting this folder. This removes the Windows
Scheduled Task as well as the optional sender files; the core mailer and its behavior
remain unchanged.
