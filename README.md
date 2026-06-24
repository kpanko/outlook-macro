# Outlook aged mail cleanup

This repo contains the original Outlook VBA macro and a Microsoft Graph PowerShell replacement.

The PowerShell replacement is split into:

- `Move-AgedMail.ps1`: command wrapper for scheduled tasks and debugging.
- `Move-AgedMail.psm1`: implementation module used by the wrapper and tests.
- `Move-AgedMail.Tests.ps1`: mocked Pester tests.

The macro moves mail older than 40 calendar days:

- `GTD-incubate` -> `cleanup/inbox`
- Sent Items -> `cleanup/sent`

The PowerShell version does the same mailbox moves through Microsoft Graph, so
it does not depend on classic Outlook, COM automation, or VBA.

## One-time setup

From this repository's root directory, run:

```powershell
.\Move-AgedMail.ps1 -InstallMissingGraphModule -WhatIf
```

Sign in when prompted. The script requests the delegated `Mail.ReadWrite`
permission so it can list and move messages in your mailbox.

If the `-WhatIf` output looks right, run it for real:

```powershell
.\Move-AgedMail.ps1
```

After the first successful sign-in, later runs under the same Windows user can
usually reuse the cached Microsoft Graph token.

## Run tests

The script has mocked Pester tests that do not touch your mailbox:

```powershell
Invoke-Pester -Path .\Move-AgedMail.Tests.ps1
```

## Schedule it at logon

This is the best local setup if you only care about cleanup on mornings when
you actually log in. It runs as your Windows user, so it can reuse your
Microsoft Graph sign-in cache.

From this repository's root directory, create the task from an elevated
PowerShell session that you started with **Run as Administrator**:

```powershell
$script = Join-Path (Get-Location).Path "Move-AgedMail.ps1"
$pwsh = (Get-Command pwsh.exe).Source
$taskUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$argument = @(
    "-NoProfile"
    "-ExecutionPolicy Bypass"
    "-File `"$script`""
) -join " "
$action = New-ScheduledTaskAction -Execute $pwsh -Argument $argument
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $taskUser
$trigger.Delay = "PT2M"
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal `
    -UserId $taskUser `
    -LogonType Interactive `
    -RunLevel Limited
$description = "Move old Outlook mail to cleanup folders at logon."
Register-ScheduledTask `
    -TaskName "Move aged Outlook mail" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description $description
```

In Task Scheduler, keep the task set to **Run only when user is logged on**.
That avoids a lot of credential/token weirdness and lets Microsoft Graph prompt
interactively if it ever needs you to sign in again.

If `Register-ScheduledTask` returns **Access is denied**, make sure there is
not already a task with the same name that was created by another user. Delete
that old task from an elevated PowerShell session or use a different task name,
then run the command above again. The explicit `$principal` keeps the task
running as your normal Windows user even though registration needs elevation.
The `-User $taskUser` value on the trigger keeps the task from firing when
another user logs in.

## Schedule it daily

If you later want the task to run daily while your Windows session is active,
this creates a daily Windows Scheduled Task for 7:00 AM:

```powershell
$script = Join-Path (Get-Location).Path "Move-AgedMail.ps1"
$pwsh = (Get-Command pwsh.exe).Source
$taskUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$argument = @(
    "-NoProfile"
    "-ExecutionPolicy Bypass"
    "-File `"$script`""
) -join " "
$action = New-ScheduledTaskAction -Execute $pwsh -Argument $argument
$trigger = New-ScheduledTaskTrigger -Daily -At 7:00AM
$principal = New-ScheduledTaskPrincipal `
    -UserId $taskUser `
    -LogonType Interactive `
    -RunLevel Limited
$description = "Move old Outlook mail to cleanup folders."
Register-ScheduledTask `
    -TaskName "Move aged Outlook mail" `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Description $description
```

If you need cleanup to run while you are logged out, Task Scheduler has to use
stored credentials or a non-interactive logon type. If that later fails because
the sign-in token cannot refresh non-interactively, the better long-term
version is a small Power Automate scheduled cloud flow or an Azure
Automation/Function job using an app registration.

## Notes

- The replacement operates on Graph `message` items. That matches normal mail
  in the current macro, but it may not move unusual non-message Outlook items
  exactly the same way VBA/MAPI did.
- The age check uses `sentDateTime`, matching the macro's use of `SentOn` for mail items.
- You can change folder paths and the age threshold with parameters, for example:

```powershell
.\Move-AgedMail.ps1 `
    -AgeDays 30 `
    -IncubateFolderPath "GTD-incubate" `
    -CleanupInboxFolderPath "cleanup/inbox" `
    -CleanupSentFolderPath "cleanup/sent"
```
