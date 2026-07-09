# Power BI Workspace Manager

A Windows desktop app (PowerShell + WPF) to manage Power BI workspaces,
datasets, and dataflows in bulk.

## What it does

- Lists the workspaces you belong to
- Drills into each workspace to show its datasets (semantic models) and dataflows
- Multi-select, including "Select All" and per-workspace select (check the
  WORKSPACE box to select everything inside it)
- Takes ownership of selected datasets
- Enables or disables scheduled refresh on selected datasets and dataflows
- Shows a live log and a "Dry run" mode so you can preview before changing anything

## How it works (why it is reliable)

Sign-in runs once at launch, in the console window, before the app window opens.
That is the only step that uses the Power BI module. Everything after that uses
plain web requests with your sign-in token, so the app never calls the auth layer
from inside the live window - which is what causes GUI freezes.

## Requirements

- Windows with Windows PowerShell 5.1 (built in)
- The MicrosoftPowerBIMgmt module (installed automatically on first run if missing)
- Power BI permission: to take ownership or change a dataset's refresh you must be
  an Admin or Member of that workspace. The app lists only workspaces you can act
  on.

## How to run

1. Put `PowerBI-Manager.ps1` and `Run-PowerBI-Manager.cmd` in the same folder.
2. Double-click `Run-PowerBI-Manager.cmd`. A console window opens first.
3. Sign in when the Power BI sign-in window or browser appears. The app window
   opens right after.
4. Click "Load Workspaces". The progress bar and log show it reading each one.
5. Check the items you want, then click Take Ownership, Enable Refresh, or
   Disable Refresh. Leave "Dry run" checked to preview first.

To switch accounts, or if you get 401 errors after about an hour (token expiry),
close the app and run it again.

## Important API limitations (Microsoft limits, not app bugs)

- Dataflow ownership: no public REST endpoint exists. Selecting a dataflow and
  clicking Take Ownership logs a clear note and skips it. Take over a dataflow in
  the Service: open the workspace, find the dataflow, click "...", then "Take over".
- Dataflow refresh state cannot be read via the API, so the tree does not show
  on/off for dataflows. Enable and Disable still work.
- Disabling refresh turns the schedule off; it does not delete the schedule.

## Troubleshooting

- Sign-in problems appear in the console window and a popup, with the exact error.
  To sanity-check outside the app, open PowerShell and run
  `Connect-PowerBIServiceAccount`. If that succeeds, the app will too.
- A FAIL line in the log includes the exact Power BI error for that item (for
  example a 401 for auth or 403 for permissions), and other items keep going.
- The console window is part of the app; keep it open. You can minimize it.
