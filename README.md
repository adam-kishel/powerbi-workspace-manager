# Power BI Workspace Manager

A Windows desktop app (PowerShell + WPF) to manage and inspect Power BI
workspaces, datasets, and dataflows in bulk.

## What it does

- Lists the workspaces you belong to, alphabetically
- Drills into each workspace to show its datasets (semantic models) and dataflows,
  also alphabetical
- Search box filters the workspace list by name; Select All then selects only the
  workspaces currently shown
- Multi-select, including per-workspace select (check the WORKSPACE box)
- Take ownership of selected datasets
- Enable or disable scheduled refresh on selected datasets and dataflows
- View run history, produce a failure report, and view detailed info for selected
  items (all read-only), with results shown in the log AND exported to CSV
- "Dry run" mode previews changes before you make them

## How it works (why it is reliable)

Sign-in runs once at launch, in the console window, before the app window opens.
That is the only step that uses the Power BI module. Everything after that uses
plain web requests with your sign-in token, so the app never calls the auth layer
from inside the live window - which is what causes GUI freezes.

## Requirements

- Windows with Windows PowerShell 5.1 (built in)
- The MicrosoftPowerBIMgmt module (installed automatically on first run if missing)
- Power BI permission: acting on an item requires you to be an Admin or Member of
  its workspace. The app lists only workspaces you can act on.

## How to run

1. Put `PowerBI-Manager.ps1` and `Run-PowerBI-Manager.cmd` in the same folder.
2. Double-click `Run-PowerBI-Manager.cmd`. A console window opens first.
3. Sign in when the Power BI sign-in window or browser appears. The app opens after.
4. Click "Load Workspaces".
5. Check the items you want, then use the action or report buttons.

To switch accounts, or if you get 401 errors after about an hour (token expiry),
close the app and run it again.

## Actions (make changes)

- Take Ownership - datasets only (see dataflow note below).
- Enable Refresh / Disable Refresh - datasets and dataflows.
- Dry run is ON by default. Uncheck it to make real changes.

## Reports (read-only, export to CSV)

Results appear in the log and are written to a `PBI-Reports` folder created next to
the app, as timestamped CSV files you can open in Excel.

- View History - last 10 runs per selected item. CSV columns: Workspace, ItemType,
  ItemName, Status, StartTime, EndTime, DurationSec, RefreshType, RequestId,
  ErrorCode, ErrorMessage. The log shows the most recent run per item; the CSV has
  all runs.
- Failure Report - only the failed runs from the last 10 per item, with the dataset
  error code/message included. Every failure is listed in the log and the CSV.
- View Info - data sources (server/database or url/path), dataset parameters,
  refresh schedule details (days/times/timezone), and users/permissions. One row per
  fact in the CSV; the log shows a per-item count summary.
- Check Google Sheets - flags selected items that use a Google Sheets data source.
  Because Power BI reports Google Sheets as a connector (Extension) source, the check
  scans the datasource type, connector kind, and path/url for a Google Sheets signal.
  Logs each hit and writes a YES/no row per item with the matching detail.

## Filtering to failures and refreshing (Find Recent Failures + Refresh Selected)

- Find Recent Failures - reads the newest run of each currently-checked item and
  narrows the selection to only the items whose most recent refresh FAILED (everything
  else is unchecked). Select a scope first (search + Select All, or check a workspace),
  then click it. For large selections this makes one call per item, so it can take a
  while; scope with search to keep it quick.
- Refresh Selected - triggers an on-demand refresh for every checked item. Datasets use
  POST /refreshes; dataflows use POST /refreshes?processType=default. Honors Dry run,
  so preview first. To refresh one at a time, uncheck all but one; to do them all,
  leave them all checked. Note: refresh capacity limits apply (shared capacity allows
  only 8 refreshes per day per dataset, including scheduled ones).

## Important API limitations (Microsoft limits, not app bugs)

- Dataflow ownership: no public REST endpoint. You do NOT need ownership to
  enable/disable a dataflow's refresh (workspace Admin/Member is enough), so this
  rarely matters. To actually edit or reassign a dataflow, use the Service:
  workspace > dataflow > "..." > Take over.
- Dataflow refresh state and schedule details cannot be read via the API (no GET
  endpoint), so the tree does not show on/off for dataflows and Info shows schedule
  as not-readable for them. Enable/Disable still work.
- Dataflow failures have no error message in the API - the failure report shows that
  a dataflow run failed and when, but not why (that detail is only in the CSV you can
  download from the dataflow's Refresh History page in the Service).
- A linked dataflow (one that refreshes automatically after an upstream flow) has no
  schedule of its own; trying to Enable it returns a weekDays error. That is expected.
- The dataset disable payload sends only `enabled:false`; the dataflow disable payload
  sends `enabled:false` plus empty `times` and a `localTimeZoneId`, because the two
  endpoints require different shapes.

## Troubleshooting

- Sign-in problems appear in the console window and a popup, with the exact error.
  To sanity-check outside the app, open PowerShell and run
  `Connect-PowerBIServiceAccount`. If that succeeds, the app will too.
- A FAIL line in the log includes the exact Power BI error for that item, and other
  items keep going.
- If a report shows "(error)" rows for Users or Parameters, that item likely needs a
  higher permission level for that specific read; the rest of the report is still valid.
- The console window is part of the app; keep it open. You can minimize it.
