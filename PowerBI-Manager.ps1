#Requires -Version 5.1
<#
.SYNOPSIS
    Power BI Workspace Manager - browse workspaces, drill into datasets and
    dataflows, multi-select, take ownership of datasets, and enable/disable
    scheduled refresh.

.DESCRIPTION
    Sign-in happens once at launch (on the console, before the window opens),
    the only step that uses the Power BI module / MSAL. Everything after that
    uses plain REST calls with the acquired token, run on the UI thread with the
    window pumped between items. This avoids the MSAL-on-a-live-window deadlock
    that freezes GUI apps.

    Workspaces and their datasets/dataflows are listed alphabetically. The search
    box filters the workspace list by name; Select All then selects only the
    workspaces currently shown.

    KNOWN API LIMITATIONS (Microsoft limits, not app bugs):
      * Dataflow ownership has no public REST endpoint; the app reports and skips.
      * Dataflows have no GET refresh-schedule endpoint, so on/off state is not
        shown. Enable/Disable still work, EXCEPT a linked dataflow (one that
        refreshes automatically after an upstream flow) has no schedule of its
        own and cannot be "enabled" - Power BI rejects it. That is expected.
      * The access token lasts about one hour. For a very long session you may
        see 401 errors; close and reopen the app to sign in again.
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$WarningPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Power BI Workspace Manager" Height="800" Width="920"
  WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="170"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0" Orientation="Horizontal">
      <Button x:Name="BtnLoad" Content="Load Workspaces" Width="160" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <TextBlock x:Name="LblStatus" Text="Starting..." VerticalAlignment="Center" Margin="10,0,0,0" Foreground="Gray"/>
    </StackPanel>

    <DockPanel Grid.Row="1" Margin="0,8,0,0">
      <TextBlock Text="Search workspaces:" DockPanel.Dock="Left" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <TextBox x:Name="SearchBox" Height="24" VerticalContentAlignment="Center"/>
    </DockPanel>

    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,8">
      <Button x:Name="BtnSelectAll" Content="Select All" Width="100" Margin="0,0,8,0" Padding="4" IsEnabled="False"/>
      <Button x:Name="BtnClearAll" Content="Clear All" Width="100" Margin="0,0,8,0" Padding="4" IsEnabled="False"/>
      <TextBlock Text="Tip: check a WORKSPACE box to select everything inside it. Select All applies to filtered results." VerticalAlignment="Center" Margin="12,0,0,0" Foreground="Gray"/>
    </StackPanel>

    <Border Grid.Row="3" BorderBrush="LightGray" BorderThickness="1">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <TreeView x:Name="Tree" BorderThickness="0"/>
      </ScrollViewer>
    </Border>

    <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,8,0,0">
      <Button x:Name="BtnTakeOwner" Content="Take Ownership" Width="150" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <Button x:Name="BtnEnable" Content="Enable Refresh" Width="150" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <Button x:Name="BtnDisable" Content="Disable Refresh" Width="150" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <CheckBox x:Name="ChkWhatIf" Content="Dry run (WhatIf - preview only, no changes)" VerticalAlignment="Center" Margin="12,0,0,0" IsChecked="True"/>
    </StackPanel>

    <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,8,0,0">
      <Button x:Name="BtnHistory" Content="View History" Width="130" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <Button x:Name="BtnFailures" Content="Failure Report" Width="130" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <Button x:Name="BtnInfo" Content="View Info" Width="130" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <Button x:Name="BtnGSheets" Content="Check Google Sheets" Width="160" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <TextBlock Text="Read-only. Results log below and export to a CSV in the PBI-Reports folder next to the app." VerticalAlignment="Center" Margin="12,0,0,0" Foreground="Gray"/>
    </StackPanel>

    <StackPanel Grid.Row="6" Orientation="Horizontal" Margin="0,8,0,0">
      <Button x:Name="BtnFindFailures" Content="Find Recent Failures" Width="170" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <Button x:Name="BtnRefresh" Content="Refresh Selected" Width="150" Margin="0,0,8,0" Padding="6" IsEnabled="False"/>
      <TextBlock Text="Find Recent Failures narrows checked items to those whose newest refresh failed. Refresh Selected honors Dry run." VerticalAlignment="Center" Margin="12,0,0,0" Foreground="Gray"/>
    </StackPanel>

    <ProgressBar Grid.Row="7" x:Name="Progress" Height="16" Margin="0,8,0,0" Minimum="0" Maximum="100" Value="0"/>

    <TextBox Grid.Row="8" x:Name="LogBox" Margin="0,8,0,0" IsReadOnly="True"
      VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" FontFamily="Consolas" FontSize="12"/>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$BtnLoad      = $Window.FindName('BtnLoad')
$LblStatus    = $Window.FindName('LblStatus')
$SearchBox    = $Window.FindName('SearchBox')
$BtnSelectAll = $Window.FindName('BtnSelectAll')
$BtnClearAll  = $Window.FindName('BtnClearAll')
$Tree         = $Window.FindName('Tree')
$BtnTakeOwner = $Window.FindName('BtnTakeOwner')
$BtnEnable    = $Window.FindName('BtnEnable')
$BtnDisable   = $Window.FindName('BtnDisable')
$BtnHistory   = $Window.FindName('BtnHistory')
$BtnFailures  = $Window.FindName('BtnFailures')
$BtnInfo      = $Window.FindName('BtnInfo')
$BtnGSheets   = $Window.FindName('BtnGSheets')
$BtnFindFailures = $Window.FindName('BtnFindFailures')
$BtnRefresh   = $Window.FindName('BtnRefresh')
$ChkWhatIf    = $Window.FindName('ChkWhatIf')
$Progress     = $Window.FindName('Progress')
$LogBox       = $Window.FindName('LogBox')

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
$script:Token             = $null
$script:Loaded            = $false
$script:ItemCheckBoxes    = New-Object System.Collections.ArrayList
$script:WorkspaceCheckBox = New-Object System.Collections.ArrayList

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Log {
    param([string]$Message)
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $LogBox.AppendText('[' + $stamp + '] ' + $Message + [Environment]::NewLine)
    $LogBox.ScrollToEnd()
}

function Invoke-DoEvents {
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    $cb = [System.Windows.Threading.DispatcherOperationCallback]{
        param($f)
        $f.Continue = $false
        return $null
    }
    [void]$Window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $cb, $frame)
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Set-Busy {
    param([bool]$Busy)
    if ($Busy) {
        $BtnLoad.IsEnabled      = $false
        $BtnSelectAll.IsEnabled = $false
        $BtnClearAll.IsEnabled  = $false
        $BtnTakeOwner.IsEnabled = $false
        $BtnEnable.IsEnabled    = $false
        $BtnDisable.IsEnabled   = $false
        $BtnHistory.IsEnabled   = $false
        $BtnFailures.IsEnabled  = $false
        $BtnInfo.IsEnabled      = $false
        $BtnGSheets.IsEnabled   = $false
        $BtnFindFailures.IsEnabled = $false
        $BtnRefresh.IsEnabled   = $false
    } else {
        $BtnLoad.IsEnabled      = ($script:Token -ne $null)
        $BtnSelectAll.IsEnabled = $script:Loaded
        $BtnClearAll.IsEnabled  = $script:Loaded
        $BtnTakeOwner.IsEnabled = $script:Loaded
        $BtnEnable.IsEnabled    = $script:Loaded
        $BtnDisable.IsEnabled   = $script:Loaded
        $BtnHistory.IsEnabled   = $script:Loaded
        $BtnFailures.IsEnabled  = $script:Loaded
        $BtnInfo.IsEnabled      = $script:Loaded
        $BtnGSheets.IsEnabled   = $script:Loaded
        $BtnFindFailures.IsEnabled = $script:Loaded
        $BtnRefresh.IsEnabled   = $script:Loaded
    }
}

function Get-FreshToken {
    try {
        $raw = Get-PowerBIAccessToken -WarningAction SilentlyContinue -ErrorAction Stop
        if ($raw -is [hashtable]) {
            if ($raw.ContainsKey('Authorization')) { $tok = [string]$raw['Authorization'] }
            else { $tok = [string]($raw.Values | Select-Object -First 1) }
        } else {
            $tok = [string]$raw
        }
        if ([string]::IsNullOrWhiteSpace($tok)) { return $null }
        if ($tok -notlike 'Bearer *') { $tok = 'Bearer ' + $tok }
        return $tok
    } catch {
        return $null
    }
}

function Wait-WithUI {
    # Sleep while keeping the window responsive (pumps events every 200ms).
    param([int]$Seconds)
    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end) {
        Start-Sleep -Milliseconds 200
        Invoke-DoEvents
    }
}

function Invoke-PBI {
    param([string]$Method, [string]$Uri, $Body)
    $headers = @{ Authorization = $script:Token }
    $maxRetries = 5
    $attempt = 0
    while ($true) {
        $attempt = $attempt + 1
        try {
            if ($Body) {
                $d = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $Body -ContentType 'application/json' -ErrorAction Stop
            } else {
                $d = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop
            }
            return @{ ok = $true; data = $d }
        } catch {
            $status = 0
            $retryAfter = 0
            $resp = $_.Exception.Response
            if ($resp -ne $null) {
                try { $status = [int]$resp.StatusCode } catch { }
                try {
                    $ra = $resp.Headers['Retry-After']
                    if ($ra) { $retryAfter = [int]$ra }
                } catch { }
            }

            # 429 = rate limited. Honor Retry-After (capped) and retry.
            if ($status -eq 429 -and $attempt -le $maxRetries) {
                if ($retryAfter -le 0) { $retryAfter = 20 }
                if ($retryAfter -gt 60) { $retryAfter = 60 }
                Write-Log ('  Rate limited (429). Waiting ' + $retryAfter + 's, then retrying (attempt ' + $attempt + ' of ' + $maxRetries + ')...')
                Wait-WithUI $retryAfter
                continue
            }

            $msg = $_.Exception.Message
            try {
                if ($resp -ne $null) {
                    $stream = $resp.GetResponseStream()
                    $srdr = New-Object System.IO.StreamReader($stream)
                    $bodyTxt = $srdr.ReadToEnd()
                    $srdr.Close()
                    if (-not [string]::IsNullOrWhiteSpace($bodyTxt)) { $msg = $bodyTxt }
                }
            } catch { }
            if ($status -eq 429) { $msg = 'Rate limited (429) after ' + $maxRetries + ' retries. Try a smaller selection or wait a few minutes.' }
            return @{ ok = $false; message = $msg }
        }
    }
}

# ---------------------------------------------------------------------------
# Tree building
# ---------------------------------------------------------------------------
function Add-WorkspaceNode {
    param($w)
    $wsId = $w.id
    $wsName = $w.name

    $wsCb = New-Object System.Windows.Controls.CheckBox
    $wsCb.Content = 'WORKSPACE: ' + $wsName
    $wsCb.FontWeight = 'Bold'
    $childList = New-Object System.Collections.ArrayList

    $wsItem = New-Object System.Windows.Controls.TreeViewItem
    $wsItem.Header = $wsCb
    $wsItem.IsExpanded = $false
    $wsItem.Tag = $wsName

    $dsR = Invoke-PBI 'Get' ('https://api.powerbi.com/v1.0/myorg/groups/' + $wsId + '/datasets') $null
    if ($dsR.ok) {
        foreach ($ds in @($dsR.data.value | Sort-Object -Property name)) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = 'Dataset:  ' + $ds.name
            $cb.Tag = [pscustomobject]@{ Type = 'Dataset'; Id = $ds.id; Name = $ds.name; WorkspaceId = $wsId; WorkspaceName = $wsName }
            [void]$script:ItemCheckBoxes.Add($cb)
            [void]$childList.Add($cb)
            $ci = New-Object System.Windows.Controls.TreeViewItem
            $ci.Header = $cb
            [void]$wsItem.Items.Add($ci)
        }
    } else {
        Write-Log ('  Could not read datasets in "' + $wsName + '": ' + $dsR.message)
    }

    $dfR = Invoke-PBI 'Get' ('https://api.powerbi.com/v1.0/myorg/groups/' + $wsId + '/dataflows') $null
    if ($dfR.ok) {
        foreach ($df in @($dfR.data.value | Sort-Object -Property name)) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = 'Dataflow: ' + $df.name
            $cb.Tag = [pscustomobject]@{ Type = 'Dataflow'; Id = $df.objectId; Name = $df.name; WorkspaceId = $wsId; WorkspaceName = $wsName }
            [void]$script:ItemCheckBoxes.Add($cb)
            [void]$childList.Add($cb)
            $ci = New-Object System.Windows.Controls.TreeViewItem
            $ci.Header = $cb
            [void]$wsItem.Items.Add($ci)
        }
    } else {
        Write-Log ('  Could not read dataflows in "' + $wsName + '": ' + $dfR.message)
    }

    $wsCb.Tag = $childList
    $wsCb.Add_Checked({   foreach ($c in $this.Tag) { $c.IsChecked = $true } })
    $wsCb.Add_Unchecked({ foreach ($c in $this.Tag) { $c.IsChecked = $false } })
    [void]$script:WorkspaceCheckBox.Add($wsCb)
    [void]$Tree.Items.Add($wsItem)
}

function Load-Workspaces {
    $Tree.Items.Clear()
    $script:ItemCheckBoxes.Clear()
    $script:WorkspaceCheckBox.Clear()
    $SearchBox.Text = ''

    Write-Log 'Retrieving workspaces...'
    Invoke-DoEvents

    $workspaces = New-Object System.Collections.ArrayList
    $top = 100
    $skip = 0
    while ($true) {
        $uri = 'https://api.powerbi.com/v1.0/myorg/groups?$top=' + $top + '&$skip=' + $skip
        $r = Invoke-PBI 'Get' $uri $null
        if (-not $r.ok) { Write-Log ('ERROR getting workspaces: ' + $r.message); break }
        $batch = @($r.data.value)
        foreach ($w in $batch) { [void]$workspaces.Add($w) }
        if ($batch.Count -lt $top) { break }
        $skip = $skip + $top
        Invoke-DoEvents
    }

    $sorted = @($workspaces | Sort-Object -Property name)
    Write-Log ('Found ' + $sorted.Count + ' workspace(s). Reading datasets and dataflows...')
    $Progress.Maximum = [Math]::Max($sorted.Count, 1)
    $Progress.Value = 0
    Invoke-DoEvents

    foreach ($w in $sorted) {
        Add-WorkspaceNode $w
        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    Write-Log 'Load complete.'
}

function Get-SelectedItems {
    $selected = New-Object System.Collections.ArrayList
    foreach ($cb in $script:ItemCheckBoxes) {
        if ($cb.IsChecked -eq $true) { [void]$selected.Add($cb.Tag) }
    }
    return $selected
}

function Do-Action {
    param([string]$Action)
    $items = Get-SelectedItems
    if ($items.Count -eq 0) {
        Write-Log 'Nothing selected. Check one or more datasets or dataflows first.'
        return
    }
    $whatIf = ($ChkWhatIf.IsChecked -eq $true)
    $mode = 'LIVE'
    if ($whatIf) { $mode = 'DRY RUN' }
    Write-Log ('--- ' + $Action + ' on ' + $items.Count + ' item(s) [' + $mode + '] ---')

    Set-Busy $true
    $Progress.Maximum = [Math]::Max($items.Count, 1)
    $Progress.Value = 0
    $ok = 0
    $fail = 0
    $skipped = 0

    foreach ($it in $items) {
        $label = $it.Type + ' "' + $it.Name + '" (workspace: ' + $it.WorkspaceName + ')'

        if ($Action -eq 'TakeOwner' -and $it.Type -eq 'Dataflow') {
            Write-Log ('SKIP  ' + $label + ' -- dataflow ownership is not available in the Power BI public API. Take over manually in the Service.')
            $skipped = $skipped + 1
            $Progress.Value = $Progress.Value + 1
            Invoke-DoEvents
            continue
        }

        if ($whatIf) {
            Write-Log ('[WhatIf] would ' + $Action + ' -> ' + $label)
            $Progress.Value = $Progress.Value + 1
            Invoke-DoEvents
            continue
        }

        if ($it.Type -eq 'Dataset') {
            $base = 'https://api.powerbi.com/v1.0/myorg/groups/' + $it.WorkspaceId + '/datasets/' + $it.Id
        } else {
            $base = 'https://api.powerbi.com/v1.0/myorg/groups/' + $it.WorkspaceId + '/dataflows/' + $it.Id
        }

        if ($Action -eq 'TakeOwner') {
            $r = Invoke-PBI 'Post' ($base + '/Default.TakeOver') $null
        } elseif ($Action -eq 'Enable') {
            $r = Invoke-PBI 'Patch' ($base + '/refreshSchedule') '{ "value": { "enabled": true } }'
        } else {
            # Datasets reject extra fields on disable; dataflows REQUIRE them.
            if ($it.Type -eq 'Dataflow') {
                $body = '{ "value": { "enabled": false, "times": [], "localTimeZoneId": "UTC" } }'
            } else {
                $body = '{ "value": { "enabled": false } }'
            }
            $r = Invoke-PBI 'Patch' ($base + '/refreshSchedule') $body
        }

        if ($r.ok) {
            Write-Log ('OK    ' + $Action + ' -> ' + $label)
            $ok = $ok + 1
        } else {
            Write-Log ('FAIL  ' + $Action + ' -> ' + $label + ' :: ' + $r.message)
            $fail = $fail + 1
        }
        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    Write-Log ('Done. Success: ' + $ok + '  Failed: ' + $fail + '  Skipped: ' + $skipped)
    Set-Busy $false
}

# ---------------------------------------------------------------------------
# Read-only reporting (run history, failures, info) + CSV export
# ---------------------------------------------------------------------------
function Get-ItemBase {
    param($it)
    if ($it.Type -eq 'Dataset') {
        return 'https://api.powerbi.com/v1.0/myorg/groups/' + $it.WorkspaceId + '/datasets/' + $it.Id
    } else {
        return 'https://api.powerbi.com/v1.0/myorg/groups/' + $it.WorkspaceId + '/dataflows/' + $it.Id
    }
}

function Get-DurationSec {
    param($startTime, $endTime)
    if ($startTime -and $endTime) {
        try { return [int]([datetime]$endTime - [datetime]$startTime).TotalSeconds } catch { return '' }
    }
    return ''
}

function Save-Csv {
    param($Rows, [string]$Prefix)
    if ($Rows.Count -eq 0) { Write-Log 'Nothing to export (no rows).'; return $null }
    $base = $PSScriptRoot
    if ([string]::IsNullOrEmpty($base)) { $base = (Get-Location).Path }
    $dir = Join-Path $base 'PBI-Reports'
    if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $path = Join-Path $dir ($Prefix + '_' + $stamp + '.csv')
    try {
        $Rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
        return $path
    } catch {
        Write-Log ('Could not write CSV: ' + $_.Exception.Message)
        return $null
    }
}

function Get-Runs {
    # Returns normalized run objects for one item (newest first).
    param($it, [int]$TopN)
    $out = New-Object System.Collections.ArrayList
    $base = Get-ItemBase $it

    if ($it.Type -eq 'Dataset') {
        $r = Invoke-PBI 'Get' ($base + '/refreshes?$top=' + $TopN) $null
        if (-not $r.ok) { Write-Log ('  History error for "' + $it.Name + '": ' + $r.message); return $out }
        foreach ($run in @($r.data.value)) {
            $code = ''
            $desc = ''
            if ($run.serviceExceptionJson) {
                try {
                    $p = $run.serviceExceptionJson | ConvertFrom-Json
                    $code = [string]$p.errorCode
                    $desc = [string]$p.errorDescription
                } catch {
                    $desc = [string]$run.serviceExceptionJson
                }
            }
            [void]$out.Add([pscustomobject]@{
                Status = $run.status; StartTime = $run.startTime; EndTime = $run.endTime;
                RefreshType = $run.refreshType; RequestId = $run.requestId;
                ErrorCode = $code; ErrorMessage = $desc
            })
        }
    } else {
        # Dataflows use transactions; no error message is exposed by the API.
        $r = Invoke-PBI 'Get' ($base + '/transactions') $null
        if (-not $r.ok) { Write-Log ('  History error for "' + $it.Name + '": ' + $r.message); return $out }
        $count = 0
        foreach ($run in @($r.data.value)) {
            if ($count -ge $TopN) { break }
            [void]$out.Add([pscustomobject]@{
                Status = $run.status; StartTime = $run.startTime; EndTime = $run.endTime;
                RefreshType = $run.refreshType; RequestId = $run.id;
                ErrorCode = ''; ErrorMessage = ''
            })
            $count = $count + 1
        }
    }
    return $out
}

function Do-History {
    param([bool]$FailuresOnly)
    $items = Get-SelectedItems
    if ($items.Count -eq 0) { Write-Log 'Nothing selected. Check one or more items first.'; return }
    $topN = 10
    $title = 'History'
    if ($FailuresOnly) { $title = 'Failure Report' }
    Write-Log ('--- ' + $title + ' on ' + $items.Count + ' item(s) (last ' + $topN + ' runs each) ---')

    Set-Busy $true
    $Progress.Maximum = [Math]::Max($items.Count, 1)
    $Progress.Value = 0
    $rows = New-Object System.Collections.ArrayList
    $failCount = 0

    foreach ($it in $items) {
        $runs = Get-Runs $it $topN
        $shownForItem = $false
        foreach ($run in $runs) {
            $isFail = ([string]$run.Status -match 'Fail')
            if ($FailuresOnly -and -not $isFail) { continue }

            [void]$rows.Add([pscustomobject]@{
                Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                Status = $run.Status; StartTime = $run.StartTime; EndTime = $run.EndTime;
                DurationSec = (Get-DurationSec $run.StartTime $run.EndTime);
                RefreshType = $run.RefreshType; RequestId = $run.RequestId;
                ErrorCode = $run.ErrorCode; ErrorMessage = $run.ErrorMessage
            })

            if ($isFail) { $failCount = $failCount + 1 }

            if ($FailuresOnly) {
                $emsg = $run.ErrorCode
                if (-not $emsg) { $emsg = $run.ErrorMessage }
                Write-Log ('  FAIL ' + $it.Type + ' "' + $it.Name + '" @ ' + $run.StartTime + ' :: ' + $emsg)
            } elseif (-not $shownForItem) {
                # For full history, log just the most recent run per item; CSV has the rest.
                Write-Log ('  ' + $it.Type + ' "' + $it.Name + '": latest = ' + $run.Status + ' @ ' + $run.StartTime)
                $shownForItem = $true
            }
        }
        if ($FailuresOnly -and -not $shownForItem) { }
        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    $prefix = 'History'
    if ($FailuresOnly) { $prefix = 'FailureReport' }
    $path = Save-Csv $rows $prefix
    if ($path) { Write-Log ('CSV saved: ' + $path) }
    if ($FailuresOnly) {
        Write-Log ('Done. ' + $failCount + ' failed run(s) across ' + $items.Count + ' item(s).')
    } else {
        Write-Log ('Done. ' + $rows.Count + ' run(s) exported across ' + $items.Count + ' item(s).')
    }
    Set-Busy $false
}

function Do-Info {
    $items = Get-SelectedItems
    if ($items.Count -eq 0) { Write-Log 'Nothing selected. Check one or more items first.'; return }
    Write-Log ('--- Info on ' + $items.Count + ' item(s) ---')

    Set-Busy $true
    $Progress.Maximum = [Math]::Max($items.Count, 1)
    $Progress.Value = 0
    $rows = New-Object System.Collections.ArrayList

    foreach ($it in $items) {
        $base = Get-ItemBase $it
        $nSrc = 0
        $nParam = 0
        $nUser = 0

        # Data sources (datasets and dataflows)
        $r = Invoke-PBI 'Get' ($base + '/datasources') $null
        if ($r.ok) {
            foreach ($src in @($r.data.value)) {
                $cd = $src.connectionDetails
                $v1 = ''
                $v2 = ''
                if ($cd) {
                    if ($cd.server) { $v1 = [string]$cd.server } elseif ($cd.url) { $v1 = [string]$cd.url }
                    if ($cd.database) { $v2 = [string]$cd.database } elseif ($cd.path) { $v2 = [string]$cd.path }
                }
                [void]$rows.Add([pscustomobject]@{
                    Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                    InfoType = 'DataSource'; Name = [string]$src.datasourceType; Value1 = $v1; Value2 = $v2; Value3 = ''
                })
                $nSrc = $nSrc + 1
            }
        } else {
            [void]$rows.Add([pscustomobject]@{ Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name; InfoType = 'DataSource'; Name = '(error)'; Value1 = $r.message; Value2 = ''; Value3 = '' })
        }

        if ($it.Type -eq 'Dataset') {
            # Parameters
            $r = Invoke-PBI 'Get' ($base + '/parameters') $null
            if ($r.ok) {
                foreach ($p in @($r.data.value)) {
                    [void]$rows.Add([pscustomobject]@{
                        Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                        InfoType = 'Parameter'; Name = [string]$p.name; Value1 = [string]$p.type; Value2 = [string]$p.currentValue; Value3 = ''
                    })
                    $nParam = $nParam + 1
                }
            }
            # Refresh schedule details
            $r = Invoke-PBI 'Get' ($base + '/refreshSchedule') $null
            if ($r.ok) {
                $sc = $r.data
                $days = (@($sc.days) -join ',')
                $times = (@($sc.times) -join ',')
                [void]$rows.Add([pscustomobject]@{
                    Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                    InfoType = 'Schedule'; Name = 'enabled=' + [string]$sc.enabled; Value1 = $days; Value2 = $times; Value3 = [string]$sc.localTimeZoneId
                })
            }
        } else {
            [void]$rows.Add([pscustomobject]@{ Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name; InfoType = 'Schedule'; Name = 'n/a'; Value1 = 'dataflow schedule not readable via API'; Value2 = ''; Value3 = '' })
        }

        # Users / permissions
        $r = Invoke-PBI 'Get' ($base + '/users') $null
        if ($r.ok) {
            foreach ($u in @($r.data.value)) {
                $right = ''
                if ($u.datasetUserAccessRight) { $right = [string]$u.datasetUserAccessRight }
                elseif ($u.dataflowUserAccessRight) { $right = [string]$u.dataflowUserAccessRight }
                $who = $u.identifier
                if (-not $who) { $who = $u.displayName }
                [void]$rows.Add([pscustomobject]@{
                    Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                    InfoType = 'User'; Name = [string]$who; Value1 = [string]$u.principalType; Value2 = $right; Value3 = ''
                })
                $nUser = $nUser + 1
            }
        } else {
            [void]$rows.Add([pscustomobject]@{ Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name; InfoType = 'User'; Name = '(error)'; Value1 = $r.message; Value2 = ''; Value3 = '' })
        }

        Write-Log ('  ' + $it.Type + ' "' + $it.Name + '": ' + $nSrc + ' source(s), ' + $nParam + ' param(s), ' + $nUser + ' user(s)')
        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    $path = Save-Csv $rows 'Info'
    if ($path) { Write-Log ('CSV saved: ' + $path) }
    Write-Log ('Done. ' + $rows.Count + ' info row(s) exported across ' + $items.Count + ' item(s).')
    Set-Busy $false
}

function Test-IsGoogleSheets {
    # Returns the matching detail string if a datasource looks like Google Sheets, else $null.
    param($src)
    $cd = $src.connectionDetails
    $candidates = New-Object System.Collections.ArrayList
    [void]$candidates.Add([string]$src.datasourceType)
    if ($cd) {
        [void]$candidates.Add([string]$cd.kind)
        [void]$candidates.Add([string]$cd.path)
        [void]$candidates.Add([string]$cd.url)
    }
    foreach ($c in $candidates) {
        if ([string]::IsNullOrEmpty($c)) { continue }
        $lc = $c.ToLower()
        if ($lc -match 'googlesheet' -or $lc -match 'google sheet' -or $lc.Contains('docs.google.com/spreadsheets') -or $lc.Contains('spreadsheets.google.com')) {
            return $c
        }
    }
    return $null
}

function Do-CheckGoogleSheets {
    $items = Get-SelectedItems
    if ($items.Count -eq 0) { Write-Log 'Nothing selected. Check one or more items first.'; return }
    Write-Log ('--- Check Google Sheets on ' + $items.Count + ' item(s) ---')

    Set-Busy $true
    $Progress.Maximum = [Math]::Max($items.Count, 1)
    $Progress.Value = 0
    $rows = New-Object System.Collections.ArrayList
    $hits = 0

    foreach ($it in $items) {
        $base = Get-ItemBase $it
        $r = Invoke-PBI 'Get' ($base + '/datasources') $null
        if (-not $r.ok) {
            Write-Log ('  Could not read datasources for "' + $it.Name + '": ' + $r.message)
            $Progress.Value = $Progress.Value + 1
            Invoke-DoEvents
            continue
        }

        $matches = New-Object System.Collections.ArrayList
        foreach ($src in @($r.data.value)) {
            $hit = Test-IsGoogleSheets $src
            if ($hit) { [void]$matches.Add($hit) }
        }

        if ($matches.Count -gt 0) {
            $hits = $hits + 1
            $detail = (($matches | Select-Object -Unique) -join ' | ')
            Write-Log ('  GOOGLE SHEETS: ' + $it.Type + ' "' + $it.Name + '" (workspace: ' + $it.WorkspaceName + ')')
            [void]$rows.Add([pscustomobject]@{
                Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                UsesGoogleSheets = 'YES'; MatchDetail = $detail
            })
        } else {
            [void]$rows.Add([pscustomobject]@{
                Workspace = $it.WorkspaceName; ItemType = $it.Type; ItemName = $it.Name;
                UsesGoogleSheets = 'no'; MatchDetail = ''
            })
        }

        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    $path = Save-Csv $rows 'GoogleSheetsCheck'
    if ($path) { Write-Log ('CSV saved: ' + $path) }
    Write-Log ('Done. ' + $hits + ' of ' + $items.Count + ' item(s) use Google Sheets.')
    Set-Busy $false
}

function Get-LatestStatus {
    # Returns @{ status = <string>; message = <string> } for an item's newest run.
    param($it)
    $base = Get-ItemBase $it
    if ($it.Type -eq 'Dataset') {
        $r = Invoke-PBI 'Get' ($base + '/refreshes?$top=1') $null
        if (-not $r.ok) { return @{ status = 'Error'; message = $r.message } }
        $runs = @($r.data.value)
        if ($runs.Count -eq 0) { return @{ status = 'None'; message = '' } }
        return @{ status = [string]$runs[0].status; message = '' }
    } else {
        $r = Invoke-PBI 'Get' ($base + '/transactions') $null
        if (-not $r.ok) { return @{ status = 'Error'; message = $r.message } }
        $runs = @($r.data.value)
        if ($runs.Count -eq 0) { return @{ status = 'None'; message = '' } }
        $latest = $runs | Sort-Object -Property startTime -Descending | Select-Object -First 1
        return @{ status = [string]$latest.status; message = '' }
    }
}

function Do-FindRecentFailures {
    $items = Get-SelectedItems
    if ($items.Count -eq 0) {
        Write-Log 'Nothing selected. Select a scope first (search + Select All, or check a workspace), then Find Recent Failures.'
        return
    }
    Write-Log ('--- Find Recent Failures among ' + $items.Count + ' selected item(s) ---')

    Set-Busy $true
    $Progress.Maximum = [Math]::Max($items.Count, 1)
    $Progress.Value = 0
    $failedKeys = @{}
    $failCount = 0

    foreach ($it in $items) {
        $res = Get-LatestStatus $it
        $st = $res.status
        if ($st -eq 'Error') {
            Write-Log ('  ? ' + $it.Type + ' "' + $it.Name + '": could not read status: ' + $res.message)
        } elseif ($st -match 'Fail') {
            $failCount = $failCount + 1
            $failedKeys[($it.WorkspaceId + '|' + $it.Type + '|' + $it.Id)] = $true
            Write-Log ('  FAILED: ' + $it.Type + ' "' + $it.Name + '" (workspace: ' + $it.WorkspaceName + ')')
        }
        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    # Keep only the failed items checked; uncheck everything else.
    foreach ($cb in $script:ItemCheckBoxes) {
        $tag = $cb.Tag
        $key = $tag.WorkspaceId + '|' + $tag.Type + '|' + $tag.Id
        if ($failedKeys.ContainsKey($key)) { $cb.IsChecked = $true } else { $cb.IsChecked = $false }
    }

    Write-Log ('Done. ' + $failCount + ' item(s) had a failed most recent refresh; only those are now selected.')
    if ($failCount -gt 0) {
        Write-Log 'Use "Refresh Selected" to re-run them all, or uncheck to do them one at a time.'
    }
    Set-Busy $false
}

function Do-RefreshSelected {
    $items = Get-SelectedItems
    if ($items.Count -eq 0) {
        Write-Log 'Nothing selected. Check the item(s) you want to refresh.'
        return
    }
    $whatIf = ($ChkWhatIf.IsChecked -eq $true)
    $mode = 'LIVE'
    if ($whatIf) { $mode = 'DRY RUN' }
    Write-Log ('--- Refresh on ' + $items.Count + ' item(s) [' + $mode + '] ---')

    Set-Busy $true
    $Progress.Maximum = [Math]::Max($items.Count, 1)
    $Progress.Value = 0
    $ok = 0
    $fail = 0

    foreach ($it in $items) {
        $label = $it.Type + ' "' + $it.Name + '" (workspace: ' + $it.WorkspaceName + ')'
        if ($whatIf) {
            Write-Log ('[WhatIf] would refresh -> ' + $label)
            $Progress.Value = $Progress.Value + 1
            Invoke-DoEvents
            continue
        }

        $base = Get-ItemBase $it
        if ($it.Type -eq 'Dataset') {
            $r = Invoke-PBI 'Post' ($base + '/refreshes') '{ "notifyOption": "NoNotification" }'
        } else {
            $r = Invoke-PBI 'Post' ($base + '/refreshes?processType=default') '{ "notifyOption": "NoNotification" }'
        }

        if ($r.ok) {
            Write-Log ('OK    refresh started -> ' + $label)
            $ok = $ok + 1
        } else {
            Write-Log ('FAIL  refresh -> ' + $label + ' :: ' + $r.message)
            $fail = $fail + 1
        }
        $Progress.Value = $Progress.Value + 1
        Invoke-DoEvents
    }

    Write-Log ('Done. Started: ' + $ok + '  Failed: ' + $fail)
    Set-Busy $false
}

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------
$SearchBox.Add_TextChanged({
    $q = $SearchBox.Text
    if ([string]::IsNullOrWhiteSpace($q)) {
        foreach ($it in $Tree.Items) { $it.Visibility = [System.Windows.Visibility]::Visible }
        return
    }
    $ql = $q.ToLower()
    foreach ($it in $Tree.Items) {
        $name = ([string]$it.Tag).ToLower()
        if ($name.Contains($ql)) {
            $it.Visibility = [System.Windows.Visibility]::Visible
        } else {
            $it.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }
})

$BtnLoad.Add_Click({
    if ($script:Token -eq $null) {
        Write-Log 'Not signed in. Close and restart the app to sign in.'
        return
    }
    Set-Busy $true
    try {
        Load-Workspaces
        $script:Loaded = $true
    } catch {
        Write-Log ('LOAD ERROR: ' + $_.Exception.Message)
    } finally {
        Set-Busy $false
    }
})

$BtnSelectAll.Add_Click({
    # Only select workspaces currently visible (respects the search filter).
    foreach ($it in $Tree.Items) {
        if ($it.Visibility -eq [System.Windows.Visibility]::Visible) {
            $it.Header.IsChecked = $true
        }
    }
})

$BtnClearAll.Add_Click({
    foreach ($cb in $script:WorkspaceCheckBox) { $cb.IsChecked = $false }
    foreach ($cb in $script:ItemCheckBoxes)    { $cb.IsChecked = $false }
})

$BtnTakeOwner.Add_Click({ Do-Action 'TakeOwner' })
$BtnEnable.Add_Click({    Do-Action 'Enable' })
$BtnDisable.Add_Click({   Do-Action 'Disable' })

$BtnHistory.Add_Click({  Do-History $false })
$BtnFailures.Add_Click({ Do-History $true })
$BtnInfo.Add_Click({     Do-Info })
$BtnGSheets.Add_Click({  Do-CheckGoogleSheets })
$BtnFindFailures.Add_Click({ Do-FindRecentFailures })
$BtnRefresh.Add_Click({  Do-RefreshSelected })

# ---------------------------------------------------------------------------
# Sign in ONCE on the console thread, before the window opens.
# ---------------------------------------------------------------------------
function Connect-AtStartup {
    try {
        Write-Host 'Power BI Workspace Manager'
        Write-Host '--------------------------'
        Write-Host 'Checking for the MicrosoftPowerBIMgmt module...'
        if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
            Write-Host 'Not found. Installing for current user (first run only; please wait)...'
            Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
        }
        Write-Host 'Loading module...'
        Import-Module MicrosoftPowerBIMgmt -ErrorAction Stop
        Write-Host ''
        Write-Host 'A Power BI sign-in window or browser will now open.'
        Write-Host 'Complete the sign-in there. The app window opens right after.'
        Write-Host ''
        Connect-PowerBIServiceAccount -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
        $script:Token = Get-FreshToken
        if ($script:Token -eq $null) {
            throw 'Signed in, but no access token was returned. Please restart and try again.'
        }
        Write-Host 'Signed in. Opening the app...'
        return $true
    } catch {
        $emsg = $_.Exception.Message
        Write-Host ('ERROR: ' + $emsg)
        $popup = 'Could not sign in to Power BI:' + [Environment]::NewLine + [Environment]::NewLine + $emsg + [Environment]::NewLine + [Environment]::NewLine + 'The app will still open. Fix the issue and restart it.'
        [void][System.Windows.MessageBox]::Show($popup, 'Sign-in failed', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

$connected = Connect-AtStartup
if ($connected) {
    $LblStatus.Text = 'Connected. Click "Load Workspaces".'
    $LblStatus.Foreground = 'Green'
    $BtnLoad.IsEnabled = $true
    Write-Log 'Connected to Power BI. Click "Load Workspaces" to begin.'
} else {
    $LblStatus.Text = 'Sign-in failed - restart the app'
    $LblStatus.Foreground = 'Red'
    Write-Log 'Sign-in failed at startup. Close and restart the app to try again.'
}
Write-Log 'Safety: "Dry run" is ON by default - uncheck it to make real changes.'
[void]$Window.ShowDialog()
