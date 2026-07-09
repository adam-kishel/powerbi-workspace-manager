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
  Title="Power BI Workspace Manager" Height="740" Width="920"
  WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
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

    <ProgressBar Grid.Row="5" x:Name="Progress" Height="16" Margin="0,8,0,0" Minimum="0" Maximum="100" Value="0"/>

    <TextBox Grid.Row="6" x:Name="LogBox" Margin="0,8,0,0" IsReadOnly="True"
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
    } else {
        $BtnLoad.IsEnabled      = ($script:Token -ne $null)
        $BtnSelectAll.IsEnabled = $script:Loaded
        $BtnClearAll.IsEnabled  = $script:Loaded
        $BtnTakeOwner.IsEnabled = $script:Loaded
        $BtnEnable.IsEnabled    = $script:Loaded
        $BtnDisable.IsEnabled   = $script:Loaded
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

function Invoke-PBI {
    param([string]$Method, [string]$Uri, $Body)
    $headers = @{ Authorization = $script:Token }
    try {
        if ($Body) {
            $d = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $Body -ContentType 'application/json' -ErrorAction Stop
        } else {
            $d = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop
        }
        return @{ ok = $true; data = $d }
    } catch {
        $msg = $_.Exception.Message
        try {
            $resp = $_.Exception.Response
            if ($resp -ne $null) {
                $stream = $resp.GetResponseStream()
                $srdr = New-Object System.IO.StreamReader($stream)
                $bodyTxt = $srdr.ReadToEnd()
                $srdr.Close()
                if (-not [string]::IsNullOrWhiteSpace($bodyTxt)) { $msg = $bodyTxt }
            }
        } catch { }
        return @{ ok = $false; message = $msg }
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
