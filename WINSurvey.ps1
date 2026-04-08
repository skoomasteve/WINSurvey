# Load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Build the Form ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'WINSurvey | SS'
$form.AutoScaleMode = 'Font'
$form.ClientSize = New-Object System.Drawing.Size(590,420)
$form.StartPosition = 'CenterScreen'
$form.AutoScroll = $true

# Server label
$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = 'Server name (optional):'
$lblServer.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($lblServer)

# Server textbox
$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(180,18)
$txtServer.Width = 300
$form.Controls.Add($txtServer)

# File label
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = 'Import CSV or TXT (without headers):'
$lblFile.Location = New-Object System.Drawing.Point(20,55)
$form.Controls.Add($lblFile)

# File path textbox
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(180,52)
$txtFile.Width = 220
$txtFile.ReadOnly = $true
$form.Controls.Add($txtFile)

# Browse button
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(410,50)
$form.Controls.Add($btnBrowse)

$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Filter = 'CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt'
$fileDialog.Multiselect = $false

$btnBrowse.Add_Click({
    if ($fileDialog.ShowDialog() -eq 'OK') {
        $txtFile.Text = $fileDialog.FileName
    }
})

# CSV export checkbox
$chkCsv = New-Object System.Windows.Forms.CheckBox
$chkCsv.Text = 'Output CSV to Desktop?'
$chkCsv.AutoSize = $true
$chkCsv.Location = New-Object System.Drawing.Point(180,90)
$form.Controls.Add($chkCsv)

# ---------- Datapoint Selection ----------
$grpData = New-Object System.Windows.Forms.GroupBox
$grpData.Text = 'Datapoints to query'
$grpData.Location = New-Object System.Drawing.Point(20,130)
$grpData.Size = New-Object System.Drawing.Size(480,200)
$form.Controls.Add($grpData)

$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Parent = $grpData
$flow.Location = New-Object System.Drawing.Point(10,20)
$flow.Size = New-Object System.Drawing.Size(460,170)
$flow.FlowDirection = 'TopDown'
$flow.WrapContents = $false
$flow.AutoScroll = $true
$grpData.Controls.Add($flow)

function New-DataCheckbox {
    param ([string]$Text)
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $Text
    $cb.Checked = $true
    $cb.AutoSize = $true
    $cb.Width = 440
    return $cb
}

$chkOS    = New-DataCheckbox 'OS Version'
$chkSQL   = New-DataCheckbox 'SQL Installed / Instances'
$chkIIS   = New-DataCheckbox 'IIS Installed / Sites'
$chkUsers = New-DataCheckbox 'User Folders'
$chkTasks = New-DataCheckbox 'Scheduled Tasks'
$chkPing  = New-DataCheckbox 'ICMP Ping'
$chkPorts = New-DataCheckbox 'Open Web Ports (80,443,8443,8080,8000,25)'

$flow.Controls.AddRange(@(
    $chkOS,$chkSQL,$chkIIS,$chkUsers,$chkTasks,$chkPing,$chkPorts
))

# OK / Cancel buttons
$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Text = 'OK'
$btnOK.Location = New-Object System.Drawing.Point(180,350)
$btnOK.Add_Click({ $form.Tag = 'OK'; $form.Close() })
$form.Controls.Add($btnOK)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(275,350)
$btnCancel.Add_Click({ $form.Tag = 'Cancel'; $form.Close() })
$form.Controls.Add($btnCancel)

$form.ShowDialog() | Out-Null

# ---------- Input Validation ----------
if ($form.Tag -ne 'OK' -or
    ([string]::IsNullOrWhiteSpace($txtServer.Text) -and
     [string]::IsNullOrWhiteSpace($txtFile.Text))) {
    return
}

$ExportCsv = $chkCsv.Checked
$DoPing    = $chkPing.Checked
$DoPorts   = $chkPorts.Checked

$Servers = @()
if ($txtServer.Text) { $Servers += $txtServer.Text.Trim() }

if ($txtFile.Text) {
    if ($txtFile.Text.ToLower().EndsWith('.txt')) {
        $Servers += Get-Content $txtFile.Text
    }
    else {
        $Servers += Import-Csv -Path $txtFile.Text -Header 'Host' |
            Select-Object -ExpandProperty Host
    }
}

$Servers = $Servers | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
$TotalServers = $Servers.Count
$Current = 0

Write-Host "Starting server inventory for $TotalServers server(s)..."

# ---------- Query Servers ----------
# ---------- Query Servers ----------
$AllResults = foreach ($Server in $Servers) {

    $Current++
    Write-Host "[$Current/$TotalServers] Querying $Server..."

    # --- Initialize per-host summary tracking ---
    $PingStatus     = 'Unknown'
    $TTLValue       = ''
    $OSHeuristic    = 'Unknown'
    $OpenPortCount  = 0
    $WinRMStatus    = 'NotAttempted'

    # --- ICMP Ping (LOCAL) ---
    if ($DoPing) {
        $ping = Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue
        if ($ping) {
            $ttl = $ping.TimeToLive

            if ($ttl -ge 65) {
                $osGuess = 'Windows-like (TTL heuristic)'
            }
            elseif ($ttl -ge 50) {
                $osGuess = 'Linux/Unix-like (TTL heuristic)'
            }
            else {
                $osGuess = 'Unknown / Network device'
            }

            $PingStatus  = 'Online'
            $TTLValue    = $ttl
            $OSHeuristic = $osGuess
        }
        else {
            $PingStatus = 'No response'
        }
    }

    # --- SUMMARY ROW (EMITTED FIRST, PLACEHOLDER COUNTS) ---
    [pscustomobject]@{
        ComputerName = $Server
        DataCategory = 'Summary'
        Name         = '=== HOST SUMMARY ==='
        Value        = "████ $Server ████ | Ping=$PingStatus | TTL=$TTLValue | OS=$OSHeuristic | PortsOpen=pending | WinRM=pending"
    }

    # --- Network detail row (ICMP) ---
    if ($DoPing) {
        if ($PingStatus -eq 'Online') {
            [pscustomobject]@{
                ComputerName = $Server
                DataCategory = 'Network'
                Name         = 'ICMP Ping'
                Value        = "Online | TTL=$TTLValue | OS Guess=$OSHeuristic"
            }
        }
        else {
            [pscustomobject]@{
                ComputerName = $Server
                DataCategory = 'Network'
                Name         = 'ICMP Ping'
                Value        = 'No response'
            }
        }
    }

    try {
        # --- Run all WinRM-based inventory and network port checks ---
        $Result = Invoke-Command -ComputerName $Server -ErrorAction Stop -ScriptBlock {

            $Rows = @()
            $LocalOpenPortCount = 0
            $Computer = $env:COMPUTERNAME

            # ---------- OS ----------
            if ($using:chkOS.Checked) {
                $OS = Get-CimInstance Win32_OperatingSystem
                $Rows += [pscustomobject]@{
                    ComputerName = $Computer
                    DataCategory = 'OS'
                    Name         = 'Version'
                    Value        = $OS.Caption
                }
            }

            # ---------- SQL ----------
            if ($using:chkSQL.Checked) {
                $SqlServices = Get-Service |
                    Where-Object { $_.Name -like 'MSSQL*' -and $_.Name -ne 'MSSQLFDLauncher' }

                $Rows += [pscustomobject]@{
                    ComputerName = $Computer
                    DataCategory = 'SQL'
                    Name         = 'Installed'
                    Value        = if ($SqlServices) { 'Yes' } else { 'No' }
                }

                foreach ($Svc in $SqlServices) {
                    $Instance = if ($Svc.Name -eq 'MSSQLSERVER') {
                        'MSSQLSERVER (Default)'
                    } else {
                        $Svc.Name -replace '^MSSQL\$', ''
                    }

                    $Rows += [pscustomobject]@{
                        ComputerName = $Computer
                        DataCategory = 'SQL'
                        Name         = 'Instance'
                        Value        = $Instance
                    }
                }
            }

            # ---------- IIS ----------
            if ($using:chkIIS.Checked) {
                $IIS = Get-WindowsFeature Web-Server -ErrorAction SilentlyContinue
                $Installed = $IIS -and $IIS.InstallState -eq 'Installed'

                $Rows += [pscustomobject]@{
                    ComputerName = $Computer
                    DataCategory = 'IIS'
                    Name         = 'Installed'
                    Value        = if ($Installed) { 'Yes' } else { 'No' }
                }

                if ($Installed) {
                    Import-Module WebAdministration
                    foreach ($Site in Get-Website) {
                        $Rows += [pscustomobject]@{
                            ComputerName = $Computer
                            DataCategory = 'IIS'
                            Name         = 'Site'
                            Value        = $Site.Name
                        }
                    }
                }
            }

            # ---------- User Profiles ----------
            if ($using:chkUsers.Checked) {
                Get-ChildItem C:\Users -Directory |
                    Where-Object { $_.Name -notin 'Public','Default','Default User','All Users','Administrator' } |
                    ForEach-Object {
                        $Rows += [pscustomobject]@{
                            ComputerName = $Computer
                            DataCategory = 'UserFolders'
                            Name         = 'Folder'
                            Value        = $_.Name
                        }
                    }
            }

            # ---------- Scheduled Tasks ----------
            if ($using:chkTasks.Checked) {
                Get-ScheduledTask |
                    Where-Object { $_.TaskPath -notlike '\Microsoft\*' -and $_.Principal.UserId } |
                    ForEach-Object {
                        $Rows += [pscustomobject]@{
                            ComputerName = $Computer
                            DataCategory = 'ScheduledTask'
                            Name         = $_.TaskName
                            Value        = $_.Principal.UserId
                        }
                    }
            }

            # ---------- Network Ports ----------
            if ($using:DoPorts) {
                foreach ($port in 80,443,8443,8080,8000,25) {
                    $open = Test-NetConnection -ComputerName $Computer -Port $port -InformationLevel Quiet
                    if ($open) { $LocalOpenPortCount++ }

                    $Rows += [pscustomobject]@{
                        ComputerName = $Computer
                        DataCategory = 'Network'
                        Name         = "Port $port"
                        Value        = if ($open) { 'Open' } else { 'Closed' }
                    }
                }
            }

            return [pscustomobject]@{
                Rows          = $Rows
                OpenPortCount = $LocalOpenPortCount
            }
        }

        $WinRMStatus   = 'Success'
        $OpenPortCount = $Result.OpenPortCount
        $Result.Rows
    }
    catch {
        $WinRMStatus = 'Failed'
        [pscustomobject]@{
            ComputerName = $Server
            DataCategory = 'ERROR'
            Name         = 'QueryFailed'
            Value        = $_.Exception.Message
        }
    }

    # --- FINAL SUMMARY ROW (CORRECT COUNTS, EMITTED LAST) ---
    [pscustomobject]@{
        ComputerName = $Server
        DataCategory = 'Summary'
        Name         = '=== HOST SUMMARY (FINAL) ==='
        Value        = "████ $Server ████ | Ping=$PingStatus | TTL=$TTLValue | OS=$OSHeuristic | PortsOpen=$OpenPortCount | WinRM=$WinRMStatus"
    }
}

# ---------- Optional CSV Export ----------
if ($ExportCsv -and $AllResults) {
    $Desktop = [Environment]::GetFolderPath('Desktop')
    $AllResults | Export-Csv (
        Join-Path $Desktop "ServerInventory_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
    ) -NoTypeInformation
}

# ---------- Output ----------
if ($AllResults) {
    $AllResults | Out-GridView -Title 'Server Discovery Inventory'
}
