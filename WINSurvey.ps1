# Load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- WinRM Timeout Configuration ----------
# 120 seconds per host to prevent hangs (PowerShell 5.1 compatible)
$WinRMSessionOptions = New-PSSessionOption -OperationTimeout 120000

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


# Generate Domain Machines button
$btnGenerateAD = New-Object System.Windows.Forms.Button
$btnGenerateAD.Text = 'Generate DomainMachines.txt'
$btnGenerateAD.Size = New-Object System.Drawing.Size(220,26)
$btnGenerateAD.Location = New-Object System.Drawing.Point(240,200)
$form.Controls.Add($btnGenerateAD)



$btnGenerateAD.Add_Click({

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            'ActiveDirectory module is not available on this machine. Install RSAT',
            'Error',
            'OK',
            'Error'
        )
        return
    }

    $DaysBack = 10
    $CutoffDate = (Get-Date).AddDays(-$DaysBack)

    try {
        $Computers = Get-ADComputer `
            -Filter { Enabled -eq $true -and LastLogonTimeStamp -gt $CutoffDate } `
            -Properties Name, LastLogonTimeStamp
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            'Failed to query Active Directory.',
            'Error',
            'OK',
            'Error'
        )
        return
    }

    if (-not $Computers) {
        [System.Windows.Forms.MessageBox]::Show(
            "No active domain computers found in the last $DaysBack days.",
            'Information',
            'OK',
            'Information'
        )
        return
    }

    # Resolve output path (same user context)
    $Desktop = [Environment]::GetFolderPath('Desktop')
    $OutputPath = Join-Path $Desktop 'DomainMachines.txt'

    $Computers |
        Select-Object -ExpandProperty Name |
        Sort-Object -Unique |
        Set-Content -Path $OutputPath -Encoding ASCII

    [System.Windows.Forms.MessageBox]::Show(
        "DomainMachines.txt created on Desktop.`r`n`r`nMachines found (active in past 10 days): $($Computers.Count)",
        'Success',
        'OK',
        'Information'
    )
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
$AllResults = foreach ($Server in $Servers) {

    $Current++
    Write-Host "[$Current/$TotalServers] Querying $Server..."

    $PingStatus = 'Unknown'
    $TTLValue   = ''
    $OSHeuristic = 'Unknown'
    $OpenPortCount = 'Not Scanned'
    $WinRMStatus = 'NotAttempted'


    # --- ICMP Ping (PS 5.1 safe) ---
    if ($DoPing) {
        $pingUp = Test-Connection -ComputerName $Server -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($pingUp) {
            $ping = Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue
            $TTLValue = $ping.TimeToLive
            $PingStatus = 'Online'

            if ($TTLValue -ge 65) { $OSHeuristic = 'Windows-like (TTL heuristic)' }
            elseif ($TTLValue -ge 50) { $OSHeuristic = 'Linux-like (TTL heuristic)' }
            else { $OSHeuristic = 'Unknown / Network device' }
        }
        else {
            $PingStatus = 'No response'
        }
    }
    # --- Build scan plan summary (intent, not results) ---
$ScanItems = @()

if ($chkPing.Checked)  { $ScanItems += 'Ping' }
if ($chkOS.Checked)    { $ScanItems += 'OS' }
if ($chkSQL.Checked)   { $ScanItems += 'SQL' }
if ($chkIIS.Checked)   { $ScanItems += 'IIS' }
if ($chkUsers.Checked) { $ScanItems += 'Users' }
if ($chkTasks.Checked) { $ScanItems += 'Tasks' }
if ($chkPorts.Checked) { $ScanItems += 'Ports' }

$ScanPlan =
    if ($ScanItems.Count -gt 0) {
        $ScanItems -join ', '
    }
    else {
        'Nothing (all checks disabled)'
    }
 # --- BEGIN HOST (intent summary) ---
[pscustomobject]@{
    ComputerName = "┌ BEGIN $Server Scan | Summary ┐"
    DataCategory = 'Summary'
    Name         = '┌ BEGIN HOST ┐'
    Value        = "██ $Server ██ | Scan: $ScanPlan"
}

    try {
        $Result = Invoke-Command `
            -ComputerName $Server `
            -SessionOption $WinRMSessionOptions `
            -ErrorAction Stop `
            -ScriptBlock {

                $Rows = @()
                $LocalOpenPortCount = $null
                $Computer = $env:COMPUTERNAME

                if ($using:chkOS.Checked) {
                    $OS = Get-CimInstance Win32_OperatingSystem
                    $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='OS'; Name='Version'; Value=$OS.Caption }
                }

                if ($using:chkSQL.Checked) {
                    $Sql = Get-Service | Where-Object { $_.Name -like 'MSSQL*' -and $_.Name -ne 'MSSQLFDLauncher' }
                    $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='SQL'; Name='Installed'; Value= if ($Sql){'Yes'}else{'No'} }

                    foreach ($Svc in $Sql) {
                        $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='SQL'; Name='Instance'; Value=$Svc.Name }
                    }
                }

                if ($using:chkIIS.Checked) {
                    $IIS = Get-WindowsFeature Web-Server -ErrorAction SilentlyContinue
                    $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='IIS'; Name='Installed'; Value= if ($IIS.InstallState -eq 'Installed'){'Yes'}else{'No'} }
                }

                if ($using:chkUsers.Checked) {
                    Get-ChildItem C:\Users -Directory |
                        Where-Object { $_.Name -notin 'Public','Default','Default User','All Users','Administrator' } |
                        ForEach-Object {
                            $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='UserFolders'; Name='Folder'; Value=$_.Name }
                        }
                }

                if ($using:chkTasks.Checked) {
                    Get-ScheduledTask |
                        Where-Object { $_.TaskPath -notlike '\Microsoft\*' -and $_.Principal.UserId } |
                        ForEach-Object {
                            $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='ScheduledTask'; Name=$_.TaskName; Value=$_.Principal.UserId }
                        }
                }

                if ($using:DoPorts) {
                    foreach ($port in 80,443,8443,8080,8000,25) {
                        $open = Test-NetConnection -ComputerName $Computer -Port $port -InformationLevel Quiet
                        if ($open) { $LocalOpenPortCount++ }
                        $Rows += [pscustomobject]@{ ComputerName=$Computer; DataCategory='Network'; Name="Port $port"; Value= if ($open){'Open'}else{'Closed'} }
                    }
                }

                return [pscustomobject]@{ Rows=$Rows; OpenPortCount=$LocalOpenPortCount }
            }

        $WinRMStatus = 'Success'
       
if ($DoPorts) {
    $OpenPortCount = $Result.OpenPortCount
}

        $Result.Rows
    }
    catch {
        $WinRMStatus = 'Failed'
        [pscustomobject]@{
            ComputerName = $Server
            DataCategory = 'ERROR'
            Name = 'QueryFailed'
            Value = $_.Exception.Message
        }
    }

    # --- END HOST ---
    [pscustomobject]@{
        ComputerName = "└ END $Server Scan | Summary: ┘"
        DataCategory = 'Summary'
        Name = '└ END HOST ┘'
        Value = "██ $Server ██ | PortsOpen=$OpenPortCount | WinRM=$WinRMStatus"
    }
}


$Desktop = [Environment]::GetFolderPath('Desktop')
$AllResults | Export-Csv (
    Join-Path $Desktop "ServerInventory_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
) -NoTypeInformation


# ---------- Output ----------
if ($AllResults) {
    $AllResults | Out-GridView -Title 'Server Discovery Inventory'
}
