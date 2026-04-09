

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
$form.ClientSize = New-Object System.Drawing.Size(590,520)
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
$btnGenerateAD.Location = New-Object System.Drawing.Point(220,125)
$form.Controls.Add($btnGenerateAD)

$btnGenerateAD.Add_Click({
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            'ActiveDirectory module is not available on this machine. Install RSAT.',
            'Error','OK','Error'
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
            'Error','OK','Error'
        )
        return
    }

    if (-not $Computers) {
        [System.Windows.Forms.MessageBox]::Show(
            "No active domain computers found in the last $DaysBack days.",
            'Information','OK','Information'
        )
        return
    }

    $Desktop = [Environment]::GetFolderPath('Desktop')
    $OutputPath = Join-Path $Desktop 'DomainMachines.txt'

    $Computers |
        Select-Object -ExpandProperty Name |
        Sort-Object -Unique |
        Set-Content -Path $OutputPath -Encoding ASCII

    [System.Windows.Forms.MessageBox]::Show(
        "DomainMachines.txt created on Desktop.`r`nMachines found (active in past $DaysBack days): $($Computers.Count)",
        'Success','OK','Information'
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
$grpData.Size = New-Object System.Drawing.Size(480,300)
$form.Controls.Add($grpData)

$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Parent = $grpData
$flow.Location = New-Object System.Drawing.Point(10,20)
$flow.Size = New-Object System.Drawing.Size(460,270)
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

$chkOS          = New-DataCheckbox 'Windows Version'
$chkSQL         = New-DataCheckbox 'SQL Installed / Instances'
$chkIIS         = New-DataCheckbox 'IIS Installed / Sites'
$chkUsersGroups = New-DataCheckbox 'Users and Groups (Last Login, Admins, RDP)'
$chkUsers       = New-DataCheckbox 'User Folders'
$chkTasks       = New-DataCheckbox 'Scheduled Tasks'
$chkPing        = New-DataCheckbox 'ICMP Ping'
$chkPorts       = New-DataCheckbox 'Open Web Ports (80,443,8443,8080,8000,25)'

$flow.Controls.AddRange(@(
    $chkOS,$chkSQL,$chkIIS,$chkUsersGroups,$chkUsers,$chkTasks,$chkPing,$chkPorts
))

# OK / Cancel buttons
$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Text = 'OK'
$btnOK.Location = New-Object System.Drawing.Point(180,460)
$btnOK.Add_Click({ $form.Tag = 'OK'; $form.Close() })
$form.Controls.Add($btnOK)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(275,460)
$btnCancel.Add_Click({ $form.Tag = 'Cancel'; $form.Close() })
$form.Controls.Add($btnCancel)

$form.ShowDialog() | Out-Null

# ---------- Input Validation ----------
if ($form.Tag -ne 'OK' -or
    ([string]::IsNullOrWhiteSpace($txtServer.Text) -and
     [string]::IsNullOrWhiteSpace($txtFile.Text))) {
    return
}

$ExportCsv      = $chkCsv.Checked
$DoPing         = $chkPing.Checked
$DoPorts        = $chkPorts.Checked
$DoUsersGroups  = $chkUsersGroups.Checked

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

$Servers = $Servers | Where-Object { $_ } | Sort-Object -Unique
$TotalServers = $Servers.Count
$Current = 0

Write-Host "Starting server inventory for $TotalServers server(s)..."

# ---------- Query Servers ----------
# ---------- Query Servers ----------
# ---------- Query Servers ----------
$AllResults = foreach ($Server in $Servers) {

    $Current++
    Write-Host "[$Current/$TotalServers] Querying $Server..."

    # -------------------------
    # Per-host summary state
    # -------------------------
    $PingStatus        = 'Unknown'
    $OnlineState       = 'Unknown'
    $TTLValue          = ''
    $OSHeuristic       = 'Unknown'
    $OpenPortCount     = 'Not Scanned'
    $WinRMStatus       = 'NotAttempted'
    $IISSummaryState   = 'N/A'
    $SQLSummaryState   = 'N/A'
    $LastLoggedUser    = 'Unknown'

    # -------------------------
    # BEGIN HOST — Scan Intent
    # -------------------------
    $ScanItems = @()
    if ($DoPing)         { $ScanItems += 'Ping' }
    if ($chkOS.Checked)  { $ScanItems += 'OS' }
    if ($chkSQL.Checked) { $ScanItems += 'SQL' }
    if ($chkIIS.Checked) { $ScanItems += 'IIS' }
    if ($chkUsers.Checked) { $ScanItems += 'UserFolders' }
    if ($chkTasks.Checked) { $ScanItems += 'ScheduledTasks' }
    if ($DoUsersGroups)  { $ScanItems += 'UsersAndGroups' }
    if ($DoPorts)        { $ScanItems += 'Ports' }

    [pscustomobject]@{
        ComputerName = "┌$Server | Preliminary Scan Summary ┐"
        DataCategory = 'Summary'
        Name         = '┌ BEGIN HOST ┐'
        Value        = "██ $Server ██ | Scan: $($ScanItems -join ', ')"
    }

    # -------------------------
    # ICMP Ping (local)
    # -------------------------
    if ($DoPing) {
        $PingUp = Test-Connection -ComputerName $Server -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($PingUp) {
            $Ping = Test-Connection -ComputerName $Server -Count 1 -ErrorAction SilentlyContinue
            $TTLValue    = $Ping.TimeToLive
            $PingStatus  = 'Online'
            $OnlineState = 'Online'

            if     ($TTLValue -ge 65) { $OSHeuristic = 'Windows-like (TTL heuristic)' }
            elseif ($TTLValue -ge 50) { $OSHeuristic = 'Linux-like (TTL heuristic)' }
            else                      { $OSHeuristic = 'Unknown / Network device' }

            [pscustomobject]@{
                ComputerName = $Server
                DataCategory = 'Network'
                Name         = 'ICMP Ping'
                Value        = "Online | TTL=$TTLValue | OS Guess=$OSHeuristic"
            }
        }
        else {
            $PingStatus  = 'No response'
            $OnlineState = 'Offline'

            [pscustomobject]@{
                ComputerName = $Server
                DataCategory = 'Network'
                Name         = 'ICMP Ping'
                Value        = 'No response'
            }
        }
    }
# -------------------------
# External Port Scan (LOCAL, fast with timeout)
# -------------------------
if ($DoPorts) {

    $PortsToScan   = @(80,443,8443,8080,8000,25)
    $PortTimeoutMs = 500   # <<< ADJUST HERE (milliseconds)
    $OpenPortCount = 0

    foreach ($Port in $PortsToScan) {

        $IsOpen = $false
        $Client = New-Object System.Net.Sockets.TcpClient

        try {
            $AsyncResult = $Client.BeginConnect($Server, $Port, $null, $null)

            if ($AsyncResult.AsyncWaitHandle.WaitOne($PortTimeoutMs, $false)) {
                $Client.EndConnect($AsyncResult)
                $IsOpen = $true
                $OpenPortCount++
            }
            else {
                $Client.Close()
            }
        }
        catch {
            $Client.Close()
        }

        [pscustomobject]@{
            ComputerName = $Server
            DataCategory = 'Network'
            Name         = "Port $Port"
            Value        = if ($IsOpen) { 'Open' } else { 'Closed' }
        }
    }
}

    # -------------------------
    # WinRM Data Collection
    # -------------------------
    try {
        $Result = Invoke-Command `
            -ComputerName $Server `
            -SessionOption $WinRMSessionOptions `
            -ErrorAction Stop `
            -ScriptBlock {

                $Rows = @()
                $Computer = $env:COMPUTERNAME

                $IISOut     = 'N/A'
                $SQLOut     = 'N/A'
                $LastUserOut = 'Unknown'

                if ($using:chkOS.Checked) {
                    $OS = Get-CimInstance Win32_OperatingSystem
                    $Rows += [pscustomobject]@{
                        ComputerName = $Computer
                        DataCategory = 'OS'
                        Name         = 'Version'
                        Value        = $OS.Caption
                    }
                }

                if ($using:chkSQL.Checked) {
                    $SqlServices = Get-Service |
                        Where-Object { $_.Name -like 'MSSQL*' -and $_.Name -ne 'MSSQLFDLauncher' }

                    $SQLOut = if ($SqlServices) { 'Enabled' } else { 'Absent' }

                    $Rows += [pscustomobject]@{
                        ComputerName = $Computer
                        DataCategory = 'SQL'
                        Name         = 'Installed'
                        Value        = if ($SqlServices) { 'Yes' } else { 'No' }
                    }
                }

                if ($using:chkIIS.Checked) {
                    $IIS = Get-WindowsFeature Web-Server -ErrorAction SilentlyContinue
                    $IISOut = if ($IIS -and $IIS.InstallState -eq 'Installed') { 'Active' } else { 'Inactive' }

                    $Rows += [pscustomobject]@{
                        ComputerName = $Computer
                        DataCategory = 'IIS'
                        Name         = 'Installed'
                        Value        = if ($IISOut -eq 'Active') { 'Yes' } else { 'No' }
                    }
                }

                if ($using:chkUsersGroups.Checked) {
                    try {
                        $Reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI'
                        if ($Reg.LastLoggedOnUser) {
                            $LastUserOut = $Reg.LastLoggedOnUser
                        }
                    } catch {}

                    $Rows += [pscustomobject]@{
                        ComputerName = $Computer
                        DataCategory = 'UsersAndGroups'
                        Name         = 'Last Logged-On User'
                        Value        = $LastUserOut
                    }
                }

                return [pscustomobject]@{
                    Rows     = $Rows
                    IISState = $IISOut
                    SQLState = $SQLOut
                    LastUser = $LastUserOut
                }
            }

        $WinRMStatus     = 'Success'
        $IISSummaryState = $Result.IISState
        $SQLSummaryState = $Result.SQLState
        $LastLoggedUser  = $Result.LastUser

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

    # -------------------------
    # END HOST — Result Summary
    # -------------------------
    $PortsSummary = if ($DoPorts) { "PortsOpen=$OpenPortCount" } else { "PortsOpen=Not Scanned" }

    [pscustomobject]@{
        ComputerName = "└ $Server | End Scan Summary ┘"
        DataCategory = 'Summary'
        Name         = '└ END HOST ┘'
        Value        = "██ $Server ██ | State=$OnlineState | IIS=$IISSummaryState | SQL=$SQLSummaryState | LastUser=$LastLoggedUser | $PortsSummary | WinRM=$WinRMStatus"
    }
}

# ---------- Write CSV ----------
$Desktop = [Environment]::GetFolderPath('Desktop')
$AllResults | Export-Csv (
    Join-Path $Desktop "ServerInventory_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
) -NoTypeInformation

# ---------- Output ----------
$AllResults | Out-GridView -Title 'Server Discovery Inventory'
