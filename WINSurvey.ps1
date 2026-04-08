#SS_2026 - MIT License

# Load assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Build the Form ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Server Discovery Inventory'
$form.AutoScaleMode = 'Font'
$form.ClientSize = New-Object System.Drawing.Size(520,420)
$form.StartPosition = 'CenterScreen'

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
$lblFile.Text = 'Import CSV or TXT (headers must be removed):'
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
$grpData.Size = New-Object System.Drawing.Size(480,120)
$form.Controls.Add($grpData)

$chkOS = New-Object System.Windows.Forms.CheckBox
$chkOS.Text = 'OS Version'
$chkOS.Checked = $true
$chkOS.Location = New-Object System.Drawing.Point(15,25)
$grpData.Controls.Add($chkOS)

$chkSQL = New-Object System.Windows.Forms.CheckBox
$chkSQL.Text = 'SQL Installed / Instances'
$chkSQL.Checked = $true
$chkSQL.Location = New-Object System.Drawing.Point(150,25)
$grpData.Controls.Add($chkSQL)

$chkIIS = New-Object System.Windows.Forms.CheckBox
$chkIIS.Text = 'IIS Installed / Sites'
$chkIIS.Checked = $true
$chkIIS.Location = New-Object System.Drawing.Point(320,25)
$grpData.Controls.Add($chkIIS)

$chkUsers = New-Object System.Windows.Forms.CheckBox
$chkUsers.Text = 'User Folders'
$chkUsers.Checked = $true
$chkUsers.Location = New-Object System.Drawing.Point(15,55)
$grpData.Controls.Add($chkUsers)

$chkTasks = New-Object System.Windows.Forms.CheckBox
$chkTasks.Text = 'Scheduled Tasks'
$chkTasks.Checked = $true
$chkTasks.Location = New-Object System.Drawing.Point(150,55)
$grpData.Controls.Add($chkTasks)

$chkPing = New-Object System.Windows.Forms.CheckBox
$chkPing.Text = 'ICMP Ping'
$chkPing.Checked = $true
$chkPing.Location = New-Object System.Drawing.Point(320,55)
$grpData.Controls.Add($chkPing)

$chkPorts = New-Object System.Windows.Forms.CheckBox
$chkPorts.Text = 'Open Web Ports (80,443,8443,8080,8000,25)'
$chkPorts.Checked = $true
$chkPorts.Location = New-Object System.Drawing.Point(15,85)
$grpData.Controls.Add($chkPorts)

# OK / Cancel buttons
$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Text = 'OK'
$btnOK.Location = New-Object System.Drawing.Point(180,280)
$btnOK.Add_Click({ $form.Tag = 'OK'; $form.Close() })
$form.Controls.Add($btnOK)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(275,280)
$btnCancel.Add_Click({ $form.Tag = 'Cancel'; $form.Close() })
$form.Controls.Add($btnCancel)

$form.ShowDialog() | Out-Null

# ---------- Input Validation ----------
if ($form.Tag -ne 'OK' -or
    (:IsNullOrWhiteSpace($txtServer.Text) -and
     :IsNullOrWhiteSpace($txtFile.Text))) {
    return
}

$ExportCsv = $chkCsv.Checked
$DoOS      = $chkOS.Checked
$DoSQL     = $chkSQL.Checked
$DoIIS     = $chkIIS.Checked
$DoUsers   = $chkUsers.Checked
$DoTasks   = $chkTasks.Checked
$DoPing    = $chkPing.Checked
$DoPorts   = $chkPorts.Checked

$Servers = @()

# ---------- Collect Hostnames ----------
if ($txtServer.Text) { $Servers += $txtServer.Text.Trim() }

if ($txtFile.Text) {
    if ($txtFile.Text.ToLower().EndsWith('.txt')) {
        $Servers += Get-Content $txtFile.Text
    } else {
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

    try {
        Invoke-Command -ComputerName $Server -ErrorAction Stop -ScriptBlock {

            $Rows = @()
            $Computer = $env:COMPUTERNAME

            if ($using:DoPing) {
                $ping = Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue
                if ($ping) {
                    $Rows += [pscustomobject]@{
                        ComputerName=$Computer; DataCategory='Network'
                        Name='ICMP Ping'; Value="Online (TTL=$($ping.TimeToLive))"
                    }
                } else {
                    $Rows += [pscustomobject]@{
                        ComputerName=$Computer; DataCategory='Network'
                        Name='ICMP Ping'; Value='No response'
                    }
                }
            }

            if ($using:DoPorts) {
                foreach ($port in 80,443,8443,8080,8000,25) {
                    $open = Test-NetConnection -ComputerName $Computer -Port $port -InformationLevel Quiet
                    $Rows += [pscustomobject]@{
                        ComputerName=$Computer; DataCategory='Network'
                        Name="Port $port"
                        Value= if ($open) {'Open'} else {'Closed'}
                    }
                }
            }

            return $Rows
        }
    }
    catch {
        [pscustomobject]@{
            ComputerName=$Server; DataCategory='ERROR'
            Name='QueryFailed'; Value=$_.Exception.Message
        }
    }
}

# ---------- Optional CSV Export ----------
if ($ExportCsv -and $AllResults) {
    $Desktop = :GetFolderPath('Desktop')
    $AllResults | Export-Csv (Join-Path $Desktop "ServerInventory_$(Get-Date -Format yyyyMMdd_HHmmss).csv") -NoTypeInformation
}

# ---------- Output ----------
if ($AllResults) {
    $AllResults | Out-GridView -Title 'Server Discovery Inventory'
}
