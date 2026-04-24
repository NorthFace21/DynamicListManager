<#
   .SYNOPSIS
   Dynamic List Suite - Main V2 (Single Script)
   .DESCRIPTION
   - Dashboard and tools are hosted in one process.
   - No external script launching.
   - Attribute Manager is loaded in a dedicated tab for easier troubleshooting.
#>

# --- HIDE HOSTING CONSOLE ---
$ConsoleVisibility = @{
    namespace = 'Win32'
    name = 'Console'
    member = '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
}
try {
    Add-Type -MemberDefinition $ConsoleVisibility.member -Name $ConsoleVisibility.name -Namespace $ConsoleVisibility.namespace
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) { [Win32.Console]::ShowWindow($hwnd, 0) }
} catch {}

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

# ==========================================
# CONFIGURATION
# ==========================================
$SqlInstance = "SQL16-DIVERS"
$Database = "DBTEST"

# Keep embedded tools aligned with host SQL target.
$env:DYN_SQL_INSTANCE = $SqlInstance
$env:DYN_SQL_DATABASE = $Database

# ==========================================
# SQL HELPER
# ==========================================
function Invoke-Sql {
    param(
        [Parameter(Mandatory=$true)][string]$Query,
        [switch]$ReturnScalar
    )

    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    $Conn = $null

    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;Connect Timeout=5;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
        $Conn.Open()

        $Cmd = $Conn.CreateCommand()
        $Cmd.CommandText = $Query
        $Cmd.CommandTimeout = 10

        if ($ReturnScalar) {
            $ResultInfo.Data = $Cmd.ExecuteScalar()
        } else {
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd
            $DS = New-Object System.Data.DataSet
            $Adapter.Fill($DS) | Out-Null
            if ($DS.Tables.Count -gt 0) {
                $ResultInfo.Data = $DS.Tables[0]
            } else {
                $ResultInfo.Data = New-Object System.Data.DataTable
            }
        }

        $ResultInfo.Success = $true
    } catch {
        $ResultInfo.Message = $_.Exception.Message
    } finally {
        if ($Conn) {
            $Conn.Close()
            $Conn.Dispose()
        }
    }

    return $ResultInfo
}

# ==========================================
# MAIN FORM
# ==========================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Dynamic List Suite - V2"
$Form.Size = New-Object System.Drawing.Size(1100, 860)
$Form.MinimumSize = New-Object System.Drawing.Size(980, 760)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$StatusStrip = New-Object System.Windows.Forms.StatusStrip
$StatusStrip.SizingGrip = $false
$StatusMain = New-Object System.Windows.Forms.ToolStripStatusLabel
$StatusMain.Spring = $true
$StatusMain.TextAlign = "MiddleLeft"
$StatusMain.Text = "Ready"
$StatusClock = New-Object System.Windows.Forms.ToolStripStatusLabel
$StatusClock.Text = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$StatusStrip.Items.AddRange(@($StatusMain, $StatusClock))
$Form.Controls.Add($StatusStrip)

$ClockTimer = New-Object System.Windows.Forms.Timer
$ClockTimer.Interval = 1000
$ClockTimer.Add_Tick({ $StatusClock.Text = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") })
$ClockTimer.Start()

$MainTabs = New-Object System.Windows.Forms.TabControl
$MainTabs.Dock = "Fill"
$Form.Controls.Add($MainTabs)

$TabDashboard = New-Object System.Windows.Forms.TabPage
$TabDashboard.Text = "Dashboard"
$TabDashboard.BackColor = [System.Drawing.Color]::FromArgb(237, 241, 246)
$MainTabs.TabPages.Add($TabDashboard)

# ==========================================
# DASHBOARD LAYOUT
# ==========================================
$DashHeader = New-Object System.Windows.Forms.Panel
$DashHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$DashHeader.Height = 78
$DashHeader.BackColor = [System.Drawing.Color]::FromArgb(25, 57, 89)
$TabDashboard.Controls.Add($DashHeader)

$DashTitle = New-Object System.Windows.Forms.Label
$DashTitle.Text = "Dynamic List Manager"
$DashTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14, [System.Drawing.FontStyle]::Bold)
$DashTitle.ForeColor = "White"
$DashTitle.Location = New-Object System.Drawing.Point(20, 12)
$DashTitle.AutoSize = $true
$DashHeader.Controls.Add($DashTitle)

$DashSub = New-Object System.Windows.Forms.Label
$DashSub.Text = "Single-script tabbed dashboard"
$DashSub.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$DashSub.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 240)
$DashSub.Location = New-Object System.Drawing.Point(22, 42)
$DashSub.AutoSize = $true
$DashHeader.Controls.Add($DashSub)

$DashBody = New-Object System.Windows.Forms.Panel
$DashBody.Dock = [System.Windows.Forms.DockStyle]::Fill
$TabDashboard.Controls.Add($DashBody)

$DashFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$DashFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
$DashFlow.Padding = New-Object System.Windows.Forms.Padding(24, 26, 18, 18)
$DashFlow.WrapContents = $true
$DashFlow.AutoScroll = $false
$DashFlow.BackColor = [System.Drawing.Color]::FromArgb(237, 241, 246)
$DashBody.Controls.Add($DashFlow)

$DashTopSpacer = New-Object System.Windows.Forms.Panel
$DashTopSpacer.Size = New-Object System.Drawing.Size(900, 10)
$DashTopSpacer.Margin = New-Object System.Windows.Forms.Padding(0)
$DashTopSpacer.BackColor = [System.Drawing.Color]::Transparent
$DashFlow.Controls.Add($DashTopSpacer)

$TabDashboard.Controls.SetChildIndex($DashBody, 1)
$TabDashboard.Controls.SetChildIndex($DashHeader, 0)

# ==========================================
# EMBEDDED TOOL SCRIPTS (SINGLE-FILE MODE)
# ==========================================
# region EmbeddedToolScripts
$EmbeddedScript_AttributeManager = @'
<#
   .SYNOPSIS
   Attribute Manager - Embedded Form
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing

$SqlInstance = if ($env:DYN_SQL_INSTANCE) { $env:DYN_SQL_INSTANCE } else { "SQL16-DIVERS" }
$Database    = if ($env:DYN_SQL_DATABASE) { $env:DYN_SQL_DATABASE } else { "DBTEST" }

function Invoke-Sql {
    param($Query, $ReturnScalar=$false)

    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    $Conn = $null

    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;Connect Timeout=5;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
        $Conn.Open()

        $Cmd = $Conn.CreateCommand()
        $Cmd.CommandText = $Query
        $Cmd.CommandTimeout = 10

        if ($ReturnScalar) {
            $ResultInfo.Data = $Cmd.ExecuteScalar()
        } else {
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd
            $DS = New-Object System.Data.DataSet
            $Adapter.Fill($DS) | Out-Null
            if ($DS.Tables.Count -gt 0) {
                $ResultInfo.Data = $DS.Tables[0]
            } else {
                $ResultInfo.Data = New-Object System.Data.DataTable
            }
        }
        $ResultInfo.Success = $true
    }
    catch {
        $ResultInfo.Message = "Server=$SqlInstance; Database=$Database; Error=$($_.Exception.Message)"
    }
    finally {
        if ($Conn) {
            $Conn.Close()
            $Conn.Dispose()
        }
    }

    return $ResultInfo
}

function New-AttributeManagerForm {
    $InvokeSqlFn = ${function:Invoke-Sql}.GetNewClosure()

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "AD Attribute Mapper"
    $Form.Size = New-Object System.Drawing.Size(560, 650)
    $Form.StartPosition = "CenterScreen"

    $StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $StatusLabel.Text = "Ready - Server: $SqlInstance / DB: $Database"
    $StatusStrip.Items.Add($StatusLabel)
    $Form.Controls.Add($StatusStrip)

    $GrpInput = New-Object System.Windows.Forms.GroupBox
    $GrpInput.Text = "Attribute Editor"
    $GrpInput.Location = New-Object System.Drawing.Point(12, 10)
    $GrpInput.Size = New-Object System.Drawing.Size(520, 150)
    $Form.Controls.Add($GrpInput)

    $LblFriendly = New-Object System.Windows.Forms.Label
    $LblFriendly.Text = "Friendly Name:"
    $LblFriendly.Location = New-Object System.Drawing.Point(15, 25)
    $GrpInput.Controls.Add($LblFriendly)

    $TxtFriendly = New-Object System.Windows.Forms.TextBox
    $TxtFriendly.Location = New-Object System.Drawing.Point(15, 45)
    $TxtFriendly.Size = New-Object System.Drawing.Size(220, 25)
    $GrpInput.Controls.Add($TxtFriendly)

    $LblAD = New-Object System.Windows.Forms.Label
    $LblAD.Text = "AD Attribute:"
    $LblAD.Location = New-Object System.Drawing.Point(250, 25)
    $GrpInput.Controls.Add($LblAD)

    $TxtAD = New-Object System.Windows.Forms.TextBox
    $TxtAD.Location = New-Object System.Drawing.Point(250, 45)
    $TxtAD.Size = New-Object System.Drawing.Size(220, 25)
    $GrpInput.Controls.Add($TxtAD)

    $LblType = New-Object System.Windows.Forms.Label
    $LblType.Text = "Type:"
    $LblType.Location = New-Object System.Drawing.Point(15, 80)
    $GrpInput.Controls.Add($LblType)

    $CmbType = New-Object System.Windows.Forms.ComboBox
    $CmbType.Location = New-Object System.Drawing.Point(15, 100)
    $CmbType.Size = New-Object System.Drawing.Size(100, 25)
    $CmbType.DropDownStyle = "DropDownList"
    [void]$CmbType.Items.AddRange(@("String", "Int", "Bool"))
    $CmbType.SelectedIndex = 0
    $GrpInput.Controls.Add($CmbType)

    $TxtID = New-Object System.Windows.Forms.TextBox
    $TxtID.Visible = $false
    $GrpInput.Controls.Add($TxtID)

    $BtnAdd = New-Object System.Windows.Forms.Button
    $BtnAdd.Text = "Save New"
    $BtnAdd.Location = New-Object System.Drawing.Point(140, 98)
    $BtnAdd.Size = New-Object System.Drawing.Size(80, 28)
    $BtnAdd.BackColor = "LightGreen"
    $GrpInput.Controls.Add($BtnAdd)

    $BtnUpdate = New-Object System.Windows.Forms.Button
    $BtnUpdate.Text = "Update"
    $BtnUpdate.Location = New-Object System.Drawing.Point(230, 98)
    $BtnUpdate.Size = New-Object System.Drawing.Size(80, 28)
    $BtnUpdate.Enabled = $false
    $GrpInput.Controls.Add($BtnUpdate)

    $BtnDelete = New-Object System.Windows.Forms.Button
    $BtnDelete.Text = "Delete"
    $BtnDelete.Location = New-Object System.Drawing.Point(320, 98)
    $BtnDelete.Size = New-Object System.Drawing.Size(80, 28)
    $BtnDelete.Enabled = $false
    $BtnDelete.BackColor = "LightSalmon"
    $GrpInput.Controls.Add($BtnDelete)

    $BtnClear = New-Object System.Windows.Forms.Button
    $BtnClear.Text = "Clear"
    $BtnClear.Location = New-Object System.Drawing.Point(410, 98)
    $BtnClear.Size = New-Object System.Drawing.Size(80, 28)
    $GrpInput.Controls.Add($BtnClear)

    $BtnRefresh = New-Object System.Windows.Forms.Button
    $BtnRefresh.Text = "Refresh Table"
    $BtnRefresh.Location = New-Object System.Drawing.Point(12, 170)
    $BtnRefresh.Size = New-Object System.Drawing.Size(520, 30)
    $BtnRefresh.BackColor = "LightBlue"
    $Form.Controls.Add($BtnRefresh)

    $Grid = New-Object System.Windows.Forms.DataGridView
    $Grid.Location = New-Object System.Drawing.Point(12, 210)
    $Grid.Size = New-Object System.Drawing.Size(520, 360)
    $Grid.ReadOnly = $true
    $Grid.SelectionMode = "FullRowSelect"
    $Grid.MultiSelect = $false
    $Grid.AllowUserToAddRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.AutoSizeColumnsMode = "Fill"
    $Form.Controls.Add($Grid)

    $DoSetStatus = {
        param($Msg, $IsError=$false)
        $StatusLabel.Text = "[$([DateTime]::Now.ToString('HH:mm:ss'))] $Msg"
        if ($IsError) { $StatusLabel.ForeColor = "Red" } else { $StatusLabel.ForeColor = "Black" }
    }.GetNewClosure()

    $DoRefreshGrid = {
        $null = $DoSetStatus.Invoke("Querying Database...", $false)
        $Grid.DataSource = $null
        $Res = $InvokeSqlFn.Invoke("SELECT AttrID, FriendlyName, ADAttribute, DataType FROM Sys_AttributeMap", $false)

        if ($Res.Success) {
            $Grid.DataSource = $Res.Data
            if ($Grid.Columns["AttrID"]) { $Grid.Columns["AttrID"].Visible = $false }
            $Count = if ($Res.Data) { $Res.Data.Rows.Count } else { 0 }
            if ($Count -eq 0) {
                $null = $DoSetStatus.Invoke("Connected. Rows Loaded: 0 (table currently empty)", $false)
            } else {
                $null = $DoSetStatus.Invoke("Connected. Rows Loaded: $Count", $false)
            }
        } else {
            $null = $DoSetStatus.Invoke("Error: $($Res.Message)", $true)
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to load attribute data from SQL.`n`nServer: $SqlInstance`nDatabase: $Database`n`nError:`n$($Res.Message)",
                "SQL Load Error",
                0,
                16
            ) | Out-Null
        }
    }.GetNewClosure()

    $DoClearForm = {
        $TxtID.Text = ""
        $TxtFriendly.Text = ""
        $TxtAD.Text = ""
        $CmbType.SelectedIndex = 0
        $BtnAdd.Enabled = $true
        $BtnUpdate.Enabled = $false
        $BtnDelete.Enabled = $false
        $Grid.ClearSelection()
    }.GetNewClosure()

    $BtnAdd.Add_Click(({
        if ($TxtFriendly.Text -and $TxtAD.Text) {
            $Q = "INSERT INTO Sys_AttributeMap (FriendlyName, ADAttribute, DataType) VALUES ('$($TxtFriendly.Text)', '$($TxtAD.Text)', '$($CmbType.SelectedItem)')"
            $Res = $InvokeSqlFn.Invoke($Q, $false)
            if ($Res.Success) {
                $null = $DoRefreshGrid.Invoke()
                $null = $DoClearForm.Invoke()
            } else {
                $null = $DoSetStatus.Invoke("Insert Failed: $($Res.Message)", $true)
            }
        }
    }).GetNewClosure())

    $BtnUpdate.Add_Click(({
        if ($TxtID.Text) {
            $Q = "UPDATE Sys_AttributeMap SET FriendlyName='$($TxtFriendly.Text)', ADAttribute='$($TxtAD.Text)', DataType='$($CmbType.SelectedItem)' WHERE AttrID=$($TxtID.Text)"
            $Res = $InvokeSqlFn.Invoke($Q, $false)
            if ($Res.Success) {
                $null = $DoRefreshGrid.Invoke()
                $null = $DoClearForm.Invoke()
            } else {
                $null = $DoSetStatus.Invoke("Update Failed: $($Res.Message)", $true)
            }
        }
    }).GetNewClosure())

    $BtnDelete.Add_Click(({
        if ($TxtID.Text) {
            if ([System.Windows.Forms.MessageBox]::Show("Delete this mapping?", "Confirm", "YesNo") -eq "Yes") {
                $Q = "DELETE FROM Sys_AttributeMap WHERE AttrID=$($TxtID.Text)"
                $Res = $InvokeSqlFn.Invoke($Q, $false)
                if ($Res.Success) {
                    $null = $DoRefreshGrid.Invoke()
                    $null = $DoClearForm.Invoke()
                } else {
                    $null = $DoSetStatus.Invoke("Delete Failed: $($Res.Message)", $true)
                }
            }
        }
    }).GetNewClosure())

    $BtnClear.Add_Click(({ $null = $DoClearForm.Invoke() }).GetNewClosure())
    $BtnRefresh.Add_Click(({ $null = $DoRefreshGrid.Invoke() }).GetNewClosure())

    $Grid.Add_CellClick(({
        if ($Grid.SelectedRows.Count -gt 0) {
            $Row = $Grid.SelectedRows[0]
            $TxtID.Text = $Row.Cells["AttrID"].Value
            $TxtFriendly.Text = $Row.Cells["FriendlyName"].Value
            $TxtAD.Text = $Row.Cells["ADAttribute"].Value
            $CmbType.SelectedItem = $Row.Cells["DataType"].Value

            $BtnAdd.Enabled = $false
            $BtnUpdate.Enabled = $true
            $BtnDelete.Enabled = $true
        }
    }).GetNewClosure())

    # In embedded hosting, Form.Load is not always reliable, so refresh immediately.
    $null = $DoRefreshGrid.Invoke()
    return $Form
}

return (New-AttributeManagerForm)
'@

$EmbeddedScript_CreateNewList = @'
<# 
   .SYNOPSIS
   Dynamic List Architect (v2.5) - Final Stable Release
   .DESCRIPTION
   Fixes Hashtable type casting errors during Save.
   Includes Nested Logic, Auto-Refresh, and Safe Preview.
#>

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing

# ==========================================
# CONFIGURATION
# ==========================================
$SqlInstance = if ($env:DYN_SQL_INSTANCE) { $env:DYN_SQL_INSTANCE } else { "SQL16-DIVERS" }
$Database    = if ($env:DYN_SQL_DATABASE) { $env:DYN_SQL_DATABASE } else { "DBTEST" }
# ==========================================

$CurrentUser = $env:USERNAME
$SyncIntervals = @(5, 15, 30, 60, 120, 240, 480, 1440)

# --- SQL HELPER (ROBUST) ---
function Invoke-Sql {
    param(
        [Parameter(Mandatory=$true)][string]$Query, 
        [Parameter(Mandatory=$false)][hashtable]$Parameters = @{}, 
        [switch]$ReturnScalar, 
        [switch]$ReturnId
    )
    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    
    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
        $Conn.Open()
        $Cmd = $Conn.CreateCommand(); $Cmd.CommandText = $Query
        
        # Robust Parameter Addition
        if ($Parameters) {
            foreach ($Key in $Parameters.Keys) {
                $Val = if ($Parameters[$Key] -eq $null) { [DBNull]::Value } else { $Parameters[$Key] }
                $Cmd.Parameters.AddWithValue("@$Key", $Val) | Out-Null
            }
        }

        if ($ReturnScalar) { $ResultInfo.Data = $Cmd.ExecuteScalar() } 
        elseif ($ReturnId) { $ResultInfo.Data = $Cmd.ExecuteScalar() }
        else {
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd
            $DS = New-Object System.Data.DataSet; $Adapter.Fill($DS) | Out-Null; $ResultInfo.Data = $DS.Tables[0]
        }
        $Conn.Close(); $ResultInfo.Success = $true
    }
    catch { 
        $ResultInfo.Message = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("SQL Error: $($_.Exception.Message)", "Database Error", 0, 16) 
    }
    return $ResultInfo
}

# --- AD USER PICKER ---
function Show-UserPicker {
    $PickForm = New-Object System.Windows.Forms.Form
    $PickForm.Text = "Find User"; $PickForm.Size = "500, 400"; $PickForm.StartPosition = "CenterParent"
    $TxtSearch = New-Object System.Windows.Forms.TextBox; $TxtSearch.Location = "10, 10"; $TxtSearch.Size = "380, 25"
    $BtnSearch = New-Object System.Windows.Forms.Button; $BtnSearch.Text = "Search"; $BtnSearch.Location = "400, 9"; $BtnSearch.Size = "70, 26"
    $ListResults = New-Object System.Windows.Forms.ListBox; $ListResults.Location = "10, 45"; $ListResults.Size = "460, 270"
    $BtnSelect = New-Object System.Windows.Forms.Button; $BtnSelect.Text = "Select User"; $BtnSelect.Location = "150, 325"; $BtnSelect.Size = "200, 30"; $BtnSelect.DialogResult = "OK"
    $PickForm.Controls.AddRange(@($TxtSearch, $BtnSearch, $ListResults, $BtnSelect)); $PickForm.AcceptButton = $BtnSearch

    $BtnSearch.Add_Click({
        $ListResults.Items.Clear()
        $Term = $TxtSearch.Text
        if ($Term.Length -gt 2) {
            try {
                $Users = Get-ADUser -Filter "anr -eq '$Term'" -Properties EmailAddress | Select Name, SamAccountName
                foreach ($U in $Users) { $ListResults.Items.Add("$($U.Name) [$($U.SamAccountName)]") | Out-Null }
            } catch { }
        }
    })
    if ($PickForm.ShowDialog() -eq "OK" -and $ListResults.SelectedItem) { return $ListResults.SelectedItem.ToString() }
    return $null
}

# --- RULE DIALOG ---
function Show-RuleDialog {
    param($ExistingRule = $null)
    $D = New-Object System.Windows.Forms.Form; $D.Text = if ($ExistingRule) { "Modify Rule" } else { "Add Rule" }
    $D.Size = "420, 220"; $D.StartPosition="CenterParent"
    
    $LblA = New-Object System.Windows.Forms.Label; $LblA.Text="Attribute:"; $LblA.Location="10,20"; $LblA.AutoSize=$true
    $CmbAttr = New-Object System.Windows.Forms.ComboBox; $CmbAttr.Location="10,40"; $CmbAttr.Size="380,25"; $CmbAttr.DropDownStyle="DropDownList"
    
    $Attrs = Invoke-Sql -Query "SELECT FriendlyName, ADAttribute FROM Sys_AttributeMap"
    if($Attrs.Success) { foreach($r in $Attrs.Data) { $CmbAttr.Items.Add("$($r.FriendlyName)|$($r.ADAttribute)")|Out-Null } }
    
    $LblO = New-Object System.Windows.Forms.Label; $LblO.Text="Operator:"; $LblO.Location="10,80"; $LblO.AutoSize=$true
    $CmbOp = New-Object System.Windows.Forms.ComboBox; $CmbOp.Location="10,100"; $CmbOp.Size="100,25"; $CmbOp.DropDownStyle="DropDownList"
    $CmbOp.Items.AddRange(@("eq", "ne", "like", "notlike")); $CmbOp.SelectedIndex=0

    $LblV = New-Object System.Windows.Forms.Label; $LblV.Text="Value:"; $LblV.Location="120,80"; $LblV.AutoSize=$true
    $TxtVal = New-Object System.Windows.Forms.TextBox; $TxtVal.Location="120,100"; $TxtVal.Size="270,25"
    
    $BtnOk = New-Object System.Windows.Forms.Button; $BtnOk.Text="OK"; $BtnOk.Location="140,140"; $BtnOk.Size="100,30"; $BtnOk.DialogResult="OK"
    $D.Controls.AddRange(@($LblA, $CmbAttr, $LblO, $CmbOp, $LblV, $TxtVal, $BtnOk)); $D.AcceptButton=$BtnOk

    if ($ExistingRule) {
        for($i=0; $i -lt $CmbAttr.Items.Count; $i++) {
            if ($CmbAttr.Items[$i].ToString().EndsWith("|$($ExistingRule.Attribute)")) { $CmbAttr.SelectedIndex = $i; break }
        }
        $CmbOp.SelectedItem = $ExistingRule.Operator; $TxtVal.Text = $ExistingRule.Value
    } elseif ($CmbAttr.Items.Count -gt 0) { $CmbAttr.SelectedIndex = 0 }

    if ($D.ShowDialog() -eq "OK") {
        return @{ Attribute = $CmbAttr.SelectedItem.ToString().Split('|')[1]; Friendly = $CmbAttr.SelectedItem.ToString().Split('|')[0]; Operator = $CmbOp.SelectedItem; Value = $TxtVal.Text }
    }
    return $null
}

# --- MAIN FORM ---
$Form = New-Object System.Windows.Forms.Form; $Form.Text = "Dynamic List Architect v2.5"; $Form.Size = "1000, 800"; $Form.StartPosition = "CenterScreen"
$TabControl = New-Object System.Windows.Forms.TabControl; $TabControl.Dock = "Fill"; $Form.Controls.Add($TabControl)
$TabGeneral = New-Object System.Windows.Forms.TabPage "1. Definition & Logic"; $TabExceptions = New-Object System.Windows.Forms.TabPage "2. Exceptions"
$TabControl.Controls.AddRange(@($TabGeneral, $TabExceptions))

# -- HEADER --
$PnlHeader = New-Object System.Windows.Forms.Panel; $PnlHeader.Location = "10, 10"; $PnlHeader.Size = "950, 110"; $PnlHeader.BorderStyle = "FixedSingle"; $TabGeneral.Controls.Add($PnlHeader)
$LblName = New-Object System.Windows.Forms.Label; $LblName.Text="List Name:"; $LblName.Location="10,15"; $LblName.AutoSize=$true
$TxtName = New-Object System.Windows.Forms.TextBox; $TxtName.Location="80,12"; $TxtName.Size="250,25"
$LblType = New-Object System.Windows.Forms.Label; $LblType.Text="Type:"; $LblType.Location="350,15"; $LblType.AutoSize=$true
$CmbType = New-Object System.Windows.Forms.ComboBox; $CmbType.Location="390,12"; $CmbType.Size="120,25"; $CmbType.DropDownStyle="DropDownList"; $CmbType.Items.AddRange(@("Security", "MailEnabled")); $CmbType.SelectedIndex=0
$LblSMTP = New-Object System.Windows.Forms.Label; $LblSMTP.Text="SMTP:"; $LblSMTP.Location="530,15"; $LblSMTP.AutoSize=$true
$TxtSMTP = New-Object System.Windows.Forms.TextBox; $TxtSMTP.Location="580,12"; $TxtSMTP.Size="200,25"; $TxtSMTP.Enabled=$false
$ChkActive = New-Object System.Windows.Forms.CheckBox; $ChkActive.Text="Sync Enabled"; $ChkActive.Location="820,12"; $ChkActive.Checked=$true

$LblDesc = New-Object System.Windows.Forms.Label; $LblDesc.Text="Desc:"; $LblDesc.Location="10,50"; $LblDesc.AutoSize=$true
$TxtDesc = New-Object System.Windows.Forms.TextBox; $TxtDesc.Location="80,47"; $TxtDesc.Size="430,25"
$LblOwner = New-Object System.Windows.Forms.Label; $LblOwner.Text="Owner:"; $LblOwner.Location="530,50"; $LblOwner.AutoSize=$true
$TxtOwner = New-Object System.Windows.Forms.TextBox; $TxtOwner.Location="580,47"; $TxtOwner.Size="150,25"; $TxtOwner.Text=$CurrentUser
$LblSync = New-Object System.Windows.Forms.Label; $LblSync.Text="Interval (Min):"; $LblSync.Location="740,50"; $LblSync.AutoSize=$true
$CmbSync = New-Object System.Windows.Forms.ComboBox; $CmbSync.Location="820,47"; $CmbSync.Size="100,25"; $CmbSync.DropDownStyle="DropDownList"
foreach($m in $SyncIntervals){ $CmbSync.Items.Add($m)|Out-Null }; $CmbSync.SelectedIndex=3
$PnlHeader.Controls.AddRange(@($LblName, $TxtName, $LblType, $CmbType, $LblSMTP, $TxtSMTP, $ChkActive, $LblDesc, $TxtDesc, $LblOwner, $TxtOwner, $LblSync, $CmbSync))

# -- LOGIC TREE --
$GrpLogic = New-Object System.Windows.Forms.GroupBox; $GrpLogic.Text="Logic Engine (Right-Click: Add/Delete | Double-Click: Edit)"; $GrpLogic.Location="10, 130"; $GrpLogic.Size="950, 300"; $TabGeneral.Controls.Add($GrpLogic)
$TreeLogic = New-Object System.Windows.Forms.TreeView; $TreeLogic.Dock="Fill"; $TreeLogic.Font=New-Object System.Drawing.Font("Segoe UI", 10); $GrpLogic.Controls.Add($TreeLogic)

$CtxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ItemEdit = $CtxMenu.Items.Add("✎ Edit / Toggle Gate"); $CtxMenu.Items.Add("-")
$ItemAddGroupAnd = $CtxMenu.Items.Add("Add Group (AND)"); $ItemAddGroupOr = $CtxMenu.Items.Add("Add Group (OR)"); $CtxMenu.Items.Add("-")
$ItemAddRule = $CtxMenu.Items.Add("Add Attribute Rule"); $CtxMenu.Items.Add("-")
$ItemDelete = $CtxMenu.Items.Add("Delete Item")
$TreeLogic.ContextMenuStrip = $CtxMenu

$RootNode = $TreeLogic.Nodes.Add("ROOT (AND)"); $RootNode.Tag = @{ Type = "GROUP"; Gate = "AND" }; $TreeLogic.SelectedNode = $RootNode

# -- PREVIEW SECTION --
$GrpPreview = New-Object System.Windows.Forms.GroupBox
$GrpPreview.Text = "Preview Results [ Total Match: 0 Users ]"
$GrpPreview.Location = "10, 440"
$GrpPreview.Size = "950, 260"
$GrpPreview.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)
$TabGeneral.Controls.Add($GrpPreview)

$PnlPrevTools = New-Object System.Windows.Forms.Panel; $PnlPrevTools.Dock="Top"; $PnlPrevTools.Height=35
$GrpPreview.Controls.Add($PnlPrevTools)

$BtnPreview = New-Object System.Windows.Forms.Button; $BtnPreview.Text="Run Logic & Refresh Preview"; $BtnPreview.Dock="Left"; $BtnPreview.Width=200; $BtnPreview.BackColor="LightBlue"
$ChkAutoPreview = New-Object System.Windows.Forms.CheckBox; $ChkAutoPreview.Text="Auto-Refresh on Change"; $ChkAutoPreview.Dock="Left"; $ChkAutoPreview.Width=160; $ChkAutoPreview.Checked=$true; $ChkAutoPreview.Padding=New-Object System.Windows.Forms.Padding(10,0,0,0)
$PnlPrevTools.Controls.AddRange(@($ChkAutoPreview, $BtnPreview))

$GridPreview = New-Object System.Windows.Forms.DataGridView; $GridPreview.Dock="Fill"; $GridPreview.AutoSizeColumnsMode="Fill"
$GrpPreview.Controls.Add($GridPreview)

# -- EXCEPTIONS TAB --
$SplitExc = New-Object System.Windows.Forms.SplitContainer; $SplitExc.Dock="Fill"; $SplitExc.Orientation="Vertical"; $TabExceptions.Controls.Add($SplitExc)
$GrpInc = New-Object System.Windows.Forms.GroupBox; $GrpInc.Text="Forced Inclusions"; $GrpInc.Dock="Fill"; $SplitExc.Panel1.Controls.Add($GrpInc)
$LstInc = New-Object System.Windows.Forms.ListBox; $LstInc.Dock="Fill"
$BtnIncAdd = New-Object System.Windows.Forms.Button; $BtnIncAdd.Text="+ Add"; $BtnIncAdd.Dock="Top"
$BtnIncRem = New-Object System.Windows.Forms.Button; $BtnIncRem.Text="- Remove"; $BtnIncRem.Dock="Bottom"
$GrpInc.Controls.AddRange(@($LstInc, $BtnIncAdd, $BtnIncRem))

$GrpExc = New-Object System.Windows.Forms.GroupBox; $GrpExc.Text="Forced Exclusions"; $GrpExc.Dock="Fill"; $SplitExc.Panel2.Controls.Add($GrpExc)
$LstExc = New-Object System.Windows.Forms.ListBox; $LstExc.Dock="Fill"
$BtnExcAdd = New-Object System.Windows.Forms.Button; $BtnExcAdd.Text="+ Block"; $BtnExcAdd.Dock="Top"
$BtnExcRem = New-Object System.Windows.Forms.Button; $BtnExcRem.Text="- Remove"; $BtnExcRem.Dock="Bottom"
$GrpExc.Controls.AddRange(@($LstExc, $BtnExcAdd, $BtnExcRem))

# -- BOTTOM ACTION --
$BtnSave = New-Object System.Windows.Forms.Button; $BtnSave.Text="💾 SAVE CONFIGURATION"; $BtnSave.Dock="Bottom"; $BtnSave.Height=40; $BtnSave.BackColor="LightGreen"; $Form.Controls.Add($BtnSave)


# ==========================================
# LOGIC: FILTER & PREVIEW
# ==========================================

function Get-TreeFilter($Node) {
    $Data = $Node.Tag
    if ($Data.Type -eq "GROUP") {
        $ChildFilters = @()
        foreach ($Child in $Node.Nodes) {
            $F = Get-TreeFilter $Child
            if (-not [string]::IsNullOrWhiteSpace($F)) { $ChildFilters += $F }
        }
        if ($ChildFilters.Count -eq 0) { return $null }
        $Gate = if($Data.Gate -eq "AND"){"-and"}else{"-or"}
        $Joined = $ChildFilters -join " $Gate "
        return "($Joined)"
    } else {
        $SafeVal = $Data.Value -replace "'", "''"
        return "($($Data.Attribute) -$($Data.Operator) '$SafeVal')"
    }
}

function Update-Preview {
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $GrpPreview.Text = "Preview Results [ Querying AD... ]"
    $Form.Refresh()

    try {
        $Filter = Get-TreeFilter $RootNode
        
        $BaseUsers = @()
        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            $BaseUsers = @(Get-ADUser -Filter $Filter -Properties EmailAddress, Title, Department -ErrorAction Stop | 
                           Select SamAccountName, Name, EmailAddress, Title)
        }

        $ExcSAMs = @(); foreach($x in $LstExc.Items){ $ExcSAMs += $x.ToString().Split('[')[1].Trim(']') }
        $IncSAMs = @(); foreach($x in $LstInc.Items){ $IncSAMs += $x.ToString().Split('[')[1].Trim(']') }

        $DisplayList = New-Object System.Collections.ArrayList
        
        foreach ($User in $BaseUsers) {
            if ($ExcSAMs -notcontains $User.SamAccountName) {
                $DisplayList.Add($User) | Out-Null
            }
        }

        foreach ($Sam in $IncSAMs) {
            $IsPresent = $false
            foreach ($Existing in $DisplayList) { if ($Existing.SamAccountName -eq $Sam) { $IsPresent=$true; break } }
            
            if (-not $IsPresent) {
                try {
                    $ManUser = Get-ADUser -Identity $Sam -Properties EmailAddress, Title | Select SamAccountName, Name, EmailAddress, Title
                    $DisplayList.Add($ManUser) | Out-Null
                } catch {}
            }
        }
        
        $GridPreview.DataSource = $null
        $GridPreview.DataSource = $DisplayList
        $GrpPreview.Text = "Preview Results [ Total Match: $($DisplayList.Count) Users ]"

    } catch {
        $GrpPreview.Text = "Preview Error: $($_.Exception.Message)"
    }
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
}

function Trigger-AutoRefresh {
    if ($ChkAutoPreview.Checked) { Update-Preview }
}

function Get-TreeJSON($Node) {
    $Data = $Node.Tag
    if ($Data.Type -eq "GROUP") {
        $Children = @()
        foreach ($Child in $Node.Nodes) { $Children += Get-TreeJSON $Child }
        return @{ type="group"; gate=$Data.Gate; children=$Children }
    } else {
        return @{ type="rule"; attribute=$Data.Attribute; operator=$Data.Operator; value=$Data.Value }
    }
}

# ==========================================
# EVENT HANDLERS
# ==========================================

# Tree Changes
$ItemAddGroupAnd.Add_Click({ 
    $N = $TreeLogic.SelectedNode.Nodes.Add("Group (AND)"); $N.Tag=@{Type="GROUP"; Gate="AND"}; 
    $TreeLogic.SelectedNode.Expand(); Trigger-AutoRefresh 
})
$ItemAddGroupOr.Add_Click({ 
    $N = $TreeLogic.SelectedNode.Nodes.Add("Group (OR)"); $N.Tag=@{Type="GROUP"; Gate="OR"}; 
    $TreeLogic.SelectedNode.Expand(); Trigger-AutoRefresh 
})
$ItemAddRule.Add_Click({
    $Rule = Show-RuleDialog
    if ($Rule) {
        $N = $TreeLogic.SelectedNode.Nodes.Add("$($Rule.Friendly) $($Rule.Operator) '$($Rule.Value)'")
        $N.Tag = @{Type="RULE"; Attribute=$Rule.Attribute; Operator=$Rule.Operator; Value=$Rule.Value}
        $TreeLogic.SelectedNode.Expand(); Trigger-AutoRefresh
    }
})

$EditHandler = {
    $Node = $TreeLogic.SelectedNode
    $Data = $Node.Tag
    $Changed = $false

    if ($Data.Type -eq "GROUP") {
        $NewGate = if ($Data.Gate -eq "AND") { "OR" } else { "AND" }
        $Node.Tag.Gate = $NewGate
        if ($Node.Text -match "ROOT") { $Node.Text = "ROOT ($NewGate)" } else { $Node.Text = "Group ($NewGate)" }
        $Changed = $true
    }
    elseif ($Data.Type -eq "RULE") {
        $NewRule = Show-RuleDialog -ExistingRule $Data
        if ($NewRule) {
            $Node.Text = "$($NewRule.Friendly) $($NewRule.Operator) '$($NewRule.Value)'"
            $Node.Tag = @{Type="RULE"; Attribute=$NewRule.Attribute; Operator=$NewRule.Operator; Value=$NewRule.Value}
            $Changed = $true
        }
    }
    if ($Changed) { Trigger-AutoRefresh }
}

$ItemEdit.Add_Click($EditHandler)
$TreeLogic.Add_NodeMouseDoubleClick({ param($s, $e) $EditHandler.Invoke() })
$ItemDelete.Add_Click({ if ($TreeLogic.SelectedNode -ne $RootNode) { $TreeLogic.SelectedNode.Remove(); Trigger-AutoRefresh } })

# Exception Changes (Trigger Refresh)
$BtnIncAdd.Add_Click({ $U = Show-UserPicker; if($U){ $LstInc.Items.Add($U); Trigger-AutoRefresh } })
$BtnExcAdd.Add_Click({ $U = Show-UserPicker; if($U){ $LstExc.Items.Add($U); Trigger-AutoRefresh } })
$BtnIncRem.Add_Click({ if($LstInc.SelectedIndex -ge 0){$LstInc.Items.RemoveAt($LstInc.SelectedIndex); Trigger-AutoRefresh } })
$BtnExcRem.Add_Click({ if($LstExc.SelectedIndex -ge 0){$LstExc.Items.RemoveAt($LstExc.SelectedIndex); Trigger-AutoRefresh } })

# Manual Preview
$BtnPreview.Add_Click({ Update-Preview })

# Save Logic (FIXED)
$BtnSave.Add_Click({
    if (-not $TxtName.Text) { [System.Windows.Forms.MessageBox]::Show("Name required"); return }
    
    # 1. Prepare Data
    $JsonTree = Get-TreeJSON $RootNode | ConvertTo-Json -Depth 10 -Compress
    
    # TYPE SAFETY FIX: Explicit casting to prevent hashtable confusion
    $Params = @{
        Name  = $TxtName.Text
        Desc  = $TxtDesc.Text
        Type  = $CmbType.SelectedItem.ToString()
        SMTP  = if ($CmbType.SelectedItem -eq "MailEnabled" -and $TxtSMTP.Text) { $TxtSMTP.Text } else { [DBNull]::Value }
        Rule  = $JsonTree
        Own   = $TxtOwner.Text
        Sync  = [int]$CmbSync.SelectedItem # FORCE INT
        Act   = [bool]$ChkActive.Checked   # FORCE BOOL
    }

    # 2. Check Exists
    $QCheck = "SELECT DistListID FROM DL_MasterList WHERE Name = @Name"
    $CheckRes = Invoke-Sql -Query $QCheck -Parameters @{Name=$TxtName.Text} -ReturnScalar
    $ID = $CheckRes.Data
    
    if ($ID) {
        # UPDATE
        $Params.Add("ID", $ID)
        $Q = "UPDATE DL_MasterList SET Description=@Desc, ListType=@Type, PrimarySMTPAddress=@SMTP, RuleDefinition=@Rule, ManagedBy=@Own, SyncIntervalMinutes=@Sync, SyncEnabled=@Act WHERE DistListID=@ID"
        Invoke-Sql -Query $Q -Parameters $Params
    } else {
        # INSERT
        $Q = "INSERT INTO DL_MasterList (Name, Description, ListType, PrimarySMTPAddress, RuleDefinition, ManagedBy, SyncIntervalMinutes, SyncEnabled) VALUES (@Name, @Desc, @Type, @SMTP, @Rule, @Own, @Sync, @Act); SELECT SCOPE_IDENTITY()"
        $Res = Invoke-Sql -Query $Q -Parameters $Params -ReturnId
        $ID = $Res.Data
    }

    # 3. Save Exceptions
    Invoke-Sql -Query "DELETE FROM DL_Inclusions WHERE DistListID=@ID; DELETE FROM DL_Exclusions WHERE DistListID=@ID" -Parameters @{ID=$ID}
    
    foreach ($Item in $LstInc.Items) { 
        Invoke-Sql -Query "INSERT INTO DL_Inclusions (DistListID, UserIdentity) VALUES (@ID, @Uid)" -Parameters @{ID=$ID; Uid=$Item.ToString()} 
    }
    foreach ($Item in $LstExc.Items) { 
        Invoke-Sql -Query "INSERT INTO DL_Exclusions (DistListID, UserIdentity) VALUES (@ID, @Uid)" -Parameters @{ID=$ID; Uid=$Item.ToString()} 
    }

    [System.Windows.Forms.MessageBox]::Show("Saved successfully! ID: $ID")
})

$CmbType.Add_SelectedIndexChanged({ $TxtSMTP.Enabled = ($CmbType.SelectedItem -eq "MailEnabled") })
$Form.ShowDialog()
'@
$EmbeddedScript_ManageList = @'
<# 
   .SYNOPSIS
   Dynamic List Editor (v5.6) - Full Audit Fix
   .DESCRIPTION
   Ensures 'NewValue' in the audit log contains the actual JSON content for Exceptions, not placeholders.
#>

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing

# ==========================================
# CONFIGURATION
# ==========================================
$SqlInstance = if ($env:DYN_SQL_INSTANCE) { $env:DYN_SQL_INSTANCE } else { "SQL16-DIVERS" }
$Database    = if ($env:DYN_SQL_DATABASE) { $env:DYN_SQL_DATABASE } else { "DBTEST" }
# ==========================================

$CurrentUser   = $env:USERNAME
$CurrentListID = 0
$SyncIntervals = @(5, 15, 30, 60, 120, 240, 480, 1440)
$HasAuditCol   = $false 

# --- SQL HELPER ---
function Invoke-Sql {
    param([Parameter(Mandatory=$true)][string]$Query, [hashtable]$Parameters = @{}, [switch]$ReturnScalar, [switch]$ReturnId, [switch]$NoErrorUI)
    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    $Conn = $null
    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;Connect Timeout=5;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString); $Conn.Open()
        $Cmd = $Conn.CreateCommand(); $Cmd.CommandText = $Query; $Cmd.CommandTimeout = 5 
        if ($Parameters) { foreach ($Key in $Parameters.Keys) { $Val = if ($Parameters[$Key] -eq $null) { [DBNull]::Value } else { $Parameters[$Key] }; $Cmd.Parameters.AddWithValue("@$Key", $Val) | Out-Null } }
        if ($ReturnScalar) { $ResultInfo.Data = $Cmd.ExecuteScalar() } elseif ($ReturnId) { $ResultInfo.Data = $Cmd.ExecuteScalar() }
        else { $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd; $DS = New-Object System.Data.DataSet; $Adapter.Fill($DS) | Out-Null; $ResultInfo.Data = $DS.Tables[0] }
        $ResultInfo.Success = $true
    }
    catch { $ResultInfo.Message = $_.Exception.Message; if (-not $NoErrorUI) { [System.Windows.Forms.MessageBox]::Show("SQL Error: $($_.Exception.Message)", "Error", 0, 16) } }
    finally { if ($Conn) { $Conn.Close(); $Conn.Dispose() } }
    return $ResultInfo
}

# --- CHECK SCHEMA ---
function Check-AuditSchema {
    $Res = Invoke-Sql "SELECT COUNT(*) FROM sys.columns WHERE object_id = OBJECT_ID('Log_ConfigHistory') AND name = 'ChangeDescription'" -ReturnScalar -NoErrorUI
    if ($Res.Success -and $Res.Data -eq 1) { return $true }
    return $false
}

# ==========================================
# UI HELPERS
# ==========================================
function Show-InputBox {
    param($Title, $Prompt)
    $F = New-Object System.Windows.Forms.Form; $F.Text=$Title; $F.Size="400, 220"; $F.StartPosition="CenterParent"; $F.FormBorderStyle="FixedDialog"; $F.MaximizeBox=$false
    $L = New-Object System.Windows.Forms.Label; $L.Text=$Prompt; $L.Location="10, 15"; $L.AutoSize=$true; $F.Controls.Add($L)
    $T = New-Object System.Windows.Forms.TextBox; $T.Location="10, 40"; $T.Size="365, 80"; $T.Multiline=$true; $F.Controls.Add($T)
    $BtnOK = New-Object System.Windows.Forms.Button; $BtnOK.Text="Commit"; $BtnOK.Location="255, 130"; $BtnOK.DialogResult="OK"; $BtnOK.BackColor="LightGreen"; $F.Controls.Add($BtnOK)
    $BtnCancel = New-Object System.Windows.Forms.Button; $BtnCancel.Text="Cancel"; $BtnCancel.Location="165, 130"; $BtnCancel.DialogResult="Cancel"; $F.Controls.Add($BtnCancel)
    $F.AcceptButton=$BtnOK; $F.CancelButton=$BtnCancel
    if ($F.ShowDialog() -eq "OK") { return $T.Text }
    return $null
}

function Show-UserPicker {
    $PickForm = New-Object System.Windows.Forms.Form; $PickForm.Text="Find User"; $PickForm.Size="500, 400"; $PickForm.StartPosition="CenterParent"
    $TxtSearch = New-Object System.Windows.Forms.TextBox; $TxtSearch.Location="10, 10"; $TxtSearch.Size="380, 25"
    $BtnSearch = New-Object System.Windows.Forms.Button; $BtnSearch.Text="Search"; $BtnSearch.Location="400, 9"; $BtnSearch.Size="70, 26"
    $ListResults = New-Object System.Windows.Forms.ListBox; $ListResults.Location="10, 45"; $ListResults.Size="460, 270"
    $BtnSelect = New-Object System.Windows.Forms.Button; $BtnSelect.Text="Select User"; $BtnSelect.Location="150, 325"; $BtnSelect.Size="200, 30"; $BtnSelect.DialogResult="OK"
    $PickForm.Controls.AddRange(@($TxtSearch, $BtnSearch, $ListResults, $BtnSelect)); $PickForm.AcceptButton = $BtnSearch
    $BtnSearch.Add_Click({ $ListResults.Items.Clear(); $Term = $TxtSearch.Text; if ($Term.Length -gt 2) { try { $Users = Get-ADUser -Filter "anr -eq '$Term'" -Properties EmailAddress | Select Name, SamAccountName; foreach ($U in $Users) { $ListResults.Items.Add("$($U.Name) [$($U.SamAccountName)]") | Out-Null } } catch {} } })
    if ($PickForm.ShowDialog() -eq "OK" -and $ListResults.SelectedItem) { return $ListResults.SelectedItem.ToString() }
    return $null
}

function Show-ListPicker {
    $P = New-Object System.Windows.Forms.Form; $P.Text = "Select List to Edit"; $P.Size = "900, 500"; $P.StartPosition="CenterScreen"
    $G = New-Object System.Windows.Forms.DataGridView; $G.Dock="Fill"; $G.SelectionMode="FullRowSelect"; $G.MultiSelect=$false; $G.ReadOnly=$true; $G.AutoSizeColumnsMode="Fill"
    $P.Controls.Add($G)
    $Res = Invoke-Sql "SELECT DistListID, Name, Description, ListType, LockedBy, LastSyncDate FROM DL_MasterList WITH (NOLOCK)"
    if ($Res.Success) { $G.DataSource = $Res.Data }
    $BtnOpen = New-Object System.Windows.Forms.Button; $BtnOpen.Text="🔓 CHECK OUT & EDIT"; $BtnOpen.Dock="Bottom"; $BtnOpen.Height=50; $BtnOpen.BackColor="LightBlue"; $BtnOpen.DialogResult="OK"; $P.Controls.Add($BtnOpen)
    if ($P.ShowDialog() -eq "OK" -and $G.SelectedRows.Count -gt 0) { return $G.SelectedRows[0] }
    return $null
}

function Show-RuleDialog {
    param($ExistingRule = $null)
    $D = New-Object System.Windows.Forms.Form; $D.Text = if ($ExistingRule) { "Modify Rule" } else { "Add Rule" }
    $D.Size = "450, 250"; $D.StartPosition="CenterParent"
    $LblA = New-Object System.Windows.Forms.Label; $LblA.Text="Attribute:"; $LblA.Location="10,20"; $LblA.AutoSize=$true
    $CmbAttr = New-Object System.Windows.Forms.ComboBox; $CmbAttr.Location="10,40"; $CmbAttr.Size="400,25"; $CmbAttr.DropDownStyle="DropDownList"
    $Attrs = Invoke-Sql "SELECT FriendlyName, ADAttribute FROM Sys_AttributeMap WITH (NOLOCK)"
    if($Attrs.Success) { foreach($r in $Attrs.Data) { $CmbAttr.Items.Add("$($r.FriendlyName)|$($r.ADAttribute)")|Out-Null } }
    $LblO = New-Object System.Windows.Forms.Label; $LblO.Text="Operator:"; $LblO.Location="10,80"; $LblO.AutoSize=$true
    $CmbOp = New-Object System.Windows.Forms.ComboBox; $CmbOp.Location="10,100"; $CmbOp.Size="120,25"; $CmbOp.DropDownStyle="DropDownList"
    $CmbOp.Items.AddRange(@("eq", "ne", "like", "notlike")); $CmbOp.SelectedIndex=0
    $LblV = New-Object System.Windows.Forms.Label; $LblV.Text="Value:"; $LblV.Location="140,80"; $LblV.AutoSize=$true
    $TxtVal = New-Object System.Windows.Forms.TextBox; $TxtVal.Location="140,100"; $TxtVal.Size="270,25"
    $BtnOk = New-Object System.Windows.Forms.Button; $BtnOk.Text="OK"; $BtnOk.Location="150,150"; $BtnOk.Size="120,35"; $BtnOk.DialogResult="OK"
    $D.Controls.AddRange(@($LblA, $CmbAttr, $LblO, $CmbOp, $LblV, $TxtVal, $BtnOk)); $D.AcceptButton=$BtnOk
    if ($ExistingRule) { for($i=0; $i -lt $CmbAttr.Items.Count; $i++) { if ($CmbAttr.Items[$i].ToString().EndsWith("|$($ExistingRule.Attribute)")) { $CmbAttr.SelectedIndex = $i; break } }; $CmbOp.SelectedItem = $ExistingRule.Operator; $TxtVal.Text = $ExistingRule.Value } elseif ($CmbAttr.Items.Count -gt 0) { $CmbAttr.SelectedIndex = 0 }
    if ($D.ShowDialog() -eq "OK") { return @{ Attribute = $CmbAttr.SelectedItem.ToString().Split('|')[1]; Friendly = $CmbAttr.SelectedItem.ToString().Split('|')[0]; Operator = $CmbOp.SelectedItem; Value = $TxtVal.Text } }
    return $null
}

# ==========================================
# MAIN FORM UI
# ==========================================
$Form = New-Object System.Windows.Forms.Form; $Form.Text = "List Editor (v5.6)"; $Form.Size = "1000, 850"; $Form.StartPosition = "CenterScreen"
$StatusStrip = New-Object System.Windows.Forms.StatusStrip; $StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $StatusLabel.Text = "Ready"; $StatusStrip.Items.Add($StatusLabel); $Form.Controls.Add($StatusStrip)
$TabControl = New-Object System.Windows.Forms.TabControl; $TabControl.Dock = "Fill"; $Form.Controls.Add($TabControl)
$TabGeneral = New-Object System.Windows.Forms.TabPage "1. Logic Definition"; $TabExceptions = New-Object System.Windows.Forms.TabPage "2. Exceptions (In/Out)"
$TabControl.Controls.AddRange(@($TabGeneral, $TabExceptions))

# Header
$PnlHeader = New-Object System.Windows.Forms.Panel; $PnlHeader.Location = "10, 10"; $PnlHeader.Size = "950, 110"; $PnlHeader.BorderStyle = "FixedSingle"; $TabGeneral.Controls.Add($PnlHeader)
$LblName = New-Object System.Windows.Forms.Label; $LblName.Text="List Name:"; $LblName.Location="10,15"; $LblName.AutoSize=$true; $PnlHeader.Controls.Add($LblName)
$TxtName = New-Object System.Windows.Forms.TextBox; $TxtName.Location="120,12"; $TxtName.Size="210,25"; $TxtName.ReadOnly=$true; $TxtName.BackColor="WhiteSmoke"; $PnlHeader.Controls.Add($TxtName)
$LblType = New-Object System.Windows.Forms.Label; $LblType.Text="Type:"; $LblType.Location="350,15"; $LblType.AutoSize=$true; $PnlHeader.Controls.Add($LblType)
$CmbType = New-Object System.Windows.Forms.ComboBox; $CmbType.Location="390,12"; $CmbType.Size="120,25"; $CmbType.DropDownStyle="DropDownList"; $CmbType.Items.AddRange(@("Security", "MailEnabled")); $CmbType.SelectedIndex=0; $PnlHeader.Controls.Add($CmbType)
$LblSMTP = New-Object System.Windows.Forms.Label; $LblSMTP.Text="SMTP:"; $LblSMTP.Location="530,15"; $LblSMTP.AutoSize=$true; $PnlHeader.Controls.Add($LblSMTP)
$TxtSMTP = New-Object System.Windows.Forms.TextBox; $TxtSMTP.Location="580,12"; $TxtSMTP.Size="200,25"; $TxtSMTP.Enabled=$false; $PnlHeader.Controls.Add($TxtSMTP)
$ChkActive = New-Object System.Windows.Forms.CheckBox; $ChkActive.Text="Sync Enabled"; $ChkActive.Location="820,12"; $ChkActive.Checked=$true; $PnlHeader.Controls.Add($ChkActive)
$LblDesc = New-Object System.Windows.Forms.Label; $LblDesc.Text="Desc:"; $LblDesc.Location="10,50"; $LblDesc.AutoSize=$true; $PnlHeader.Controls.Add($LblDesc)
$TxtDesc = New-Object System.Windows.Forms.TextBox; $TxtDesc.Location="120,47"; $TxtDesc.Size="390,25"; $PnlHeader.Controls.Add($TxtDesc)
$LblOwner = New-Object System.Windows.Forms.Label; $LblOwner.Text="Owner:"; $LblOwner.Location="530,50"; $LblOwner.AutoSize=$true; $PnlHeader.Controls.Add($LblOwner)
$TxtOwner = New-Object System.Windows.Forms.TextBox; $TxtOwner.Location="580,47"; $TxtOwner.Size="150,25"; $TxtOwner.Text=$CurrentUser; $PnlHeader.Controls.Add($TxtOwner)
$LblSync = New-Object System.Windows.Forms.Label; $LblSync.Text="Interval (Min):"; $LblSync.Location="740,50"; $LblSync.AutoSize=$true; $PnlHeader.Controls.Add($LblSync)
$CmbSync = New-Object System.Windows.Forms.ComboBox; $CmbSync.Location="820,47"; $CmbSync.Size="100,25"; $CmbSync.DropDownStyle="DropDownList"; foreach($m in $SyncIntervals){ $CmbSync.Items.Add($m)|Out-Null }; $CmbSync.SelectedIndex=3; $PnlHeader.Controls.Add($CmbSync)

# Logic
$GrpLogic = New-Object System.Windows.Forms.GroupBox; $GrpLogic.Text="Logic Rules"; $GrpLogic.Location="10, 130"; $GrpLogic.Size="950, 300"; $TabGeneral.Controls.Add($GrpLogic)
$TreeLogic = New-Object System.Windows.Forms.TreeView; $TreeLogic.Dock="Fill"; $TreeLogic.Font=New-Object System.Drawing.Font("Segoe UI", 10); $GrpLogic.Controls.Add($TreeLogic)
$CtxMenu = New-Object System.Windows.Forms.ContextMenuStrip; $ItemEdit = $CtxMenu.Items.Add("Edit / Toggle"); $CtxMenu.Items.Add("-"); $ItemAddGroupAnd = $CtxMenu.Items.Add("Add Group (AND)"); $ItemAddGroupOr = $CtxMenu.Items.Add("Add Group (OR)"); $CtxMenu.Items.Add("-"); $ItemAddRule = $CtxMenu.Items.Add("Add Rule"); $CtxMenu.Items.Add("-"); $ItemDelete = $CtxMenu.Items.Add("Delete"); $TreeLogic.ContextMenuStrip = $CtxMenu

# Preview
$GrpPreview = New-Object System.Windows.Forms.GroupBox; $GrpPreview.Text = "Preview"; $GrpPreview.Location = "10, 440"; $GrpPreview.Size = "950, 260"; $GrpPreview.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10); $TabGeneral.Controls.Add($GrpPreview)
$PnlPrevTools = New-Object System.Windows.Forms.Panel; $PnlPrevTools.Dock="Top"; $PnlPrevTools.Height=35; $GrpPreview.Controls.Add($PnlPrevTools)
$BtnPreview = New-Object System.Windows.Forms.Button; $BtnPreview.Text="Refresh Preview"; $BtnPreview.Dock="Left"; $BtnPreview.Width=200; $BtnPreview.BackColor="LightBlue"
$ChkAutoPreview = New-Object System.Windows.Forms.CheckBox; $ChkAutoPreview.Text="Auto-Refresh"; $ChkAutoPreview.Dock="Left"; $ChkAutoPreview.Width=160; $ChkAutoPreview.Checked=$true; $PnlPrevTools.Controls.AddRange(@($ChkAutoPreview, $BtnPreview))
$GridPreview = New-Object System.Windows.Forms.DataGridView; $GridPreview.Dock="Fill"; $GridPreview.AutoSizeColumnsMode="Fill"; $GrpPreview.Controls.Add($GridPreview)

# Exceptions
$SplitExc = New-Object System.Windows.Forms.SplitContainer; $SplitExc.Dock="Fill"; $SplitExc.Orientation="Vertical"; $TabExceptions.Controls.Add($SplitExc)
$GrpInc = New-Object System.Windows.Forms.GroupBox; $GrpInc.Text="Forced Inclusions"; $GrpInc.Dock="Fill"; $SplitExc.Panel1.Controls.Add($GrpInc)
$LstInc = New-Object System.Windows.Forms.ListBox; $LstInc.Dock="Fill"; $BtnIncAdd = New-Object System.Windows.Forms.Button; $BtnIncAdd.Text="+ Add"; $BtnIncAdd.Dock="Top"; $BtnIncRem = New-Object System.Windows.Forms.Button; $BtnIncRem.Text="- Remove"; $BtnIncRem.Dock="Bottom"; $GrpInc.Controls.AddRange(@($LstInc, $BtnIncAdd, $BtnIncRem))
$GrpExc = New-Object System.Windows.Forms.GroupBox; $GrpExc.Text="Forced Exclusions"; $GrpExc.Dock="Fill"; $SplitExc.Panel2.Controls.Add($GrpExc)
$LstExc = New-Object System.Windows.Forms.ListBox; $LstExc.Dock="Fill"; $BtnExcAdd = New-Object System.Windows.Forms.Button; $BtnExcAdd.Text="+ Block"; $BtnExcAdd.Dock="Top"; $BtnExcRem = New-Object System.Windows.Forms.Button; $BtnExcRem.Text="- Remove"; $BtnExcRem.Dock="Bottom"; $GrpExc.Controls.AddRange(@($LstExc, $BtnExcAdd, $BtnExcRem))

$BtnSave = New-Object System.Windows.Forms.Button; $BtnSave.Text="💾 SAVE CHANGES"; $BtnSave.Dock="Bottom"; $BtnSave.Height=50; $BtnSave.BackColor="LightGreen"; $BtnSave.Font=New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $Form.Controls.Add($BtnSave)

# ==========================================
# LOGIC
# ==========================================

function Rebuild-Tree($Json) {
    $TreeLogic.Nodes.Clear()
    if ([string]::IsNullOrWhiteSpace($Json)) { $Root=$TreeLogic.Nodes.Add("ROOT (AND)"); $Root.Tag=@{Type="GROUP"; Gate="AND"}; $TreeLogic.SelectedNode=$Root; return }
    try {
        $Data = $Json | ConvertFrom-Json
        $Root = $TreeLogic.Nodes.Add("ROOT ($($Data.gate))"); $Root.Tag = @{Type="GROUP"; Gate=$Data.gate}; $TreeLogic.SelectedNode = $Root
        if ($Data.children) { foreach ($child in $Data.children) { Render-Node -ParentNode $Root -Data $child } }
        $TreeLogic.ExpandAll()
    } catch { $Root=$TreeLogic.Nodes.Add("ROOT (AND)"); $Root.Tag=@{Type="GROUP"; Gate="AND"} }
}

function Render-Node($ParentNode, $Data) {
    if ($Data.type -eq "group") {
        $N = $ParentNode.Nodes.Add("Group ($($Data.gate))"); $N.Tag = @{Type="GROUP"; Gate=$Data.gate}
        foreach ($c in $Data.children) { Render-Node -ParentNode $N -Data $c }
    } else {
        $N = $ParentNode.Nodes.Add("$($Data.attribute) $($Data.operator) '$($Data.value)'")
        $N.Tag = @{Type="RULE"; Attribute=$Data.attribute; Operator=$Data.operator; Value=$Data.value}
    }
}

function Get-TreeJSON($Node) {
    $Data = $Node.Tag
    if ($Data.Type -eq "GROUP") {
        $Children = @(); foreach ($Child in $Node.Nodes) { $Children += Get-TreeJSON $Child }
        return @{ type="group"; gate=$Data.Gate; children=$Children }
    } else { return @{ type="rule"; attribute=$Data.Attribute; operator=$Data.Operator; value=$Data.Value } }
}

# --- SAVE LOGIC (CORRECTED AUDIT) ---
$BtnSave.Add_Click({
    if ($CurrentListID -eq 0) { return }
    
    $Desc = ""
    if ($HasAuditCol) {
        $Desc = Show-InputBox -Title "Commit Update" -Prompt "Please describe changes (Required):"
        if ([string]::IsNullOrWhiteSpace($Desc)) { [System.Windows.Forms.MessageBox]::Show("Save Cancelled: Description required."); return }
    }
    
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $StatusLabel.Text = "Step 1/5: Serializing..."
        $Form.Refresh()
        
        # Rules JSON
        $NodeRoot = $TreeLogic.Nodes[0]
        $NewJson = Get-TreeJSON $NodeRoot | ConvertTo-Json -Depth 10 -Compress

        # Exceptions JSON (New State from UI)
        $NewIncList = @(); foreach($x in $LstInc.Items){ $NewIncList += $x.ToString() }
        $NewExcList = @(); foreach($x in $LstExc.Items){ $NewExcList += $x.ToString() }
        $NewUnifiedObj = @{ Inclusions = $NewIncList; Exclusions = $NewExcList }
        $NewExcJson = $NewUnifiedObj | ConvertTo-Json -Compress -Depth 5

        $StatusLabel.Text = "Step 2/5: Fetching Backup..."
        $Form.Refresh()
        
        # Get Old Rule
        $OldRuleData = (Invoke-Sql "SELECT RuleDefinition FROM DL_MasterList WITH (NOLOCK) WHERE DistListID=$CurrentListID").Data
        $OldRule = if ($OldRuleData) { $OldRuleData.RuleDefinition } else { "" }
        
        # Get Old Exceptions (From DB)
        $RawInc = (Invoke-Sql "SELECT UserIdentity FROM DL_Inclusions WITH (NOLOCK) WHERE DistListID=$CurrentListID").Data
        $RawExc = (Invoke-Sql "SELECT UserIdentity FROM DL_Exclusions WITH (NOLOCK) WHERE DistListID=$CurrentListID").Data
        
        $ListInc = @(); if($RawInc){ foreach($r in $RawInc){ $ListInc += $r.UserIdentity } }
        $ListExc = @(); if($RawExc){ foreach($r in $RawExc){ $ListExc += $r.UserIdentity } }
        
        $OldUnifiedObj = @{ Inclusions = $ListInc; Exclusions = $ListExc }
        $OldExcJson = $OldUnifiedObj | ConvertTo-Json -Compress -Depth 5

        # STEP 3: Write Logs
        $StatusLabel.Text = "Step 3/5: Writing Logs..."
        $Form.Refresh()
        
        $LogP = @{ ID=$CurrentListID; User=$CurrentUser; OldR=$OldRule; NewR=$NewJson; OldE=$OldExcJson; NewE=$NewExcJson; Txt=$Desc }
        
        if ($HasAuditCol) {
            Invoke-Sql "INSERT INTO Log_ConfigHistory (DistListID, AdminUser, ModifiedTable, ChangeType, OldValue, NewValue, ChangeDescription) VALUES (@ID, @User, 'DL_MasterList', 'UPDATE', @OldR, @NewR, @Txt)" -Parameters $LogP
            # FIXED: Uses @NewE instead of text
            Invoke-Sql "INSERT INTO Log_ConfigHistory (DistListID, AdminUser, ModifiedTable, ChangeType, OldValue, NewValue, ChangeDescription) VALUES (@ID, @User, 'DL_Exceptions_Combined', 'UPDATE', @OldE, @NewE, @Txt)" -Parameters $LogP
        } else {
            Invoke-Sql "INSERT INTO Log_ConfigHistory (DistListID, AdminUser, ModifiedTable, ChangeType, OldValue, NewValue) VALUES (@ID, @User, 'DL_MasterList', 'UPDATE', @OldR, @NewR)" -Parameters $LogP
        }

        # STEP 4: Update Master
        $StatusLabel.Text = "Step 4/5: Updating Master..."
        $Form.Refresh()
        
        $UpdateP = @{
            ID=$CurrentListID; Desc=$TxtDesc.Text; Type=$CmbType.SelectedItem.ToString();
            SMTP=if($CmbType.SelectedItem -eq "MailEnabled" -and $TxtSMTP.Text){$TxtSMTP.Text}else{[DBNull]::Value};
            Rule=$NewJson; Own=$TxtOwner.Text; Sync=[int]$CmbSync.SelectedItem; Act=[bool]$ChkActive.Checked
        }
        Invoke-Sql "UPDATE DL_MasterList SET Description=@Desc, ListType=@Type, PrimarySMTPAddress=@SMTP, RuleDefinition=@Rule, ManagedBy=@Own, SyncIntervalMinutes=@Sync, SyncEnabled=@Act WHERE DistListID=@ID" -Parameters $UpdateP

        # STEP 5: Update Exceptions (BULK)
        $StatusLabel.Text = "Step 5/5: Saving Exceptions..."
        $Form.Refresh()
        
        Invoke-Sql "DELETE FROM DL_Inclusions WHERE DistListID=$CurrentListID; DELETE FROM DL_Exclusions WHERE DistListID=$CurrentListID"
        
        if ($LstInc.Items.Count -gt 0) {
            $Values = New-Object System.Collections.ArrayList
            foreach ($Item in $LstInc.Items) { $Safe = $Item.ToString().Replace("'", "''"); $Values.Add("($CurrentListID, '$Safe')") | Out-Null }
            for ($i=0; $i -lt $Values.Count; $i+=1000) { $Batch = ($Values | Select-Object -Skip $i -First 1000) -join ","; Invoke-Sql "INSERT INTO DL_Inclusions (DistListID, UserIdentity) VALUES $Batch" }
        }
        if ($LstExc.Items.Count -gt 0) {
            $Values = New-Object System.Collections.ArrayList
            foreach ($Item in $LstExc.Items) { $Safe = $Item.ToString().Replace("'", "''"); $Values.Add("($CurrentListID, '$Safe')") | Out-Null }
            for ($i=0; $i -lt $Values.Count; $i+=1000) { $Batch = ($Values | Select-Object -Skip $i -First 1000) -join ","; Invoke-Sql "INSERT INTO DL_Exclusions (DistListID, UserIdentity) VALUES $Batch" }
        }

        [System.Windows.Forms.MessageBox]::Show("Saved Successfully!")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_")
    } finally {
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $StatusLabel.Text = "Ready"
    }
})

# PREVIEW
function Get-TreeFilter($Node) {
    $Data = $Node.Tag
    if ($Data.Type -eq "GROUP") {
        $ChildFilters = @(); foreach ($Child in $Node.Nodes) { $F = Get-TreeFilter $Child; if (-not [string]::IsNullOrWhiteSpace($F)) { $ChildFilters += $F } }
        if ($ChildFilters.Count -eq 0) { return $null }
        $Gate = if($Data.Gate -eq "AND"){"-and"}else{"-or"}; $Joined = $ChildFilters -join " $Gate "; return "($Joined)"
    } else { $SafeVal = $Data.Value -replace "'", "''"; return "($($Data.Attribute) -$($Data.Operator) '$SafeVal')" }
}

function Update-Preview {
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $GrpPreview.Text = "Preview [ Querying... ]"; $Form.Refresh()
    try {
        if ($TreeLogic.Nodes.Count -eq 0) { $GrpPreview.Text = "Preview [ Empty Tree ]"; return }
        $Filter = Get-TreeFilter $TreeLogic.Nodes[0]
        $BaseUsers = @()
        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            $BaseUsers = @(Get-ADUser -Filter $Filter -Properties EmailAddress, Title, Department -ErrorAction Stop | Select SamAccountName, Name, EmailAddress, Title)
        }
        $ExcSAMs = @(); foreach($x in $LstExc.Items){ $ExcSAMs += $x.ToString().Split('[')[1].Trim(']') }
        $IncSAMs = @(); foreach($x in $LstInc.Items){ $IncSAMs += $x.ToString().Split('[')[1].Trim(']') }
        $DisplayList = New-Object System.Collections.ArrayList
        foreach ($User in $BaseUsers) { if ($ExcSAMs -notcontains $User.SamAccountName) { $DisplayList.Add($User) | Out-Null } }
        foreach ($Sam in $IncSAMs) {
            $IsPresent = $false; foreach ($Existing in $DisplayList) { if ($Existing.SamAccountName -eq $Sam) { $IsPresent=$true; break } }
            if (-not $IsPresent) { try { $DisplayList.Add((Get-ADUser -Identity $Sam -Properties EmailAddress, Title | Select SamAccountName, Name, EmailAddress, Title)) } catch {} }
        }
        $GridPreview.DataSource = $null; $GridPreview.DataSource = $DisplayList
        $GrpPreview.Text = "Preview [ Matches: $($DisplayList.Count) ]"
    } catch { $GrpPreview.Text = "Preview Error" }
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
}
function Trigger-AutoRefresh { if ($ChkAutoPreview.Checked) { Update-Preview } }

# EVENTS
$CmbType.Add_SelectedIndexChanged({ $TxtSMTP.Enabled = ($CmbType.SelectedItem -eq "MailEnabled") })
$BtnPreview.Add_Click({ Update-Preview })
$BtnIncAdd.Add_Click({ $U=Show-UserPicker; if($U){$LstInc.Items.Add($U); Trigger-AutoRefresh} }); $BtnExcAdd.Add_Click({ $U=Show-UserPicker; if($U){$LstExc.Items.Add($U); Trigger-AutoRefresh} })
$BtnIncRem.Add_Click({ if($LstInc.SelectedIndex -ge 0){$LstInc.Items.RemoveAt($LstInc.SelectedIndex); Trigger-AutoRefresh} }); $BtnExcRem.Add_Click({ if($LstExc.SelectedIndex -ge 0){$LstExc.Items.RemoveAt($LstExc.SelectedIndex); Trigger-AutoRefresh} })
$ItemAddGroupAnd.Add_Click({ $N=$TreeLogic.SelectedNode.Nodes.Add("Group (AND)"); $N.Tag=@{Type="GROUP"; Gate="AND"}; $TreeLogic.SelectedNode.Expand(); Trigger-AutoRefresh })
$ItemAddGroupOr.Add_Click({ $N=$TreeLogic.SelectedNode.Nodes.Add("Group (OR)"); $N.Tag=@{Type="GROUP"; Gate="OR"}; $TreeLogic.SelectedNode.Expand(); Trigger-AutoRefresh })
$ItemAddRule.Add_Click({ $R=Show-RuleDialog; if($R){ $N=$TreeLogic.SelectedNode.Nodes.Add("$($R.Friendly) $($R.Operator) '$($R.Value)'"); $N.Tag=@{Type="RULE"; Attribute=$R.Attribute; Operator=$R.Operator; Value=$R.Value}; $TreeLogic.SelectedNode.Expand(); Trigger-AutoRefresh } })
$ItemDelete.Add_Click({ if($TreeLogic.SelectedNode -ne $TreeLogic.Nodes[0]){ $TreeLogic.SelectedNode.Remove(); Trigger-AutoRefresh } })
$EditHandler = { $N=$TreeLogic.SelectedNode; $D=$N.Tag; $Changed=$false; if($D.Type -eq "GROUP"){ $New=if($D.Gate -eq "AND"){"OR"}else{"AND"}; $N.Tag.Gate=$New; if($N.Text -match "ROOT"){$N.Text="ROOT ($New)"}else{$N.Text="Group ($New)"}; $Changed=$true }elseif($D.Type -eq "RULE"){ $R=Show-RuleDialog -ExistingRule $D; if($R){ $N.Text="$($R.Friendly) $($R.Operator) '$($R.Value)'"; $N.Tag=@{Type="RULE"; Attribute=$R.Attribute; Operator=$R.Operator; Value=$R.Value}; $Changed=$true } } if($Changed){Trigger-AutoRefresh} }
$ItemEdit.Add_Click($EditHandler); $TreeLogic.Add_NodeMouseDoubleClick({ param($s,$e)$EditHandler.Invoke() })

function Load-Selection {
    $Row = Show-ListPicker; if (-not $Row) { $Form.Close(); return }
    $ID = $Row.Cells["DistListID"].Value; $Locker = $Row.Cells["LockedBy"].Value
    if ($Locker -and $Locker.ToString() -ne "" -and $Locker -ne $CurrentUser) { [System.Windows.Forms.MessageBox]::Show("List Locked by '$Locker'.", "Locked", 0, 16); $Form.Close(); return }
    Invoke-Sql "UPDATE DL_MasterList SET LockedBy = '$CurrentUser', LockedAt = GETDATE() WHERE DistListID = $ID" | Out-Null
    $script:CurrentListID = $ID
    $Data = (Invoke-Sql "SELECT * FROM DL_MasterList WITH (NOLOCK) WHERE DistListID=$ID").Data
    $TxtName.Text = $Data.Name; $TxtDesc.Text = $Data.Description; $CmbType.SelectedItem = $Data.ListType; $TxtSMTP.Text = if($Data.PrimarySMTPAddress -is [DBNull]){""}else{$Data.PrimarySMTPAddress}; $TxtOwner.Text = $Data.ManagedBy; $CmbSync.SelectedItem = [int]$Data.SyncIntervalMinutes; $ChkActive.Checked = [bool]$Data.SyncEnabled
    Rebuild-Tree $Data.RuleDefinition
    $LstInc.Items.Clear(); $LstExc.Items.Clear()
    (Invoke-Sql "SELECT UserIdentity FROM DL_Inclusions WITH (NOLOCK) WHERE DistListID=$ID").Data | ForEach { $LstInc.Items.Add($_.UserIdentity) }
    (Invoke-Sql "SELECT UserIdentity FROM DL_Exclusions WITH (NOLOCK) WHERE DistListID=$ID").Data | ForEach { $LstExc.Items.Add($_.UserIdentity) }
    $StatusLabel.Text = "Editing: $($Data.Name) (ID: $ID)"; $Form.Refresh(); Update-Preview
}

$DoCheckAuditSchema = ${function:Check-AuditSchema}.GetNewClosure()
$DoLoadSelection = ${function:Load-Selection}.GetNewClosure()

$Form.Add_FormClosing({ if ($CurrentListID -gt 0) { Invoke-Sql "UPDATE DL_MasterList SET LockedBy = NULL WHERE DistListID = $CurrentListID AND LockedBy = '$CurrentUser'" | Out-Null } })
$Form.Add_Load(({
    $script:HasAuditCol = $DoCheckAuditSchema.Invoke()
    $DoLoadSelection.Invoke()
}).GetNewClosure())
$Form.ShowDialog()
'@
$EmbeddedScript_AuditDashboard = @'
<# 
   .SYNOPSIS
   Dynamic List Auditor (v2.2) - Layout & Schema Fix
   .DESCRIPTION
   - Uses SplitContainers to prevent Grid Headers from being hidden.
   - Queries correct table: 'Log_MemberActivity' using 'ActionDate'.
#>

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing

# ==========================================
# CONFIGURATION
# ==========================================
$SqlInstance = if ($env:DYN_SQL_INSTANCE) { $env:DYN_SQL_INSTANCE } else { "SQL16-DIVERS" }
$Database    = if ($env:DYN_SQL_DATABASE) { $env:DYN_SQL_DATABASE } else { "DBTEST" }
# ==========================================

# --- SQL HELPER ---
function Invoke-Sql {
    param([string]$Query, [hashtable]$Parameters = @{}, [switch]$ReturnScalar)
    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;Connect Timeout=5;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString); $Conn.Open()
        $Cmd = $Conn.CreateCommand(); $Cmd.CommandText = $Query
        foreach ($Key in $Parameters.Keys) {
            $Val = if ($Parameters[$Key] -eq $null) { [DBNull]::Value } else { $Parameters[$Key] }
            $Cmd.Parameters.AddWithValue("@$Key", $Val) | Out-Null
        }
        if ($ReturnScalar) { $ResultInfo.Data = $Cmd.ExecuteScalar() } 
        else { $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd; $DS = New-Object System.Data.DataSet; $Adapter.Fill($DS) | Out-Null; $ResultInfo.Data = $DS.Tables[0] }
        $Conn.Close(); $ResultInfo.Success = $true
    }
    catch { [System.Windows.Forms.MessageBox]::Show("SQL Error: $($_.Exception.Message)") }
    return $ResultInfo
}

# --- USER PICKER DIALOG ---
function Show-UserPicker {
    $PickForm = New-Object System.Windows.Forms.Form; $PickForm.Text="Find AD User"; $PickForm.Size="500, 400"; $PickForm.StartPosition="CenterParent"
    $TxtS = New-Object System.Windows.Forms.TextBox; $TxtS.Location="10, 10"; $TxtS.Size="380, 25"
    $BtnS = New-Object System.Windows.Forms.Button; $BtnS.Text="Search"; $BtnS.Location="400, 9"; $BtnS.Size="70, 26"
    $LstR = New-Object System.Windows.Forms.ListBox; $LstR.Location="10, 45"; $LstR.Size="460, 270"
    $BtnOk = New-Object System.Windows.Forms.Button; $BtnOk.Text="Select"; $BtnOk.Location="150, 325"; $BtnOk.Size="200, 30"; $BtnOk.DialogResult="OK"
    
    $PickForm.Controls.AddRange(@($TxtS, $BtnS, $LstR, $BtnOk)); $PickForm.AcceptButton = $BtnS

    $BtnS.Add_Click({
        $LstR.Items.Clear(); $Term = $TxtS.Text
        if ($Term.Length -gt 2) {
            try {
                $Users = Get-ADUser -Filter "anr -eq '$Term'" -Properties EmailAddress | Select Name, SamAccountName
                foreach ($U in $Users) { $LstR.Items.Add("$($U.Name) [$($U.SamAccountName)]") | Out-Null }
            } catch {}
        }
    })
    
    if ($PickForm.ShowDialog() -eq "OK" -and $LstR.SelectedItem) { 
        return $LstR.SelectedItem.ToString().Split('[')[1].Trim(']') 
    }
    return $null
}

# --- DRILL-DOWN: LIST MEMBERS ---
function Show-ListDetails($ListID, $ListName) {
    $DForm = New-Object System.Windows.Forms.Form; $DForm.Text = "Members of: $ListName"; $DForm.Size = "800, 600"; $DForm.StartPosition = "CenterParent"
    
    # LAYOUT FIX: Use SplitContainer instead of Dock=Top to guarantee separation
    $Split = New-Object System.Windows.Forms.SplitContainer; $Split.Dock="Fill"; $Split.Orientation="Horizontal"; $Split.SplitterDistance=60; $Split.IsSplitterFixed=$true; $DForm.Controls.Add($Split)
    
    # Tools (Top Panel)
    $PnlTools = $Split.Panel1
    $PnlTools.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $LblS = New-Object System.Windows.Forms.Label; $LblS.Text="Filter Member:"; $LblS.AutoSize=$true; $LblS.Location="10,20"
    $TxtS = New-Object System.Windows.Forms.TextBox; $TxtS.Location="100,17"; $TxtS.Size="250,25"
    $BtnPick = New-Object System.Windows.Forms.Button; $BtnPick.Text="🔍 Picker"; $BtnPick.Location="360,15"; $BtnPick.Size="80,28"; $BtnPick.BackColor="LightBlue"
    $BtnFind = New-Object System.Windows.Forms.Button; $BtnFind.Text="Filter"; $BtnFind.Location="450,15"; $BtnFind.Size="80,28"
    $LblCount = New-Object System.Windows.Forms.Label; $LblCount.Text="Count: 0"; $LblCount.Location="550,20"; $LblCount.AutoSize=$true; $LblCount.Font=New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    
    $PnlTools.Controls.AddRange(@($LblS, $TxtS, $BtnPick, $BtnFind, $LblCount))
    
    # Grid (Bottom Panel)
    $MGrid = New-Object System.Windows.Forms.DataGridView; $MGrid.Dock="Fill"; $MGrid.ReadOnly=$true; $MGrid.AutoSizeColumnsMode="Fill"; $Split.Panel2.Controls.Add($MGrid)

    $LoadAction = {
        $Filter = $TxtS.Text
        $FullQ = "SELECT UserIdentity, DisplayName, AddedDate FROM DL_Members WHERE DistListID = @ID"
        $P = @{ID=$ListID}
        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            $FullQ += " AND (UserIdentity LIKE @F OR DisplayName LIKE @F)"
            $P.F = "%$Filter%"
        }
        $FullQ += " ORDER BY UserIdentity"
        
        $Res = Invoke-Sql $FullQ -Parameters $P
        if ($Res.Success) {
            $MGrid.DataSource = $Res.Data
            $LblCount.Text = "Count: $($Res.Data.Rows.Count)"
        }
    }
    
    $BtnPick.Add_Click({ $U = Show-UserPicker; if ($U) { $TxtS.Text = $U; $LoadAction.Invoke() } })
    $BtnFind.Add_Click({ $LoadAction.Invoke() })
    $TxtS.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $LoadAction.Invoke() } })
    
    $LoadAction.Invoke()
    $DForm.ShowDialog()
}

# ==========================================
# MAIN FORM
# ==========================================
$Form = New-Object System.Windows.Forms.Form; $Form.Text = "Dynamic List Auditor v2.2"; $Form.Size = "1200, 800"; $Form.StartPosition = "CenterScreen"

$TabControl = New-Object System.Windows.Forms.TabControl; $TabControl.Dock = "Fill"; $Form.Controls.Add($TabControl)
$TabDash = New-Object System.Windows.Forms.TabPage "1. Dashboard"; $TabUser = New-Object System.Windows.Forms.TabPage "2. User Inspector"; $TabHist = New-Object System.Windows.Forms.TabPage "3. Activity History"
$TabControl.Controls.AddRange(@($TabDash, $TabUser, $TabHist))

# ==========================================
# TAB 1: DASHBOARD
# ==========================================
# LAYOUT FIX: Outer Split for Stats, Inner Split for Grid
$SplitDashOuter = New-Object System.Windows.Forms.SplitContainer; $SplitDashOuter.Dock="Fill"; $SplitDashOuter.Orientation="Horizontal"; $SplitDashOuter.SplitterDistance=110; $SplitDashOuter.IsSplitterFixed=$true; $TabDash.Controls.Add($SplitDashOuter)

# Top: Stats
$PnlStats = $SplitDashOuter.Panel1; $PnlStats.BackColor="WhiteSmoke"

function Add-StatCard($Title, $Value, $Color, $X) {
    $P = New-Object System.Windows.Forms.Panel; $P.Size="200, 80"; $P.Location="$X, 15"; $P.BackColor="White"; $P.BorderStyle="FixedSingle"
    $L1 = New-Object System.Windows.Forms.Label; $L1.Text=$Title; $L1.Location="10,10"; $L1.AutoSize=$true; $L1.ForeColor="Gray"
    $L2 = New-Object System.Windows.Forms.Label; $L2.Text=$Value; $L2.Location="10,35"; $L2.AutoSize=$true; $L2.ForeColor=$Color; $L2.Font=New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $P.Controls.AddRange(@($L1, $L2)); $PnlStats.Controls.Add($P)
    return $L2
}
$LblStatTotal = Add-StatCard "Total Lists" "0" "Black" 20
$LblStatActive = Add-StatCard "Active Syncing" "0" "Green" 240
$LblStatInactive = Add-StatCard "Inactive / Draft" "0" "Gray" 460
$LblStatMembers = Add-StatCard "Total Memberships" "0" "Blue" 680

$BtnRefreshDash = New-Object System.Windows.Forms.Button; $BtnRefreshDash.Text="↻ Refresh"; $BtnRefreshDash.Location="900, 35"; $BtnRefreshDash.Size="100, 40"; $BtnRefreshDash.BackColor="LightBlue"; $PnlStats.Controls.Add($BtnRefreshDash)

# Bottom: Grid + Hint (Use another split to guarantee hint visibility)
$SplitDashInner = New-Object System.Windows.Forms.SplitContainer; $SplitDashInner.Dock="Fill"; $SplitDashInner.Orientation="Horizontal"; $SplitDashInner.SplitterDistance=30; $SplitDashInner.IsSplitterFixed=$true; $SplitDashOuter.Panel2.Controls.Add($SplitDashInner)

$LblHint = New-Object System.Windows.Forms.Label; $LblHint.Text="ℹ️ Double-click a row to view members"; $LblHint.Dock="Fill"; $LblHint.TextAlign="MiddleLeft"; $LblHint.BackColor="Info"; $SplitDashInner.Panel1.Controls.Add($LblHint)

$GridLists = New-Object System.Windows.Forms.DataGridView; $GridLists.Dock="Fill"; $GridLists.ReadOnly=$true; $GridLists.AutoSizeColumnsMode="Fill"; $GridLists.SelectionMode="FullRowSelect"; $SplitDashInner.Panel2.Controls.Add($GridLists)


# ==========================================
# TAB 2: USER INSPECTOR
# ==========================================
# LAYOUT FIX: SplitContainer
$SplitUser = New-Object System.Windows.Forms.SplitContainer; $SplitUser.Dock="Fill"; $SplitUser.Orientation="Horizontal"; $SplitUser.SplitterDistance=60; $SplitUser.IsSplitterFixed=$true; $TabUser.Controls.Add($SplitUser)

# Top: Search
$PnlUserTop = $SplitUser.Panel1
$LblUserSearch = New-Object System.Windows.Forms.Label; $LblUserSearch.Text="Search User (Identity):"; $LblUserSearch.Location="20,20"; $LblUserSearch.AutoSize=$true; $PnlUserTop.Controls.Add($LblUserSearch)
$TxtUserSearch = New-Object System.Windows.Forms.TextBox; $TxtUserSearch.Location="150,17"; $TxtUserSearch.Size="250,25"; $PnlUserTop.Controls.Add($TxtUserSearch)
$BtnUserPick = New-Object System.Windows.Forms.Button; $BtnUserPick.Text="..."; $BtnUserPick.Location="405, 16"; $BtnUserPick.Size="40, 27"; $BtnUserPick.BackColor="LightBlue"; $PnlUserTop.Controls.Add($BtnUserPick)
$BtnUserSearch = New-Object System.Windows.Forms.Button; $BtnUserSearch.Text="Find Memberships"; $BtnUserSearch.Location="460,15"; $BtnUserSearch.Size="150,30"; $BtnUserSearch.BackColor="LightGreen"; $PnlUserTop.Controls.Add($BtnUserSearch)

# Bottom: Grid
$GridUser = New-Object System.Windows.Forms.DataGridView; $GridUser.Dock="Fill"; $GridUser.ReadOnly=$true; $GridUser.AutoSizeColumnsMode="Fill"; $SplitUser.Panel2.Controls.Add($GridUser)


# ==========================================
# TAB 3: HISTORY
# ==========================================
# LAYOUT FIX: SplitContainer
$SplitHist = New-Object System.Windows.Forms.SplitContainer; $SplitHist.Dock="Fill"; $SplitHist.Orientation="Horizontal"; $SplitHist.SplitterDistance=60; $SplitHist.IsSplitterFixed=$true; $TabHist.Controls.Add($SplitHist)

# Top: Tools
$PnlHistTop = $SplitHist.Panel1
$LblHList = New-Object System.Windows.Forms.Label; $LblHList.Text="Select List:"; $LblHList.Location="20,20"; $LblHList.AutoSize=$true; $PnlHistTop.Controls.Add($LblHList)
$CmbHList = New-Object System.Windows.Forms.ComboBox; $CmbHList.Location="100,17"; $CmbHList.Size="250,25"; $CmbHList.DropDownStyle="DropDownList"; $PnlHistTop.Controls.Add($CmbHList)

$LblHUser = New-Object System.Windows.Forms.Label; $LblHUser.Text="User:"; $LblHUser.Location="370,20"; $LblHUser.AutoSize=$true; $PnlHistTop.Controls.Add($LblHUser)
$TxtHUser = New-Object System.Windows.Forms.TextBox; $TxtHUser.Location="420,17"; $TxtHUser.Size="150,25"; $PnlHistTop.Controls.Add($TxtHUser)
$BtnHUserPick = New-Object System.Windows.Forms.Button; $BtnHUserPick.Text="..."; $BtnHUserPick.Location="575, 16"; $BtnHUserPick.Size="40, 27"; $BtnHUserPick.BackColor="LightBlue"; $PnlHistTop.Controls.Add($BtnHUserPick)

$BtnHistSearch = New-Object System.Windows.Forms.Button; $BtnHistSearch.Text="Search Logs"; $BtnHistSearch.Location="640,15"; $BtnHistSearch.Size="120,30"; $BtnHistSearch.BackColor="LightSalmon"; $PnlHistTop.Controls.Add($BtnHistSearch)

# Bottom: Grid
$GridHist = New-Object System.Windows.Forms.DataGridView; $GridHist.Dock="Fill"; $GridHist.ReadOnly=$true; $GridHist.AutoSizeColumnsMode="Fill"; $SplitHist.Panel2.Controls.Add($GridHist)


# ==========================================
# LOGIC
# ==========================================

# 1. DASHBOARD LOGIC
$BtnRefreshDash.Add_Click({
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $Counts = (Invoke-Sql "SELECT (SELECT COUNT(*) FROM DL_MasterList) as Total, (SELECT COUNT(*) FROM DL_MasterList WHERE SyncEnabled=1) as Active, (SELECT COUNT(*) FROM DL_MasterList WHERE SyncEnabled=0) as Inactive, (SELECT COUNT(*) FROM DL_Members) as Members").Data
    if ($Counts) { $LblStatTotal.Text=$Counts.Total; $LblStatActive.Text=$Counts.Active; $LblStatInactive.Text=$Counts.Inactive; $LblStatMembers.Text=$Counts.Members }

    $Q = "SELECT L.DistListID, L.Name, L.Description, CASE WHEN L.SyncEnabled=1 THEN 'Active' ELSE 'Inactive' END as Status, COUNT(M.MemberID) as [Members], L.LastSyncDate FROM DL_MasterList L LEFT JOIN DL_Members M ON L.DistListID = M.DistListID GROUP BY L.DistListID, L.Name, L.Description, L.SyncEnabled, L.LastSyncDate ORDER BY L.Name"
    $Res = Invoke-Sql $Q
    if ($Res.Success) { 
        $GridLists.DataSource = $Res.Data 
        foreach ($Row in $GridLists.Rows) { if ($Row.Cells["Status"].Value -eq "Inactive") { $Row.DefaultCellStyle.ForeColor = "Gray" } }
    }
    
    $CmbHList.Items.Clear(); $CmbHList.Items.Add("All Lists") | Out-Null
    foreach ($Row in $Res.Data) { $CmbHList.Items.Add("$($Row.Name)|$($Row.DistListID)") | Out-Null }; $CmbHList.SelectedIndex = 0
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

$GridLists.Add_CellDoubleClick({
    if ($GridLists.SelectedRows.Count -gt 0) {
        $ID = $GridLists.SelectedRows[0].Cells["DistListID"].Value
        $Name = $GridLists.SelectedRows[0].Cells["Name"].Value
        Show-ListDetails $ID $Name
    }
})

# 2. USER INSPECTOR
$BtnUserPick.Add_Click({ $U = Show-UserPicker; if ($U) { $TxtUserSearch.Text = $U; $BtnUserSearch.PerformClick() } })
$BtnUserSearch.Add_Click({
    if (-not $TxtUserSearch.Text) { return }
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $Term = $TxtUserSearch.Text
    $Q = "SELECT L.Name, L.Description, CASE WHEN L.SyncEnabled=1 THEN 'Active' ELSE 'Inactive' END as Status, M.AddedDate, M.UserIdentity FROM DL_Members M JOIN DL_MasterList L ON M.DistListID = L.DistListID WHERE M.UserIdentity = @Exact OR M.UserIdentity LIKE @Like OR M.DisplayName LIKE @Like ORDER BY L.Name"
    $Res = Invoke-Sql $Q -Parameters @{Exact=$Term; Like="%$Term%"}
    $GridUser.DataSource = $Res.Data
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# 3. HISTORY (Updated to show List Name)
$BtnHUserPick.Add_Click({ $U = Show-UserPicker; if ($U) { $TxtHUser.Text = $U } })

$BtnHistSearch.Add_Click({
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # We use 'A' as the alias for the Activity table in the WHERE clause
    $Where = "WHERE 1=1"
    $P = @{}
    
    if ($TxtHUser.Text) { 
        $Where += " AND (A.UserIdentity = @Exact OR A.UserIdentity LIKE @Like)"
        $P.Exact = $TxtHUser.Text
        $P.Like = "%$($TxtHUser.Text)%" 
    }
    
    if ($CmbHList.SelectedIndex -gt 0) { 
        $Where += " AND A.DistListID = @ID"
        $P.ID = $CmbHList.SelectedItem.ToString().Split('|')[1] 
    }
    
    # JOIN DL_MasterList to get the Name column
    $Q = @"
    SELECT TOP 500 
        A.ActivityID, 
        L.Name AS [List Name], 
        A.DistListID, 
        A.UserIdentity, 
        A.ActionType, 
        A.ActivityDate
    FROM Log_MemberActivity A
    LEFT JOIN DL_MasterList L ON A.DistListID = L.DistListID
    $Where 
    ORDER BY A.ActivityDate DESC
"@

    $Res = Invoke-Sql $Q -Parameters $P
    $GridHist.DataSource = $Res.Data
    
    # Color Formatting
    foreach ($Row in $GridHist.Rows) {
        if ($Row.Cells["ActionType"].Value -eq "Added") { $Row.DefaultCellStyle.ForeColor = "Green" }
        elseif ($Row.Cells["ActionType"].Value -eq "Removed") { $Row.DefaultCellStyle.ForeColor = "Red" }
    }
    
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

$Form.Add_Load({ $BtnRefreshDash.PerformClick() })
$Form.ShowDialog()
'@
$EmbeddedScript_RestoreManager = @'
<# 
   .SYNOPSIS
   Dynamic List Restore Manager (v5.2) - Perfect Layout Persistence
   .DESCRIPTION
   - Fixes Layout Loading order to prevent "collapsed" panes.
   - Includes Legacy Data handling and Active Diff logic.
#>

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing

# ==========================================
# CONFIGURATION
# ==========================================
$SqlInstance = if ($env:DYN_SQL_INSTANCE) { $env:DYN_SQL_INSTANCE } else { "SQL16-DIVERS" }
$Database    = if ($env:DYN_SQL_DATABASE) { $env:DYN_SQL_DATABASE } else { "DBTEST" }
$ConfigFile  = "$env:APPDATA\DynamicListRestoreLayout_v5_2.json"

# --- SQL HELPER ---
function Invoke-Sql {
    param([string]$Query, [hashtable]$Parameters = @{}, [switch]$ReturnScalar)
    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;Connect Timeout=5;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
        $Conn.Open()
        $Cmd = $Conn.CreateCommand(); $Cmd.CommandText = $Query
        foreach ($Key in $Parameters.Keys) {
            $Val = if ($Parameters[$Key] -eq $null) { [DBNull]::Value } else { $Parameters[$Key] }
            $Cmd.Parameters.AddWithValue("@$Key", $Val) | Out-Null
        }
        if ($ReturnScalar) { $ResultInfo.Data = $Cmd.ExecuteScalar() } 
        else {
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd
            $DS = New-Object System.Data.DataSet; $Adapter.Fill($DS) | Out-Null; $ResultInfo.Data = $DS.Tables[0]
        }
        $Conn.Close(); $ResultInfo.Success = $true
    }
    catch { $ResultInfo.Message = $_.Exception.Message; [System.Windows.Forms.MessageBox]::Show("SQL Error: $($_.Exception.Message)") }
    return $ResultInfo
}

# ==========================================
# STATE
# ==========================================
$CurrentListID = 0
$CurrentName   = ""
$LiveState     = @{ Rules=""; Inc=@(); Exc=@() }

# ==========================================
# LAYOUT MANAGER (FIXED)
# ==========================================
function Save-Layout {
    $Layout = @{
        Width      = $Form.Width
        Height     = $Form.Height
        Top        = $Form.Top
        Left       = $Form.Left
        SplitMain  = $SplitMain.SplitterDistance
        SplitV     = $SplitPreviewV.SplitterDistance
        SplitRules = $SplitRules.SplitterDistance
        SplitEx    = $SplitEx.SplitterDistance
        Maximized  = ($Form.WindowState -eq 'Maximized')
    }
    $Layout | ConvertTo-Json | Out-File $ConfigFile -Force
}

function Load-Layout {
    $DefaultApplied = $false
    
    if (Test-Path $ConfigFile) {
        try {
            $Layout = Get-Content $ConfigFile | ConvertFrom-Json
            
            # 1. Restore Window Size
            if ($Layout.Width -and $Layout.Height) { 
                $Form.Width = $Layout.Width
                $Form.Height = $Layout.Height 
            }
            if ($Layout.Top) { $Form.Top = $Layout.Top }
            if ($Layout.Left) { $Form.Left = $Layout.Left }
            if ($Layout.Maximized) { $Form.WindowState = 'Maximized' }

            # 2. Apply Splitters (DELAYED via Shown event to allow controls to size first)
            $Form.Add_Shown({
                $Form.SuspendLayout()
                try {
                    # Apply Main Splitter First
                    if ($Layout.SplitMain -gt 50 -and $Layout.SplitMain -lt ($Form.Width - 50)) { 
                        $SplitMain.SplitterDistance = $Layout.SplitMain 
                    }
                    $SplitMain.Panel2.ResumeLayout() # Force resize of right pane
                    
                    # Apply Vertical Splitter
                    if ($Layout.SplitV -gt 50 -and $Layout.SplitV -lt ($SplitMain.Panel2.Height - 50)) { 
                        $SplitPreviewV.SplitterDistance = $Layout.SplitV 
                    }
                    
                    # Apply Inner Splitters
                    if ($Layout.SplitRules -gt 50) { $SplitRules.SplitterDistance = $Layout.SplitRules }
                    if ($Layout.SplitEx -gt 50) { $SplitEx.SplitterDistance = $Layout.SplitEx }
                } catch {
                    # If restoration fails, fallback to balanced
                    Apply-BalancedLayout
                }
                $Form.ResumeLayout()
            })
            $DefaultApplied = $true
        } catch {}
    }
    
    if (-not $DefaultApplied) {
        $Form.Add_Shown({ Apply-BalancedLayout })
    }
}

function Apply-BalancedLayout {
    $Form.SuspendLayout()
    $SplitMain.SplitterDistance = 450
    # Force update so Panel2 has width/height
    $SplitMain.Panel2.ResumeLayout() 
    
    $HalfH = [int]($SplitMain.Panel2.Height / 2)
    if ($HalfH -gt 50) { $SplitPreviewV.SplitterDistance = $HalfH }
    
    $HalfW = [int]($SplitMain.Panel2.Width / 2)
    if ($HalfW -gt 50) { 
        $SplitRules.SplitterDistance = $HalfW 
        $SplitEx.SplitterDistance = $HalfW 
    }
    $Form.ResumeLayout()
}

# ==========================================
# HELPER DIALOGS
# ==========================================
function Show-ListPicker {
    $P = New-Object System.Windows.Forms.Form; $P.Text = "Select List to Restore"; $P.Size = "900, 600"; $P.StartPosition="CenterScreen"
    $G = New-Object System.Windows.Forms.DataGridView; $G.Dock="Fill"; $G.SelectionMode="FullRowSelect"; $G.MultiSelect=$false; $G.ReadOnly=$true; $G.AutoSizeColumnsMode="Fill"
    $P.Controls.Add($G)
    $Res = Invoke-Sql "SELECT DistListID, Name, Description, LastSyncDate FROM DL_MasterList WITH (NOLOCK)"
    if ($Res.Success) { $G.DataSource = $Res.Data }
    $BtnOpen = New-Object System.Windows.Forms.Button; $BtnOpen.Text="View History"; $BtnOpen.Dock="Bottom"; $BtnOpen.Height=45; $BtnOpen.BackColor="LightBlue"; $BtnOpen.DialogResult="OK"; $P.Controls.Add($BtnOpen)
    if ($P.ShowDialog() -eq "OK" -and $G.SelectedRows.Count -gt 0) { return $G.SelectedRows[0] }
    return $null
}

# ==========================================
# MAIN FORM UI
# ==========================================
$Form = New-Object System.Windows.Forms.Form; $Form.Text = "Restore Manager (v5.2 Layout Fixed)"; $Form.Size = "1400, 900"; $Form.StartPosition = "CenterScreen"

# 1. BOTTOM BAR
$PnlBottom = New-Object System.Windows.Forms.Panel; $PnlBottom.Dock="Bottom"; $PnlBottom.Height=60; $PnlBottom.BackColor="WhiteSmoke"; $Form.Controls.Add($PnlBottom)
$BtnChange = New-Object System.Windows.Forms.Button; $BtnChange.Text="📂 Change List"; $BtnChange.Location="10, 10"; $BtnChange.Size="150, 40"; $BtnChange.BackColor="LightGray"; $PnlBottom.Controls.Add($BtnChange)
$BtnRestore = New-Object System.Windows.Forms.Button; $BtnRestore.Text="⚠ RESTORE SELECTED VERSION"; $BtnRestore.Dock="Right"; $BtnRestore.Width=250; $BtnRestore.BackColor="LightSalmon"; $BtnRestore.Enabled=$false; $PnlBottom.Controls.Add($BtnRestore)

# 2. TOP BAR
$PnlTop = New-Object System.Windows.Forms.Panel; $PnlTop.Dock="Top"; $PnlTop.Height=40; $PnlTop.Padding=New-Object System.Windows.Forms.Padding(10); $PnlTop.BackColor="AliceBlue"; $Form.Controls.Add($PnlTop)
$LblTarget = New-Object System.Windows.Forms.Label; $LblTarget.Text="No List Selected"; $LblTarget.Dock="Fill"; $LblTarget.TextAlign="MiddleLeft"; $LblTarget.Font=New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold); $PnlTop.Controls.Add($LblTarget)

# 3. MAIN SPLIT
$SplitMain = New-Object System.Windows.Forms.SplitContainer; $SplitMain.Dock="Fill"; $SplitMain.SplitterDistance=450; $Form.Controls.Add($SplitMain)

# LEFT: History Grid
$GrpGrid = New-Object System.Windows.Forms.GroupBox; $GrpGrid.Text="Version History"; $GrpGrid.Dock="Fill"; $GrpGrid.Padding=New-Object System.Windows.Forms.Padding(5, 20, 5, 5); $SplitMain.Panel1.Controls.Add($GrpGrid)
$Grid = New-Object System.Windows.Forms.DataGridView; $Grid.Dock="Fill"; $Grid.SelectionMode="FullRowSelect"; $Grid.MultiSelect=$false; $Grid.ReadOnly=$true; $Grid.AutoSizeColumnsMode="Fill"; $Grid.RowHeadersVisible=$false; $GrpGrid.Controls.Add($Grid)

# RIGHT: Preview Area
$PnlPreviewRoot = New-Object System.Windows.Forms.Panel; $PnlPreviewRoot.Dock="Fill"; $PnlPreviewRoot.Padding=New-Object System.Windows.Forms.Padding(5); $SplitMain.Panel2.Controls.Add($PnlPreviewRoot)

$SplitPreviewV = New-Object System.Windows.Forms.SplitContainer; $SplitPreviewV.Dock="Fill"; $SplitPreviewV.Orientation="Horizontal"; $SplitPreviewV.SplitterDistance=400; $PnlPreviewRoot.Controls.Add($SplitPreviewV)

# -- RULES (Top) --
$SplitRules = New-Object System.Windows.Forms.SplitContainer; $SplitRules.Dock="Fill"; $SplitRules.Orientation="Vertical"; $SplitRules.SplitterDistance=400; $SplitPreviewV.Panel1.Controls.Add($SplitRules)
$GrpLive = New-Object System.Windows.Forms.GroupBox; $GrpLive.Text="Current Live Rules"; $GrpLive.Dock="Fill"; $GrpLive.Padding=New-Object System.Windows.Forms.Padding(5,15,5,5); $SplitRules.Panel1.Controls.Add($GrpLive); $TreeLive = New-Object System.Windows.Forms.TreeView; $TreeLive.Dock="Fill"; $TreeLive.BackColor="WhiteSmoke"; $GrpLive.Controls.Add($TreeLive)
$GrpBackup = New-Object System.Windows.Forms.GroupBox; $GrpBackup.Text="Backup Rules (Snapshot)"; $GrpBackup.Dock="Fill"; $GrpBackup.Padding=New-Object System.Windows.Forms.Padding(5,15,5,5); $SplitRules.Panel2.Controls.Add($GrpBackup); $TreeBackup = New-Object System.Windows.Forms.TreeView; $TreeBackup.Dock="Fill"; $TreeBackup.BackColor="Azure"; $GrpBackup.Controls.Add($TreeBackup)

# -- EXCEPTIONS (Bottom) --
$SplitEx = New-Object System.Windows.Forms.SplitContainer; $SplitEx.Dock="Fill"; $SplitEx.Orientation="Vertical"; $SplitEx.SplitterDistance=400; $SplitPreviewV.Panel2.Controls.Add($SplitEx)
$GrpInc = New-Object System.Windows.Forms.GroupBox; $GrpInc.Text="Inclusions Diff"; $GrpInc.Dock="Fill"; $GrpInc.Padding=New-Object System.Windows.Forms.Padding(5,15,5,5); $SplitEx.Panel1.Controls.Add($GrpInc); $RtbInc = New-Object System.Windows.Forms.RichTextBox; $RtbInc.Dock="Fill"; $RtbInc.ReadOnly=$true; $GrpInc.Controls.Add($RtbInc)
$GrpExc = New-Object System.Windows.Forms.GroupBox; $GrpExc.Text="Exclusions Diff"; $GrpExc.Dock="Fill"; $GrpExc.Padding=New-Object System.Windows.Forms.Padding(5,15,5,5); $SplitEx.Panel2.Controls.Add($GrpExc); $RtbExc = New-Object System.Windows.Forms.RichTextBox; $RtbExc.Dock="Fill"; $RtbExc.ReadOnly=$true; $GrpExc.Controls.Add($RtbExc)

# ==========================================
# LOGIC FUNCTIONS
# ==========================================
function Build-Tree($TreeView, $Json) {
    $TreeView.Nodes.Clear()
    if ([string]::IsNullOrWhiteSpace($Json)) { $TreeView.Nodes.Add("No Logic Defined"); return }
    try {
        $Data = $Json | ConvertFrom-Json
        $Root = $TreeView.Nodes.Add("ROOT ($($Data.gate))"); $Root.Expand()
        if ($Data.children) { foreach ($child in $Data.children) { Render-Node -ParentNode $Root -Data $child } }
        $TreeView.ExpandAll()
    } catch { $TreeView.Nodes.Add("Error parsing JSON") }
}

function Render-Node($ParentNode, $Data) {
    if ($Data.type -eq "group") {
        $N = $ParentNode.Nodes.Add("Group ($($Data.gate))")
        foreach ($c in $Data.children) { Render-Node -ParentNode $N -Data $c }
    } else { $N = $ParentNode.Nodes.Add("$($Data.attribute) $($Data.operator) '$($Data.value)'") }
}

function Write-Diff($Rtb, $BackupList, $LiveList) {
    $Rtb.Clear()
    $InBackup = New-Object System.Collections.Generic.HashSet[string]
    if($BackupList){ $BackupList | ForEach { $InBackup.Add($_) | Out-Null } }
    
    $InLive = New-Object System.Collections.Generic.HashSet[string]
    if($LiveList){ $LiveList | ForEach { $InLive.Add($_) | Out-Null } }
    
    $DiffFound = $false
    foreach ($Item in $InBackup) { if (-not $InLive.Contains($Item)) { Append-Color $Rtb "+ $Item (Will be ADDED)" "Green"; $DiffFound=$true } }
    foreach ($Item in $InLive) { if (-not $InBackup.Contains($Item)) { Append-Color $Rtb "- $Item (Will be REMOVED)" "Red"; $DiffFound=$true } }
    
    if (-not $DiffFound) { if ($InBackup.Count -eq 0) { Append-Color $Rtb "(Both Empty)" "Gray" } else { Append-Color $Rtb "= No Changes Needed" "Black" } }
}

function Append-Color($Rtb, $Text, $ColorName) {
    $Rtb.SelectionStart = $Rtb.TextLength; $Rtb.SelectionLength = 0
    $Rtb.SelectionColor = [System.Drawing.Color]::FromName($ColorName)
    $Rtb.AppendText($Text + "`n"); $Rtb.SelectionColor = $Rtb.ForeColor
}

function Fetch-LiveState($ID) {
    $RawData = (Invoke-Sql "SELECT RuleDefinition FROM DL_MasterList WHERE DistListID=$ID").Data
    $LiveState.Rules = if ($RawData) { $RawData.RuleDefinition } else { "" }
    Build-Tree $TreeLive $LiveState.Rules
    
    $RawInc = (Invoke-Sql "SELECT UserIdentity FROM DL_Inclusions WHERE DistListID=$ID").Data
    $LiveState.Inc = @(); if ($RawInc) { foreach($r in $RawInc){ $LiveState.Inc += $r.UserIdentity } }

    $RawExc = (Invoke-Sql "SELECT UserIdentity FROM DL_Exclusions WHERE DistListID=$ID").Data
    $LiveState.Exc = @(); if ($RawExc) { foreach($r in $RawExc){ $LiveState.Exc += $r.UserIdentity } }
}

function Load-History($ListID, $ListName) {
    $script:CurrentListID = $ListID; $script:CurrentName = $ListName; $LblTarget.Text = "Target: $ListName (ID: $ListID)"
    Fetch-LiveState $ListID
    $Res = Invoke-Sql "SELECT LogID, ChangeDate, AdminUser, ChangeDescription, NewValue FROM Log_ConfigHistory WITH (NOLOCK) WHERE DistListID = $ListID AND ModifiedTable = 'DL_MasterList' ORDER BY ChangeDate DESC"
    if ($Res.Success) {
        $Grid.DataSource = $Res.Data
        $Grid.Columns["NewValue"].Visible = $false
        if($Grid.Columns["LogID"]){$Grid.Columns["LogID"].Width = 50}
        if($Grid.Columns["ChangeDate"]){$Grid.Columns["ChangeDate"].Width = 120}
        if($Grid.Columns["AdminUser"]){$Grid.Columns["AdminUser"].Width = 90}
    }
}

# ==========================================
# EVENTS
# ==========================================
$Grid.Add_SelectionChanged({
    if ($Grid.SelectedRows.Count -gt 0) {
        $Row = $Grid.SelectedRows[0]
        $RuleJson = $Row.Cells["NewValue"].Value
        $DateRef  = $Row.Cells["ChangeDate"].Value
        
        Build-Tree $TreeBackup $RuleJson
        
        $QExc = "SELECT TOP 1 NewValue FROM Log_ConfigHistory WITH (NOLOCK) WHERE DistListID = $CurrentListID AND ModifiedTable = 'DL_Exceptions_Combined' AND ABS(DATEDIFF(second, ChangeDate, @DateRef)) <= 5"
        $ResWrapper = Invoke-Sql $QExc -Parameters @{DateRef=$DateRef} -ReturnScalar
        $SnapshotJson = $ResWrapper.Data
        
        $BackInc = @(); $BackExc = @(); $ValidSnapshot = $false
        
        if ($SnapshotJson -and $SnapshotJson -is [string] -and $SnapshotJson.Trim().StartsWith("{")) {
            try {
                $Obj = $SnapshotJson | ConvertFrom-Json
                if ($Obj.Inclusions) { $BackInc = $Obj.Inclusions }
                if ($Obj.Exclusions) { $BackExc = $Obj.Exclusions }
                $ValidSnapshot = $true
            } catch { }
        } 
        
        if ($ValidSnapshot) { $GrpInc.Text = "Inclusions Diff (Snapshot Found)"; $GrpExc.Text = "Exclusions Diff (Snapshot Found)" }
        else { $GrpInc.Text = "Inclusions Diff [Legacy/No Snapshot]"; $GrpExc.Text = "Exclusions Diff [Legacy/No Snapshot]" }
        
        Fetch-LiveState $CurrentListID 
        Write-Diff $RtbInc $BackInc $LiveState.Inc
        Write-Diff $RtbExc $BackExc $LiveState.Exc
        
        $BtnRestore.Enabled = $true
    } else {
        $BtnRestore.Enabled = $false; $TreeBackup.Nodes.Clear(); $RtbInc.Clear(); $RtbExc.Clear()
    }
})

$BtnChange.Add_Click({ $Row = Show-ListPicker; if ($Row) { Load-History $Row.Cells["DistListID"].Value $Row.Cells["Name"].Value } })

$BtnRestore.Add_Click({
    $Row = $Grid.SelectedRows[0]
    $DateRef = $Row.Cells["ChangeDate"].Value; $Desc = $Row.Cells["ChangeDescription"].Value; $BackupRule = $Row.Cells["NewValue"].Value 
    
    if ([System.Windows.Forms.MessageBox]::Show("Confirm Restore from [$DateRef]?`nLive configuration will be overwritten.", "Confirm", "YesNo", "Warning") -eq "Yes") {
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $CurrData = (Invoke-Sql "SELECT RuleDefinition FROM DL_MasterList WHERE DistListID=$CurrentListID").Data
        $CurrRule = $CurrData.RuleDefinition
        $RestoreMsg = "RESTORED from backup [$DateRef]. Orig Desc: $Desc"
        Invoke-Sql "INSERT INTO Log_ConfigHistory (DistListID, AdminUser, ModifiedTable, ChangeType, OldValue, NewValue, ChangeDescription) VALUES ($CurrentListID, '$env:USERNAME', 'DL_MasterList', 'RESTORE', '$CurrRule', '$BackupRule', '$RestoreMsg')"
        Invoke-Sql "UPDATE DL_MasterList SET RuleDefinition = '$BackupRule', LastSyncDate = NULL WHERE DistListID=$CurrentListID"
        
        $QExc = "SELECT TOP 1 NewValue FROM Log_ConfigHistory WITH (NOLOCK) WHERE DistListID = $CurrentListID AND ModifiedTable = 'DL_Exceptions_Combined' AND ABS(DATEDIFF(second, ChangeDate, @DateRef)) <= 5"
        $ResWrapper = Invoke-Sql $QExc -Parameters @{DateRef=$DateRef} -ReturnScalar
        $SnapshotJson = $ResWrapper.Data
        
        if ($SnapshotJson -and $SnapshotJson.Trim().StartsWith("{")) {
            Invoke-Sql "DELETE FROM DL_Inclusions WHERE DistListID=$CurrentListID; DELETE FROM DL_Exclusions WHERE DistListID=$CurrentListID"
            try {
                $Obj = $SnapshotJson | ConvertFrom-Json
                if ($Obj.Inclusions) { 
                    $V = New-Object System.Collections.ArrayList
                    foreach($i in $Obj.Inclusions){ $Safe=$i.ToString().Replace("'","''"); $V.Add("($CurrentListID, '$Safe')")|Out-Null }
                    if($V.Count -gt 0){ $B=$V -join ","; Invoke-Sql "INSERT INTO DL_Inclusions (DistListID, UserIdentity) VALUES $B" }
                }
                if ($Obj.Exclusions) {
                    $V = New-Object System.Collections.ArrayList
                    foreach($i in $Obj.Exclusions){ $Safe=$i.ToString().Replace("'","''"); $V.Add("($CurrentListID, '$Safe')")|Out-Null }
                    if($V.Count -gt 0){ $B=$V -join ","; Invoke-Sql "INSERT INTO DL_Exclusions (DistListID, UserIdentity) VALUES $B" }
                }
            } catch {}
        }
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.MessageBox]::Show("Restore Complete."); Load-History $CurrentListID $CurrentName
    }
})

$DoLoadLayout = ${function:Load-Layout}.GetNewClosure()
$DoShowListPicker = ${function:Show-ListPicker}.GetNewClosure()
$DoLoadHistory = ${function:Load-History}.GetNewClosure()

$Form.Add_FormClosing({ Save-Layout })
$Form.Add_Load(({
    $DoLoadLayout.Invoke()
    try {
        $StartRow = $DoShowListPicker.Invoke()
        if ($StartRow) {
            $DoLoadHistory.Invoke($StartRow.Cells["DistListID"].Value, $StartRow.Cells["Name"].Value)
        } else {
            $Form.Close()
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_")
        $Form.Close()
    }
}).GetNewClosure())

$Form.ShowDialog()
'@
# endregion EmbeddedToolScripts

function Set-CardVisual {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Forms.Button]$Card,
        [Parameter(Mandatory=$true)][bool]$Hover
    )

    if ($Hover) {
        $Card.BackColor = [System.Drawing.Color]::FromArgb(243, 248, 255)
        $Card.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 130, 190)
    } else {
        $Card.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
        $Card.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(205, 212, 220)
    }
}

function Get-OrCreateTabPage {
    param(
        [Parameter(Mandatory=$true)][string]$TabName,
        [Parameter(Mandatory=$true)][scriptblock]$Factory
    )

    foreach ($Page in $MainTabs.TabPages) {
        if ($Page.Text -eq $TabName) {
            $MainTabs.SelectedTab = $Page
            return $Page
        }
    }

    $NewPage = & $Factory
    $MainTabs.TabPages.Add($NewPage)
    $MainTabs.SelectedTab = $NewPage
    return $NewPage
}

function Get-EmbeddedFormFromText {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptText
    )

    if ([string]::IsNullOrWhiteSpace($ScriptText)) {
        throw "Embedded script text is empty."
    }

    $Patched = $ScriptText
    $Patched = [regex]::Replace($Patched, '(?m)^\s*\[void\]\s*\$Form\.ShowDialog\(\)\s*$', 'return $Form')
    $Patched = [regex]::Replace($Patched, '(?m)^\s*\$Form\.ShowDialog\(\)\s*$', 'return $Form')

    # Run each embedded tool in an isolated scope to avoid variable/function collisions across tabs.
    $Wrapped = "& {`n$Patched`n}"
    $ScriptBlock = [ScriptBlock]::Create($Wrapped)
    $Result = & $ScriptBlock
    $FormResult = @($Result) | Where-Object { $_ -is [System.Windows.Forms.Form] } | Select-Object -Last 1

    if ($null -eq $FormResult) {
        throw "No Form object was returned by embedded script."
    }

    return $FormResult
}

function New-EmbeddedToolTab {
    param(
        [Parameter(Mandatory=$true)][string]$ToolName,
        [Parameter(Mandatory=$true)][string]$ScriptContent
    )

    $Page = New-Object System.Windows.Forms.TabPage
    $Page.Text = $ToolName
    $Page.BackColor = [System.Drawing.Color]::White

    $TopBar = New-Object System.Windows.Forms.Panel
    $TopBar.Dock = [System.Windows.Forms.DockStyle]::Top
    $TopBar.Height = 44
    $TopBar.BackColor = [System.Drawing.Color]::FromArgb(242, 245, 248)
    $Page.Controls.Add($TopBar)

    $LblTitle = New-Object System.Windows.Forms.Label
    $LblTitle.Text = "$ToolName"
    $LblTitle.AutoSize = $true
    $LblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $LblTitle.Location = New-Object System.Drawing.Point(12, 12)
    $TopBar.Controls.Add($LblTitle)

    $BtnCloseTab = New-Object System.Windows.Forms.Button
    $BtnCloseTab.Text = "Close Tab"
    $BtnCloseTab.Size = New-Object System.Drawing.Size(100, 26)
    $BtnCloseTab.Anchor = "Top,Right"
    $BtnCloseTab.Location = New-Object System.Drawing.Point(955, 9)
    $TopBar.Controls.Add($BtnCloseTab)

    $Body = New-Object System.Windows.Forms.Panel
    $Body.Dock = [System.Windows.Forms.DockStyle]::Fill
    $Body.Padding = New-Object System.Windows.Forms.Padding(18)
    $Page.Controls.Add($Body)
    $TopBar.BringToFront()

    $LblInfo = New-Object System.Windows.Forms.Label
    $LblInfo.Text = "Loading embedded tool..."
    $LblInfo.AutoSize = $true
    $LblInfo.Location = New-Object System.Drawing.Point(12, 16)
    $Body.Controls.Add($LblInfo)

    try {
        $EmbeddedForm = Get-EmbeddedFormFromText -ScriptText $ScriptContent
        $EmbeddedForm.TopLevel = $false
        $EmbeddedForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $EmbeddedForm.Dock = [System.Windows.Forms.DockStyle]::Fill
        $EmbeddedForm.Padding = New-Object System.Windows.Forms.Padding(0)
        $EmbeddedForm.Margin = New-Object System.Windows.Forms.Padding(0)

        $Body.Controls.Clear()
        $Body.Padding = New-Object System.Windows.Forms.Padding(0)
        $Body.Controls.Add($EmbeddedForm)
        $EmbeddedForm.Show()

        if ($ToolName -eq "Attribute Manager") {
            # Explicitly trigger table refresh after embedding; Load events are inconsistent in hosted mode.
            $RefreshButton = $null
            $AttributeGrid = $null
            $Stack = New-Object System.Collections.Stack
            $Stack.Push($EmbeddedForm)
            while ($Stack.Count -gt 0 -and (-not $RefreshButton -or -not $AttributeGrid)) {
                $Node = $Stack.Pop()
                foreach ($Child in $Node.Controls) {
                    if ($Child -is [System.Windows.Forms.Button] -and $Child.Text -eq "Refresh Table") {
                        $RefreshButton = $Child
                    }
                    if ($Child -is [System.Windows.Forms.DataGridView] -and -not $AttributeGrid) {
                        $AttributeGrid = $Child
                    }
                    if ($Child.Controls.Count -gt 0) {
                        $Stack.Push($Child)
                    }
                }
            }

            # Hard-bind from host SQL to guarantee rows are shown even if embedded event wiring fails.
            if ($AttributeGrid) {
                $HostRes = Invoke-Sql -Query "SELECT AttrID, FriendlyName, ADAttribute, DataType FROM Sys_AttributeMap"
                if ($HostRes.Success) {
                    $AttributeGrid.DataSource = $HostRes.Data
                    if ($AttributeGrid.Columns["AttrID"]) { $AttributeGrid.Columns["AttrID"].Visible = $false }
                }
            }

            if ($RefreshButton) {
                $RefreshButton.PerformClick()
            }
        }

        $StatusMain.Text = "Opened: $ToolName"
    } catch {
        $Body.Controls.Clear()
        $ErrLabel = New-Object System.Windows.Forms.Label
        $ErrLabel.AutoSize = $true
        $ErrLabel.ForeColor = [System.Drawing.Color]::Maroon
        $ErrLabel.Text = "Failed to load embedded tool: $($_.Exception.Message)"
        $ErrLabel.Location = New-Object System.Drawing.Point(12, 16)
        $Body.Controls.Add($ErrLabel)
        $StatusMain.Text = "Failed: $ToolName"
    }

    $BtnCloseTab.Add_Click(({
        $MainTabs.TabPages.Remove($Page)
        $MainTabs.SelectedTab = $TabDashboard
    }).GetNewClosure())

    return $Page
}

function Add-DashboardCard {
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string]$Description,
        [Parameter(Mandatory=$true)][string]$ColorName,
        [Parameter(Mandatory=$true)][scriptblock]$OnClick
    )

    $Card = New-Object System.Windows.Forms.Button
    $Card.Size = New-Object System.Drawing.Size(310, 164)
    $Card.FlatStyle = "Flat"
    $Card.FlatAppearance.BorderSize = 1
    $Card.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(234, 241, 252)
    $Card.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(243, 248, 255)
    $Card.Margin = New-Object System.Windows.Forms.Padding(12)
    $Card.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Card.Text = ""
    Set-CardVisual -Card $Card -Hover $false

    $Card.Add_Click($OnClick)
    $Card.Add_MouseEnter({ Set-CardVisual -Card $Card -Hover $true }.GetNewClosure())
    $Card.Add_MouseLeave({ Set-CardVisual -Card $Card -Hover $false }.GetNewClosure())

    $Bar = New-Object System.Windows.Forms.Panel
    $Bar.Size = New-Object System.Drawing.Size(306, 7)
    $Bar.Location = New-Object System.Drawing.Point(2, 6)
    $Bar.BackColor = [System.Drawing.Color]::FromName($ColorName)

    $Lbl = New-Object System.Windows.Forms.Label
    $Lbl.Text = $Title
    $Lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $Lbl.ForeColor = [System.Drawing.Color]::FromArgb(32, 45, 58)
    $Lbl.Location = New-Object System.Drawing.Point(12, 26)
    $Lbl.AutoSize = $true

    $LblD = New-Object System.Windows.Forms.Label
    $LblD.Text = $Description
    $LblD.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $LblD.ForeColor = [System.Drawing.Color]::FromArgb(87, 97, 108)
    $LblD.Location = New-Object System.Drawing.Point(12, 56)
    $LblD.Size = New-Object System.Drawing.Size(286, 40)

    $Line = New-Object System.Windows.Forms.Label
    $Line.BackColor = [System.Drawing.Color]::FromArgb(219, 225, 231)
    $Line.Location = New-Object System.Drawing.Point(12, 102)
    $Line.Size = New-Object System.Drawing.Size(286, 1)

    $LblHint = New-Object System.Windows.Forms.Label
    $LblHint.Text = "Click to open tab"
    $LblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $LblHint.ForeColor = [System.Drawing.Color]::FromArgb(117, 125, 132)
    $LblHint.Location = New-Object System.Drawing.Point(12, 116)
    $LblHint.AutoSize = $true

    $Forwards = @($Bar, $Lbl, $LblD, $Line, $LblHint)
    foreach ($Ctrl in $Forwards) {
        $Ctrl.Add_Click($OnClick)
        $Ctrl.Add_MouseEnter({ Set-CardVisual -Card $Card -Hover $true }.GetNewClosure())
        $Ctrl.Add_MouseLeave({ Set-CardVisual -Card $Card -Hover $false }.GetNewClosure())
    }

    $Card.Controls.AddRange(@($Bar, $Lbl, $LblD, $Line, $LblHint))
    $DashFlow.Controls.Add($Card)
}

Add-DashboardCard -Title "Attribute Manager" -Description "Manage attribute mappings in Sys_AttributeMap table." -ColorName "Goldenrod" -OnClick ({
    $StatusMain.Text = "Opening: Attribute Manager"
    $null = Get-OrCreateTabPage -TabName "Attribute Manager" -Factory ({ New-EmbeddedToolTab -ToolName "Attribute Manager" -ScriptContent $EmbeddedScript_AttributeManager }.GetNewClosure())
}.GetNewClosure())

Add-DashboardCard -Title "Create New List" -Description "Open Create New List module tab." -ColorName "SeaGreen" -OnClick ({
    $StatusMain.Text = "Opening: Create New List"
    $null = Get-OrCreateTabPage -TabName "Create New List" -Factory ({ New-EmbeddedToolTab -ToolName "Create New List" -ScriptContent $EmbeddedScript_CreateNewList }.GetNewClosure())
}.GetNewClosure())

Add-DashboardCard -Title "Manage List" -Description "Open Manage List module tab." -ColorName "Teal" -OnClick ({
    $StatusMain.Text = "Opening: Manage List"
    $null = Get-OrCreateTabPage -TabName "Manage List" -Factory ({ New-EmbeddedToolTab -ToolName "Manage List" -ScriptContent $EmbeddedScript_ManageList }.GetNewClosure())
}.GetNewClosure())

Add-DashboardCard -Title "Audit Dashboard" -Description "Open Audit Dashboard module tab." -ColorName "CornflowerBlue" -OnClick ({
    $StatusMain.Text = "Opening: Audit Dashboard"
    $null = Get-OrCreateTabPage -TabName "Audit Dashboard" -Factory ({ New-EmbeddedToolTab -ToolName "Audit Dashboard" -ScriptContent $EmbeddedScript_AuditDashboard }.GetNewClosure())
}.GetNewClosure())

Add-DashboardCard -Title "Restore Manager" -Description "Open Restore Manager module tab." -ColorName "Salmon" -OnClick ({
    $StatusMain.Text = "Opening: Restore Manager"
    $null = Get-OrCreateTabPage -TabName "Restore Manager" -Factory ({ New-EmbeddedToolTab -ToolName "Restore Manager" -ScriptContent $EmbeddedScript_RestoreManager }.GetNewClosure())
}.GetNewClosure())

[void]$Form.ShowDialog()
