<# 
   .SYNOPSIS
   Attribute Manager (v4.0) - Finalized with Status Bar & Hard Refresh
   .DESCRIPTION
   Manages the Sys_AttributeMap table. 
   Dependencies: None (Uses native .NET SQL Client).
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

    # 1. STATUS STRIP (Bottom Bar)
    $StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $StatusLabel.Text = "Ready - Server: $SqlInstance / DB: $Database"
    $StatusStrip.Items.Add($StatusLabel)
    $Form.Controls.Add($StatusStrip)

    # 2. EDITOR PANEL
    $GrpInput = New-Object System.Windows.Forms.GroupBox
    $GrpInput.Text = "Attribute Editor"
    $GrpInput.Location = New-Object System.Drawing.Point(12, 10)
    $GrpInput.Size = New-Object System.Drawing.Size(520, 150)
    $Form.Controls.Add($GrpInput)

    # Inputs
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

    # -- BUTTONS --
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

    # 3. GRID CONTROLS
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

            [System.Windows.Forms.MessageBox]::Show(
                "Tip: You can override SQL target without editing code by setting environment variables before launching:`n`nDYN_SQL_INSTANCE=your-server-or-server\instance`nDYN_SQL_DATABASE=your-database",
                "SQL Connection Tip",
                0,
                64
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

    $Form.Add_Load(({ $null = $DoRefreshGrid.Invoke() }).GetNewClosure())
    return $Form
}

if ($MyInvocation.InvocationName -ne '.') {
    [void](New-AttributeManagerForm).ShowDialog()
}
