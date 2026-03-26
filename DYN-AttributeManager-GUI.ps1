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
$SqlInstance = "SQL16-DIVERS" 
$Database    = "DBTEST"      
# ==========================================

# --- SQL HELPER ---
function Invoke-Sql {
    param($Query, $ReturnScalar=$false)
    
    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }

    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
        $Conn.Open()
        
        $Cmd = $Conn.CreateCommand()
        $Cmd.CommandText = $Query
        
        if ($ReturnScalar) {
            $ResultInfo.Data = $Cmd.ExecuteScalar()
        } else {
            $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Cmd
            $DS = New-Object System.Data.DataSet
            $Adapter.Fill($DS) | Out-Null
            $ResultInfo.Data = $DS.Tables[0]
        }
        $Conn.Close()
        $ResultInfo.Success = $true
    }
    catch {
        $ResultInfo.Message = $_.Exception.Message
    }
    
    return $ResultInfo
}

function New-AttributeManagerForm {
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "AD Attribute Mapper"
    $Form.Size = New-Object System.Drawing.Size(560, 650)
    $Form.StartPosition = "CenterScreen"

    # 1. STATUS STRIP (Bottom Bar)
    $StatusStrip = New-Object System.Windows.Forms.StatusStrip
    $StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $StatusLabel.Text = "Ready"
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

    function Set-Status {
        param($Msg, $IsError=$false)
        $StatusLabel.Text = "[$([DateTime]::Now.ToString('HH:mm:ss'))] $Msg"
        if ($IsError) { $StatusLabel.ForeColor = "Red" } else { $StatusLabel.ForeColor = "Black" }
    }

    function Refresh-Grid {
        Set-Status "Querying Database..."
        $Grid.DataSource = $null
        $Res = Invoke-Sql "SELECT AttrID, FriendlyName, ADAttribute, DataType FROM Sys_AttributeMap"

        if ($Res.Success) {
            $Grid.DataSource = $Res.Data
            if ($Grid.Columns["AttrID"]) { $Grid.Columns["AttrID"].Visible = $false }

            $Count = if ($Res.Data) { $Res.Data.Rows.Count } else { 0 }
            Set-Status "Connected. Rows Loaded: $Count"
        } else {
            Set-Status "Error: $($Res.Message)" $true
        }
    }

    function Clear-Form {
        $TxtID.Text = ""
        $TxtFriendly.Text = ""
        $TxtAD.Text = ""
        $CmbType.SelectedIndex = 0
        $BtnAdd.Enabled = $true
        $BtnUpdate.Enabled = $false
        $BtnDelete.Enabled = $false
        $Grid.ClearSelection()
    }

    $BtnAdd.Add_Click({
        if ($TxtFriendly.Text -and $TxtAD.Text) {
            $Q = "INSERT INTO Sys_AttributeMap (FriendlyName, ADAttribute, DataType) VALUES ('$($TxtFriendly.Text)', '$($TxtAD.Text)', '$($CmbType.SelectedItem)')"
            $Res = Invoke-Sql $Q
            if ($Res.Success) { Refresh-Grid; Clear-Form } else { Set-Status "Insert Failed: $($Res.Message)" $true }
        }
    })

    $BtnUpdate.Add_Click({
        if ($TxtID.Text) {
            $Q = "UPDATE Sys_AttributeMap SET FriendlyName='$($TxtFriendly.Text)', ADAttribute='$($TxtAD.Text)', DataType='$($CmbType.SelectedItem)' WHERE AttrID=$($TxtID.Text)"
            $Res = Invoke-Sql $Q
            if ($Res.Success) { Refresh-Grid; Clear-Form } else { Set-Status "Update Failed: $($Res.Message)" $true }
        }
    })

    $BtnDelete.Add_Click({
        if ($TxtID.Text) {
            if ([System.Windows.Forms.MessageBox]::Show("Delete this mapping?", "Confirm", "YesNo") -eq "Yes") {
                $Q = "DELETE FROM Sys_AttributeMap WHERE AttrID=$($TxtID.Text)"
                $Res = Invoke-Sql $Q
                if ($Res.Success) { Refresh-Grid; Clear-Form } else { Set-Status "Delete Failed: $($Res.Message)" $true }
            }
        }
    })

    $BtnClear.Add_Click({ Clear-Form })
    $BtnRefresh.Add_Click({ Refresh-Grid })

    $Grid.Add_CellClick({
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
    })

    $Form.Add_Load({ Refresh-Grid })
    return $Form
}

if ($MyInvocation.InvocationName -ne '.') {
    [void](New-AttributeManagerForm).ShowDialog()
}
