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
$SqlInstance = "SQL16-DIVERS" 
$Database    = "DBTEST"      
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

$Form.Add_FormClosing({ Save-Layout })
$Form.Add_Load({ Load-Layout; try { $StartRow = Show-ListPicker; if ($StartRow) { Load-History $StartRow.Cells["DistListID"].Value $StartRow.Cells["Name"].Value } else { $Form.Close() } } catch { [System.Windows.Forms.MessageBox]::Show("Error: $_"); $Form.Close() } })

$Form.ShowDialog()
