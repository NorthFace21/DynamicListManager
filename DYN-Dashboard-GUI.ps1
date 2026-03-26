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
$SqlInstance = "SQL16-DIVERS" 
$Database    = "DBTEST"      
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
