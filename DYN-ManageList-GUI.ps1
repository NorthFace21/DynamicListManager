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
$SqlInstance = "SQL16-DIVERS" 
$Database    = "DBTEST"      
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

$Form.Add_FormClosing({ if ($CurrentListID -gt 0) { Invoke-Sql "UPDATE DL_MasterList SET LockedBy = NULL WHERE DistListID = $CurrentListID AND LockedBy = '$CurrentUser'" | Out-Null } })
$Form.Add_Load({ $script:HasAuditCol = Check-AuditSchema; Load-Selection })
$Form.ShowDialog()

