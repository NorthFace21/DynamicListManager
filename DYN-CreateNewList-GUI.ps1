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
$SqlInstance = "SQL16-DIVERS" 
$Database    = "DBTEST"      
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
