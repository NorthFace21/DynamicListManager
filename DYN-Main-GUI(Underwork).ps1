<# 
   .SYNOPSIS
   Dynamic List Management Suite (DYN-Main v5.2)
   .DESCRIPTION
   - FIX: Reliable Path Handling using .Tag property.
   - FIX: Layout spacing to prevent cards from being cut off.
   - FIX: Bottom panel resized to ensure Start Button is visible.
#>

# --- 1. HIDE HOSTING CONSOLE ---
$ConsoleVisibility = @{
    namespace  = 'Win32'
    name       = 'Console'
    member     = '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
}
try {
    Add-Type -MemberDefinition $ConsoleVisibility['member'] -Name $ConsoleVisibility['name'] -Namespace $ConsoleVisibility['namespace']
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) { [Win32.Console]::ShowWindow($hwnd, 0) }
} catch {}

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================================
# 2. PATH HUNTER (Robust Resolution)
# ==========================================
$PossiblePaths = @(
    $PSScriptRoot,
    (Split-Path -Parent $MyInvocation.MyCommand.Path),
    (Get-Location).Path
)

$ScriptDir = $null
foreach ($Path in $PossiblePaths) {
    if ($Path -and (Test-Path (Join-Path $Path "DYN-UpdateEngine.ps1"))) {
        $ScriptDir = $Path
        break
    }
}
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

# ==========================================
# CONFIGURATION
# ==========================================
$EngineScript = Join-Path $ScriptDir "DYN-UpdateEngine.ps1"

$Tools = @{
    "01" = @{ Title="Attribute Manager"; File="DYN-AttributeManager-GUI.ps1"; Color="Goldenrod";      Desc="Map AD Attributes to SQL" }
    "02" = @{ Title="Create New List";   File="DYN-CreateNewList-GUI.ps1";    Color="SeaGreen";       Desc="Wizard for new distributions" }
    "03" = @{ Title="Manage List";       File="DYN-ManageList-GUI.ps1";           Color="Teal";           Desc="Edit Rules & Exceptions" }
    "04" = @{ Title="Audit Dashboard";   File="DYN-Dashboard-GUI.ps1";        Color="CornflowerBlue"; Desc="Stats, Membership & History" }
    "05" = @{ Title="Restore Manager";   File="DYN-Restore-GUI.ps1";          Color="Salmon";         Desc="Rollback Configs & Undo" }
}

# ==========================================
# MAIN FORM SETUP
# ==========================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Dynamic List Suite (Running from: $ScriptDir)"
$Form.Size = "1000, 850"
$Form.StartPosition = "CenterScreen"
$Form.BackColor = "WhiteSmoke"

# --- MAIN SPLIT ---
$Split = New-Object System.Windows.Forms.SplitContainer
$Split.Dock = "Fill"
$Split.Orientation = "Horizontal"
$Split.SplitterDistance = 550  # REDUCED to give bottom panel more space
$Split.FixedPanel = "Panel2"   
$Form.Controls.Add($Split)

# ==========================================
# TOP PANEL: TOOLS
# ==========================================
$PnlTools = $Split.Panel1
$PnlTools.BackColor = "#F0F0F0"

$LblTitle = New-Object System.Windows.Forms.Label
$LblTitle.Text = "GUI TOOLS"
$LblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$LblTitle.ForeColor = "DimGray"
$LblTitle.Dock = "Top"
$LblTitle.Height = 50
$LblTitle.Padding = New-Object System.Windows.Forms.Padding(20,15,0,0)
$PnlTools.Controls.Add($LblTitle)

$Flow = New-Object System.Windows.Forms.FlowLayoutPanel
$Flow.Dock = "Fill"
$Flow.Padding = New-Object System.Windows.Forms.Padding(30, 10, 0, 0) # Added Top Padding
$Flow.AutoScroll = $true
$PnlTools.Controls.Add($Flow)

# SHARED CLICK HANDLER (Relies on .Tag)
$CardClickHandler = {
    $Path = $this.Tag
    # If clicked child control (Label/Panel), bubble up to Button Tag
    if ($null -eq $Path) { $Path = $this.Parent.Tag }

    if (Test-Path $Path) {
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Path`""
    } else {
        [System.Windows.Forms.MessageBox]::Show("Script Missing:`n$Path", "Error", 0, 16)
    }
}

function Add-ToolCard($Title, $Data) {
    $FullPath = Join-Path $ScriptDir $Data.File
    
    $Btn = New-Object System.Windows.Forms.Button
    $Btn.Size = "300, 150" # Taller for safety
    $Btn.FlatStyle = "Flat"
    $Btn.BackColor = "White"
    $Btn.Margin = New-Object System.Windows.Forms.Padding(15)
    $Btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Btn.Tag = $FullPath # Store Path Here
    $Btn.Add_Click($CardClickHandler)

    # 1. Color Bar
    $Bar = New-Object System.Windows.Forms.Panel
    $Bar.Size = "298, 6"
    $Bar.Location = "1, 1"
    $Bar.BackColor = [System.Drawing.Color]::FromName($Data.Color)
    $Bar.Add_Click($CardClickHandler)
    
    # 2. Title
    $Lbl = New-Object System.Windows.Forms.Label
    $Lbl.Text = $Title
    $Lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $Lbl.Location = "10, 20"
    $Lbl.AutoSize = $true
    $Lbl.Add_Click($CardClickHandler)

    # 3. Description
    $LblD = New-Object System.Windows.Forms.Label
    $LblD.Text = $Data.Desc
    $LblD.ForeColor = "Gray"
    $LblD.Location = "10, 45"
    $LblD.Size = "280, 20"
    $LblD.Add_Click($CardClickHandler)
    
    # 4. Separator
    $Line = New-Object System.Windows.Forms.Label
    $Line.BackColor = "LightGray"
    $Line.Location = "10, 75"
    $Line.Size = "280, 1"
    
    # 5. Path Info
    $LblPath = New-Object System.Windows.Forms.Label
    $LblPath.Text = "File: $FullPath"
    $LblPath.Font = New-Object System.Drawing.Font("Consolas", 8)
    $LblPath.ForeColor = "DimGray"
    $LblPath.Location = "10, 85"
    $LblPath.Size = "280, 55" # Taller to prevent cut-off
    $LblPath.Add_Click($CardClickHandler)

    $Btn.Controls.AddRange(@($Bar, $Lbl, $LblD, $Line, $LblPath))
    $Flow.Controls.Add($Btn)
}

$Tools.GetEnumerator() | Sort-Object Name | ForEach-Object { Add-ToolCard $_.Value.Title $_.Value }

# ==========================================
# BOTTOM PANEL: ENGINE
# ==========================================
$PnlEngine = $Split.Panel2
$PnlEngine.BackColor = "#E0E0E0"
$PnlEngine.Padding = New-Object System.Windows.Forms.Padding(20)

$GrpEngine = New-Object System.Windows.Forms.GroupBox
$GrpEngine.Text = "Sync Engine Controller"
$GrpEngine.Dock = "Fill"
$GrpEngine.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$PnlEngine.Controls.Add($GrpEngine)

# Layout Panel for Engine Group
$PnlEngInner = New-Object System.Windows.Forms.Panel
$PnlEngInner.Dock = "Fill"
$PnlEngInner.Padding = New-Object System.Windows.Forms.Padding(20)
$GrpEngine.Controls.Add($PnlEngInner)

$LblEngInfo = New-Object System.Windows.Forms.Label
$LblEngInfo.Text = "Script: $EngineScript"
$LblEngInfo.Dock = "Top"
$LblEngInfo.Height = 30
$PnlEngInner.Controls.Add($LblEngInfo)

$BtnRun = New-Object System.Windows.Forms.Button
$BtnRun.Text = "▶ LAUNCH SYNC ENGINE"
$BtnRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$BtnRun.BackColor = "LightGreen"
$BtnRun.ForeColor = "DarkGreen"
$BtnRun.Dock = "Top"
$BtnRun.Height = 50
$BtnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
$PnlEngInner.Controls.Add($BtnRun)

$LblNote = New-Object System.Windows.Forms.Label
$LblNote.Text = "* Opens a separate PowerShell console window."
$LblNote.ForeColor = "DimGray"
$LblNote.Dock = "Top"
$LblNote.Height = 30
$LblNote.Padding = New-Object System.Windows.Forms.Padding(0,5,0,0)
$PnlEngInner.Controls.Add($LblNote)

# ==========================================
# ACTIONS
# ==========================================
$BtnRun.Add_Click({
    if (-not (Test-Path $EngineScript)) {
        [System.Windows.Forms.MessageBox]::Show("Engine script not found!`n$EngineScript", "Error", 0, 16)
        return
    }
    try {
        Start-Process powershell -ArgumentList "-NoExit -ExecutionPolicy Bypass -File `"$EngineScript`"" -WorkingDirectory $ScriptDir
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Launch Failed:`n$_", "Error", 0, 16)
    }
})

# ==========================================
# RUN
# ==========================================
$Form.ShowDialog()
