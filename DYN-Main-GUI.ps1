<#
   .SYNOPSIS
   Dynamic List Management Suite (DYN-Main v5.3)
   .DESCRIPTION
   - Improved card visuals with hover highlights and cleaner spacing.
   - Added status bar + clock for better runtime feedback.
   - Added in-app engine mode with live output (no separate PowerShell window required).
#>

# --- 1. HIDE HOSTING CONSOLE ---
$ConsoleVisibility = @{
    namespace  = 'Win32'
    name       = 'Console'
    member     = '[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
}
try {
    Add-Type -MemberDefinition $ConsoleVisibility.member -Name $ConsoleVisibility.name -Namespace $ConsoleVisibility.namespace
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) { [Win32.Console]::ShowWindow($hwnd, 0) }
} catch {}

# --- LOAD ASSEMBLIES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================================
# PATH HUNTER
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
$script:EngineProcess = $null
$ScriptHost = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell" }

$Tools = @{
    "01" = @{ Title="Attribute Manager"; File="DYN-AttributeManager-GUI.ps1"; Color="Goldenrod";      Desc="Map AD Attributes to SQL" }
    "02" = @{ Title="Create New List";   File="DYN-CreateNewList-GUI.ps1";    Color="SeaGreen";       Desc="Wizard for new distributions" }
    "03" = @{ Title="Manage List";       File="DYN-ManageList-GUI.ps1";       Color="Teal";           Desc="Edit rules and exceptions" }
    "04" = @{ Title="Audit Dashboard";   File="DYN-Dashboard-GUI.ps1";        Color="CornflowerBlue"; Desc="Stats, membership and history" }
    "05" = @{ Title="Restore Manager";   File="DYN-Restore-GUI.ps1";          Color="Salmon";         Desc="Rollback configs and undo" }
}

# ==========================================
# MAIN FORM SETUP
# ==========================================
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Dynamic List Suite"
$Form.Size = "1060, 860"
$Form.MinimumSize = "980, 760"
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$Split = New-Object System.Windows.Forms.SplitContainer
$Split.Dock = "Fill"
$Split.Orientation = "Horizontal"
$Split.SplitterDistance = 540
$Split.FixedPanel = "Panel2"
$Form.Controls.Add($Split)

# ==========================================
# STATUS BAR
# ==========================================
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

# ==========================================
# TOP PANEL: TOOLS
# ==========================================
$PnlTools = $Split.Panel1
$PnlTools.BackColor = [System.Drawing.Color]::FromArgb(237, 241, 246)

$Flow = New-Object System.Windows.Forms.FlowLayoutPanel
$Flow.Dock = "Fill"
$Flow.Padding = New-Object System.Windows.Forms.Padding(22, 16, 14, 16)
$Flow.WrapContents = $true
$Flow.AutoScroll = $true
$PnlTools.Controls.Add($Flow)

$Header = New-Object System.Windows.Forms.Panel
$Header.Dock = "Top"
$Header.Height = 74
$Header.BackColor = [System.Drawing.Color]::FromArgb(25, 57, 89)
$PnlTools.Controls.Add($Header)

$LblTitle = New-Object System.Windows.Forms.Label
$LblTitle.Text = "Dynamic List Manager"
$LblTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14, [System.Drawing.FontStyle]::Bold)
$LblTitle.ForeColor = "White"
$LblTitle.Location = "20, 12"
$LblTitle.AutoSize = $true
$Header.Controls.Add($LblTitle)

$LblSub = New-Object System.Windows.Forms.Label
$LblSub.Text = "Launch tools from one dashboard"
$LblSub.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$LblSub.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 240)
$LblSub.Location = "22, 42"
$LblSub.AutoSize = $true
$Header.Controls.Add($LblSub)

function Set-CardVisual {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Forms.Button]$Card,
        [Parameter(Mandatory=$true)][bool]$Hover
    )

    $BaseColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $HoverColor = [System.Drawing.Color]::FromArgb(243, 248, 255)

    if ($Hover) {
        $Card.BackColor = $HoverColor
        $Card.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 130, 190)
    } else {
        $Card.BackColor = $BaseColor
        $Card.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(205, 212, 220)
    }
}

$CardClickHandler = {
    $Path = $this.Tag
    if ($null -eq $Path -and $this.Parent -is [System.Windows.Forms.Control]) { $Path = $this.Parent.Tag }

    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    if (Test-Path $Path) {
        $FileName = [System.IO.Path]::GetFileName($Path)
        $StatusMain.Text = "Launching: $FileName"

        if ($FileName -eq "DYN-AttributeManager-GUI.ps1") {
            try {
                . $Path
                if (-not (Get-Command New-AttributeManagerForm -ErrorAction SilentlyContinue)) {
                    throw "New-AttributeManagerForm function was not found after loading script."
                }

                $ChildForm = New-AttributeManagerForm
                $ChildForm.StartPosition = "CenterParent"
                [void]$ChildForm.ShowDialog($Form)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Attribute Manager launch failed:`n$($_.Exception.Message)", "Error", 0, 16)
            }
            return
        }

        try {
            $SafePath = $Path.Replace("'", "''")
            $Cmd = "& ([ScriptBlock]::Create((Get-Content -LiteralPath '$SafePath' -Raw -Encoding UTF8)))"
            Start-Process -FilePath $ScriptHost -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-Command", $Cmd) -WindowStyle Hidden -WorkingDirectory $ScriptDir
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Tool launch failed:`n$($_.Exception.Message)", "Error", 0, 16)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Script Missing:`n$Path", "Error", 0, 16)
    }
}

$CardMouseEnter = {
    $Card = if ($this -is [System.Windows.Forms.Button]) { $this } else { $this.Parent }
    if ($Card -is [System.Windows.Forms.Button]) { Set-CardVisual -Card $Card -Hover $true }
}

$CardMouseLeave = {
    $Card = if ($this -is [System.Windows.Forms.Button]) { $this } else { $this.Parent }
    if ($Card -is [System.Windows.Forms.Button]) { Set-CardVisual -Card $Card -Hover $false }
}

function Add-ToolCard {
    param([string]$Title, [hashtable]$Data)

    $FullPath = Join-Path $ScriptDir $Data.File

    $Btn = New-Object System.Windows.Forms.Button
    $Btn.Size = "310, 158"
    $Btn.FlatStyle = "Flat"
    $Btn.FlatAppearance.BorderSize = 1
    $Btn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(234, 241, 252)
    $Btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(243, 248, 255)
    $Btn.Margin = New-Object System.Windows.Forms.Padding(12)
    $Btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Btn.Tag = $FullPath
    $Btn.Text = ""
    Set-CardVisual -Card $Btn -Hover $false
    $Btn.Add_Click($CardClickHandler)
    $Btn.Add_MouseEnter($CardMouseEnter)
    $Btn.Add_MouseLeave($CardMouseLeave)

    $Bar = New-Object System.Windows.Forms.Panel
    $Bar.Size = "306, 7"
    $Bar.Location = "2, 2"
    $Bar.BackColor = [System.Drawing.Color]::FromName($Data.Color)
    $Bar.Add_Click($CardClickHandler)
    $Bar.Add_MouseEnter($CardMouseEnter)
    $Bar.Add_MouseLeave($CardMouseLeave)

    $Lbl = New-Object System.Windows.Forms.Label
    $Lbl.Text = $Title
    $Lbl.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $Lbl.ForeColor = [System.Drawing.Color]::FromArgb(32, 45, 58)
    $Lbl.Location = "12, 22"
    $Lbl.AutoSize = $true
    $Lbl.Add_Click($CardClickHandler)
    $Lbl.Add_MouseEnter($CardMouseEnter)
    $Lbl.Add_MouseLeave($CardMouseLeave)

    $LblD = New-Object System.Windows.Forms.Label
    $LblD.Text = $Data.Desc
    $LblD.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $LblD.ForeColor = [System.Drawing.Color]::FromArgb(87, 97, 108)
    $LblD.Location = "12, 50"
    $LblD.Size = "286, 20"
    $LblD.Add_Click($CardClickHandler)
    $LblD.Add_MouseEnter($CardMouseEnter)
    $LblD.Add_MouseLeave($CardMouseLeave)

    $Line = New-Object System.Windows.Forms.Label
    $Line.BackColor = [System.Drawing.Color]::FromArgb(219, 225, 231)
    $Line.Location = "12, 78"
    $Line.Size = "286, 1"

    $LblPath = New-Object System.Windows.Forms.Label
    $LblPath.Text = "File: $($Data.File)"
    $LblPath.Font = New-Object System.Drawing.Font("Consolas", 8)
    $LblPath.ForeColor = [System.Drawing.Color]::FromArgb(97, 107, 118)
    $LblPath.Location = "12, 89"
    $LblPath.Size = "286, 18"
    $LblPath.Add_Click($CardClickHandler)
    $LblPath.Add_MouseEnter($CardMouseEnter)
    $LblPath.Add_MouseLeave($CardMouseLeave)

    $LblHint = New-Object System.Windows.Forms.Label
    $LblHint.Text = "Click to open"
    $LblHint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $LblHint.ForeColor = [System.Drawing.Color]::FromArgb(117, 125, 132)
    $LblHint.Location = "12, 125"
    $LblHint.AutoSize = $true
    $LblHint.Add_Click($CardClickHandler)
    $LblHint.Add_MouseEnter($CardMouseEnter)
    $LblHint.Add_MouseLeave($CardMouseLeave)

    $Btn.Controls.AddRange(@($Bar, $Lbl, $LblD, $Line, $LblPath, $LblHint))
    $Flow.Controls.Add($Btn)
}

$Tools.GetEnumerator() | Sort-Object Name | ForEach-Object { Add-ToolCard -Title $_.Value.Title -Data $_.Value }

# ==========================================
# BOTTOM PANEL: ENGINE
# ==========================================
$PnlEngine = $Split.Panel2
$PnlEngine.BackColor = [System.Drawing.Color]::FromArgb(228, 233, 238)
$PnlEngine.Padding = New-Object System.Windows.Forms.Padding(14)

$GrpEngine = New-Object System.Windows.Forms.GroupBox
$GrpEngine.Text = "Sync Engine Controller"
$GrpEngine.Dock = "Fill"
$GrpEngine.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$PnlEngine.Controls.Add($GrpEngine)

$TblEngine = New-Object System.Windows.Forms.TableLayoutPanel
$TblEngine.Dock = "Fill"
$TblEngine.RowCount = 4
$TblEngine.ColumnCount = 1
$TblEngine.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$TblEngine.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$TblEngine.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
$TblEngine.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$GrpEngine.Controls.Add($TblEngine)

$LblEngInfo = New-Object System.Windows.Forms.Label
$LblEngInfo.Text = "Script: $EngineScript"
$LblEngInfo.Dock = "Fill"
$LblEngInfo.TextAlign = "MiddleLeft"
$TblEngine.Controls.Add($LblEngInfo, 0, 0)

$ChkConsoleMode = New-Object System.Windows.Forms.CheckBox
$ChkConsoleMode.Text = "Open in separate console window"
$ChkConsoleMode.Dock = "Fill"
$ChkConsoleMode.Checked = $false
$TblEngine.Controls.Add($ChkConsoleMode, 0, 1)

$PnlButtons = New-Object System.Windows.Forms.Panel
$PnlButtons.Dock = "Fill"
$TblEngine.Controls.Add($PnlButtons, 0, 2)

$BtnRun = New-Object System.Windows.Forms.Button
$BtnRun.Text = "LAUNCH SYNC ENGINE"
$BtnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$BtnRun.BackColor = [System.Drawing.Color]::FromArgb(201, 236, 209)
$BtnRun.ForeColor = [System.Drawing.Color]::FromArgb(18, 102, 36)
$BtnRun.Size = "190, 32"
$BtnRun.Location = "0, 5"
$BtnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
$PnlButtons.Controls.Add($BtnRun)

$BtnStop = New-Object System.Windows.Forms.Button
$BtnStop.Text = "STOP ENGINE"
$BtnStop.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$BtnStop.BackColor = [System.Drawing.Color]::FromArgb(247, 209, 209)
$BtnStop.ForeColor = [System.Drawing.Color]::FromArgb(128, 30, 30)
$BtnStop.Size = "120, 32"
$BtnStop.Location = "200, 5"
$BtnStop.Enabled = $false
$BtnStop.Cursor = [System.Windows.Forms.Cursors]::Hand
$PnlButtons.Controls.Add($BtnStop)

$TxtEngineLog = New-Object System.Windows.Forms.RichTextBox
$TxtEngineLog.Dock = "Fill"
$TxtEngineLog.ReadOnly = $true
$TxtEngineLog.BackColor = [System.Drawing.Color]::FromArgb(30, 33, 36)
$TxtEngineLog.ForeColor = [System.Drawing.Color]::FromArgb(225, 231, 237)
$TxtEngineLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$TxtEngineLog.WordWrap = $false
$TblEngine.Controls.Add($TxtEngineLog, 0, 3)

function Write-EngineLog {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $TxtEngineLog.AppendText("[$((Get-Date).ToString('HH:mm:ss'))] $Message`r`n")
    $TxtEngineLog.SelectionStart = $TxtEngineLog.TextLength
    $TxtEngineLog.ScrollToCaret()
}

function Start-EngineInApp {
    if ($script:EngineProcess -and -not $script:EngineProcess.HasExited) {
        Write-EngineLog "Engine is already running."
        return
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ScriptHost
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$EngineScript`""
    $psi.WorkingDirectory = $ScriptDir
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $script:EngineProcess = New-Object System.Diagnostics.Process
    $script:EngineProcess.StartInfo = $psi
    $script:EngineProcess.EnableRaisingEvents = $true

    $script:EngineProcess.add_OutputDataReceived({
        param($sender, $e)
        if ($e.Data) {
            $msg = $e.Data
            $null = $Form.BeginInvoke([Action[string]]{
                param($line)
                Write-EngineLog $line
            }, @($msg))
        }
    })

    $script:EngineProcess.add_ErrorDataReceived({
        param($sender, $e)
        if ($e.Data) {
            $msg = "ERROR: $($e.Data)"
            $null = $Form.BeginInvoke([Action[string]]{
                param($line)
                Write-EngineLog $line
            }, @($msg))
        }
    })

    $script:EngineProcess.add_Exited({
        $null = $Form.BeginInvoke([Action]{
            Write-EngineLog "Engine process exited."
            $StatusMain.Text = "Engine stopped"
            $BtnRun.Enabled = $true
            $BtnStop.Enabled = $false
        })
    })

    $started = $script:EngineProcess.Start()
    if ($started) {
        $script:EngineProcess.BeginOutputReadLine()
        $script:EngineProcess.BeginErrorReadLine()
        Write-EngineLog "Engine started in-app."
        $StatusMain.Text = "Engine running in-app"
        $BtnRun.Enabled = $false
        $BtnStop.Enabled = $true
    } else {
        Write-EngineLog "Failed to start engine process."
    }
}

function Stop-EngineInApp {
    if ($script:EngineProcess -and -not $script:EngineProcess.HasExited) {
        try {
            $script:EngineProcess.Kill()
            $script:EngineProcess.WaitForExit(2000) | Out-Null
            Write-EngineLog "Engine was stopped by user."
        } catch {
            Write-EngineLog "Failed to stop engine: $($_.Exception.Message)"
        }
    }

    $BtnRun.Enabled = $true
    $BtnStop.Enabled = $false
    $StatusMain.Text = "Ready"
}

# ==========================================
# ACTIONS
# ==========================================
$BtnRun.Add_Click({
    if (-not (Test-Path $EngineScript)) {
        [System.Windows.Forms.MessageBox]::Show("Engine script not found:`n$EngineScript", "Error", 0, 16)
        return
    }

    if ($ChkConsoleMode.Checked) {
        try {
            Start-Process -FilePath $ScriptHost -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-File", $EngineScript) -WorkingDirectory $ScriptDir
            $StatusMain.Text = "Engine launched in console mode"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Launch Failed:`n$($_.Exception.Message)", "Error", 0, 16)
        }
    } else {
        Start-EngineInApp
    }
})

$BtnStop.Add_Click({ Stop-EngineInApp })

$Form.Add_FormClosing({
    if ($script:EngineProcess -and -not $script:EngineProcess.HasExited) {
        try {
            $script:EngineProcess.Kill()
            $script:EngineProcess.WaitForExit(1500) | Out-Null
        } catch {}
    }
})

# ==========================================
# RUN
# ==========================================
[void]$Form.ShowDialog()
