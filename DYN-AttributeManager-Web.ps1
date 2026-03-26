<#
   .SYNOPSIS
   Attribute Manager Web Dashboard (PowerShell Universal)
   .DESCRIPTION
   Full web interface for managing the Sys_AttributeMap table, converted from WinForm GUI.
   Dependencies: PowerShell Universal module (Install-Module Universal)
#>

# --- LOAD ASSEMBLIES/MODULES ---
Import-Module Universal

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

# ==========================================
# DASHBOARD CONTENT (Now as a Page for PSU Server)
# ==========================================
$Page = New-UDPage -Name "AD Attribute Manager" -Content {

    # Status Display
    New-UDElement -Id "status" -Content {
        New-UDAlert -Severity info -Text "Ready" -Id "statusAlert"
    }

    # Attribute Editor Panel
    New-UDCard -Title "Attribute Editor" -Content {
        New-UDRow {
            New-UDColumn -Size 6 {
                New-UDTextbox -Id "txtFriendly" -Label "Friendly Name" -Placeholder "Enter friendly name"
            }
            New-UDColumn -Size 6 {
                New-UDTextbox -Id "txtAD" -Label "AD Attribute" -Placeholder "Enter AD attribute"
            }
        }
        New-UDRow {
            New-UDColumn -Size 6 {
                New-UDSelect -Id "cmbType" -Label "Type" -Option {
                    New-UDSelectOption -Name "String" -Value "String"
                    New-UDSelectOption -Name "Int" -Value "Int"
                    New-UDSelectOption -Name "Bool" -Value "Bool"
                } -DefaultValue "String"
            }
            New-UDColumn -Size 6 {
                New-UDTextbox -Id "txtID" -Label "ID" -Disabled -Value ""
            }
        }
        New-UDRow {
            New-UDColumn -Size 2 {
                New-UDButton -Id "btnAdd" -Text "Save New" -OnClick {
                    $friendly = (Get-UDElement -Id "txtFriendly").Value
                    $ad = (Get-UDElement -Id "txtAD").Value
                    $type = (Get-UDElement -Id "cmbType").Value

                    if ($friendly -and $ad) {
                        $Q = "INSERT INTO Sys_AttributeMap (FriendlyName, ADAttribute, DataType) VALUES ('$friendly', '$ad', '$type')"
                        $Res = Invoke-Sql $Q
                        if ($Res.Success) {
                            Sync-UDElement -Id "attributeTable"
                            Clear-Form
                            Set-Status "New attribute added successfully"
                        } else {
                            Set-Status "Insert Failed: $($Res.Message)" -Error
                        }
                    } else {
                        Set-Status "Please fill in Friendly Name and AD Attribute" -Error
                    }
                } -BackgroundColor "LightGreen"
            }
            New-UDColumn -Size 2 {
                New-UDButton -Id "btnUpdate" -Text "Update" -OnClick {
                    $id = (Get-UDElement -Id "txtID").Value
                    $friendly = (Get-UDElement -Id "txtFriendly").Value
                    $ad = (Get-UDElement -Id "txtAD").Value
                    $type = (Get-UDElement -Id "cmbType").Value

                    if ($id) {
                        $Q = "UPDATE Sys_AttributeMap SET FriendlyName='$friendly', ADAttribute='$ad', DataType='$type' WHERE AttrID=$id"
                        $Res = Invoke-Sql $Q
                        if ($Res.Success) {
                            Sync-UDElement -Id "attributeTable"
                            Clear-Form
                            Set-Status "Attribute updated successfully"
                        } else {
                            Set-Status "Update Failed: $($Res.Message)" -Error
                        }
                    } else {
                        Set-Status "No attribute selected for update" -Error
                    }
                } -Disabled
            }
            New-UDColumn -Size 2 {
                New-UDButton -Id "btnDelete" -Text "Delete" -OnClick {
                    $id = (Get-UDElement -Id "txtID").Value
                    if ($id) {
                        # In web, we can use a confirmation modal, but for simplicity, direct delete
                        $Q = "DELETE FROM Sys_AttributeMap WHERE AttrID=$id"
                        $Res = Invoke-Sql $Q
                        if ($Res.Success) {
                            Sync-UDElement -Id "attributeTable"
                            Clear-Form
                            Set-Status "Attribute deleted successfully"
                        } else {
                            Set-Status "Delete Failed: $($Res.Message)" -Error
                        }
                    } else {
                        Set-Status "No attribute selected for deletion" -Error
                    }
                } -Disabled -BackgroundColor "LightSalmon"
            }
            New-UDColumn -Size 2 {
                New-UDButton -Id "btnClear" -Text "Clear" -OnClick {
                    Clear-Form
                }
            }
            New-UDColumn -Size 2 {
                New-UDButton -Id "btnRefresh" -Text "↻ Refresh Table" -OnClick {
                    Sync-UDElement -Id "attributeTable"
                    Set-Status "Table refreshed"
                } -BackgroundColor "LightBlue"
            }
        }
    }

    # Table Display
    New-UDCard -Title "Attribute Mappings" -Content {
        New-UDElement -Id "attributeTable" -Content {
            $Res = Invoke-Sql "SELECT AttrID, FriendlyName, ADAttribute, DataType FROM Sys_AttributeMap"

            if ($Res.Success) {
                $Data = $Res.Data | ForEach-Object {
                    [PSCustomObject]@{
                        AttrID       = $_.AttrID
                        FriendlyName = $_.FriendlyName
                        ADAttribute  = $_.ADAttribute
                        DataType     = $_.DataType
                    }
                }

                New-UDTable -Data $Data -Columns @(
                    New-UDTableColumn -Property "AttrID" -Title "ID" -Hidden
                    New-UDTableColumn -Property "FriendlyName" -Title "Friendly Name"
                    New-UDTableColumn -Property "ADAttribute" -Title "AD Attribute"
                    New-UDTableColumn -Property "DataType" -Title "Type"
                ) -ShowPagination -PageSize 10 -OnRowClick {
                    $Session:SelectedID = $EventData.AttrID
                    $Session:SelectedFriendly = $EventData.FriendlyName
                    $Session:SelectedAD = $EventData.ADAttribute
                    $Session:SelectedType = $EventData.DataType

                    Sync-UDElement -Id "txtID"
                    Sync-UDElement -Id "txtFriendly"
                    Sync-UDElement -Id "txtAD"
                    Sync-UDElement -Id "cmbType"
                    Sync-UDElement -Id "btnAdd"
                    Sync-UDElement -Id "btnUpdate"
                    Sync-UDElement -Id "btnDelete"
                }
            } else {
                New-UDAlert -Severity error -Text "Failed to load data: $($Res.Message)"
            }
        }
    }
}

# Helper Functions (defined in dashboard scope)
function Set-Status {
    param($Message, [switch]$Error)
    $Severity = if ($Error) { "error" } else { "info" }
    Set-UDElement -Id "statusAlert" -Properties @{
        Text = "[$([DateTime]::Now.ToString('HH:mm:ss'))] $Message"
        Severity = $Severity
    }
}

function Clear-Form {
    Set-UDElement -Id "txtID" -Properties @{ Value = "" }
    Set-UDElement -Id "txtFriendly" -Properties @{ Value = "" }
    Set-UDElement -Id "txtAD" -Properties @{ Value = "" }
    Set-UDElement -Id "cmbType" -Properties @{ Value = "String" }
    Set-UDElement -Id "btnAdd" -Properties @{ Disabled = $false }
    Set-UDElement -Id "btnUpdate" -Properties @{ Disabled = $true }
    Set-UDElement -Id "btnDelete" -Properties @{ Disabled = $true }
    $Session:SelectedID = $null
}

# Publish the page to the PSU server
Publish-UDPage -Page $Page</content>
<parameter name="filePath">vscode-vfs://github%2B7b2276223a312c22726566223a7b2274797065223a342c226964223a22574542227d7d/NorthFace21/DynamicListManager/DYN-AttributeManager-Web.ps1