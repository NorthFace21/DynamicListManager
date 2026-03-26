<# 
   .SYNOPSIS
   Dynamic List Sync Engine (v1.0)
   .DESCRIPTION
   Headless backend worker.
   1. Provisions AD Groups in correct OUs based on Type.
   2. Enables Mail attributes via Exchange (if MailEnabled).
   3. Calculates membership based on JSON Rules + Exceptions.
   4. Applies Delta changes (Add/Remove) and logs to SQL.
#>
# ==========================================
# GLOBAL CONFIGURATION
# ==========================================
param (     $SqlInstance = "SQL16-DIVERS",
            [string]$Database = "DBTEST",
            [string]$ExChangeserver = "RB1EX19DB01",
            [string]$OUSecurity = "OU=DynamicSecurity,OU=DynalistV2,OU=RCGT Groups,DC=rcgt,DC=net",
            [string]$OUMail = "OU=DynamicDistribution,OU=DynalistV2,OU=RCGT Groups,DC=rcgt,DC=net",
            [int]$MaxChangesPerRun = "500"
    )

# --- LOAD MODULES ---
Import-Module ActiveDirectory

$Config = @{
    # Where to create the groups
    OUSecurity    = $OUSecurity
    OUMail        = $OUMail
    
    # Safety limits
    MaxChangesPerRun = $MaxChangesPerRun  # Stop if removing 500 users at once (prevent disasters)
}
# ==========================================

# --- SQL HELPER ---
function Sync-SQLMembers {
    param(
        [int]$ListID,
        [array]$CurrentADMembers # The final calculated array of SAMAccountNames that SHOULD be in the group
    )

    # 1. Fetch who SQL *thinks* is in the list right now
    $Q_Get = "SELECT UserIdentity FROM DL_Members WHERE DistListID = @ID"
    $SqlMembers = (Invoke-Sql $Q_Get -Parameters @{ID=$ListID}).Data.UserIdentity
    if ($SqlMembers -eq $null) { $SqlMembers = @() }

    # 2. Calculate Deltas
    $ToAdd = $CurrentADMembers | Where-Object { $_ -notin $SqlMembers }
    $ToRemove = $SqlMembers | Where-Object { $_ -notin $CurrentADMembers }

    if ($ToAdd.Count -eq 0 -and $ToRemove.Count -eq 0) { return } # No changes

    # 3. Process ADDITIONS
    if ($ToAdd.Count -gt 0) {
        $ValuesMember = New-Object System.Collections.ArrayList
        
        foreach ($User in $ToAdd) {
            $SafeUser = $User.ToString().Replace("'", "''")
            # Add to Snapshot
            $ValuesMember.Add("($ListID, '$SafeUser', 'SyncEngine')") | Out-Null
        }
        
        # Bulk Insert
        if ($ValuesMember.Count -gt 0) {
            $BatchM = $ValuesMember -join ","
            Invoke-Sql "INSERT INTO DL_Members (DistListID, UserIdentity, DisplayName) VALUES $BatchM" | Out-Null
        }
    }

    # 4. Process REMOVALS
    if ($ToRemove.Count -gt 0) {
        foreach ($User in $ToRemove) {
            $SafeUser = $User.ToString().Replace("'", "''")
           
            # Delete from Snapshot (One by one is safer for deletes, or use IN clause for speed)
            Invoke-Sql "DELETE FROM DL_Members WHERE DistListID = $ListID AND UserIdentity = '$SafeUser'" | Out-Null
        }
        
    }
    
    Write-Host "SQL Sync Complete: +$($ToAdd.Count) / -$($ToRemove.Count)" -ForegroundColor Cyan
}

function Invoke-Sql {
    param([string]$Query, [hashtable]$Parameters = @{}, [switch]$ReturnScalar, [switch]$ReturnId)
    $ResultInfo = @{ Success = $false; Data = $null; Message = "" }
    try {
        $ConnString = "Server=$SqlInstance;Database=$Database;Integrated Security=True;"
        $Conn = New-Object System.Data.SqlClient.SqlConnection($ConnString)
        $Conn.Open()
        $Cmd = $Conn.CreateCommand(); $Cmd.CommandText = $Query
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
    catch { $ResultInfo.Message = $_.Exception.Message; Write-Error "SQL Error: $($_.Exception.Message)" }
    return $ResultInfo
}

# --- LOGGING HELPER ---
function Log-Activity {
    param($ListID, $User, $Action, $Reason)
    $Q = "INSERT INTO Log_MemberActivity (DistListID, UserIdentity, ActionType, Reason) VALUES (@ID, @User, @Act, @Rsn)"
    Invoke-Sql -Query $Q -Parameters @{ID=$ListID; User=$User; Act=$Action; Rsn=$Reason} | Out-Null
    
    $Color = if ($Action -eq "ADD") { "Green" } else { "Red" }
    Write-Host "   [$Action] $User ($Reason)" -ForegroundColor $Color
}

# ==========================================
# PHASE 1: LOGIC PARSER (JSON -> AD FILTER)
# ==========================================
function Convert-JsonToFilter {
    param($JsonString)
    
    if ([string]::IsNullOrWhiteSpace($JsonString)) { return $null }
    
    try {
        $Tree = $JsonString | ConvertFrom-Json
        return Parse-Node -Node $Tree
    } catch {
        Write-Error "JSON Parse Error: $_"
        return $null
    }
}

function Parse-Node {
    param($Node)
    
    # Recursive Group
    if ($Node.type -eq "group") {
        $ChildFilters = @()
        foreach ($Child in $Node.children) {
            $F = Parse-Node -Node $Child
            if (-not [string]::IsNullOrWhiteSpace($F)) { $ChildFilters += $F }
        }
        if ($ChildFilters.Count -eq 0) { return $null }
        
        $Gate = if ($Node.gate -eq "AND") { "-and" } else { "-or" }
        return "($($ChildFilters -join " $Gate "))"
    } 
    # Rule
    else {
        $SafeVal = $Node.value -replace "'", "''" # Escape for AD Filter
        # Map operator if needed, mostly they match
        return "($($Node.attribute) -$($Node.operator) '$SafeVal')"
    }
}

# ==========================================
# PHASE 2: PROVISIONING & SYNC
# ==========================================

function Start-Engine {
    Write-Host "--- ENGINE START: $(Get-Date) ---" -ForegroundColor Cyan
    
    # 1. Get Due Lists (SyncEnabled = 1 AND (LastSync is NULL OR LastSync + Interval < Now))
    $Q = @"
    SELECT DistListID, Name, Description, ListType, PrimarySMTPAddress, RuleDefinition, SyncIntervalMinutes 
    FROM DL_MasterList 
    WHERE SyncEnabled = 1 
    AND (LastSyncDate IS NULL OR DATEADD(minute, SyncIntervalMinutes, LastSyncDate) < GETDATE())
"@
    $Lists = Invoke-Sql $Q
    
    if (-not $Lists.Success -or $Lists.Data.Rows.Count -eq 0) {
        Write-Host "No lists due for sync."
        return
    }

    foreach ($Row in $Lists.Data) {
        Process-List -ListConfig $Row
    }
    
    Write-Host "--- ENGINE COMPLETE ---" -ForegroundColor Cyan
}

function Process-List {
    param($ListConfig)
    
    $ID = $ListConfig.DistListID
    $Name = $ListConfig.Name
    Write-Host "Processing: $Name (ID: $ID)..." -NoNewline

    # ----------------------------------
    # STEP A: PROVISIONING (Create/Config)
    # ----------------------------------
    $ADGroup = $null
    try {
        $ADGroup = Get-ADGroup -Identity $Name -ErrorAction Stop
    } catch {
        # Squelch error. Variable remains null, triggering the creation block below.
    }
    $TargetOU = if ($ListConfig.ListType -eq "MailEnabled") { $Config.OUMail } else { $Config.OUSecurity }
    
    if (-not $ADGroup) {
        Write-Host " [CREATING]" -ForegroundColor Yellow
        try {
            # Create Group
            $Scope = "Universal" # Required for Mail Enabled usually
            New-ADGroup -Name $Name -SamAccountName $Name -GroupCategory "Security" -GroupScope $Scope -Path $TargetOU -Description $ListConfig.Description
            
            # Wait for replication/creation
            Start-Sleep -Seconds 120
            $ADGroup = Get-ADGroup -Identity $Name
        } catch {
            Write-Error "Failed to create group '$Name': $_"
            return
        }
    }

    # Exchange Provisioning (Placeholder)
    # If Type is MailEnabled, checks availability of Enable-DistributionGroup
    if ($ListConfig.ListType -eq "MailEnabled") {
        try {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$($ExChangeserver)/PowerShell/ -Authentication Kerberos
            $importex16 = Import-PSSession $Session -commandname Enable-DistributionGroup,Get-DistributionGroup -AllowClobber
            } catch {}
        if (Get-Command Enable-DistributionGroup -ErrorAction SilentlyContinue) {
            try {
                # Check if already mail enabled
                if (-not (Get-DistributionGroup $Name -ErrorAction SilentlyContinue)) {
                    $alias = $ListConfig.PrimarySMTPAddress.Split("@")[0]

                    Enable-DistributionGroup -Identity $Name -PrimarySmtpAddress $ListConfig.PrimarySMTPAddress -Alias $alias -ErrorAction Stop
                    Write-Host " [MAIL ENABLED]" -ForegroundColor Magenta
                }
            } catch {
                Write-Warning "Exchange Error: $_"
            }
        } else {
            # Write-Warning "Exchange Cmdlets not available. Skipping mail attributes."
        }
    }

    # ----------------------------------
    # STEP B: CALCULATE MEMBERSHIP
    # ----------------------------------
    $Filter = Convert-JsonToFilter -JsonString $ListConfig.RuleDefinition
    
    # 1. Get Calculated Members (AD Query)
    $CalculatedUsers = New-Object System.Collections.Generic.HashSet[string]
    $FinalListSam = New-Object System.Collections.Generic.HashSet[string]
    if ($Filter) {
        try {
            $Users = Get-ADUser -Filter $Filter -Properties EmailAddress
            foreach ($u in $Users) { 
                    $CalculatedUsers.Add($u.DistinguishedName) | Out-Null 
                    $FinalListSam.Add($u.SamAccountName) | Out-Null
                }
        } catch {
            Write-Error "AD Filter Failed for '$Name': $_"
            return # Abort this list
        }
    }

    # 2. Apply Exceptions
    # Fetch Inclusions
    $Inc = Invoke-Sql "SELECT UserIdentity FROM DL_Inclusions WHERE DistListID=$ID"
    if ($Inc.Success) {
        foreach ($Row in $Inc.Data) {
            try {
                $sam = ($Row.UserIdentity.split("[")[1]).replace("]","")
                $U = Get-ADUser -Identity $sam # Resolve SAM to DN
                $CalculatedUsers.Add($U.DistinguishedName) | Out-Null
                $FinalListSam.Add($u.SamAccountName) | Out-Null
            } catch {}
        }
    }

    # Fetch Exclusions
    $Exc = Invoke-Sql "SELECT UserIdentity FROM DL_Exclusions WHERE DistListID=$ID"
    if ($Exc.Success) {
        foreach ($Row in $Exc.Data) {
            try {
                $sam = ($Row.UserIdentity.split("[")[1]).replace("]","")
                $U = Get-ADUser -Identity $sam
                if ($CalculatedUsers.Contains($U.DistinguishedName)) {
                    $CalculatedUsers.Remove($U.DistinguishedName) | Out-Null
                    $FinalListSam.Remove($U.SamAccountName) | Out-Null
                }
            } catch {}
        }
    }

    # ----------------------------------
    # STEP C: DELTA SYNC (Diff)
    # ----------------------------------
    $CurrentMembers = Get-ADGroupMember -Identity $Name -Recursive:$false | Where-Object {$_.objectClass -eq "user"}
    $CurrentDNs = New-Object System.Collections.Generic.HashSet[string]
    foreach ($m in $CurrentMembers) { $CurrentDNs.Add($m.DistinguishedName) | Out-Null }

    # Calculate Diff
    $ToRem = $CurrentMembers | Where-Object { -not $CalculatedUsers.Contains($_.DistinguishedName) }
    $ToAdd = $CalculatedUsers | Where-Object { -not $CurrentDNs.Contains($_) }

    # Safety Check
    if ($ToRem.Count -gt $Config.MaxChangesPerRun) {
        Write-Error "Safety Triggered: Attempting to remove $($ToRem.Count) users. Limit is $($Config.MaxChangesPerRun). Aborting."
        return
    }

    Write-Host " [Syncing: +$($ToAdd.Count) / -$($ToRem.Count)]" -ForegroundColor Green

    # EXECUTE REMOVES
    foreach ($User in $ToRem) {
        Remove-ADGroupMember -Identity $Name -Members $User -Confirm:$false
        Log-Activity $ID $User.SamAccountName "REMOVE" "SyncEngine"
    }

    # EXECUTE ADDS
    $AddBatch = @()
    foreach ($DN in $ToAdd) {
        $AddBatch += $DN
        # Log individually for audit, but add in batch for speed? 
        # We'll just log here.
        # Need to get SAM for logging nicely
        $SAM = (get-aduser $dn).SamAccountName
        Log-Activity $ID $SAM "ADD" "SyncEngine"
    }
    
    if ($AddBatch.Count -gt 0) {
        # Batch add for performance
        Add-ADGroupMember -Identity $Name -Members $AddBatch
    }
    # ----------------------------------
    # STEP D: Update Members List.
    # ----------------------------------
    Sync-SQLMembers -ListID $ID -CurrentADMembers $FinalListSam
    # ----------------------------------
    # STEP E: UPDATE TIMESTAMP
    # ----------------------------------
    Invoke-Sql "UPDATE DL_MasterList SET LastSyncDate = GETDATE() WHERE DistListID=$ID" | Out-Null
    Write-Host " Done."
}

# --- RUN ---
Start-Engine

