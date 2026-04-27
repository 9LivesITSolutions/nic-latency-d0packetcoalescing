#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Disables packet coalescing on physical network adapters — local or remote via WinRM.

.DESCRIPTION
    Disables packet coalescing on all physical NICs (WiFi + Ethernet) that support
    the setting, regardless of manufacturer. Uses direct registry writes followed by
    a scheduled task running as SYSTEM to restart affected NICs — ensuring the driver
    reloads the new value independently of WinRM session state.

    Execution modes (auto-detected from parameters):
      - No parameters    : local mode — runs on current machine, logs to Event Viewer
      - -ComputerName    : remote mode — targets explicit computer list via WinRM
      - -OUPath          : remote mode — targets all computers in an AD OU via WinRM

    Remote workflow per computer:
      1. Write registry values for all matching coalescing keys
      2. Register a one-shot scheduled task (SYSTEM) to restart affected NICs
      3. Wait for NIC restart + WinRM reconnect
      4. Read and report final status via Get-NetAdapterPowerManagement
      5. Cleanup scheduled task and temp files

    PowerShell version handling:
      - PS 7+  : remote targets processed in parallel (ForEach-Object -Parallel)
      - PS 5.1 : remote targets processed sequentially

.PARAMETER ComputerName
    One or more remote computer names to target via WinRM.

.PARAMETER OUPath
    Active Directory OU distinguishedName. All computers in the OU will be targeted.
    Example: "OU=Workstations,DC=cmcap,DC=local"

.PARAMETER Credential
    Optional PSCredential for remote authentication.
    If omitted, uses the current session credentials.

.PARAMETER ThrottleLimit
    Maximum number of parallel WinRM connections (PS7+ only). Default: 10.

.PARAMETER ExportCsv
    Optional path to export results as CSV.
    Example: "C:\Logs\coalescing-fix.csv"

.PARAMETER WhatIf
    Dry run - reports current status without applying any changes.

.EXAMPLE
    .\Disable-NICPacketCoalescing.ps1
.EXAMPLE
    .\Disable-NICPacketCoalescing.ps1 -ComputerName "PC01","PC02","CHAR-EM53"
.EXAMPLE
    .\Disable-NICPacketCoalescing.ps1 -OUPath "OU=Workstations,DC=cmcap,DC=local"
.EXAMPLE
    .\Disable-NICPacketCoalescing.ps1 -OUPath "OU=Workstations,DC=cmcap,DC=local" -WhatIf -ExportCsv "C:\Logs\audit.csv"
.EXAMPLE
    $cred = Get-Credential
    .\Disable-NICPacketCoalescing.ps1 -ComputerName "PC01","PC02" -Credential $cred -ThrottleLimit 5

.NOTES
    Author  : 9 Lives IT Solutions
    Version : 4.2.0
    Tested  : PowerShell 5.1 / 7.x - Windows 10/11 domain-joined workstations
    GPO     : Computer Configuration > Windows Settings > Scripts > Startup (local mode)
    WinRM   : Must be enabled on remote targets

    Registry key mapping (Intel driver variants):
      *PacketCoalescing    - Intel AX201, AX211 and variants  (0 = disabled)
      *D0PacketCoalescing  - Intel AX200, AC9560 and variants (0 = disabled)
      DMACoalescing        - Intel I225, I219 Ethernet        (0 = disabled)

.LINK
    https://github.com/9lives-it/nic-latency-d0packetcoalescing
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string[]]$ComputerName,
    [string]$OUPath,
    [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
    [int]$ThrottleLimit = 10,
    [string]$ExportCsv
)

# ---------------------------------------------------------------------------
# SCRIPTBLOCK - APPLY (remote)
# 1. Scans registry and writes 0 to all matching coalescing keys
# 2. Writes adapter list to temp file (avoids quoting hell in scheduled task)
# 3. Encodes restart script as Base64 and registers one-shot SYSTEM task
# Logs every step to C:\Windows\Temp\9Lives-NICFix.log
# ---------------------------------------------------------------------------
$ApplyBlock = {
    $classKey   = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
    $targetKeys = @("*PacketCoalescing", "*D0PacketCoalescing", "DMACoalescing")
    $taskName   = "9Lives-NICRestart"
    $logFile    = "C:\Windows\Temp\9Lives-NICFix.log"
    $listFile   = "C:\Windows\Temp\9Lives-NICList.txt"
    $fixedDescs = @()

    function Write-Log {
        param([string]$Msg)
        Add-Content -Path $logFile -Value "$(Get-Date -Format 'HH:mm:ss') | $Msg" -ErrorAction SilentlyContinue
    }

    Set-Content $logFile "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] === Disable-NICPacketCoalescing START ===" -ErrorAction SilentlyContinue

    # Step 1 - scan registry and write values
    Write-Log "STEP 1 - Scanning registry for coalescing keys"
    Get-ChildItem $classKey -ErrorAction SilentlyContinue | ForEach-Object {
        $regPath = $_.PSPath
        $props   = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
        if (-not $props.DriverDesc) { return }
        foreach ($key in $targetKeys) {
            $prop = $props.PSObject.Properties | Where-Object { $_.Name -eq $key }
            if ($prop) {
                Write-Log "  FOUND [$($props.DriverDesc)] key=[$key] value=[$($prop.Value)]"
                if ($prop.Value -ne "0") {
                    try {
                        Set-ItemProperty -Path $regPath -Name $key -Value "0" -ErrorAction Stop
                        $recheck = (Get-ItemProperty $regPath -ErrorAction SilentlyContinue).$key
                        Write-Log "  WRITTEN -> recheck=[$recheck]"
                        $fixedDescs += $props.DriverDesc
                    } catch {
                        Write-Log "  WRITE FAILED: $_"
                    }
                } else {
                    Write-Log "  SKIP - already 0"
                }
                break
            }
        }
    }

    Write-Log "STEP 1 DONE - $($fixedDescs.Count) adapter(s) patched: $($fixedDescs -join ', ')"

    if ($fixedDescs.Count -eq 0) {
        Write-Log "No adapters to restart - exiting"
        return
    }

    # Step 2 - write adapter list to temp file (read by scheduled task)
    $fixedDescs | Set-Content $listFile -ErrorAction SilentlyContinue
    Write-Log "STEP 2 - NIC list written to $listFile"

    # Step 3 - build restart script and encode as Base64 to avoid quoting issues
    $taskScript = @'
$logFile  = "C:\Windows\Temp\9Lives-NICFix.log"
$listFile = "C:\Windows\Temp\9Lives-NICList.txt"
function Write-Log { param([string]$Msg); Add-Content -Path $logFile -Value "$(Get-Date -Format 'HH:mm:ss') | $Msg" -ErrorAction SilentlyContinue }
Write-Log "TASK - Started as SYSTEM"
if (Test-Path $listFile) {
    Get-Content $listFile | ForEach-Object {
        $desc    = $_.Trim()
        $adapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -eq $desc }
        if ($adapter) {
            Write-Log "TASK - Restarting [$desc]"
            Restart-NetAdapter -Name $adapter.Name -Confirm:$false
            Write-Log "TASK - Restart done [$desc]"
        } else {
            Write-Log "TASK - Adapter not found [$desc]"
        }
    }
    Remove-Item $listFile -Force -ErrorAction SilentlyContinue
} else {
    Write-Log "TASK - NICList file not found at $listFile"
}
Write-Log "TASK - Finished"
'@

    $encodedCmd = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($taskScript))

    # Step 4 - register scheduled task
    Write-Log "STEP 3 - Registering scheduled task [$taskName]"
    try {
        $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCmd"
        $trigger   = New-ScheduledTaskTrigger -Once -At (Get-Date).AddSeconds(3)
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 1)
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
            -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        Write-Log "STEP 3 DONE - Task registered, fires in 3s"
    } catch {
        Write-Log "STEP 3 FAILED - $_"
    }
}

# ---------------------------------------------------------------------------
# SCRIPTBLOCK - READ STATUS (second WinRM session, after NIC restart)
# Reads and prints log written by ApplyBlock + scheduled task
# Reads live adapter status via Get-NetAdapterPowerManagement
# Cleans up scheduled task and temp log
# ---------------------------------------------------------------------------
$ReadBlock = {
    $taskName = "9Lives-NICRestart"
    $logFile  = "C:\Windows\Temp\9Lives-NICFix.log"
    $rows     = @()

    if (Test-Path $logFile) {
        Get-Content $logFile -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "  [LOG] $_"
        }
    } else {
        Write-Host "  [LOG] Log file not found - ApplyBlock may not have run"
    }

    Get-NetAdapter -Physical | ForEach-Object {
        $adapterDesc = $_.InterfaceDescription
        try {
            $pm     = Get-NetAdapterPowerManagement -Name $_.Name -ErrorAction Stop
            $status = switch ($pm.D0PacketCoalescing) {
                "Disabled"    { "OK - Disabled"        }
                "Enabled"     { "WARN - Still Enabled" }
                "Unsupported" { "N/A - Unsupported"    }
                default       { "N/A - Unsupported"    }
            }
        } catch {
            $status = "ERROR: $($_.Exception.Message)"
        }
        $rows += [PSCustomObject]@{
            Computer = $env:COMPUTERNAME
            Adapter  = $adapterDesc
            Status   = $status
        }
    }

    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item $logFile -Force -ErrorAction SilentlyContinue
    return $rows
}

# ---------------------------------------------------------------------------
# LOCAL SCRIPTBLOCK - apply + restart + read in single pass (no WinRM)
# ---------------------------------------------------------------------------
$LocalBlock = {
    param([bool]$DryRun)

    $classKey   = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
    $targetKeys = @("*PacketCoalescing", "*D0PacketCoalescing", "DMACoalescing")
    $rows       = @()
    $fixedDescs = @()

    Get-ChildItem $classKey -ErrorAction SilentlyContinue | ForEach-Object {
        $regPath = $_.PSPath
        $props   = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
        if (-not $props.DriverDesc) { return }
        if ($props.DriverDesc -like "*Virtual*" -or $props.DriverDesc -like "*Miniport*") { return }
        foreach ($key in $targetKeys) {
            $prop = $props.PSObject.Properties | Where-Object { $_.Name -eq $key }
            if ($prop) {
                if ($prop.Value -eq "0") {
                    $status = "ALREADY_OK"
                } elseif ($DryRun) {
                    $status = "WHATIF - Would disable (current: $($prop.Value))"
                } else {
                    Set-ItemProperty -Path $regPath -Name $key -Value "0" -ErrorAction SilentlyContinue
                    $fixedDescs += $props.DriverDesc
                    $status = "FIXED"
                }
                $rows += [PSCustomObject]@{
                    Computer = $env:COMPUTERNAME
                    Adapter  = $props.DriverDesc
                    Status   = $status
                }
                break
            }
        }
    }

    if (-not $DryRun -and $fixedDescs.Count -gt 0) {
        foreach ($desc in $fixedDescs) {
            $adapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -eq $desc }
            if ($adapter) {
                Write-Host "  Restarting [$desc]..." -ForegroundColor Yellow
                Restart-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        Start-Sleep -Seconds 5
        $rows = $rows | ForEach-Object {
            $row = $_
            if ($row.Status -eq "FIXED") {
                $adapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -eq $row.Adapter }
                if ($adapter) {
                    $pm = Get-NetAdapterPowerManagement -Name $adapter.Name -ErrorAction SilentlyContinue
                    $row.Status = if ($pm.D0PacketCoalescing -eq "Disabled") { "FIXED - Confirmed Disabled" } else { "FIXED - Restart pending" }
                }
            }
            $row
        }
    }
    return $rows
}

# ---------------------------------------------------------------------------
# MODE DETECTION
# ---------------------------------------------------------------------------
$isRemote = ($ComputerName.Count -gt 0) -or ($OUPath -ne "")
$isDryRun = [bool]$WhatIfPreference

# ---------------------------------------------------------------------------
# LOCAL MODE
# ---------------------------------------------------------------------------
if (-not $isRemote) {
    $LogSource = "9Lives-NetworkPerf"
    $LogName   = "Application"
    if (-not [System.Diagnostics.EventLog]::SourceExists($LogSource)) {
        New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    }
    $rows    = & $LocalBlock -DryRun $isDryRun
    $summary = $rows | ForEach-Object { "$($_.Status) | $($_.Adapter)" }
    $body    = "Disable-NICPacketCoalescing -- $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n`n" + ($summary -join "`n")
    Write-EventLog -LogName $LogName -Source $LogSource -EventId 1001 -EntryType Information -Message $body -ErrorAction SilentlyContinue
    $rows | Format-Table -AutoSize
    return
}

# ---------------------------------------------------------------------------
# REMOTE MODE - build computer list from OUPath if needed
# ---------------------------------------------------------------------------
if ($OUPath) {
    try {
        $ComputerName = (Get-ADComputer -Filter * -SearchBase $OUPath -ErrorAction Stop).Name
        Write-Host "[$($ComputerName.Count) computers found in OU]" -ForegroundColor Cyan
    } catch {
        Write-Error "Failed to query AD OU '$OUPath': $_"
        exit 1
    }
}

# ---------------------------------------------------------------------------
# HELPER - two-pass remote fix (PS5.1 compatible)
# Pass 1 : registry write + schedule SYSTEM task for NIC restart
# Pass 2 : reconnect after delay + read live status + cleanup
# ---------------------------------------------------------------------------
function Invoke-FixOnComputer {
    param(
        [string]$PC,
        [bool]$DryRun,
        [scriptblock]$Apply,
        [scriptblock]$Read,
        [PSCredential]$Cred,
        [int]$ReconnectDelaySec = 15
    )
    $invokeParams = @{ ComputerName = $PC; ErrorAction = "Stop" }
    if ($Cred -ne [System.Management.Automation.PSCredential]::Empty) {
        $invokeParams.Credential = $Cred
    }
    Write-Host "`n[$PC] Pass 1 - Applying fix..." -ForegroundColor Cyan
    if (-not $DryRun) {
        try {
            Invoke-Command @invokeParams -ScriptBlock $Apply
            Write-Host "[$PC] Pass 1 done - waiting ${ReconnectDelaySec}s for NIC restart..." -ForegroundColor Cyan
        } catch {
            Write-Host "[$PC] Pass 1 - WinRM dropped (expected on WiFi targets)" -ForegroundColor Yellow
        }
        Start-Sleep -Seconds $ReconnectDelaySec
    }
    Write-Host "[$PC] Pass 2 - Reading status..." -ForegroundColor Cyan
    try {
        Invoke-Command @invokeParams -ScriptBlock $Read
    } catch {
        [PSCustomObject]@{ Computer = $PC; Adapter = "N/A"; Status = "UNREACHABLE: $_" }
    }
}

# ---------------------------------------------------------------------------
# REMOTE MODE - PS7 parallel vs PS5.1 sequential
# ---------------------------------------------------------------------------
$allResults = @()

if ($PSVersionTable.PSVersion.Major -ge 7) {
    Write-Host "[PS7 - parallel execution, ThrottleLimit=$ThrottleLimit]" -ForegroundColor Cyan
    $allResults = $ComputerName | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        $PC     = $_
        $Apply  = $using:ApplyBlock
        $Read   = $using:ReadBlock
        $DryRun = $using:isDryRun
        $Cred   = $using:Credential
        $invokeParams = @{ ComputerName = $PC; ErrorAction = "Stop" }
        if ($Cred -ne [System.Management.Automation.PSCredential]::Empty) {
            $invokeParams.Credential = $Cred
        }
        if (-not $DryRun) {
            try { Invoke-Command @invokeParams -ScriptBlock $Apply } catch {}
            Start-Sleep -Seconds 15
        }
        try {
            Invoke-Command @invokeParams -ScriptBlock $Read
        } catch {
            [PSCustomObject]@{ Computer = $PC; Adapter = "N/A"; Status = "UNREACHABLE: $_" }
        }
    }
} else {
    Write-Host "[PS5.1 - sequential execution]" -ForegroundColor Cyan
    foreach ($PC in $ComputerName) {
        $allResults += Invoke-FixOnComputer -PC $PC -DryRun $isDryRun `
            -Apply $ApplyBlock -Read $ReadBlock -Cred $Credential
    }
}

# ---------------------------------------------------------------------------
# OUTPUT
# ---------------------------------------------------------------------------
Write-Host "`n=== RESULTS ===" -ForegroundColor Green
$allResults | Select-Object Computer, Adapter, Status | Format-Table -AutoSize

if ($ExportCsv) {
    $allResults | Select-Object Computer, Adapter, Status |
        Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "[Exported to $ExportCsv]" -ForegroundColor Green
}
