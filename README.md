# Disable-NICPacketCoalescing

> *Cut WiFi/Ethernet latency on Windows endpoints by disabling packet coalescing — works on any NIC, any vendor.*

---

## What is packet coalescing?

Packet coalescing batches incoming network packets instead of processing them immediately, reducing CPU usage at the cost of added latency. On workstations plugged into AC power, there is no benefit to leaving it enabled.

```
Without coalescing:  packet → immediate processing → ~1ms latency
With coalescing:     packet → buffer → batch processing → ~10-50ms latency
```

The setting is stored in the NIC driver registry key and varies by chipset vendor:

| Registry key | Chipset |
|---|---|
| `*PacketCoalescing` | Intel AX201, AX211 and variants |
| `*D0PacketCoalescing` | Intel AX200, AC9560 and variants |
| `DMACoalescing` | Intel I225, I219 Ethernet |

---

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/Disable-NICPacketCoalescing.ps1` | Main script — disables coalescing locally or remotely |
| `scripts/Check-NetworkPowerStatus.ps1` | Audit script — checks current status across remote workstations |

---

## How it works

### Local mode (no parameters)
Scans the driver registry, writes `0` to all matching coalescing keys, restarts affected NICs, and logs results to the Windows Application Event Log. Designed for GPO startup script deployment.

### Remote mode (`-ComputerName` or `-OUPath`)
Uses two separate WinRM sessions per target:

```
Pass 1 │ Write registry values
       │ Write adapter list to C:\Windows\Temp\9Lives-NICList.txt
       │ Register one-shot SYSTEM scheduled task (fires in 3s)
       │   → Task reads NICList, calls Restart-NetAdapter per adapter
       │   → Task logs every step to C:\Windows\Temp\9Lives-NICFix.log
       ↓
  [15s delay — NIC restart + WinRM reconnect]
       ↓
Pass 2 │ Reconnect via WinRM
       │ Print full execution log from 9Lives-NICFix.log
       │ Read live status via Get-NetAdapterPowerManagement
       │ Cleanup scheduled task + temp files
       ↓
  Console output with per-adapter status
```

The scheduled task runs as SYSTEM and survives any WinRM session drop caused by the NIC restart.

---

## Usage

```powershell
# Local (GPO startup)
.\Disable-NICPacketCoalescing.ps1

# Remote — explicit list
.\Disable-NICPacketCoalescing.ps1 -ComputerName "PC01","PC02","WORKSTATION01"

# Remote — entire AD OU
.\Disable-NICPacketCoalescing.ps1 -OUPath "OU=Workstations,DC=domain,DC=local"

# Dry run — no changes applied
.\Disable-NICPacketCoalescing.ps1 -OUPath "OU=Workstations,DC=domain,DC=local" -WhatIf

# Export results to CSV
.\Disable-NICPacketCoalescing.ps1 -OUPath "OU=Workstations,DC=domain,DC=local" -ExportCsv "C:\Logs\audit.csv"

# With explicit credentials and parallel throttle (PS7)
$cred = Get-Credential
.\Disable-NICPacketCoalescing.ps1 -ComputerName "PC01","PC02" -Credential $cred -ThrottleLimit 5

# Audit current status across multiple workstations
.\Check-NetworkPowerStatus.ps1 -ComputerName "PC01","PC02","PC03"
```

### Console output example

```
[PS5.1 - sequential execution]

[WORKSTATION01] Pass 1 - Applying fix...
  [LOG] 14:32:01 | === Disable-NICPacketCoalescing START ===
  [LOG] 14:32:01 | STEP 1 - Scanning registry for coalescing keys
  [LOG] 14:32:01 |   FOUND [Intel(R) Wi-Fi 6 AX201 160MHz] key=[*PacketCoalescing] value=[1]
  [LOG] 14:32:01 |   WRITTEN -> recheck=[0]
  [LOG] 14:32:01 | STEP 1 DONE - 1 adapter(s) patched: Intel(R) Wi-Fi 6 AX201 160MHz
  [LOG] 14:32:02 | STEP 2 - NIC list written to C:\Windows\Temp\9Lives-NICList.txt
  [LOG] 14:32:02 | STEP 3 - Registering scheduled task [9Lives-NICRestart]
  [LOG] 14:32:02 | STEP 3 DONE - Task registered, fires in 3s
[WORKSTATION01] Pass 1 done - waiting 15s for NIC restart...
[WORKSTATION01] Pass 2 - Reading status...
  [LOG] 14:32:05 | TASK - Started as SYSTEM
  [LOG] 14:32:05 | TASK - Restarting [Intel(R) Wi-Fi 6 AX201 160MHz]
  [LOG] 14:32:07 | TASK - Restart done [Intel(R) Wi-Fi 6 AX201 160MHz]
  [LOG] 14:32:07 | TASK - Finished

=== RESULTS ===

Computer   Adapter                          Status
--------   -------                          ------
WORKSTATION01  Intel(R) Wi-Fi 6 AX201 160MHz    OK - Disabled
WORKSTATION01  Intel(R) Ethernet I225-LM        N/A - Unsupported
```

---

## GPO Deployment (local mode)

### 1. Copy to SYSVOL

```
\\yourdomain.local\SYSVOL\yourdomain.local\Scripts\Disable-NICPacketCoalescing.ps1
```

### 2. Link via Group Policy

```
GPO Name : Workstations-NIC-Latency
  └── Computer Configuration
      └── Windows Settings
          └── Scripts (Startup/Shutdown)
              └── Startup → PowerShell Scripts
                  Script : \\yourdomain.local\SYSVOL\...\Disable-NICPacketCoalescing.ps1
```

### 3. Optional WMI filter

```sql
SELECT * FROM Win32_OperatingSystem WHERE Version LIKE "10.%"
```

---

## Event Log Reference (local mode)

| EventID | Source | Description |
|---------|--------|-------------|
| 1001 | 9Lives-NetworkPerf | Execution summary per adapter (FIXED / ALREADY_OK / WHATIF / N/A) |

```powershell
# Read log on a workstation
Get-EventLog -LogName Application -Source "9Lives-NetworkPerf" -Newest 5 |
    Select-Object TimeGenerated, Message | Format-List
```

---

## Temp files (remote mode)

| File | Purpose | Lifetime |
|------|---------|---------|
| `C:\Windows\Temp\9Lives-NICFix.log` | Step-by-step execution log | Deleted after Pass 2 |
| `C:\Windows\Temp\9Lives-NICList.txt` | Adapter list for scheduled task | Deleted by task after restart |

---

## Requirements

| Requirement | Detail |
|-------------|--------|
| OS | Windows 10 (1809+) / Windows 11 |
| PowerShell | 5.1 (sequential) or 7+ (parallel) |
| Privileges | Administrator — satisfied by GPO startup context |
| WinRM | Must be enabled on remote targets |
| AD module | Required only when using `-OUPath` |

---

## License

MIT — © 9 Lives IT Solutions
