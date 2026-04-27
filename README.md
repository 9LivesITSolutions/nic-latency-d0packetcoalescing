# GPO — Network Power Performance

> **Disable D0PacketCoalescing on all physical network adapters via GPO startup script.**

Reduces WiFi and Ethernet latency on Windows 10/11 workstations by disabling packet coalescing at the driver level — manufacturer agnostic (Intel, Realtek, Broadcom, Qualcomm, MediaTek...).

---

## What is D0PacketCoalescing?

D0 (Device state 0) packet coalescing batches incoming network packets instead of processing them immediately, reducing CPU usage at the cost of added latency (10–50ms typical). On workstations plugged into AC power, there is no benefit to leaving it enabled.

```
Without coalescing:  packet → immediate processing → ~1ms latency
With coalescing:     packet → buffer → batch processing → ~10-50ms latency
```

---

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/Fix-NetworkPowerPerformance.ps1` | GPO startup script — disables D0PacketCoalescing |
| `scripts/Check-NetworkPowerStatus.ps1` | Audit script — checks status across remote workstations |

---

## Deployment

### 1. Copy the script to SYSVOL

```
\\yourdomain.local\SYSVOL\yourdomain.local\Scripts\Fix-NetworkPowerPerformance.ps1
```

### 2. Link via Group Policy

```
GPO Name : Workstations-NetworkPerformance
  └── Computer Configuration
      └── Windows Settings
          └── Scripts (Startup/Shutdown)
              └── Startup → PowerShell Scripts
                  Script : \\yourdomain.local\SYSVOL\...\Fix-NetworkPowerPerformance.ps1
```

### 3. Optional — WMI filter (target specific hardware only)

```sql
SELECT * FROM Win32_NetworkAdapter WHERE AdapterType = "Ethernet 802.3"
```

---

## Verification

### Check Event Viewer log on a workstation

```powershell
Get-EventLog -LogName Application -Source "9Lives-NetworkPerf" -Newest 5 |
    Select-Object TimeGenerated, Message | Format-List
```

### Audit multiple workstations remotely

```powershell
$computers = @("PC01", "PC02", "PC03")
.\scripts\Check-NetworkPowerStatus.ps1 -ComputerName $computers
```

### Expected output

```
Computer   Adapter                          Status  D0PacketCoalescing
--------   -------                          ------  ------------------
PC01       Intel(R) Wi-Fi 6 AX200 160MHz    Up      Disabled
PC01       Intel(R) Ethernet I219-V         Up      Disabled
PC02       Realtek PCIe GbE Family...       Up      N/A (Unsupported)
```

---

## Event Log Reference

| EventID | Source | Description |
|---------|--------|-------------|
| 1001 | 9Lives-NetworkPerf | Script execution summary (OK / already disabled / N/A / ERR per adapter) |

---

## Compatibility

| OS | Supported |
|----|-----------|
| Windows 11 | ✅ |
| Windows 10 (1809+) | ✅ |
| Windows Server 2019/2022 | ✅ |

> Requires PowerShell 5.1+ and administrator privileges (satisfied by GPO startup context).

---

## License

MIT — © 9 Lives IT Solutions
