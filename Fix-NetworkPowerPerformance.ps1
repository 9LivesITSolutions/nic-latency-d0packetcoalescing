#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Disables D0PacketCoalescing on all physical network adapters.

.DESCRIPTION
    GPO Startup Script — Network Latency Optimization
    Disables D0PacketCoalescing on all physical network adapters (WiFi + Ethernet)
    that support the setting, regardless of manufacturer.
    Results are logged to the Windows Application Event Log.

.NOTES
    Author  : 9 Lives IT Solutions
    Version : 1.0.0
    Target  : Windows 10/11 workstations joined to a domain
    GPO     : Computer Configuration > Windows Settings > Scripts > Startup

.LINK
    https://github.com/9lives-it/gpo-network-perf
#>

$LogSource = "9Lives-NetworkPerf"
$LogName   = "Application"

# Create event log source if missing
if (-not [System.Diagnostics.EventLog]::SourceExists($LogSource)) {
    New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
}

$results = @()

Get-NetAdapter -Physical | Where-Object { $_.Status -ne "Not Present" } | ForEach-Object {
    $adapterName = $_.Name
    $adapterDesc = $_.InterfaceDescription

    try {
        $pm = Get-NetAdapterPowerManagement -Name $adapterName -ErrorAction Stop

        if ($pm.D0PacketCoalescing -eq "Enabled") {
            Set-NetAdapterPowerManagement -Name $adapterName -D0PacketCoalescing Disabled -ErrorAction Stop
            $results += "OK  | $adapterDesc -> D0PacketCoalescing disabled"
        }
        elseif ($pm.D0PacketCoalescing -eq "Disabled") {
            $results += "--  | $adapterDesc -> already disabled"
        }
        else {
            $results += "N/A | $adapterDesc -> not supported ($($pm.D0PacketCoalescing))"
        }
    }
    catch {
        $results += "ERR | $adapterDesc -> $($_.Exception.Message)"
    }
}

# Consolidated log in Event Viewer
$body = "Fix-NetworkPowerPerformance -- $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n`n" + ($results -join "`n")
Write-EventLog -LogName $LogName -Source $LogSource -EventId 1001 -EntryType Information -Message $body -ErrorAction SilentlyContinue
