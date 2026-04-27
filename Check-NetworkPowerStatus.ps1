#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Audits D0PacketCoalescing status across a list of remote workstations.

.DESCRIPTION
    Run this script from a management workstation to check the D0PacketCoalescing
    setting on all physical network adapters across multiple remote computers.

.PARAMETER ComputerName
    List of computer names to audit. Defaults to local machine.

.EXAMPLE
    .\Check-NetworkPowerStatus.ps1 -ComputerName "PC01","PC02","PC03"

.EXAMPLE
    # Audit from AD OU
    $computers = (Get-ADComputer -Filter * -SearchBase "OU=Workstations,DC=cmcap,DC=local").Name
    .\Check-NetworkPowerStatus.ps1 -ComputerName $computers

.NOTES
    Author  : 9 Lives IT Solutions
    Version : 1.0.0
#>

param(
    [string[]]$ComputerName = @($env:COMPUTERNAME)
)

Invoke-Command -ComputerName $ComputerName -ErrorAction SilentlyContinue -ScriptBlock {
    Get-NetAdapter -Physical | Where-Object { $_.Status -ne "Not Present" } | ForEach-Object {
        $pm = Get-NetAdapterPowerManagement -Name $_.Name -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            Computer           = $env:COMPUTERNAME
            Adapter            = $_.InterfaceDescription
            Status             = $_.Status
            D0PacketCoalescing = $pm.D0PacketCoalescing
        }
    }
} | Format-Table Computer, Adapter, Status, D0PacketCoalescing -AutoSize
