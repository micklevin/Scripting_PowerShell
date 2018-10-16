# PowerShell script to list version of all Exchange Servers in organization
#
# Requires PowerShell Remoting enabled
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0
#
# Configuration

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

$ExchangeServers = Get-ExchangeServer | Sort-Object Name

ForEach ($Server in $ExchangeServers) 
{
 Invoke-Command -ComputerName $Server.Name -ScriptBlock {Get-Command Exsetup.exe | ForEach-Object {$_.FileversionInfo} }
}
