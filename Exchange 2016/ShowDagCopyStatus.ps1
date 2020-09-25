<#
.SYNOPSIS
PowerShell script to continously display the DAG copy status. It shows data
for Mounted databases separately from data for replicas.

.DESCRIPTION
This script reads the database copy status for each mailbox database, and
displays it as two tables: mounted and not. Then it refreshes that table every 5 seconds.
Infinitely.

PS. This used to be a one-liner :)

.PARAMETER DagName
Name of the DAG to display the information about.

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\ShowDagCopyStatus.ps1" -DagName "DAG01"

.NOTES
Name:    ShowDagCopyStatus.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.1

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

#--------------------------------------
param
(
 [String]$DagName = 'DAG01'
)

#-------------------
try
{ Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn }
catch
{
 'ERROR: The PowerShell snap-in for Exchange was not found'
 exit 1
}

#-------------------
$TimeOut   = New-Timespan -Minutes 5
$StopWatch = [diagnostics.stopwatch]::startnew()

#-------------------
while($StopWatch.elapsed -lt $TimeOut)
{
 Write-Host "Reloading..."

 $dag  = Get-DatabaseAvailabilityGroup $DagName
 $data =
 (
  $dag |
  ForEach-Object `
  {
   $_.Servers |
   Sort-Object |
   ForEach-Object {Get-MailboxDatabaseCopyStatus -Server $_}
  }
 )

 $mounted =
 (
  $data |
  Sort-Object Status, Name |
  Where-Object {$_.Status -eq 'Mounted'} |
  Select-Object Name, *length, Status, ContentIndexState |
  Format-Table -AutoSize
 )

 $notmounted =
 (
  $data |
  Sort-Object Name, Status |
  Where-Object {$_.Status -ne 'Mounted'} |
  Select-Object Name, *length, Status, ContentIndexState |
  Format-Table -AutoSize
 )

 clear-host
 ($mounted + $notmounted)
 start-sleep -seconds 5
}
