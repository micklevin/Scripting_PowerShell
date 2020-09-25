<#
.SYNOPSIS
PowerShell script to continously display the DAG replication queue lenghts

.DESCRIPTION
This script reads the database copy status for each mailbox database, and
displays it as a table. Then it refreshes that table every 5 seconds.
Infinitely.

PS. This used to be a one-liner :)

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\ShowDagQueueLength.ps1" -DagName "DAG01"

.NOTES
Name:    ShowDagQueueLength.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.1

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

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
 Write-Host -NoNewline 'Reloading: '
 $data   = (Get-MailboxDatabase |
            Sort-Object Name |
            Get-MailboxDatabaseCopyStatus |
            Where-Object {$_.Status -ne 'Mounted'} |
            Select-Object Name, CopyQueueLength, ReplayQueueLength)
 $output = @()

 foreach($MDB in $data)
 {
  Write-Host -NoNewline '.'
  $Obj  = New-Object System.Object
  $Obj | Add-Member -MemberType NoteProperty -Name 'Database'            -Value $MDB.Name
  $Obj | Add-Member -MemberType NoteProperty -Name 'Copy Queue Length'   -Value $MDB.CopyQueueLength
  $Obj | Add-Member -MemberType NoteProperty -Name 'Replay Queue Length' -Value $MDB.ReplayQueueLength
  $output += $Obj
 }

 clear-host
 $output
 start-sleep -seconds 5
}
