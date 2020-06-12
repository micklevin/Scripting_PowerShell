<#
.SYNOPSIS
PowerShell script to report enabled custom Scheduled Tasks from all servers in the domain

.DESCRIPTION
This script lists all servers in domain and then uses Get-ScheduledTask.ps1 by Jaap Brasser to load the list of Scheduled Tasks from each server.
Built-in tasks by Microsoft are ignored. The resultant list is emailed to specified recipients

.PARAMETER SmtpServer
The DNS name or IP address of SMTP server. Note - this script does not support SMTP authentication

.PARAMETER SmtpFrom
The SMTP name and address of Sender.

.PARAMETER SmtpTo
The SMTP address of recipient(s).

.PARAMETER SmtpSubject
Alternative text for the email's Subject.

.PARAMETER SmtpBody
Alternative text for the email's body.

.NOTES
Name:    Get Scheduled Tasks from Servers.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

#--------------------------------------
param
(
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Windows Server Reports <windows.reports@domain.local>',
 [String]$SmtpTo        = 'windows.reports@domain.local',
 [String]$SmtpSubject   = 'Report on Scheduled Tasks',
 [String]$SmtpBody      = 'Attached is the list of Scheduled Tasks on all servers in the domain.'
)

#--------------------------------------
# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\Scheduled-Tasks-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'

#-------------------
Write-Progress -Id 1 -Activity 'Locating Servers in AD' -Status 'Loading...' -PercentComplete 1
$AllAdServers = Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' `
                               -Properties Name, Operatingsystem, OperatingSystemVersion, IPv4Address |
                Sort-Object -Property Operatingsystem, Name |
                Select-Object -Property Name,Operatingsystem,OperatingSystemVersion,IPv4Address
Write-Progress -Id 1 -Activity 'Locating Servers in AD' -Status 'Done' -PercentComplete 100 -Completed

if ($AllAdServers)
{
 $ServerCount   = $AllAdServers | Measure-Object | Select-Object -ExpandProperty Count
 $ServerCurrent = 0
 $ServerFailed  = 0

 $AllAdServers |
 ForEach-Object `
 {
  $ServerCurrent++
  Write-Progress -Id 1 -Activity 'Inventorying Servers' -Status "[$ServerCurrent + $ServerFailed / $ServerCount] $($_.Name) [$($_.IPv4Address)]" -PercentComplete ($ServerCurrent / $ServerCount * 100)
  try
  {
   .\Get-ScheduledTask.ps1 -ComputerName $_.Name |
   Where-Object {(-not ($_.Path -like '\Microsoft\*')) -and `
                 (-not ($_.Path -like '\Optimize Start Menu Cache*')) -and `
                 ($_.State -ne 'Disabled')}
  }
  catch
  { $ServerFailed++ }
 } |
 Select-Object ComputerName, Path, State, NextRunTime, LastRunTime, LastTaskResult |
 Export-Csv -Path $ReportFile -Encoding UTF8 -NoTypeInformation
 Write-Progress -Id 1 -Activity 'Inventorying Servers' -Status 'Done' -PercentComplete 100 -Completed

 Write-Progress -Id 1 -Activity 'Sending Report in e-mail' -Status $SmtpTo -PercentComplete 1
 Send-Mailmessage -SmtpServer $SmtpServer `
                  -From $SmtpFrom `
                  -To $SmtpTo `
                  -Subject $SmtpSubject `
                  -Body $SmtpBody `
                  -Priority High `
                  -Attachments $ReportFile
 Remove-Item $ReportFile -force
 Write-Progress -Id 1 -Activity 'Sending Report in e-mail' -Status 'Done' -PercentComplete 100 -Completed
}
