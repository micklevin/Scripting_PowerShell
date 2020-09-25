<#
.SYNOPSIS
PowerShell script to report all mailboxes, which were accessed by mobile devices.

.DESCRIPTION
This script loads all mailboxes in Exchange Organization and reports all connections from mobile devices.

.PARAMETER DaysAge
Minimal age of connections, in days.

.PARAMETER SendReport
Default is to email the report. Otherwise it will be just sitting in a CSV file.

.PARAMETER PSSnapin
Dafault name for Exchange 2016 snap-in for PowerShell.

.PARAMETER SmtpServer
The DNS name of SMTP server

.PARAMETER SmtpFrom
The SMTP name and address of Sender.

.PARAMETER SmtpTo
The SMTP address of recipient(s).

.PARAMETER SmtpCc
The SMTP address of CC recipient(s).

.PARAMETER SmtpSubject
Alternative text for the email's Subject.

.PARAMETER SmtpBody
Alternative text for the email's body.

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\ReportEasDevices.ps1" -DaysAge 7

.NOTES
Name:    ReportEasDevices.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>
param
(
 [int]$DaysAge          = 0,
 [Boolean]$SendReport   = $true,
 [string]$PSSnapin      = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Exchange Reports <exchange.reports@domain.local>',
 [String]$SmtpTo        = 'exchange.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = 'All Mobile Exchange Clients',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes accessed from mobile devices.'
)

#-------------------
# Load Exchange snap-in and environment

if(!(Get-PSSnapin | Where-Object {$_.Name -eq $PSSnapin}))
{
 if($Debug) { 'DEBUG: The PowerShell snap-in "' + $PSSnapin + '" not loaded, trying to load now...' }
 try
 { Add-PSSnapin $PSSnapin }
 catch
 {
  ('ERROR: The PowerShell snap-in "' + $PSSnapin + '" failed to load')
  exit 1
 }
 . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
 Connect-ExchangeServer -auto -AllowClobber
}
else
{ if($Debug) { 'DEBUG: The PowerShell snap-in "' + $PSSnapin + '" already loaded'} }

#-------------------
# Run-time Parameters

$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllEasDevices-' + `
              $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '-' + $CurrDate.Hour + '-' + $CurrDate.Minute + '-' + $CurrDate.Second + '.csv'
$Output     = @()

Write-Host -NoNewLine 'Loading list of all mailboxes...'
$CasMbxs = @(Get-CASMailbox -Resultsize Unlimited -WarningAction SilentlyContinue | Sort-Object Name )
Write-Host " done for $($CasMbxs.count) mailboxes"

foreach($Mbx in $CasMbxs)
{
 if($MBx.HasActiveSyncDevicePartnership)
 {
  Write-Host -NoNewLine "$($Mbx.Name)`t"
  $Stats = @(Get-ActiveSyncDeviceStatistics -Mailbox $Mbx.Identity -WarningAction SilentlyContinue -ErrorAction Stop | Sort-Object -Descending LastSuccessSync)
  if($Stats.Count -gt 0)
  {
   Write-Host -NoNewLine "$($Stats.Count)`t"

   $Info = Get-Mailbox $Mbx.Identity | Select-Object DisplayName, PrimarySMTPAddress, OrganizationalUnit

   foreach($EasDevice in $Stats)
   {
    if($null -eq $EasDevice.LastSyncAttemptTime)
    { $SyncAge = 3650 } # 10 years means "never"
    else
    { $SyncAge = ($CurrDate - $EasDevice.LastSyncAttemptTime).Days }

    if($SyncAge -ge $DaysAge)
    {
     Write-Host -NoNewLine '.'
     $Obj = New-Object PSObject
     $Obj | Add-Member NoteProperty -Name 'Display Name'        -Value $Info.DisplayName
     $Obj | Add-Member NoteProperty -Name 'Organizational Unit' -Value $Info.OrganizationalUnit
     $Obj | Add-Member NoteProperty -Name 'Email Address'       -Value $Info.PrimarySMTPAddress
     $Obj | Add-Member NoteProperty -Name 'Sync Age (Days)'     -Value $SyncAge
     $Obj | Add-Member NoteProperty -Name 'Device ID'           -Value $EasDevice.DeviceID
     $Obj | Add-Member NoteProperty -Name 'Device User Agent'   -Value $EasDevice.DeviceUserAgent
     $Obj | Add-Member NoteProperty -Name 'Device Model'        -Value $EasDevice.DeviceModel
     $Obj | Add-Member NoteProperty -Name 'Device Name'         -Value $EasDevice.DeviceFriendlyName
     $Obj | Add-Member NoteProperty -Name 'Device OS'           -Value $EasDevice.DeviceOS
     $Obj | Add-Member NoteProperty -Name 'Device Status'       -Value $EasDevice.Status
     $Obj | Add-Member NoteProperty -Name 'Device Wipable?'     -Value $EasDevice.IsRemoteWipeSupported
     $Output += $Obj
    }
   }
  }
  else
  {
   Write-Host -NoNewline '?'
  }
  Write-Host ''
 }
 else
 {
  Write-Host "$($Mbx.Name)`t-"
 }
}
Write-Host ' Done'

Write-Host -NoNewLine "Exporting information into: $ReportFile ..."
$Output | Export-Csv $ReportFile -Encoding UTF8 -NoTypeInformation
Write-Host ' Done'

if($SendReport)
{
 Write-Host -NoNewLine "Emailing this report to: $SmtpTo ..."
 if($SmtpCC -ne '')
 {
  Send-Mailmessage -SmtpServer $SmtpServer `
                   -From $SmtpFrom `
                   -To $SmtpTo `
                   -Cc $SmtpCc `
                   -Subject $SmtpSubject `
                   -Body $SmtpBody `
                   -Priority High `
                   -Attachments $ReportFile
 }
 Else
 {
  Send-Mailmessage -SmtpServer $SmtpServer `
                   -From $SmtpFrom `
                   -To $SmtpTo `
                   -Subject $SmtpSubject `
                   -Body $SmtpBody `
                   -Priority High `
                   -Attachments $ReportFile
 }
 Write-Host ' Done'
 Write-Host (' Deleting file: ' + $ReportFile + ' ...')
 Remove-Item $ReportFile -force
 Write-Host ' Done.'
}
