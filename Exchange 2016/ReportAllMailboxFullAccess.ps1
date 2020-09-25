<#
.SYNOPSIS
PowerShell script to report all mailboxes, which have Full Access granted to somebody else.

.DESCRIPTION
This script loads all mailboxes in Exchange Organization and inspects each for access granted to other mailboxes.
If there was a mailbox with Full Access granted to somebody else, then such mailbox is reported.

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
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\ReportAllMailboxFullAccess.ps1"

.NOTES
Name:    ReportAllMailboxFullAccess.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>
param
(
 [Boolean]$SendReport   = $true,
 [string]$PSSnapin      = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Exchange Reports <exchange.reports@domain.local>',
 [String]$SmtpTo        = 'exchange.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = 'Mailboxes with Full Access',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes with full access granted to other users.'
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
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllMailboxFullAccess-' + `
              $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '-' + $CurrDate.Hour + '-' + $CurrDate.Minute + '-' + $CurrDate.Second + '.csv'
$Output     = @()

Write-Host -NoNewLine 'Loading list of databases...'
$DBs = Get-MailboxServer | ForEach-Object { Get-MailboxDatabase -Server $_.Name } | Sort-Object Name | Select-Object Name | Get-Unique -AsString
Write-Host ' Done'

foreach($DB in $DBs)
{
 Write-Host -NoNewLine "Inventorying mailboxes of $($DB.Name) "
 $MBs = Get-Mailbox -Database $DB.Name -ResultSize Unlimited | Sort-Object DisplayName

 foreach($MB in $MBs)
 {
  Write-Host -NoNewLine '.'
  $MBXPerms = $MB | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and ($_.User.tostring() -ne "NT AUTHORITY\SELF")}

  if($MBXPerms)
  {
   $Stat = $MB | Get-MailboxStatistics -WarningAction silentlyContinue | Select-Object LastLogonTime

   $AdUser = Get-User $MB.SamAccountName
   try   { if($AdUser.UserAccountControl[0] -band 2) { $AdEnabled = 'Disabled' } Else { $AdEnabled = 'Enabled' }}
   catch { $AdEnabled = 'UNKNOWN' }

   $LastLogonTime = ''
   if($Stat.LastLogonTime)
   {
    if($Stat.LastLogonTime.GetType().FullName -eq 'System.DateTime')
    { $LastLogonTime = ($Stat.LastLogonTime.Year.ToString() + '-' + $Stat.LastLogonTime.Month.ToString() + '-' + $Stat.LastLogonTime.Day.ToString()) }
    else
    { $LastLogonTime = $Stat.LastLogonTime }
   }

   foreach($MBP in $MBXPerms)
   {
    $Obj  = New-Object System.Object
    $Obj | Add-Member -MemberType NoteProperty -Name 'User'                -Value $MB.DisplayName
    $Obj | Add-Member -MemberType NoteProperty -Name 'User Name'           -Value $MB.Alias
    $Obj | Add-Member -MemberType NoteProperty -Name 'SMTP'                -Value $MB.PrimarySmtpAddress
    $Obj | Add-Member -MemberType NoteProperty -Name 'Automatic Addresses' -Value $MB.EmailAddressPolicyEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name 'Forward To'          -Value $MB.ForwardingAddress
    $Obj | Add-Member -MemberType NoteProperty -Name 'Forward and Deliver' -Value $MB.DeliverToMailboxAndForward
    $Obj | Add-Member -MemberType NoteProperty -Name 'Litigation Hold'     -Value $MB.LitigationHoldEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name 'Enabled'             -Value $AdEnabled
    $Obj | Add-Member -MemberType NoteProperty -Name 'Company'             -Value $AdUser.Company
    $Obj | Add-Member -MemberType NoteProperty -Name 'Department'          -Value $AdUser.Department
    $Obj | Add-Member -MemberType NoteProperty -Name 'Country'             -Value $AdUser.CountryOrRegion
    $Obj | Add-Member -MemberType NoteProperty -Name 'Location'            -Value $AdUser.Office
    $Obj | Add-Member -MemberType NoteProperty -Name 'DB Name'             -Value $DB.Name
    $Obj | Add-Member -MemberType NoteProperty -Name 'Last Logon Date'     -Value $LastLogonTime
    $Obj | Add-Member -MemberType NoteProperty -Name 'Access User'         -Value $MBP.User
    $Obj | Add-Member -MemberType NoteProperty -Name 'Access Right'        -Value ($MBP.AccessRights -join ', ')
    $Obj | Add-Member -MemberType NoteProperty -Name 'Access Denied'       -Value $MBP.Deny
    $Obj | Add-Member -MemberType NoteProperty -Name 'Access Valid'        -Value $MBP.IsValid
    $Output += $Obj
   }
  }
 }
 Write-Host ' Done'
}

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
