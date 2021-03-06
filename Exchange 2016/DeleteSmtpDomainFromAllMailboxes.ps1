<#
.SYNOPSIS
PowerShell script to remote any SMTP addresses or particular domain from all users.

.DESCRIPTION
This script takes a domain name of particular SMTP namespace, and removes any SMTP address matching that domain name
from all users in Exchange Organization.

.PARAMETER DeleteDomain
Domain name; no @ sign in front.

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
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\DeleteSmtpDomainFromAllMailboxes.ps1" -DeleteDomain "olddomain.com"

.NOTES
Name:    DeleteSmtpDomainFromAllMailboxes.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0.1

.LINK
https://github.com/micklevin/Scripting_PowerShell

#>
param
(
 [String]$DeleteDomain  = 'olddomain.com',
 [Boolean]$SendReport   = $true,
 [string]$PSSnapin      = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Exchange Reports <exchange.reports@domain.local>',
 [String]$SmtpTo        = 'exchange.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = 'Delete Email Domain',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes which had the email address deleted.'
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
$ReportFile = [environment]::getfolderpath('mydocuments') + '\DeleteEmailDomain-' + $DeleteDomain + '-' + `
              $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '-' + $CurrDate.Hour + '-' + $CurrDate.Minute + '-' + $CurrDate.Second + '.csv'
$Output     = @()

if($DeleteDomain.Substring(1) -ne '@')
{ $DeleteDomain = '@' + $DeleteDomain }

if($SmtpSubject -eq 'Delete Email Domain')
{ $SmtpSubject = 'Delete Email Domain: ' + $DeleteDomain }

$DeleteDomain    = $DeleteDomain.ToLower()
$DeleteDomainLen = $DeleteDomain.Length

Write-Host -NoNewLine 'Loading list of databases...'
$DBs = Get-MailboxServer | ForEach-Object { Get-MailboxDatabase -Server $_.Name } | Sort-Object Name | Select-Object Name | Get-Unique -AsString
Write-Host ' Done'

foreach($DB in $DBs)
{
 Write-Host -NoNewLine "Updating mailboxes of $($DB.Name) "
 $MBs = Get-Mailbox -Database $DB.Name -ResultSize Unlimited |
        Where-Object {$_.EmailAddresses -like ('*' + $DeleteDomain)} |
        Sort-Object DisplayName

 # Process each mailbox in current database
 foreach($MB in $MBs)
 {
  $DeleteAddresses = @()
  $WillChange      = 0

  # Find all SMTP addresses matching this domain
  foreach($EmailAddress in $MB.EmailAddresses)
  {
   if($EmailAddress.Prefix.DisplayName -eq 'SMTP')
   {
    if($EmailAddress.AddressString.Length -gt $DeleteDomainLen)
    {
     $EmailAddressDomain = $EmailAddress.AddressString.SubString($EmailAddress.AddressString.Length - $DeleteDomainLen, $DeleteDomainLen).ToLower()
     if($EmailAddressDomain -eq $DeleteDomain)
     {
      if($EmailAddress.IsPrimaryAddress)
      {
       # Gee, that should NOT be happening. Make sure ahead of time, that the domain to be deleted is NOT primary!
       Write-Host -NoNewLine ('!')
      }
      else
      {
       $DeleteAddresses += $EmailAddress.AddressString
       $WillChange++
      }
     }
    }
   }
  }

  # If addresses-to-be-deleted were found
  if($WillChange -gt 0)
  {
   Write-Host -NoNewLine $WillChange
   foreach($DeleteAddress in $DeleteAddresses)
   {
    try
    { Set-Mailbox $MB.Identity -EmailAddresses @{remove=$DeleteAddress} -ErrorAction Stop }
    catch
    {
     $ErrorCode = $_.Exception.HResult
     if($ErrorCode -ne -2146233088)
     {
      Write-Host ("`n> Error:  " + $_.Exception.HResult + ' (' + $_.Exception.Message + ")`n> Mailbox: " + $MB.Name + ' (' + $MB.Alias + ')')
      Exit 0
     }
    }
   }
  }
  else
  {
   # Well, NOW it tells me there is no SMTP addresses to delete?
   Write-Host -NoNewLine '?'
  }

  # Write-Host -NoNewLine '.'
  $Obj  = New-Object System.Object
  $Obj | Add-Member -MemberType NoteProperty -Name 'User'                -Value $MB.DisplayName
  $Obj | Add-Member -MemberType NoteProperty -Name 'User Name'           -Value $MB.Alias
  $Obj | Add-Member -MemberType NoteProperty -Name 'OU'                  -Value $MB.OrganizationalUnit
  $Obj | Add-Member -MemberType NoteProperty -Name 'SMTP'                -Value $MB.PrimarySmtpAddress
  $Obj | Add-Member -MemberType NoteProperty -Name 'Addresses'           -Value ($MB.EmailAddresses -join "`n")
  $Obj | Add-Member -MemberType NoteProperty -Name 'Automatic Addresses' -Value $MB.EmailAddressPolicyEnabled
  $Obj | Add-Member -MemberType NoteProperty -Name 'Deleted Addresses'   -Value ($DeleteAddresses -join "`n")

  $Output += $Obj
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
