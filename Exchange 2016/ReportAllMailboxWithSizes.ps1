<#
.SYNOPSIS
PowerShell script to report all mailboxes, along with the sizes.

.DESCRIPTION
This script loads all mailboxes in Exchange Organization and optionally reads their sizes.
Then it sends that as a report over email.

.PARAMETER DomainName
FQDN of Active Directory domain

.PARAMETER GetSizes
Default is to read sizes for each mailbox.

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
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\ReportAllMailboxWithSizes.ps1"

.NOTES
Name:    ReportAllMailboxWithSizes.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>
param
(
 [string]$DomainName    = 'domain.local',
 [Boolean]$GetSizes     = $true,
 [Boolean]$SendReport   = $true,
 [string]$PSSnapin      = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Exchange Reports <exchange.reports@domain.local>',
 [String]$SmtpTo        = 'exchange.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = '',
 [String]$SmtpBody      = ''
)

#-------------------
Function Format-OU()
{
 param
 (
  [String]$OuString,     # domain.local/Ou1/Ou2/Ou3/Users/Mailboxes
  [string]$DomainName,   # domain.local
  [Int16]$DomainNameLen  # 11
 )

 if($OuString.Substring(0, $DomainNameLen).ToLower() -eq $DomainName.ToLower()
 { $OuString = $OuString.Substring($DomainNameLen) }

 return $OuString
}

#-------------------
Function Format-Company()
{
 param
 (
  [String]$OuString,     # domain.local/Ou1/Ou2/Ou3/Users/Mailboxes
  [Int16]$DomainNameLen  # 11
 )

 $OuString = $OuString.Substring($DomainNameLen)
 $Result   = $OuString

 foreach($Ou in $OuToCompany.Keys)
 {
  if($OuString.Length -ge $Ou.length)
  {
   if($OuString.Substring(0, $Ou.length).ToLower() -eq $Ou.ToLower())
   {
    $Result = $OuToCompany[$Ou]
    break
   }
  }
 }

 return $Result
}

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
# Environment-specific Parameters

$OuToCompany   = @{'/Warehouse/Company 1' = 'Company 1, Warehouse';
                   '/Warehouse/Company 2' = 'Company 2, Warehouse';
                   '/Office/Company 1'    = 'Company 1, Office';
                   '/Office/Company 2'    = 'Company 2, Office';
                   '/Users'               = 'N/A';}

#-------------------
# Run-time Parameters

$CurrDate      = Get-Date
$ReportFile    = [environment]::getfolderpath('mydocuments') + '\AllMailboxSizes-' + `
                 $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '-' + $CurrDate.Hour + '-' + $CurrDate.Minute + '-' + $CurrDate.Second + '.csv'
$Output        = @()
$DomainNameLen = $DomainName.Length

#-------------------

if($SmtpSubject -eq '')
{
 if($GetSizes)
 {
  $SmtpSubject   = 'Mailbox Sizes'
  $SmtpBody      = 'Attached is a list of all mailboxes with the sizes.'
 }
 else
 {
  $SmtpSubject   = 'Mailboxes'
  $SmtpBody      = 'Attached is a list of all mailboxes.'
 }
}

Write-Progress -Id 1 -Activity 'List of databases' -Status 'Loading...' -PercentComplete 1
$DBs = Get-MailboxServer | ForEach-Object { Get-MailboxDatabase -Server $_.Name } | Sort-Object Name | Select-Object Name | Get-Unique -AsString
Write-Progress -Id 1 -Activity 'List of databases' -Status 'Loaded' -PercentComplete 100 -Completed

$DBsCount   = $DBs | Measure-Object | Select-Object -ExpandProperty Count
$DBsCurrent = 0

foreach($DB in $DBs)
{
 $DBsCurrent++
 Write-Progress -Id 1 -Activity "Inventorying mailboxes" -Status $DB.Name -PercentComplete ($DBsCurrent / $DBsCount * 100)

 $MBs        = Get-Mailbox -Database $DB.Name | Sort-Object DisplayName
 $MBsCount   = $MBs | Measure-Object | Select-Object -ExpandProperty Count
 $MBsCurrent = 0

 foreach($MB in $MBs)
 {
  $MBsCurrent++
  Write-Progress -ParentId 1 -Id 2 -Activity "Mailbox" -Status $MB.DisplayName -PercentComplete ($MBsCurrent / $MBsCount * 100)
  if($GetSizes)
  {
   $Stat = $MB |
           Get-MailboxStatistics -WarningAction silentlyContinue |
           Select-Object LastLogonTime,
                         StorageLimitStatus,
                         ItemCount,
                         @{name="TotalItemSizeMB";expression={[math]::Round((($_.TotalItemSize.Value.ToString()).Split('(')[1].Split(' ')[0].Replace(',','')/1MB),2)}},
                         DeletedItemCount,
                         @{name="TotalDeletedItemSizeMB";expression={[math]::Round((($_.TotalDeletedItemSize.Value.ToString()).Split('(')[1].Split(' ')[0].Replace(',','')/1MB),2)}}
  }

  $AdUser = Get-User $MB.SamAccountName
  try
  { if($AdUser.UserAccountControl[0] -band 2) { $AdEnabled = 'Disabled' } Else { $AdEnabled = 'Enabled' }}
  catch
  { $AdEnabled = 'UNKNOWN' }

  $Obj  = New-Object System.Object
  $Obj | Add-Member -MemberType NoteProperty -Name 'User'                -Value $MB.DisplayName
  $Obj | Add-Member -MemberType NoteProperty -Name 'User Name'           -Value $MB.Alias
  $Obj | Add-Member -MemberType NoteProperty -Name 'First Name'          -Value $AdUser.FirstName
  $Obj | Add-Member -MemberType NoteProperty -Name 'Last Name'           -Value $AdUser.LastName
  $Obj | Add-Member -MemberType NoteProperty -Name 'OU'                  -Value (Format-OU -OUString $MB.OrganizationalUnit -DomainName $DomainName -DomainNameLen $DomainNameLen)
  $Obj | Add-Member -MemberType NoteProperty -Name 'OU Company'          -Value (Format-Company -OUString $MB.OrganizationalUnit -DomainNameLen $DomainNameLen)
  $Obj | Add-Member -MemberType NoteProperty -Name 'SMTP'                -Value $MB.PrimarySmtpAddress
  $Obj | Add-Member -MemberType NoteProperty -Name 'Addresses'           -Value ($MB.EmailAddresses -join ", `n")
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
  $Obj | Add-Member -MemberType NoteProperty -Name 'Default Quota'       -Value $MB.UseDatabaseQuotaDefaults
  $Obj | Add-Member -MemberType NoteProperty -Name 'Retention Policy'    -Value $MB.RetentionPolicy

  if($GetSizes)
  {
   $LastLogonTime = ''
   if($Stat.LastLogonTime)
   {
    if($Stat.LastLogonTime.GetType().FullName -eq 'System.DateTime')
    { $LastLogonTime = ($Stat.LastLogonTime.Year.ToString() + '-' + $Stat.LastLogonTime.Month.ToString() + '-' + $Stat.LastLogonTime.Day.ToString()) }
    else
    { $LastLogonTime = $Stat.LastLogonTime }
   }

   $Obj | Add-Member -MemberType NoteProperty -Name 'Last Logon Date'     -Value $LastLogonTime
   $Obj | Add-Member -MemberType NoteProperty -Name 'Limit Status'        -Value $Stat.StorageLimitStatus
   $Obj | Add-Member -MemberType NoteProperty -Name 'Mail Items'          -Value $Stat.ItemCount
   $Obj | Add-Member -MemberType NoteProperty -Name 'Mailbox Size, MB'    -Value $Stat.TotalItemSizeMB
   $Obj | Add-Member -MemberType NoteProperty -Name 'Deleted Items'       -Value $Stat.DeletedItemCount
   $Obj | Add-Member -MemberType NoteProperty -Name 'Deleted Size, MB'    -Value $Stat.TotalDeletedItemSizeMB
  }

  $Output += $Obj
 }
 Write-Progress -ParentId 1 -Id 2 -Activity "Mailbox" -Status 'Done' -PercentComplete 100 -Completed
}
Write-Progress -Id 1 -Activity "Inventorying mailboxes" -Status 'Done' -PercentComplete 100 -Completed

Write-Progress -Id 1 -Activity "Exporting information into a file" -Status $ReportFile -PercentComplete 1
$Output | Export-Csv $ReportFile -Encoding UTF8 -NoTypeInformation
Write-Progress -Id 1 -Activity "Exporting information into a file" -Status 'Done' -PercentComplete 100 -Completed

if($SendReport)
{
 Write-Progress -Id 1 -Activity "Emailing information" -Status $SmtpTo -PercentComplete 1
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
 Write-Progress -Id 1 -Activity "Emailing information" -Status 'Done' -PercentComplete 100 -Completed

 Write-Progress -Id 1 -Activity "Deleting file" -Status $ReportFile -PercentComplete 1
 Remove-Item $ReportFile -force
 Write-Progress -Id 1 -Activity "Deleting file" -Status 'Done' -PercentComplete 1
}
