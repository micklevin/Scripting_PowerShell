# PowerShell script to inventory Mailbox Databases of Exchange 2010 Organization
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0
#
# Configuration

param
(
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Exchange Reports <exchange.reports@domain.local>',
 [String]$SmtpTo        = 'exchange.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = 'Mailbox Sizes',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes with the sizes.',
 [String]$SmtpBodyTail  = ''
)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllMailboxSizes-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()

Write-Host -NoNewLine 'Loading list of databases...'
$DBs    = Get-MailboxDatabase | Sort-Object Name
Write-Host ' Done'

foreach($DB in $DBs)
{
 Write-Host -NoNewLine "Inventorying mailboxes of $($DB.Name) "
 $MBs = Get-Mailbox -Database $DB.Name | Sort-Object DisplayName

 foreach($MB in $MBs)
 {
  Write-Host -NoNewLine '.'
  $Stat = $MB |
          Get-MailboxStatistics -WarningAction silentlyContinue |
          Select DisplayName,
                 StorageLimitStatus,
                 ItemCount,
                 @{name="TotalItemSizeMB";expression={[math]::Round((($_.TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}},
                 DeletedItemCount,
                 @{name="TotalDeletedItemSizeMB";expression={[math]::Round((($_.TotalDeletedItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}

  $AD = Get-User $MB.SamAccountName
  if ($AD.UserAccountControl -band 2) { $AdEnabled = 'Disabled' } Else { $AdEnabled = 'Enabled' }
 
  $Obj  = New-Object System.Object
  $Obj | Add-Member -MemberType NoteProperty -Name 'User'                -Value $MB.DisplayName
  $Obj | Add-Member -MemberType NoteProperty -Name 'User Name'           -Value $MB.Alias
  $Obj | Add-Member -MemberType NoteProperty -Name 'OU'                  -Value $MB.OrganizationalUnit
  $Obj | Add-Member -MemberType NoteProperty -Name 'SMTP'                -Value $MB.PrimarySmtpAddress
  $Obj | Add-Member -MemberType NoteProperty -Name 'Addresses'           -Value ($MB.EmailAddresses -join ', ')
  $Obj | Add-Member -MemberType NoteProperty -Name 'Forward To'          -Value $MB.ForwardingAddress
  $Obj | Add-Member -MemberType NoteProperty -Name 'Forward and Deliver' -Value $MB.DeliverToMailboxAndForward
  $Obj | Add-Member -MemberType NoteProperty -Name 'Litigation Hold'     -Value $MB.LitigationHoldEnabled
  $Obj | Add-Member -MemberType NoteProperty -Name 'Enabled'             -Value $AdEnabled
  $Obj | Add-Member -MemberType NoteProperty -Name 'Company'             -Value $AD.Company
  $Obj | Add-Member -MemberType NoteProperty -Name 'Department'          -Value $AD.Department
  $Obj | Add-Member -MemberType NoteProperty -Name 'Country'             -Value $AD.CountryOrRegion
  $Obj | Add-Member -MemberType NoteProperty -Name 'Location'            -Value $AD.Office
  $Obj | Add-Member -MemberType NoteProperty -Name 'DB Name'             -Value $DB.Name
  $Obj | Add-Member -MemberType NoteProperty -Name 'Default Quota'       -Value $MB.UseDatabaseQuotaDefaults
  $Obj | Add-Member -MemberType NoteProperty -Name 'Limit Status'        -Value $Stat.StorageLimitStatus
  $Obj | Add-Member -MemberType NoteProperty -Name 'Mail Items'          -Value $Stat.ItemCount
  $Obj | Add-Member -MemberType NoteProperty -Name 'Mailbox Size, MB'    -Value $Stat.TotalItemSizeMB
  $Obj | Add-Member -MemberType NoteProperty -Name 'Deleted Items'       -Value $Stat.DeletedItemCount
  $Obj | Add-Member -MemberType NoteProperty -Name 'Deleted Size, MB'    -Value $Stat.TotalDeletedItemSizeMB

  $Output += $Obj
 }
 Write-Host ' Done'
}

Write-Host -NoNewLine "Exporting mailboxes into: $ReportFile ..."
$Output | Export-Csv $ReportFile -Encoding UTF8 -NoTypeInformation
Write-Host ' Done'

Write-Host -NoNewLine "Emailing this report to: $SmtpTo ..."
If ($SmtpCC -ne '')
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
