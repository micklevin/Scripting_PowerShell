# PowerShell script to inventory mailboxes with Litigation Hold in Exchange 2010 Organization
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0.1
#
# Configuration

param
(
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Exchange Reports <exchange.reports@domain.local>',
 [String]$SmtpTo        = 'exchange.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = 'Exchange Litigation Hold Report',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes with Litigation Hold enabled.',
 [String]$SmtpBodyTail  = ''
)

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllLitigationHold-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()

$WarningActionPreference = 'SilentlyContinue'

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction $WarningActionPreference

Write-Host -NoNewLine 'Loading list of databases...'
$DBs = Get-MailboxDatabase | Sort-Object Name
Write-Host ' Done'

foreach ($DB in $DBs)
{
 Write-Host -NoNewLine "Inventorying mailboxes of $($DB.Name): "

 Get-Mailbox -Database $DB.Name |
 Sort-Object DisplayName |
 Where-Object { $_.LitigationHoldEnabled -eq $True } |
 Select-Object DisplayName, Alias, PrimarySmtpAddress, LitigationHoldDate, LitigationHoldOwner |
 ForEach-Object `
 {
  $AD = Get-User $_.Alias
  if ($AD.UserAccountControl -band 2) { $AdEnabled = 'Disabled' } Else { $AdEnabled = 'Enabled' }

  $Obj = New-Object System.Object
  $Obj | Add-Member -MemberType NoteProperty -Name 'Name'            -Value $_.DisplayName
  $Obj | Add-Member -MemberType NoteProperty -Name 'Account'         -Value $_.Alias
  $Obj | Add-Member -MemberType NoteProperty -Name 'Account Enabled' -Value $AdEnabled
  $Obj | Add-Member -MemberType NoteProperty -Name 'Email'           -Value $_.PrimarySmtpAddress
  $Obj | Add-Member -MemberType NoteProperty -Name 'Hold Enabled On' -Value $_.LitigationHoldDate
  $Obj | Add-Member -MemberType NoteProperty -Name 'Hold Enabled By' -Value $_.LitigationHoldOwner
  $Output += $Obj
  Write-Host -NoNewLine '.'
 }
 Write-Host ''
}
Write-Host 'Done'

Write-Host -NoNewLine "Exporting this report into file: $ReportFile ..."
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
