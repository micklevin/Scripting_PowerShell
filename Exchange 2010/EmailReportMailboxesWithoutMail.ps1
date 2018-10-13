# PowerShell script to inventory mailboxes of Exchange 2010 Organization, which Active Directory
# accounts on some reason are missing the WindowsEmailAddress attribute
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
 [String]$SmtpSubject   = 'Mailboxes without Email',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes without the WindowsEmailAddress.',
 [String]$SmtpBodyTail  = ''
)

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllMailboxWithoutWindowsEmailAddress-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

Write-Host -NoNewLine 'Loading list of databases...'
$DBs    = Get-MailboxDatabase | Sort-Object Name
Write-Host ' Done'

foreach($DB in $DBs)
{
 Write-Host -NoNewLine "Inventorying mailboxes of $($DB.Name) "
 $MBs = Get-Mailbox -Database $DB.Name | Where-Object { (Get-User $_.Alias).WindowsEmailAddress -eq $false } | Sort-Object DisplayName

 foreach($MB in $MBs)
 {
  Write-Host -NoNewLine '.'
 
  $Obj  = New-Object System.Object
  $Obj | Add-Member -MemberType NoteProperty -Name 'User'             -Value $MB.DisplayName
  $Obj | Add-Member -MemberType NoteProperty -Name 'Alias'            -Value $MB.Alias
  $Obj | Add-Member -MemberType NoteProperty -Name 'OU'               -Value $MB.OrganizationalUnit
  $Obj | Add-Member -MemberType NoteProperty -Name 'Primary SMTP'     -Value $MB.PrimarySmtpAddress
  $Obj | Add-Member -MemberType NoteProperty -Name 'Addresses'        -Value ($MB.EmailAddresses -join ', ')

  $Output += $Obj
 }
 Write-Host ' Done'
}

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
