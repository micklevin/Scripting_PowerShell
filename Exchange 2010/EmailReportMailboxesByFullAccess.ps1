# PowerShell script to inventory mailboxes, which allowed somebody else to have Full Access in Exchange 2010 Organization
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
 [String]$SmtpSubject   = 'Mailbox Full Access',
 [String]$SmtpBody      = 'Attached is a list of all mailboxes with full access to them.',
 [String]$SmtpBodyTail  = ''
)

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllMailboxFullAccess-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

Write-Host -NoNewLine 'Loading list of mailboxes...'
$MBXs = Get-Mailbox -ResultSize Unlimited | Sort-Object Name
Write-Host ' Done'

Write-Host -NoNewLine 'Loading mailbox explicit permissions...'
$MBXPerms = $MBXs | Get-MailboxPermission | Where-Object {($_.IsInherited -eq $false) -and ($_.User.tostring() -ne "NT AUTHORITY\SELF")}
Write-Host ' Done'

Write-Host -NoNewLine "Exporting information into: $ReportFile ..."
$MBXPerms |
Select-Object Identity, User, @{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} |
Sort-Object Identity, User |
Export-Csv $ReportFile -Encoding UTF8 -NoTypeInformation
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
