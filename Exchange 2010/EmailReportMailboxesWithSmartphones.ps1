# PowerShell script to inventory mailboxes of Exchange 2010 with Smartphones
# registered, listing only the recent communication
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
 [String]$SmtpSubject   = 'Mailboxes with recent Smartphones',
 [String]$SmtpBody      = 'Attached is a list of mailboxes with Smartphones registered, listing only the recent communication.',
 [String]$SmtpBodyTail  = ''
)

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\MailboxSmartphones-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

Write-Host -NoNewLine 'Loading list of mailboxes...'
$MBs = Get-CASMailbox -ResultSize Unlimited | Sort-Object Name
Write-Host ' Done'

Write-Host 'Inspecting mailboxes:'
$ProgressWidth = 80
foreach($MB in $MBs)
{
 $Obj  = New-Object System.Object
 $Obj | Add-Member -MemberType NoteProperty -Name 'Name'             -Value $MB.Name
 $Obj | Add-Member -MemberType NoteProperty -Name 'SMTP'             -Value $MB.PrimarySmtpAddress
 $Obj | Add-Member -MemberType NoteProperty -Name 'ActiveSync'       -Value $MB.ActiveSyncEnabled
 $Obj | Add-Member -MemberType NoteProperty -Name 'OWA'              -Value $MB.OWAEnabled
 $Obj | Add-Member -MemberType NoteProperty -Name 'POP3'             -Value $MB.PopEnabled
 $Obj | Add-Member -MemberType NoteProperty -Name 'IMAP4'            -Value $MB.ImapEnabled
 $Obj | Add-Member -MemberType NoteProperty -Name 'MAPI'             -Value $MB.MapiEnabled
 $Obj | Add-Member -MemberType NoteProperty -Name 'Type'             -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'Model'            -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'OS'               -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'ID'               -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'FirstSync'        -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'LastSync'         -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'Policy'           -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'WipeSent'         -Value '-'
 $Obj | Add-Member -MemberType NoteProperty -Name 'WipedBy'          -Value '-'

 $Result = ' ' ## .

 try
 {
  try
  {
   $SPs  = Get-ActiveSyncDevice -Mailbox $MB.Identity

   if ($SPs)
   {
    try
    {
     $SPss = $SPs | Get-ActiveSyncDeviceStatistics
     $SP   = $SPss | Sort-Object -Descending LastSyncAttemptTime | Select-Object -First 1

     $Obj.Type      = $SP.DeviceType
     $Obj.Model     = $SP.DeviceModel
     $Obj.OS        = $SP.DeviceOS
     $Obj.ID        = $SP.DeviceID
     $Obj.FirstSync = $SP.FirstSyncTime
     $Obj.LastSync  = $SP.LastSuccessSync
     $Obj.Policy    = $SP.DevicePolicyApplied
     $Obj.WipeSent  = $SP.DeviceWipeSentTime
     $Obj.WipedBy   = $SP.LastDeviceWipeRequestor

     $Result = '.'
    }
    catch
    { $Result = 'X' }
   }
  }
  catch
  { $Result = 'Y' }
 }
 catch
 { $Result = '?' }

 $Output += $Obj

 Write-Host -NoNewLine $Result
 $ProgressWidth = $ProgressWidth - 1; if ($ProgressWidth -eq 0) { Write-Host ''; $ProgressWidth = 80 }
}
Write-Host "`nDone"

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
