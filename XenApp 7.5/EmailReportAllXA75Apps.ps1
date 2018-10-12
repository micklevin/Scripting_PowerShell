# PowerShell script to inventory published applications of Citrix XenApp 7.5
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
 [String]$SmtpFrom      = 'Citrix Reports <citrix.reports@domain.local>',
 [String]$SmtpTo        = 'citrix.reports@domain.local',
 [String]$SmtpCc        = '',
 [String]$SmtpSubject   = 'XenApp 7.5 Applications',
 [String]$SmtpBody      = 'Attached is a list of all applications in XenApp 6.5 Farm.',
 [String]$SmtpBodyTail  = ''
)

Add-PSSnapIn 'Citrix*'

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllXenApp75Apps-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()

Write-Host -NoNewLine 'Connecting to the XenApp 7.5 Farm...'
$XA75apps    = Get-BrokerApplication | Sort-Object PublishedName
Write-Host ' Done'

Write-Host -NoNewLine 'Inventorying applications '

foreach($XA75app in $XA75apps)
{
 Write-Host -NoNewLine '.'

 $Obj  = New-Object System.Object
 $Obj | Add-Member -MemberType NoteProperty -Name 'Application'         -Value $XA75app.PublishedName
 $Obj | Add-Member -MemberType NoteProperty -Name 'DG'                  -Value ($XA75app.AssociatedDesktopGroupUids -join '; ')
 $Obj | Add-Member -MemberType NoteProperty -Name 'Enabled'             -Value $XA75app.Enabled
 $Obj | Add-Member -MemberType NoteProperty -Name 'Executable'          -Value $XA75app.CommandLineExecutable
 $Obj | Add-Member -MemberType NoteProperty -Name 'Arguments'           -Value $XA75app.CommandLineArguments
 $Obj | Add-Member -MemberType NoteProperty -Name 'Users'               -Value ($XA75app.AssociatedUserNames -join '; ')

 $Output += $Obj

 Write-Host ' Done'
}

Write-Host -NoNewLine "Exporting information into: $ReportFile ..."
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
