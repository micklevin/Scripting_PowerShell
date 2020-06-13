<#
.SYNOPSIS
PowerShell script to report all updates on WSUS

.DESCRIPTION
This script does a heavy-lifting of reporting of status of all updates on all WSUS servers in hierarchy. Use it when
report rollup does not work.
Script loads the list of reports from each WSUS server, then loads the list of computers, and tries to match each update
to each computer's status and understand if an update is needed, installed or not installed.
Resulting 3 CSV files will help you figure which of updates are not needed by any computer, and hence could be declined
in order to save precious space in WID and SQL databases, as well as cleanup unused updates from the WSUS server's disk.

.PARAMETER SmtpServer
The DNS name or IP address of SMTP server. Note - this script does not support SMTP authentication

.PARAMETER SmtpFrom
The SMTP name and address of Sender.

.PARAMETER SmtpTo
The SMTP address of recipient(s).

.PARAMETER SmtpSubject
Alternative text for the email's Subject.

.PARAMETER SmtpBody
Alternative text for the email's body.

.NOTES
Name:    GetAllWsusUpdates.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

#--------------------------------------
param
(
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'WSUS Reports <wsus.reports@domain.local>',
 [String]$SmtpTo        = 'wsus.reports@domain.local',
 [String]$SmtpSubject   = 'Report on WSUS Updates',
 [String]$SmtpBody      = 'Attached are the reports on WSUS updates: Needed, Installed and Not Installed.'
)

#--------------------------------------
# Run-time Parameters
$CurrDate                     = Get-Date
$MyDocuments                  = [environment]::getfolderpath('mydocuments')
$Report_Updates_Needed        = $MyDocuments + '\WSUS-Updates-Needed-'        + $LookFor + '-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Report_Updates_Installed     = $MyDocuments + '\WSUS-Updates-Installed-'     + $LookFor + '-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Report_Updates_Not_Installed = $MyDocuments + '\WSUS-Updates-Not-Installed-' + $LookFor + '-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'

try
{ [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null }
catch
{ 'WSUS API already loaded' }

$WsusServers = @('WSUSUS01','WSUSDS01','WSUSDS02','WSUSDS03','WSUSDS04','WSUSDS05','WSUSDS06','WSUSDS07')

$UpdateScope                            = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.ApprovedStates             = [Microsoft.UpdateServices.Administration.ApprovedStates]::Any
$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All

$UpdatesNeeded       = @{}
$UpdatesInstalled    = @{}
$UpdatesNotInstalled = @{}
$UpdatesTitles       = @{}
$UpdatesCreated      = @{}

$WsusServers | ForEach-Object `
{
 Write-Host ("`nWSUS Server: " + $_)

 $WsusServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($_,$true, 8531)
 If ($WsusServer)
 {
  $ComputerTargets = $WsusServer.GetComputerTargets().GetEnumerator()
  $NameLen         = 30

  ForEach ($ComputerTarget in $ComputerTargets)
  {
   $NameLen = [math]::max($NameLen, $ComputerTarget.FullDomainName.Length)
   Write-Host -NoNewline (' Client ' + $ComputerTarget.FullDomainName.PadLeft($NameLen) + "`t: ")

   $UpdatesInfo = $ComputerTarget.GetUpdateInstallationInfoPerUpdate($UpdateScope)

   $UpdatesCount   = $UpdatesInfo | Measure-Object | Select-Object -ExpandProperty Count
   $UpdatesCurrent = 0
   $CurrPercent    = 0
   $PrevPercent    = 0

   ForEach ($UpdateInfo in $UpdatesInfo)
   {
    $UpdatesCurrent++
    $CurrPercent = [int](100 * ([float]($UpdatesCurrent / $UpdatesCount)))
    if ($CurrPercent -gt $PrevPercent)
    {
     Write-Host -NoNewline ('=' * ($CurrPercent -  $PrevPercent))
     $PrevPercent = $CurrPercent
    }
    Switch ($UpdateInfo.UpdateInstallationState)
    {
     'Failed'
     {
      If ($UpdatesNeeded.ContainsKey($UpdateInfo.UpdateId))
      { $UpdatesNeeded[$UpdateInfo.UpdateId]++ }
      else
      { $UpdatesNeeded.Set_Item($UpdateInfo.UpdateId, 1) }
      break;
     }
     'Downloaded'
     {
      If ($UpdatesNeeded.ContainsKey($UpdateInfo.UpdateId))
      { $UpdatesNeeded[$UpdateInfo.UpdateId]++ }
      else
      { $UpdatesNeeded.Set_Item($UpdateInfo.UpdateId, 1) }
      break;
     }
     'InstalledPendingReboot'
     {
      If ($UpdatesNeeded.ContainsKey($UpdateInfo.UpdateId))
      { $UpdatesNeeded[$UpdateInfo.UpdateId]++ }
      else
      { $UpdatesNeeded.Set_Item($UpdateInfo.UpdateId, 1) }
      break;
     }
     'Installed'
     {
      If ($UpdatesInstalled.ContainsKey($UpdateInfo.UpdateId))
      { $UpdatesInstalled[$UpdateInfo.UpdateId]++ }
      else
      { $UpdatesInstalled.Set_Item($UpdateInfo.UpdateId, 1) }
      break;
     }
     'NotInstalled'
     {
      If ($UpdatesNotInstalled.ContainsKey($UpdateInfo.UpdateId))
      { $UpdatesNotInstalled[$UpdateInfo.UpdateId]++ }
      else
      { $UpdatesNotInstalled.Set_Item($UpdateInfo.UpdateId, 1) }
      break;
     }
     'NotApplicable' { break; }
     'Unknown'       { break; }
     default
     { Write-Host -NoNewline ('[' + $UpdateInfo.UpdateInstallationState + ']') }
    }
   }
   Write-Host ''
  }
 }

 Write-Host -NoNewline "Loading Needed Update's names           : "
 $UpdatesCount   = $UpdatesNeeded.Keys | Measure-Object | Select-Object -ExpandProperty Count
 $UpdatesCurrent = 0
 $CurrPercent    = 0
 $PrevPercent    = 0
 $UpdatesNeeded.Keys | ForEach-Object `
 {
  $UpdatesCurrent++
  $CurrPercent = [int](100 * ([float]($UpdatesCurrent / $UpdatesCount)))
  if ($CurrPercent -gt $PrevPercent)
  {
   Write-Host -NoNewline ('=' * ($CurrPercent -  $PrevPercent))
   $PrevPercent = $CurrPercent
  }
  If (-not $UpdatesTitles.ContainsKey($_))
  {
   $UpdateInfo = $WsusServer.GetUpdate($_)
   $UpdatesTitles.Set_Item($_, $UpdateInfo.Title)
   $UpdatesCreated.Set_Item($_, $UpdateInfo.CreationDate)
  }
 }

 Write-Host ''
 Write-Host -NoNewline "Loading Installed Update's names        : "
 $UpdatesCount   = $UpdatesInstalled.Keys | Measure-Object | Select-Object -ExpandProperty Count
 $UpdatesCurrent = 0
 $CurrPercent    = 0
 $PrevPercent    = 0
 $UpdatesInstalled.Keys | ForEach-Object `
 {
  $UpdatesCurrent++
  $CurrPercent = [int](100 * ([float]($UpdatesCurrent / $UpdatesCount)))
  if ($CurrPercent -gt $PrevPercent)
  {
   Write-Host -NoNewline ('=' * ($CurrPercent -  $PrevPercent))
   $PrevPercent = $CurrPercent
  }
  If (-not $UpdatesTitles.ContainsKey($_))
  {
   $UpdateInfo = $WsusServer.GetUpdate($_)
   $UpdatesTitles.Set_Item($_, $UpdateInfo.Title)
   $UpdatesCreated.Set_Item($_, $UpdateInfo.CreationDate)
  }
 }

 Write-Host ''
 Write-Host -NoNewline "Loading Not Installed Update's names    : "
 $UpdatesCount   = $UpdatesNotInstalled.Keys | Measure-Object | Select-Object -ExpandProperty Count
 $UpdatesCurrent = 0
 $CurrPercent    = 0
 $PrevPercent    = 0
 $UpdatesNotInstalled.Keys | ForEach-Object `
 {
  $UpdatesCurrent++
  $CurrPercent = [int](100 * ([float]($UpdatesCurrent / $UpdatesCount)))
  if ($CurrPercent -gt $PrevPercent)
  {
   Write-Host -NoNewline ('=' * ($CurrPercent -  $PrevPercent))
   $PrevPercent = $CurrPercent
  }
  If (-not $UpdatesTitles.ContainsKey($_))
  {
   $UpdateInfo = $WsusServer.GetUpdate($_)
   $UpdatesTitles.Set_Item($_, $UpdateInfo.Title)
   $UpdatesCreated.Set_Item($_, $UpdateInfo.CreationDate)
  }
 }
 Write-Host ''
}

Write-Host -NoNewline 'Exporting Updates Needed                : '
$UpdatesNeeded.Keys |
ForEach-Object `
{
 $Row = New-Object PSObject;
 $Row | Add-Member NoteProperty -Name 'Id' -Value $_ ;
 $Row | Add-Member NoteProperty -Name 'Count' -Value $UpdatesNeeded[$_];
 $Row | Add-Member NoteProperty -Name 'Name' -Value $UpdatesTitles[$_];
 $Row | Add-Member NoteProperty -Name 'Created' -Value $UpdatesCreated[$_];
 $Row
} |
Export-Csv -Path $Report_Updates_Needed -Encoding UTF8 -NoTypeInformation
Write-Host 'Done'

Write-Host -NoNewline 'Exporting Updates Installed             : '
$UpdatesInstalled.Keys |
ForEach-Object `
{
 $Row = New-Object PSObject;
 $Row | Add-Member NoteProperty -Name 'Id' -Value $_ ;
 $Row | Add-Member NoteProperty -Name 'Count' -Value $UpdatesInstalled[$_];
 $Row | Add-Member NoteProperty -Name 'Name' -Value $UpdatesTitles[$_];
 $Row | Add-Member NoteProperty -Name 'Created' -Value $UpdatesCreated[$_];
 $Row
} |
Export-Csv -Path $Report_Updates_Installed -Encoding UTF8 -NoTypeInformation
Write-Host 'Done'

Write-Host -NoNewline 'Exporting Updates Not Installed         : '
$UpdatesNotInstalled.Keys |
ForEach-Object `
{
 $Row = New-Object PSObject;
 $Row | Add-Member NoteProperty -Name 'Id' -Value $_ ;
 $Row | Add-Member NoteProperty -Name 'Count' -Value $UpdatesNotInstalled[$_];
 $Row | Add-Member NoteProperty -Name 'Name' -Value $UpdatesTitles[$_];
 $Row | Add-Member NoteProperty -Name 'Created' -Value $UpdatesCreated[$_];
 $Row
} |
Export-Csv -Path $Report_Updates_Not_Installed -Encoding UTF8 -NoTypeInformation
Write-Host 'Done'

Write-Host -NoNewline 'Sending reports via email               : '
Send-Mailmessage -SmtpServer $SmtpServer `
                 -From $SmtpFrom `
                 -To $SmtpTo `
                 -Subject $SmtpSubject `
                 -Body $SmtpBody `
                 -Priority High `
                 -Attachments $Report_Updates_Needed,$Report_Updates_Installed,$Report_Updates_Not_Installed
Write-Host 'Done'

Remove-Item $Report_Updates_Needed -force
Remove-Item $Report_Updates_Installed -force
Remove-Item $Report_Updates_Not_Installed -force

Exit
