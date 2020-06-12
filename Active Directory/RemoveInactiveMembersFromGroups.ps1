# Version 1.0
#  input:
#   list of AD groups, inactivity days (default = 90)
#  output:
#   remove from these groups: all user accounts which are disabled, or are inactive for more than these number days; no limit by OU
#
# Attention:
#  In order to run this script from Scheduler or Command Prompt, use tripple-quote symbols to enclose the strings. And no spaces
#  between comma-separated list of groups! For example:
#  
#  powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\AD\Delete_AD_Group_Inactive_Members.ps1" -Mode "Report" -AdGroups "KELO-StandardVPN-Access","""VPN Carlota Users""","""VPN Corporate Users"""

# Configuration

param
(
 [String]$Mode          = 'List',
 [string[]]$AdGroups    = @('VPN * Users'),
 [String]$DaysInactive  = 90,
 [String]$SmtpServer    = 'kwcasarray.quadra.local',
 [String]$SmtpFrom      = 'KWITAUTO01 <KWITAUTO01@quadra.local>',
 [String]$SmtpTo        = 'corporateit@ca.kghm.com',
 [String]$SmtpCc        = ''
)

# Some input validation
if(($Mode -ne 'Remove') -and `
   ($Mode -ne 'List'))    { $Mode         = 'List'}
if($AdGroups.Count -lt 1) { $AdGroups     = @('KELO-StandardVPN-Access') }
if($DaysInactive -lt 1)   { $DaysInactive = 1 }

#-------------------
Function Format-OU()
{
 param
 (
  [String]$OUString
 )

 $OUString = $OUString.Substring(0, $OUString.IndexOf('DC') - 1)
 $OUArray  = $OUString.Split(',')
 $Result   = ''
 $Count    = 1

 $OUArray |
 ForEach-Object `
 {
  if ($Count -eq 1)
  { $Result = $Result.Insert(0, $_) }
  else
  { $Result = $Result.Insert(0, "$_,") }
  $Count++
 }

 return $Result -replace "^OU=", "/" -replace ",$" -replace ",OU=", "/"
}

#-------------------

Import-Module ActiveDirectory

# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\InactiveUsersInGroups-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$MembersToX = @()

#-------------------

$RecentDays  = (Get-date).AddDays(0 - $DaysInactive)
$GroupsTotal = $AdGroups | Measure-Object | Select-Object -ExpandProperty Count
$GroupsCount = 0

foreach ($AdGroup in $AdGroups)
{
 $AllAdGroups = (Get-ADGroup -Filter {Name -like $AdGroup} | Sort-Object Name)

 if ($AllAdGroups)
 {
  $GroupsTotal += (($AllAdGroups | Measure-Object | Select-Object -ExpandProperty Count) - 1)

  foreach ($AdGroupName in $AllAdGroups)
  {
   $GroupsCount++
   Write-Progress -Id 1 -Activity 'Inventorying groups' -status ("[$GroupsCount / $GroupsTotal] " + $AdGroupName.Name) -percentComplete ($GroupsCount / $GroupsTotal * 100)

   try
   { $GroupMembers = (Get-ADGroupMember $AdGroupName.Name | Where-Object {$_.objectClass -eq 'user'} | Sort-Object Name) }
   catch
   { ('ERROR: Could not load up members of group "' + $AdGroupName.Name + '"') }

   if ($GroupMembers)
   {
    $MembersTotal = $GroupMembers | Measure-Object | Select-Object -ExpandProperty Count
    $MembersCount = 0

    foreach ($GroupMember in $GroupMembers)
    {
     $MembersCount++
     Write-Progress -ParentId 1 -Id 2 -Activity 'Inventorying users' -status ("[$MembersCount / $MembersTotal] " + $GroupMember.Name) -percentComplete ($MembersCount / $MembersTotal * 100)

     $AdUser = Get-ADUser $GroupMember.distinguishedName -Properties LastLogonTimeStamp, Created, PasswordLastSet
     $Ou     = Format-OU -OUString (($AdUser.DistinguishedName -split '(,OU=)' | Select-Object -Skip 1) -join '') -replace '^,'

     if ($AdUser.Enabled -eq $false)
     {
      $MembersToX += New-Object PsObject -Property `
      @{
        'Group'          = $AdGroupName.Name
        'Name'           = $AdUser.Name
        'sAMAccountName' = $AdUser.SamAccountName
        'OU'             = $Ou
        'Is Enabled'     = 'No'
        'Last Logon'     = ''
        'Password Set'   = ''
        'Removed'        = $false
      }
     }
     else
     {
      if ($AdUser.Created -and ($AdUser.Created -lt $RecentDays))
      {
       if ($AdUser.LastLogonTimeStamp)
       {
        $LastLogonTimeStamp = ([datetime]::FromFileTime($AdUser.LastLogonTimeStamp))
        if ($LastLogonTimeStamp -lt $RecentDays)
        {
         $MembersToX += New-Object PsObject -Property `
         @{
           'Group'          = $AdGroupName.Name
           'Name'           = $AdUser.Name
           'sAMAccountName' = $AdUser.SamAccountName
           'OU'             = $Ou
           'Is Enabled'     = 'Yes'
           'Last Logon'     = $LastLogonTimeStamp
           'Password Set'   = ''
           'Removed'        = $false
         }
        }
       }
       else
       {
        if ($AdUser.PasswordLastSet -and ($AdUser.PasswordLastSet -lt $RecentDays))
        {
         $MembersToX += New-Object PsObject -Property `
         @{
           'Group'          = $AdGroupName.Name
           'Name'           = $AdUser.Name
           'sAMAccountName' = $AdUser.SamAccountName
           'OU'             = $Ou
           'Is Enabled'     = 'Yes'
           'Last Logon'     = ''
           'Password Set'   = $AdUser.PasswordLastSet
           'Removed'        = $false
         }
        }
       }
      }
     }
    }
    Write-Progress -ParentId 1 -Id 2 -Activity 'Inventorying users' -Status 'done' -Completed
   }
  }
 }
}
Write-Progress -Id 1 -Activity 'Inventorying groups' -Status 'done' -Completed

if ($MembersToX)
{
 if ($Mode -eq 'Remove')
 {
  $RemovedTotal = $MembersToX | Measure-Object | Select-Object -ExpandProperty Count
  $RemovedCount = 0

  $MembersToX |
  Sort-Object Group, Name |
  ForEach-Object `
  {
   $RemovedCount++
   Write-Progress -Id 1 -Activity 'Removing users from groups' -status ("[$RemovedCount / $RemovedTotal] " + $_.Group + ': ' + $_.sAMAccountName) -percentComplete ($RemovedCount / $RemovedTotal * 100)
   try
   {
    Remove-ADGroupMember -Identity $_.Group -Members $_.sAMAccountName -Confirm:$false
    $_.Removed = $true
   }
   catch
   { ('ERROR: Could not remove "' + $_.sAMAccountName + '" from "' + $_.Group + '"') }
  }
  Write-Progress -Id 1 -Activity 'Removing users from groups' -Status 'done' -Completed

  $SmtpSubject  = "REMOVED: Disabled or inactive for $DaysInactive days users of groups"
  $SmtpBody     = "Attached is a list of groups and their members which are disabled or inactive for $DaysInactive days or more.`n`nThese members have been removed from corresponding groups, if indicated in the list.`n"
 }
 Else
 {
  $SmtpSubject  = "REPORT: Disabled or inactive for $DaysInactive days users of groups"
  $SmtpBody     = "Attached is a list of groups and their members which are disabled or inactive for $DaysInactive days or more.`n`nConsider removing these members from corresponding groups.`n"
 }

 Write-Progress -Id 1 -Activity 'Sending email report' -status ('Export into: ' + $ReportFile) -percentComplete 50
 $MembersToX | Export-Csv $ReportFile -NoTypeInformation -Encoding UTF8

 Write-Progress -Id 1 -Activity 'Sending email report' -status ('Emailing to: ' + $SmtpTo) -percentComplete 100
 If ($SmtpCC -ne '')
 { Send-Mailmessage -SmtpServer $SmtpServer -From $SmtpFrom -To $SmtpTo -Cc $SmtpCc -Subject $SmtpSubject -Body $SmtpBody -Priority High -Attachments $ReportFile }
 Else
 { Send-Mailmessage -SmtpServer $SmtpServer -From $SmtpFrom -To $SmtpTo -Subject $SmtpSubject -Body $SmtpBody -Priority High -Attachments $ReportFile }
 Write-Progress -Id 1 -Activity 'Sending email report' -Status 'done' -Completed
}
Else
{ 'Nothing to report' }
