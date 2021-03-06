<#
.SYNOPSIS
PowerShell script to remove inactive and disabled users from certain AD groups

.DESCRIPTION
This script takes a list of AD groups and checks each for inactive or disabled users.
Such users then are removed from the groups.

.PARAMETER Mode
The mode of operation - list only or list and remove.

.PARAMETER AdGroups
The list of AD groups.

.PARAMETER DaysInactive
Number of days of inactivity to trigger user's removal.

.PARAMETER SmtpServer
The DNS name of SMTP server

.PARAMETER SmtpFrom
The SMTP name and address of Sender.

.PARAMETER SmtpTo
The SMTP address of recipient(s).

.PARAMETER SmtpSubject
Alternative text for the email's Subject.

.PARAMETER SmtpBody
Alternative text for the email's body.

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\AD\RemoveInactiveMembersFromGroups.ps1" -Mode "Report" -AdGroups """Group 1"""","""Group 2""","""Group 3"""

.NOTES
Name:    RemoveInactiveMembersFromGroups.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0.1

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

#--------------------------------------
param
(
 [String]$Mode          = 'List',
 [string[]]$AdGroups    = @('VPN * Users'),
 [String]$DaysInactive  = 90,
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Active Directory Reports <ad.reports@domain.local>',
 [String]$SmtpTo        = 'ad.reports@domain.local'
)

#--------------------------------------
# Some input validation
if (($Mode -ne 'Remove') -and `
    ($Mode -ne 'List'))  { $Mode         = 'List'}
if ($DaysInactive -lt 1) { $DaysInactive = 1 }

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

#--------------------------------------
# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\InactiveUsersInGroups-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$MembersToX = @()

#--------------------------------------
try
{ Import-module ActiveDirectory }
catch
{
 'ERROR: The PowerShell snap-in for Active Directory was not found'
 exit 1
}

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
