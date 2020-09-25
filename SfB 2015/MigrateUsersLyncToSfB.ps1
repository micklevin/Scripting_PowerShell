<#
.SYNOPSIS
PowerShell script to move users from Lync 2010 pool to Skype for Business 2015 pool,
and then to modify their SIP address based on an input from CSV file.

.DESCRIPTION
This script takes a list of users from the CSV file, and moves each of them to desired SfB pool.
After that is done, it overwrites the SIP address, using the data from same CSV file.
That second step is useful when SfB SIP of a user must match [new] Exchange namespace for that users.

The CSV file has to have the following columns:
username,sipaddress

For example:
username,sipaddress
jdoe,sip:jdoe@company.com

.PARAMETER CsvFile
Full patch to the CSV file.

.PARAMETER NewPool
FQDN of SfB 2015 pool server.

.PARAMETER WaitSecs
Number of seconds to wait for SfB Topology to synchronize, estimate.

.PARAMETER Test
Test run, this is a default.

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\SfB\MigrateUsersLyncToSfB.ps1" -CsvFile "C:\bin\Scripts\Test.csv" -NewPool "SFBCS01.domain.local" -WaitSecs 30 {-Test:False}

.NOTES
Name:    RemoveInactiveMembersFromGroups.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0.1

.LINK
https://github.com/micklevin/Scripting_PowerShell

#>
param
(
 [string]$CsvFile = 'C:\bin\Scripts\Test.csv',
 [string]$NewPool = 'SFBCS01.domain.local',
 [int]$WaitSecs   = 900,
 [bool]$Test      = $true
)

#--------------------------------------
try
{ Import-Module 'C:\Program Files\Common Files\Skype for Business Server 2015\Modules\SkypeForBusiness\SkypeForBusiness.psd1' -ErrorAction Stop }
catch
{
 'ERROR: The PowerShell module for Skype for Business 2015 was not found'
 exit 1
}

#--------------------------------------------------
function Test-CsManagementStoreReplicationStatus ()
{
 $return = $true
 Get-CsManagementStoreReplicationStatus | ForEach-Object `
 {
  if (-not $_.UpToDate)
  { $return = $false }
 }

 return $return
}

#--------------------------------------------------
function Wait-CsManagementStoreReplicationStatus ()
{
 if(Test-CsManagementStoreReplicationStatus)
 { Write-Host "`nManagement store does not require replication" }
 else
 {
  Write-Host -NoNewline "`nReplication of management store is still in progress. Checking every 5 seconds: "
  do
  {
   Write-Host -NoNewline '.'
   Start-Sleep -s 5
  }
  until (Test-CsManagementStoreReplicationStatus)
  Write-Host "`nReplication finished"
 }
}

#--------------------------------------------------

if($CsvFile -ne '')
{
 if(Test-Path $CsvFile)
 {
  $Users = Import-Csv -path $CsvFile -Delimiter ',' -Encoding UTF8

  if($Users -ne $null)
  {
   $UsersCount   = $Users | Measure-Object | Select-Object -ExpandProperty Count
   $SuccessCount = 0

   Wait-CsManagementStoreReplicationStatus

   Write-Host ("`nMoving " + $UsersCount + ' users')

   foreach($AUser in $Users)
   {
    if($Test)
    {
     Write-Host (' ' + $AUser.username + ' (testing)')
     Move-CsUser -Identity $AUser.username -Target $NewPool -Confirm:$false -WhatIf
    }
    else
    {
     Write-Host (' ' + $AUser.username)
     try
     {
      Move-CsUser -Identity $AUser.username -Target $NewPool -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
      $AUser | Add-Member NoteProperty Success $true
      $SuccessCount++
     }
     catch
     {
      Write-Error ('  Error:' + $_.Exception.Message)
      $AUser | Add-Member NoteProperty Success $false
     }
    }
   }

   Write-Host ('Moved ' + $SuccessCount + ' of ' + $UsersCount + ' users')

   Write-Host ("`nWaiting " + $WaitSecs + ' seconds...')
   Start-Sleep -s $WaitSecs

   Wait-CsManagementStoreReplicationStatus

   Write-Host ('Modifying SIP addresses for ' + $SuccessCount + ' users')
   foreach($AUser in $Users)
   {
    if($AUser.Success)
    {
     if($Test)
     {
      Write-Host (' ' + $AUser.sipaddress + ' (testing)')
      Set-CsUser -Identity $AUser.username -SipAddress $AUser.sipaddress  -Confirm:$false -WhatIf
     }
     else
     {
      Write-Host (' ' + $AUser.sipaddress)
      try
      {
       Set-CsUser -Identity $AUser.username -SipAddress $AUser.sipaddress  -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
      }
      catch
      { Write-Error ('  Error:' + $_.Exception.Message) }
     }
    }
    else
    { Write-Host ('-' + $AUser.sipaddress) }
   }

   Write-Host "`nDone`nPlease check the Errors and Warnings in the output above.`n"
  }
  else
  {
   Write-Host ('Error: No users found in the CSV file "' + $CsvFile + '"')
   Exit 0
  }
 }
 else
 {
  Write-Host ('Error: The CSV file "' + $CsvFile + '" does not exists')
  Exit 0
 }
}
else
{
 Write-Host 'Error: No CSV file provided'
 Exit 0
}
