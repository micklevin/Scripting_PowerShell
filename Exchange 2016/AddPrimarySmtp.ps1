<#
.SYNOPSIS
PowerShell script to add and set-as-primary the SMTP addresses for a group of users,
based on an input from CSV file.

.DESCRIPTION
This script takes a list of users from the CSV file, and adds listed SMTP address per each user.
Then it makes that SMTP address the Primary.

The CSV file has to have the following columns:
User,Login,NewPrimarySmtp

For example:

User,Login,NewPrimarySmtp
User One,user1,user.one@company.com
User Two,user2,user.two@company.com

.PARAMETER PSSnapin
Dafault name for Exchange 2016 snap-in for PowerShell.

.PARAMETER CsvFile
Full patch to the CSV file.

.PARAMETER Test
Test run, this is a default.

.PARAMETER Test
Debug (no-run), this is a default.

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\Exchange 2016\AddPrimarySmtp.ps1" -CsvFile "C:\bin\Scripts\Test.csv" {-Test:False} {-Debug:False}

.NOTES
Name:    AddPrimarySmtp.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0.1

.LINK
https://github.com/micklevin/Scripting_PowerShell

#>
param
(
 [string]$PSSnapin = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
 [string]$CsvFile  = 'C:\Migration\test.csv',
 [bool]$Test       = $true,
 [bool]$Debug      = $true
)

#-------------------
# Load Exchange snap-in and environment

if(!(Get-PSSnapin | Where-Object {$_.Name -eq $PSSnapin}))
{
 if($Debug) { 'DEBUG: The PowerShell snap-in "' + $PSSnapin + '" not loaded, trying to load now...' }
 try
 { Add-PSSnapin $PSSnapin }
 catch
 {
  ('ERROR: The PowerShell snap-in "' + $PSSnapin + '" failed to load')
  exit 1
 }
 . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
 Connect-ExchangeServer -auto -AllowClobber
}
else
{ if($Debug) { 'DEBUG: The PowerShell snap-in "' + $PSSnapin + '" already loaded'} }

#--------------------------------------------------
# Process the CSV file

if($CsvFile -ne '')
{
 if(Test-Path $CsvFile)
 {
  $Users = Import-Csv -path $CsvFile -Delimiter ',' -Encoding UTF8

  if($null -ne $Users)
  {
   $UsersCount   = $Users | Measure-Object | Select-Object -ExpandProperty Count
   $SuccessCount = 0

   Write-Host ("`nUpdating " + $UsersCount + ' users')

   foreach($AUser in $Users)
   {
    if (($null -ne $AUser.User) -and ($null -ne $AUser.Login) -and ($null -ne $AUser.NewPrimarySmtp))
    {
     $AMailbox = (Get-Mailbox $AUser.Login | Where-Object {-not ($_.EmailAddresses -like ('*:' + $AUser.NewPrimarySmtp + '*'))})
     if($AMailbox)
     {
      # There is the mailbox for this user, and given NewPrimarySmtp is not yet in the list of SMTP addresses
      $NewSmtp  = 'smtp:' + $AUser.NewPrimarySmtp
      if($Test)
      {
       # Test run, expect errors
       Write-Host -NoNewline ('(testing) ' + $AUser.User + "`t(" + $AMailbox.PrimarySmtpAddress + ' -> ' + $AUser.NewPrimarySmtp + ')')
       if($AMailbox.EmailAddressPolicyEnabled)
       {
        Write-Host "`tEmail Address Policy is On, turning it Off..."
        try
        { $AMailbox | Set-Mailbox -EmailAddressPolicyEnabled $false -ErrorAction Stop -Verbose -WhatIf }
        catch
        { Write-Error ('  Error:' + $_.Exception.Message) }
       }
       else
       { Write-Host "`tEmail Address Policy is Off" }
       $AMailbox | Set-Mailbox -EmailAddresses @{add=$NewSmtp} -WhatIf
       $AMailbox | Set-Mailbox -PrimarySmtpAddress $AUser.NewPrimarySmtp -ErrorAction SilentlyContinue -WhatIf
      }
      else
      {
       # Production run, buckle your seat belt!
       Write-Host (' ' + $AUser.User)
       if($AUser.EmailAddressPolicyEnabled)
       {
        try
        { $AMailbox | Set-Mailbox -EmailAddressPolicyEnabled $false -ErrorAction Stop }
        catch
        { Write-Error ('  Error:' + $_.Exception.Message) }
       }
       try
       {
        $AMailbox | Set-Mailbox -EmailAddresses @{add=$NewSmtp} -ErrorAction Stop
        $AMailbox | Set-Mailbox -PrimarySmtpAddress $AUser.NewPrimarySmtp -ErrorAction Stop
        $SuccessCount++
       }
       catch
       { Write-Error ('  Error:' + $_.Exception.Message) }
      }
     }
     else
     {
      # There is the mailbox for this user, and we know the given SMTP address is already there, somehow
      $AMailbox = Get-Mailbox $AUser.Login
      if($AMailbox.PrimarySmtpAddress.Address -ne $AUser.NewPrimarySmtp)
      {
       # Change the Primary SMTP address, if it does not match the desired one
       if($Test)
       {
        # Test run, expect errors
        Write-Host -NoNewline ('(testing) ' + $AUser.User + ' -> SMTP:' + $AUser.NewPrimarySmtp)
        if($AMailbox.EmailAddressPolicyEnabled)
        {
         Write-Host "`tEmail Address Policy is On, turning it Off..."
         try
         { $AMailbox | Set-Mailbox -EmailAddressPolicyEnabled $false -ErrorAction Stop -Verbose -WhatIf }
         catch
         { Write-Error ('  Error:' + $_.Exception.Message) }
        }
        else
        { Write-Host "`Email Address Policy is Off" }
        $AMailbox | Set-Mailbox -PrimarySmtpAddress $AUser.NewPrimarySmtp -ErrorAction SilentlyContinue -WhatIf
       }
       else
       {
        # Production run, buckle your seat belt!
        Write-Host (' ' + $AUser.User)
        try
        {
         $AMailbox | Set-Mailbox -EmailAddressPolicyEnabled $false -ErrorAction Stop
         $AMailbox | Set-Mailbox -PrimarySmtpAddress $AUser.NewPrimarySmtp -ErrorAction Stop
         $SuccessCount++
        }
        catch
        { Write-Error ('  Error:' + $_.Exception.Message) }
       }
      }
      else
      { if(-not $Test) {Write-Host ('(already) ' + $AUser.User + ' = ' + $AMailbox.PrimarySmtpAddress) }}
     }
    }
    else
    { Write-Error ('Error: Wrong data for user "' + $AUser.User + '" in the CSV file.') }
   }
   Write-Host ('Updated ' + $SuccessCount + ' of ' + $UsersCount + ' users')
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
