<#
.SYNOPSIS
PowerShell script to report various user and computer accounts based on predefined filters

.DESCRIPTION
This script generates and emails a report of user or computer accounts which fall under particular category, such as:
- Disabled Users
- Inactive Users
- Inactive Computers
- Users with No Password Expire option
- All Users
- All Computers
- Users of Skype for Business

.PARAMETER Debug
Enables printing debug information

.PARAMETER DC
Specifies the domain controller to use for AD queries

.PARAMETER Mode
The mode of operation - report only or report and disable.

.PARAMETER Report
Specifies what exact report to generate

.PARAMETER DaysInactive
Number of days of inactivity to trigger user's removal.

.PARAMETER SearchOU
The LDAP address of the search base

.PARAMETER ExceptGroup
Name of the AD group, which members will be excluded from processing

.PARAMETER ExceptService
Enables excluding "service" accounts from processing, based on the sAMAccountName prefix (hardcoded)

.PARAMETER ExceptOU
Enables excluding objects, located at specific OUs (hardcoded)

.PARAMETER OnlySite
Specifies the higher-level division of a company, which must be processed. The mapping of divisions to OU(s) is hardcoded

.PARAMETER SmtpServer
The DNS name of SMTP server

.PARAMETER SmtpFrom
The SMTP name and address of Sender.

.PARAMETER SmtpTo
The SMTP address of recipient(s).

.PARAMETER SmtpCc
The SMTP address of CC recipient(s).

.EXAMPLE
powershell.exe -ExecutionPolicy ByPass -Command "C:\bin\Scripts\AD\ReportAccounts.ps1" -Mode "Report" -Report "InactiveUsers"

.NOTES
Name:    ReportAccounts.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.3

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

#--------------------------------------
param
(
 [String]$Debug         = 'No',
 [String]$DC            = '',
 [String]$Mode          = 'Report',
 [String]$Report        = '',
 [String]$DaysInactive  = 90,
 [String]$SearchOU      = 'DC=domain,DC=local',
 [String]$ExceptGroup   = 'Never Inactive Users',
 [String]$ExceptService = 'No',
 [String]$ExceptOU      = 'Yes',
 [String]$OnlySite      = '',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Active Directory Reports <ad.reports@domain.local>',
 [String]$SmtpTo        = 'ad.reports@domain.local',
 [String]$SmtpCc        = ''
)

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
# Fix a glitch in date/time reporting
Function Format-Time()
{
 param
 (
  [String]$DateTime
 )

 if ($DateTime.StartsWith('1600-'))
 { $Result = '' }
 else
 { $Result = $DateTime }

 return $Result
}

#-------------------
# Check for more conditions of account's inactivity
Function Limit-AccountActive()
{
 param
 (
  $ADUser,
  $RecentDays
 )

 $Result = $true

 if ($ADUser.LastLogonTimeStamp)
 { $Result = $true }
 else
 {
  if ($ADUser.Created -and $ADUser.Created -ge $RecentDays)
  { $Result = $false }
  elseif ($ADUser.PasswordLastSet -and $ADUser.PasswordLastSet -ge $RecentDays)
  { $Result = $false }
 }

 return $Result
}

#-------------------
# Currently filters out any account name, starting with 's-', 'svc-', and 'srv-'
Function Limit-AccountName()
{
 param
 (
  [String]$Account
 )

 if ($ExceptService -ne 'Yes')
 { $Result = $true }
 else
 {
  if ($Account.StartsWith('s-',   'CurrentCultureIgnoreCase') -or
      $Account.StartsWith('svc-', 'CurrentCultureIgnoreCase') -or
      $Account.StartsWith('srv-', 'CurrentCultureIgnoreCase'))
  { $Result = $false }
  else
  { $Result = $true }
 }

 return $Result
}

#-------------------
# Filter out some OUs
Function Limit-OU ()
{
 param
 (
  [String]$OU
 )

 if ($ExceptOU -ne 'Yes')
 { $Result = $true }
 else
 {
  if ($OU.Contains('OU=Admin Accounts,') -or
      $OU.Contains('OU=Admins,') -or
      $OU.Contains('OU=Service Accounts,') -or
      $OU.Contains('OU=Vendors,') -or
      $OU.Contains('OU=Mailboxes,') -or
      $OU.EndsWith(        ',CN=Builtin,DC=domain,DC=local', 'CurrentCultureIgnoreCase') -or
      $OU.EndsWith(',OU=System Accounts,DC=domain,DC=local', 'CurrentCultureIgnoreCase'))
  { $Result = $false }
  else
  { $Result = $true }
 }

 return $Result
}

#-------------------
# Filter out Site's OUs
Function Limit-Site ()
{
 param
 (
  [String]$OU
 )

 $Result = $true

 if ($SiteOUs.Count -ne 0)
 {
  $Result = $false
  foreach ($SiteOU in $SiteOUs)
  {
   if ($SiteOU -eq '')
   {
    if (($OU -like     '*CN=Users,DC=domain,DC=local') -or
        ($OU -like '*CN=Computers,DC=domain,DC=local'))
    {
     $Result = $true
     break
    }
   }
   else
   {
    if ($OU -like ('*OU=' + $SiteOU + ',*'))
    {
     $Result = $true
     break
    }
   }
  }
 }

 return $Result
}

#-------------------
Import-Module ActiveDirectory

# Run-time Parameters
$ReportFile   = [environment]::getfolderpath('mydocuments') + '\Reports-AD-' + $Report + ' (' + $OnlySite + ').csv'
$NotMember    = (Get-ADGroup $ExceptGroup).DistinguishedName

switch ($Report)
{
 'AllUsers'          { }
 'LyncUsers'         { }
 'InactiveUsers'     { }
 'NoPasswordExpire'  { }
 'DisabledUsers'     { }
 'AllComputers'      { }
 'InactiveComputers' { }
 default {$Report = 'InactiveUsers'}
}

Switch ($OnlySite)
{
 'Division 1'   { $SiteLang = 'EN'; $SiteOUs = @('One,OU=Divisions') }
 'Division 2'   { $SiteLang = 'ES'; $SiteOUs = @('Two,OU=Divisions', 'Division 2', 'Division Two') }
 default        { $SiteLang = 'EN'; $SiteOUs = @() ; $OnlySite = 'All' }
}

# AD Search Parameters
if ($DC -eq '')
{
 $LocalDC      = $env:LOGONSERVER
 if (!$LocalDC)
 { $LocalDC = $env:COMPUTERNAME }
 elseif ($LocalDC.StartsWith("\\"))
 { $LocalDC = $LocalDC.Substring(2) }
 else
 { $LocalDC = $env:COMPUTERNAME }
}
else
{ $LocalDC = $DC }

if ($Debug -ne 'No') { ("DEBUG:`n SearchBase: $SearchOU`n Server:     $LocalDC") }

#-------------------

Switch ($Report)
{
 'AllUsers'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for all users...") }
  $SearchUsers = Get-ADUser -Filter * `
                            -SearchBase $SearchOU `
                            -Server $LocalDC `
                            -Properties Department, `
                                        Title, `
                                        mail, `
                                        MemberOf, `
                                        LastLogonTimeStamp, `
                                        PasswordLastSet, `
                                        PasswordNeverExpires, `
                                        pwdLastSet, `
                                        AccountExpirationDate, `
                                        Created, `
                                        Modified
 }
 'LyncUsers'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for Lync users...") }
  $FoundUsers = Get-ADUser -Filter {msRTCSIP-PrimaryUserAddress -ne $false} `
                            -SearchBase $SearchOU `
                            -Server $LocalDC `
                            -Properties msRTCSIP-PrimaryUserAddress, `
                                        userPrincipalName, `
                                        Department, `
                                        mail, `
                                        MemberOf, `
                                        LastLogonTimeStamp, `
                                        PasswordLastSet, `
                                        PasswordNeverExpires, `
                                        pwdLastSet, `
                                        AccountExpirationDate, `
                                        Created, `
                                        Modified |
                Select-Object @{name='Login Name';             expression={$_.sAMAccountName}},
                              @{name='User Principal Name';    expression={$_.userPrincipalName}},
                              @{name='SIP Account';            expression={$_.'msRTCSIP-PrimaryUserAddress'}},
                              Name,
                              Department,
                              @{name='E-mail';                 expression={$_.mail}},
                              @{name='Enabled';                expression={if($_.Enabled) {'Yes'} else {'No'}}},
                              @{name='Member Of';              expression={$_.MemberOf -join "`n"}},
                              @{name='Container';              expression={Format-OU -OUString (($_.distinguishedName -split '(,OU=)' | Select-Object -Skip 1) -join '') -replace '^,'}},
                              @{name='Last Login';             expression={Format-Time -DateTime ([DateTime]::FromFileTime($_.lastLogonTimestamp).ToString('yyyy-MM-dd hh:mm:ss'))}},
                              @{name='Last Password Set';      expression={$_.PasswordLastSet}},
                              @{name='Password Never Expires'; expression={if($_.PasswordNeverExpires) {'Yes'} else {'No'}}},
                              @{name="Must Change Password";   expression={if($_.pwdLastSet -eq 0) {'Yes'} else {'No'}}},
                              @{name='Account Expires';        expression={$_.AccountExpirationDate}},
                              @{name='Created';                expression={$_.Created}},
                              @{name='Modified';               expression={$_.Modified}}
 }
 'InactiveUsers'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for inactive users...") }
  $TimeSpan    = New-Timespan -Days $DaysInactive
  $RecentDays  = (Get-date).AddDays(0 - $DaysInactive)
  $SearchUsers = Search-ADAccount -UsersOnly `
                                  -AccountInactive `
                                  -TimeSpan $TimeSpan `
                                  -SearchBase $SearchOU `
                                  -Server $LocalDC |
                 Get-ADUser -Properties Department, `
                                        Title, `
                                        mail, `
                                        MemberOf, `
                                        LastLogonTimeStamp, `
                                        PasswordLastSet, `
                                        PasswordNeverExpires, `
                                        pwdLastSet, `
                                        AccountExpirationDate, `
                                        Created, `
                                        Modified |
                 Where-Object {($_.Enabled -eq $true) -and (-not ($_.MemberOf -eq $NotMember)) } |
                 Where-Object {Limit-AccountName -Account ($_.sAMAccountName)} |
                 Where-Object {Limit-AccountActive -ADUser ($_) -RecentDays ($RecentDays)} |
                 Where-Object {Limit-OU -OU ($_.distinguishedName)} |
                 Where-Object {Limit-Site -OU ($_.distinguishedName)}
 }
 'NoPasswordExpire'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for users with no password expiration...") }
  $SearchUsers = Search-ADAccount -UsersOnly `
                                  -PasswordNeverExpires `
                                  -SearchBase $SearchOU `
                                  -Server $LocalDC |
                 Get-ADUser -Properties Department, `
                                        Title, `
                                        mail, `
                                        MemberOf, `
                                        LastLogonTimeStamp, `
                                        PasswordLastSet, `
                                        PasswordNeverExpires, `
                                        pwdLastSet, `
                                        AccountExpirationDate, `
                                        Created, `
                                        Modified |
                 Where-Object {($_.Enabled -eq $true) -and (-not ($_.MemberOf -eq $NotMember)) } |
                 Where-Object {Limit-AccountName -Account ($_.sAMAccountName)} |
                 Where-Object {Limit-OU -OU ($_.distinguishedName)} |
                 Where-Object {Limit-Site -OU ($_.distinguishedName)}
 }
 'DisabledUsers'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for disabled users...") }
  $SearchUsers = Search-ADAccount -UsersOnly -AccountDisabled -SearchBase $SearchOU -Server $LocalDC |
                 Get-ADUser -Properties Department, `
                                        Title, `
                                        mail, `
                                        MemberOf, `
                                        LastLogonTimeStamp, `
                                        PasswordLastSet, `
                                        PasswordNeverExpires, `
                                        pwdLastSet, `
                                        AccountExpirationDate, `
                                        Created, `
                                        Modified |
                 Where-Object {Limit-OU -OU ($_.distinguishedName)} |
                 Where-Object {Limit-Site -OU ($_.distinguishedName)}
 }
 'AllComputers'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for all computers...") }
  $SearchComputers = Get-ADComputer -Filter * -Properties LastLogonTimeStamp, `
                                                          OperatingSystem, `
                                                          OperatingSystemServicePack, `
                                                          OperatingSystemVersion, `
                                                          Enabled `
                                              -SearchBase $SearchOU `
                                              -Server $LocalDC |
                     Where-Object {Limit-Site -OU ($_.distinguishedName)}
 }
 'InactiveComputers'
 {
  if ($Debug -ne 'No') { ("DEBUG: Searching for inactive computers...") }
  $GraceTime       = (Get-Date).Adddays(-($DaysInactive))
  $GraceFileTime   = $GraceTime.ToFileTimeUtc()
  $SearchComputers = Get-ADComputer -Filter * -Properties LastLogonTimeStamp, `
                                                          OperatingSystem, `
                                                          OperatingSystemServicePack, `
                                                          OperatingSystemVersion, `
                                                          Enabled `
                                              -SearchBase $SearchOU `
                                              -Server $LocalDC |
                     Where-Object {($_.Enabled -eq $true) -and (-not ($_.MemberOf -eq $NotMember)) -and ($_.LastLogonTimeStamp -lt $GraceFileTime) } |
                     Where-Object {Limit-Site -OU ($_.distinguishedName)}
 }
}

$SmtpSubject = ''

if ($SearchUsers -or $FoundUsers)
{
 if ($FoundUsers)
 {
  if ($Debug -ne 'No') { ("DEBUG: Exporting into: $ReportFile ...") }
  $FoundUsers |
  Sort-Object 'Login ID' |
  Export-Csv $ReportFile -NoTypeInformation -Encoding UTF8
 }
 else
 {
  if ($Debug -ne 'No') { ("DEBUG: Exporting into: $ReportFile ...") }
  $SearchUsers |
  Select-Object @{name='Login ID';               expression={$_.sAMAccountName}},
                Name,
                Department,
                Title,
                @{name='E-mail';                 expression={$_.mail}},
                @{name='Enabled';                expression={if($_.Enabled) {'Yes'} else {'No'}}},
                @{name='Member Of';              expression={$_.MemberOf -join "`n"}},
                @{name='Container';              expression={Format-OU -OUString (($_.distinguishedName -split '(,OU=)' | Select-Object -Skip 1) -join '') -replace '^,'}},
                @{name='Last Login';             expression={Format-Time -DateTime ([DateTime]::FromFileTime($_.lastLogonTimestamp).ToString('yyyy-MM-dd hh:mm:ss'))}},
                @{name='Last Password Set';      expression={$_.PasswordLastSet}},
                @{name='Password Never Expires'; expression={if($_.PasswordNeverExpires) {'Yes'} else {'No'}}},
                @{name="Must Change Password";   expression={if($_.pwdLastSet -eq 0) {'Yes'} else {'No'}}},
                @{name='Account Expires';        expression={$_.AccountExpirationDate}},
                @{name='Created';                expression={$_.Created}},
                @{name='Modified';               expression={$_.Modified}} |
  Sort-Object 'Login ID' |
  Export-Csv $ReportFile -NoTypeInformation -Encoding UTF8
 }

 if ($Mode -eq 'Disable')
 {
  $SearchUsers | ForEach-Object { Disable-ADAccount -Identity $_.sAMAccountName }

  Switch ($Report)
  {
   'InactiveUsers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = "Users inactive for $DaysInactive days have been disabled"
      $SmtpBody     = "Attached is a list of disabled users, inactive for $DaysInactive days or more."
      $SmtpBodyTail = ''
     }
     'ES'
     {
      $SmtpSubject  = "Users inactive for $DaysInactive days have been disabled"
      $SmtpBody     = "Attached is a list of disabled users, inactive for $DaysInactive days or more."
      $SmtpBodyTail = ''
     }
    }
   }
   'NoPasswordExpire'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = "Users with no password expiration have been disabled"
      $SmtpBody     = 'Attached is a list of disabled users with no password expiration.'
      $SmtpBodyTail = ''
     }
     'ES'
     {
      $SmtpSubject  = "Users with no password expiration have been disabled"
      $SmtpBody     = "Attached is a list of disabled users with no password expiration."
      $SmtpBodyTail = ''
     }
    }
   }
  }
 }
 else
 {
  Switch ($Report)
  {
   'AllUsers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = 'All users'
      $SmtpBody     = 'Attached is a list of all users.'
      $SmtpBodyTail = ''
     }
     'ES'
     {
      $SmtpSubject  = 'All users'
      $SmtpBody     = 'Attached is a list of all users.'
      $SmtpBodyTail = ''
     }
    }
   }
   'LyncUsers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = 'Lync users'
      $SmtpBody     = 'Attached is a list of Lync users.'
      $SmtpBodyTail = ''
     }
     'ES'
     {
      $SmtpSubject  = 'Lync users'
      $SmtpBody     = 'Adjunta esta una lista de Lync usuarios.'
      $SmtpBodyTail = ''
     }
    }
   }
   'InactiveUsers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = "Users inactive for $DaysInactive days"
      $SmtpBody     = "Attached is a list of users, inactive for $DaysInactive days or more."
      $SmtpBodyTail = "`nConsider disabling them.`n"
      if (Test-Path 'Report Accounts (Report Inactive Users EN).txt')
      {
       $SmtpBody     = (Get-Content 'Report Accounts (Report Inactive Users EN).txt' | Out-String).Replace('$DaysInactive', "$DaysInactive")
       $SmtpBodyTail = ''
      }
     }
     'ES'
     {
      $SmtpSubject  = "Usuarios inactivos por $DaysInactive dias"
      $SmtpBody     = "Adjunta esta una lista de usuarios, inactivos por $DaysInactive dias o mas." # "Adjunta está una lista de usuarios, inactivos por $DaysInactive días o más."
      $SmtpBodyTail = "`nConsidere desactivarlos.`n"
      if (Test-Path 'Report Accounts (Report Inactive Users ES).txt')
      {
       $SmtpBody     = (Get-Content 'Report Accounts (Report Inactive Users ES).txt' | Out-String).Replace('$DaysInactive', "$DaysInactive")
       $SmtpBodyTail = ''
      }
     }
    }
   }
   'NoPasswordExpire'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = 'Users with no password expiration'
      $SmtpBody     = 'Attached is a list of users with no password expiration.'
      $SmtpBodyTail = ''
     }
     'ES'
     {
      $SmtpSubject  = 'Users with no password expiration'
      $SmtpBody     = 'Attached is a list of users with no password expiration.'
      $SmtpBodyTail = ''
     }
    }
   }
   'DisabledUsers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = 'Disabled users'
      $SmtpBody     = 'Attached is a list of disabled users.'
      $SmtpBodyTail = ''
     }
     'ES'
     {
      $SmtpSubject  = 'Disabled users'
      $SmtpBody     = 'Attached is a list of disabled users.'
      $SmtpBodyTail = ''
     }
    }
   }
  }
 }
}

if ($SearchComputers)
{
 if ($Debug -ne 'No') { ("DEBUG: Exporting into: $ReportFile ...") }
 $SearchComputers |
 Select-Object Name,
               Enabled,
               OperatingSystem,
               OperatingSystemServicePack,
               OperatingSystemVersion,
               @{Name='Container';  Expression={Format-OU -OUString (($_.distinguishedName -split '(,OU=)' | Select-Object -Skip 1) -join '') -replace '^,'}},
               @{name='Last Login'; expression={Format-Time -DateTime ([DateTime]::FromFileTime($_.lastLogonTimestamp).ToString('yyyy-MM-dd hh:mm:ss'))}} |
 Sort-Object Name |
 Export-Csv $ReportFile -NoTypeInformation -Encoding UTF8

 if ($Mode -eq 'Disable')
 {
  $SearchComputers | ForEach-Object { Disable-ADAccount -Identity $_.sAMAccountName }

  $SmtpSubject  = "Computers inactive for $DaysInactive days have been disabled"
  $SmtpBody     = "Attached is a list of disabled Computers, inactive for $DaysInactive days or more."
  $SmtpBodyTail = ''
 }
 else
 {
  Switch ($Report)
  {
   'AllComputers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = "All computers"
      $SmtpBody     = "Attached is a list of all computers"
      $SmtpBodyTail = "`n"
     }
     'ES'
     {
      $SmtpSubject  = "All computers"
      $SmtpBody     = "Attached is a list of all computers"
      $SmtpBodyTail = "`n"
     }
    }
   }
   'InactiveComputers'
   {
    Switch ($SiteLang)
    {
     'EN'
     {
      $SmtpSubject  = "Computers inactive for $DaysInactive days"
      $SmtpBody     = "Attached is a list of computers, inactive for $DaysInactive days or more."
      $SmtpBodyTail = "`nConsider disabling them`n"
     }
     'ES'
     {
      $SmtpSubject  = "Computers inactive for $DaysInactive days"
      $SmtpBody     = "Attached is a list of computers, inactive for $DaysInactive days or more."
      $SmtpBodyTail = "`nConsider disabling them`n"
     }
    }
   }
  }
 }
}

if ($SmtpSubject -ne '')
{
 if ($Debug -ne 'No') { ("DEBUG: Sending email to: $SmtpTo ...") }
 Switch ($SiteLang)
 {
  'EN'
  {
   $SmtpBody = $SmtpBody + "`n`nSearch Parameters:`n`n" +
                           "Site:              $OnlySite`n" +
                           "Inactive Period:   $DaysInactive days`n" +
                           "Search Server:     $LocalDC`n" +
                           "Search Container:  $SearchOU`n" +
                           "Ignore members of: $NotMember`n" +
                           $SmtpBodyTail
  }
  "ES"
  {
   $SmtpBody = $SmtpBody + "`n`nParametros de busqueda:`n`n" +               # "`n`nParámetros de búsqueda:`n`n"
                           "Sitio:                  $OnlySite`n" +
                           "Periodo de inactividad: $DaysInactive days`n" +
                           "Servidor de busqueda:   $LocalDC`n" +            # "Servidor de búsqueda:   $LocalDC`n"
                           "Contenedor de busqueda: $SearchOU`n" +           # "Contenedor de búsqueda: $SearchOU`n"
                           "Ignore miembros de:     $NotMember`n" +
                           $SmtpBodyTail
  }
 }

 if ($SmtpCC -ne '')
 { Send-Mailmessage -SmtpServer "$SmtpServer" `
                    -From "$SmtpFrom" `
                    -To "$SmtpTo" `
                    -Cc "$SmtpCc" `
                    -Subject ($SmtpSubject + ' (' + $OnlySite + ')') `
                    -Body $SmtpBody `
                    -Priority High `
                    -Attachments $ReportFile `
                    -Verbose }
 else
 { Send-Mailmessage -SmtpServer $SmtpServer `
                    -From $SmtpFrom `
                    -To $SmtpTo `
                    -Subject ($SmtpSubject + ' (' + $OnlySite + ')') `
                    -Body $SmtpBody `
                    -Priority High `
                    -Attachments $ReportFile `
                    -Verbose }
}
Else
{ 'Nothing to report' }

if ($Debug -ne 'No') { ("DEBUG: Done") }
