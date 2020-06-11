# PowerShell script to set Company name and Country codes according to OU map, read from
# a CSV file (OU,Company,Country). Country is 2-letter code:
#  CA
#  CL
#  CO
#  PL
#  UK
#  US
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0
#
# Configuration

#--------------------------------------
param
(
 [String]$CsvFile = 'Config\SetCompanyAndCountryPerOu.csv',
 [String]$Domain  = 'DC=domain,DC=local'
)

#--------------------------------------
# Find AD Users with a Mailbox, under specific OU path,
# then set the Company name and Country codes to them
function AdUsersSetCompanyAndCountry ($OUs, $CompanyName, $CountryCode)
{
 $OuPath = ($OUs.Trim('/') -split '/')
 [array]::Reverse($OuPath)

 switch ($CountryCode)
 {
  'CA'    { $UserC = 'CA'; $UserCo = 'Canada';         $UserCountryCode = 124 }
  'CL'    { $UserC = 'CL'; $UserCo = 'Chile';          $UserCountryCode = 152 }
  'CO'    { $UserC = 'CO'; $UserCo = 'Colombia';       $UserCountryCode = 170 }
  'PL'    { $UserC = 'PL'; $UserCo = 'Poland';         $UserCountryCode = 616 }
  'UK'    { $UserC = 'GB'; $UserCo = 'United Kingdom'; $UserCountryCode = 826 }
  'US'    { $UserC = 'US'; $UserCo = 'United States';  $UserCountryCode = 840 }
  default { $UserC = $null }
 }

# DISABLED USERS:                     '(!(userAccountControl:1.2.840.113556.1.4.803:=2))' + `
 $OuUsers = Get-ADUser -LDAPFilter ('(&' + `
                                      '(sAMAccountName=*)' + `
                                      '(homeMDB=*)' + `
                                    ')') `
                       -SearchBase ('OU=' + ($OuPath -join ',OU=') + ',' + $Domain) `
                       -SearchScope Subtree `
                       -Properties Company, c, co, countryCode `
                       -ResultSetSize $null
 if ($OuUsers)
 {
  $OuUsers | ForEach-Object `
  {
   try
   {
    if ($null -ne $UserC)
    { Set-ADUser $_.DistinguishedName -Replace @{Company=$CompanyName;c=$UserC;co=$UserCo;countryCode=$UserCountryCode} }
    else
    { Set-ADUser $_.DistinguishedName -Replace @{Company=$CompanyName} }
   }
   catch
   { Write-Error ('  ' + $_.DistinguishedName + "`t" + 'ERROR: ' + $Error[0].Exception.GetType().FullName) }
  }
 }
}

#--------------------------------------
try
{ Import-module ActiveDirectory }
catch
{
 'ERROR: The PowerShell snap-in for Active Directory was not found'
 exit 1
}

#--------------------------------------
if ($CsvFile -ne '')
{
 if (Test-Path $CsvFile)
 {
  $Ous = Import-Csv -path $CsvFile -Delimiter ',' -Encoding UTF8

  if ($null -ne $Ous)
  {
   $OusCount     = $Ous | Measure-Object | Select-Object -ExpandProperty Count
   $CurrentCount = 0

   foreach ($Ou in $Ous)
   {
    $CurrentCount++
    Write-Progress -Activity 'Processing OU' -status $Ou.OU -percentComplete ($CurrentCount / $OusCount * 100)

    AdUsersSetCompanyAndCountry -OUs $Ou.OU -CompanyName $Ou.Company -CountryCode $Ou.Country
   }
  }
 }
}
