<#
.SYNOPSIS
 Report all still-active SSL certificates on enterprise CA servers

.DESCRIPTION
 This script reads active SSL certificates from all known enterprise CA
 servers, and emails the list of only interesting certificates, e.g:
 - Web  server certificates
 - Certificates for vSphere vCenters
 - Certificates for WiFi authentication servers
 The filtering happens based on the Certificate Template's code.
 The list is being sorted by Expiration date, and in the email body it
 mentions how soon will be the closest expiration.
 This script requires the PSPKI module, which could be installed via:
 > Install-Module -Name PSPKI
 NOTE: You need to get the Templates' codes from your Enterprise CA, for example:
  RAS and IAS Server SHA256 = 1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.9412240.9653608
  vSphere 6.x               = 1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.7225740.142572
  Web Server SHA256         = 1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.15373386.14075319
  Outlook Sign              = 1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.424909.5944443
  Outlook Encrypt           = 1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.14190114.8359825
 When adding new templates, update the filter condition in line #127

.PARAMETER LookFor
Specifies the condition of certificates to read - Active (default) or Expired

.PARAMETER Filter
Specifies if the certificates shall be filtered by Template code.

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
Name:    Report CA Certs.ps1
Author:  Mick Levin
Site:    https://github.com/micklevin/Scripting_PowerShell
Version: 1.0

.LINK
https://github.com/micklevin/Scripting_PowerShell
#>

#--------------------------------------
param
(
 [String]$LookFor       = 'Active',
 [string]$Filter        = 'Yes',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'SSL Reports <ssl.reports@domain.local>',
 [String]$SmtpTo        = 'ssl.reports@domain.local',
 [String]$SmtpSubject   = 'SSL Certificates',
 [String]$SmtpBody      = 'Attached is a list of all Enterprise SSL Certificates.'
)

#--------------------------------------
switch ($LookFor)
{
 'Expired' { $LookFilter = "NotAfter -lt $(Get-Date)" }
 'Active'  { $LookFilter = "NotAfter -ge $(Get-Date)" }
 default   { $LookFilter = "NotAfter -ge $(Get-Date)"; $LookFor = 'Active' }
}

switch ($Filter)
{
 'No'    { $FilterMe = $false }
 default { $FilterMe = $true }
}

#--------------------------------------
# Run-time Parameters
$CurrDate   = Get-Date
$ReportFile = [environment]::getfolderpath('mydocuments') + '\AllCaCertificates-' + $LookFor + '-' + $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '.csv'
$Output     = @()
$Found      = $false

#--------------------------------------
try
{ Import-Module PSPKI }
catch
{
 'ERROR: The PowerShell snap-in PSPKI was not found'
 exit 1
}

#--------------------------------------
#Get all CAs
Write-Progress -Id 1 -Activity 'Locating CA Servers' -Status 'Loading...' -PercentComplete 1
$CaServers    = Get-CA -ErrorAction Stop | Sort-Object DisplayName
Write-Progress -Id 1 -Activity 'Locating CA Servers' -Status 'Done' -PercentComplete 100 -Completed

$CaCount   = $CaServers | Measure-Object | Select-Object -ExpandProperty Count
$CaCurrent = 0

foreach ($CaServer in $CaServers)
{
 $CaCurrent++
 $CaServerName = ($CaServer | Select-Object Computername).Computername
 Write-Progress -Id 1 -Activity 'Inventorying CA Servers' -Status "Reading $CaServerName..." -PercentComplete ($CaCurrent / $CaCount * 100)

 # Get all certificates from the CA server
 try
 { $AllIssuedCerts = $CaServer | Get-IssuedRequest -Filter $LookFilter -ErrorAction Stop }
 catch
 { $AllIssuedCerts = $null }

 if ($null -ne $AllIssuedCerts)
 {
  if ($FilterMe)
  {
   $IssuedCerts = $AllIssuedCerts |
                  Select-Object RequestID, CommonName, NotBefore, NotAfter, CertificateTemplate, CertificateTemplateOid, SerialNumber |
                  Where-Object {($_.CertificateTemplate -eq '1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.9412240.9653608') -or `
                                ($_.CertificateTemplate -eq '1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.7225740.142572') -or `
                                ($_.CertificateTemplate -eq '1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.15373386.14075319') -or `
                                ($_.CertificateTemplate -eq '1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.424909.5944443') -or `
                                ($_.CertificateTemplate -eq '1.3.6.1.4.1.311.21.8.6734896.8216493.16205754.15498458.10085908.140.14190114.8359825')} |
                  Sort-Object Notafter
  }
  else
  {
   $IssuedCerts = $AllIssuedCerts |
                  Select-Object RequestID, CommonName, NotBefore, NotAfter, CertificateTemplate, CertificateTemplateOid, SerialNumber |
                  Sort-Object Notafter
  }

  if ($null -ne $IssuedCerts)
  {
   $EarliestExpiry = ([math]::abs(($CurrDate - ($IssuedCerts[0].Notafter)).Days)).ToString()

   $CertCount   = $IssuedCerts | Measure-Object | Select-Object -ExpandProperty Count
   $CertCurrent = 0

   foreach ($IssuedCert in $IssuedCerts)
   {
    $CertCurrent++
    Write-Progress -ParentId 1 -Id 2 -Activity 'Certificate' -Status ("($CertCurrent / $CertCount) " + $IssuedCert.CommonName) -PercentComplete ($CertCurrent / $CertCount * 100)

    if ($null -eq $IssuedCert.CertificateTemplateOid.FriendlyName)
    { $CertificateTemplate = $IssuedCert.CertificateTemplateOid.Value }
    else
    { $CertificateTemplate = $IssuedCert.CertificateTemplateOid.FriendlyName }

    $Obj = New-Object System.Object
    $Obj | Add-Member -MemberType NoteProperty -Name 'CA'            -Value $CaServerName
    $Obj | Add-Member -MemberType NoteProperty -Name 'Number'        -Value $IssuedCert.RequestID
    $Obj | Add-Member -MemberType NoteProperty -Name 'CN'            -Value $IssuedCert.CommonName
    $Obj | Add-Member -MemberType NoteProperty -Name 'Issued'        -Value $IssuedCert.NotBefore
    $Obj | Add-Member -MemberType NoteProperty -Name 'Expires'       -Value $IssuedCert.NotAfter
    $Obj | Add-Member -MemberType NoteProperty -Name 'Template'      -Value $CertificateTemplate
    $Obj | Add-Member -MemberType NoteProperty -Name 'Serial Number' -Value $IssuedCert.SerialNumber
    $Obj | Add-Member -MemberType NoteProperty -Name 'Template Code' -Value $IssuedCert.CertificateTemplate
    $Output += $Obj
   }

   $SmtpBody += "`n`nFor the server $CaServerName - next expiration is in $EarliestExpiry days."
   $Found     = $true
  }
 }
 Write-Progress -ParentId 1 -Id 2 -Activity 'Certificate' -Status 'Done' -PercentComplete 100 -Completed
}
Write-Progress -Id 1 -Activity 'Inventorying CA Servers' -Status 'Done' -PercentComplete 100 -Completed

if ($Found)
{
 Write-Progress -Id 1 -Activity 'Exporting information into a file' -Status $ReportFile -PercentComplete 1
 $Output | Export-Csv $ReportFile -Encoding UTF8 -NoTypeInformation
 Write-Progress -Id 1 -Activity 'Exporting information into a file' -Status 'Done' -PercentComplete 100 -Completed

 Write-Progress -Id 1 -Activity "Emailing information" -Status $SmtpTo -PercentComplete 1
 Send-Mailmessage -SmtpServer $SmtpServer `
                  -From $SmtpFrom `
                  -To $SmtpTo `
                  -Subject $SmtpSubject `
                  -Body $SmtpBody `
                  -Priority High `
                  -Attachments $ReportFile
 Write-Progress -Id 1 -Activity "Emailing information" -Status 'Done' -PercentComplete 100 -Completed

 Remove-Item $ReportFile -force
}
else
{
 Send-Mailmessage -SmtpServer $SmtpServer `
                  -From $SmtpFrom `
                  -To $SmtpTo `
                  -Subject ($SmtpSubject + ' (ERROR)') `
                  -Body 'Nothing to report - likely experienced a run-time error :(' `
                  -Priority High
}
