# PowerShell script to inventory Active Directory Sites and Links
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0
#
# Configuration

param
(
 [String]$ForestName    = 'domain.local',
 [String]$Debug         = 'No',
 [String]$SmtpServer    = 'smtp.domain.local',
 [String]$SmtpFrom      = 'Active Directory Reports <ad.reports@domain.local>',
 [String]$SmtpTo        = 'ad.reports@domain.local',
 [String]$SmtpSubject   = 'AD Sites Report',
 [String]$SmtpBody      = 'Attached is the report of AD Sites and Links.'
)

#-------------------
try
{ import-module ActiveDirectory }
catch
{
 'ERROR: The PowerShell module for Active Directory was not found'
 exit 1
}

if ($Debug -eq 'Yes') { ' Reading Sites from the Active Directory forest ...' }
$AdForest       = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest", $ForestName)
[array]$AdSites = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($AdForest).sites
if ($Debug -eq 'Yes') { ' Done.' }

# Run-time Parameters
$CurrDate      = Get-Date
$ReportFile    = [environment]::getfolderpath('mydocuments') + '\AdSites-' + `
                 $CurrDate.Year + '-' + $CurrDate.Month + '-' + $CurrDate.Day + '-' + $CurrDate.Hour + '-' + $CurrDate.Minute + '-' + $CurrDate.Second + '.csv'
if ($AdSites)
{
 if ($Debug -eq 'Yes') { (' Exporting information into file: ' + $ReportFile + ' ...') }
 $AdSites |
 Sort-Object Name |
 Select-Object Name, 
               @{name='Subnets';                       expression={($_.Subnets           | Sort-Object) -join "`n"}},
               @{name='Servers';                       expression={($_.Servers           | Sort-Object) -join "`n"}},
               @{name='Adjacent Sites';                expression={($_.AdjacentSites     | Sort-Object) -join "`n"}},
               @{name='Site Links';                    expression={($_.SiteLinks         | Sort-Object) -join "`n"}},
               @{name='Inter-Site Topology Generator'; expression={$_.InterSiteTopologyGenerator}},
               @{name='Bridgehead Servers';            expression={($_.BridgeheadServers | Sort-Object) -join "`n"}} |
 Export-Csv $ReportFile -Encoding UTF8 -NoTypeInformation
 if ($Debug -eq 'Yes') { ' Done.' }

 if ($SmtpTo -ne '')
 {
  if ($Debug -eq 'Yes') { (' Sending over email to: ' + $SmtpTo + ' ...') }
  Send-Mailmessage -SmtpServer $SmtpServer `
                   -From $SmtpFrom `
                   -To $SmtpTo `
                   -Subject $SmtpSubject `
                   -Body $SmtpBody `
                   -Priority High `
                   -Attachments $ReportFile

 if ($Debug -eq 'Yes') { (' Deleting file: ' + $ReportFile + ' ...') }
 Remove-Item $ReportFile -force
 if ($Debug -eq 'Yes') { ' Done.' }
 }
}
Else
{ if ($Debug -eq 'Yes') { 'Nothing to report.' }}
