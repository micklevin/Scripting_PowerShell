# Disable SMBv1 on Windows 2012 R2 and 2016 (or Windows 8.1 and 10)
#
# Generally it would be done via several commands:
#
# Disable SMBv1 in "Server" component:
#  Set-SmbServerConfiguration -EnableSMB1Protocol $false
#  Remove-WindowsFeature -Name FS-SMB1
#
# Disable SMBv1 in "Client" component:
#  Disable-WindowsOptionalFeature -Online -FeatureName smb1protocol
#
# NOTE: reboot is required

# Result handling
$Result     = 0
$LogMessage = ''

# Disable SMB configuration for SMBv1
try
{
 $SMBv1Configuration = Get-SmbServerConfiguration

 if ($SMBv1Configuration.EnableSMB1Protocol -eq $true)
 {
  Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force
  if ($Result -eq 0) { $Result = 1 }
 }
}
catch
{
 $Result      = -1
 $LogMessage += "`n'Get-SmbServerConfiguration' is not supported on this system."
}

# Remove SMB1 feature from OS
try
{
 $SMBv1Feature = Get-WindowsFeature -Name 'FS-SMB1'

 if ($SMBv1Feature.Installed -eq $true)
 {
  Remove-WindowsFeature -Name 'FS-SMB1'
  if ($Result -eq 0) { $Result = 1 }
 }
}
catch
{
 $Result      = -1
 $LogMessage += "`n'Get-WindowsFeature' is not supported on this system."
}

# Remove SMBv1 Optional Feature from OS
try
{
 $SMBv1OptionalFeature = (Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -eq 'SMB1Protocol'})

 If ($SMBv1OptionalFeature.State -ne 'Disabled')
 {
  Disable-WindowsOptionalFeature -Online -FeatureName 'SMB1Protocol' -NoRestart
  if ($Result -eq 0) { $Result = 1 }
 }
}
catch
{
 $Result      = -1
 $LogMessage += "`n'Get-WindowsOptionalFeature' is not supported on this system."
}

# Result handling
switch ($Result)
{
 -1      { $LogType    = 'Warning';     $LogMessage = "Script: DisableSMBv1-Windows2012R2-2016.ps1`n" + $LogMessage }
  1      { $LogType    = 'Information'; $LogMessage = "Script: DisableSMBv1-Windows2012R2-2016.ps1`nSMBv1 has been disabled" }
 default { $LogType    = 'Success';     $LogMessage = '' }
}

if ($LogType -ne 'Success')
{
 Write-EventLog -LogName 'Application' `
                -Source 'WSH' `
                -EventID 0 `
                -EntryType $LogType `
                -Message $LogMessage
}
