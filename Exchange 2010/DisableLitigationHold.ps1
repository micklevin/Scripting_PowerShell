# PowerShell script to disable Litigation Hold on multiple mailboxes
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0
#
# Configuration

$Accounts = @(
'user1',
'user2',
'userN'
);


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

foreach($Account in $Accounts) `
{
 Write-Host -NoNewLine "Processing: $Account..."
 Set-Mailbox -Identity $Account -LitigationHoldEnabled $false -ErrorAction Stop

 Write-Host ' Done'
}
