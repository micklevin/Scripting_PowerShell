# PowerShell script to create new Email Policy, using Custom Recipient Filter
#
# This example creates the policy, which is sensitive to the Country and to
# the Company attributes. For continuinity reasons it preserves the default
# company email address as well.
#
# Also, just for the testing purposes, it addes the condition of an account being
# a member of the AD group
#
# Written by:   Mick Levin
# Published at: https://github.com/micklevin/Scripting_PowerShell
#
# Version:      1.0
#
# Configuration

param
(
 [String]$CountryCode   = 'CA',
 [String]$CompanyName   = 'Company Inc.'
)


Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue

Import-module ActiveDirectory

New-EmailAddressPolicy `
 -Name 'country.company.com (Country and Company)' `
 -EnabledEmailAddressTemplates "SMTP:%g.%s@country.company.com",
                               "smtp:%g.%s@company.com" `
 -RecipientFilter {(RecipientType -eq 'UserMailbox') -and
                   (C -eq $CountryCode) -and
                   (Company -eq $CompanyName) -and
                   (MemberOfGroup -eq 'company.local/Company/Data Centers/Location/Security Groups/Infrastructure/Exchange Address Policy - country.company.com')} `
 -RecipientContainer 'OU=Company,DC=company,DC=local' `
 -Priority 1
