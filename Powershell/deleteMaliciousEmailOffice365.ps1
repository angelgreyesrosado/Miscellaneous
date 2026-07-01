$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
# Import Complicance Search libraries
Import-PSSession $Session -DisableNameChecking
$compSearchName = "Password Change requests"
New-ComplianceSearch -Name $compSearchName -ExchangeLocation all -ContentMatchQuery 'sent>=01/01/2019 AND Subject:"Password"'
#New-ComplianceSearch -Name $compSearchName -ExchangeLocation all -ContentMatchQuery 'sent>=01/01/2019 AND From:"7596558@kpnmail.nl"' # Can also do something like Subject:"Bad Subject"
Start-ComplianceSearch -Identity $compSearchName
Get-ComplianceSearch -Identity $compSearchName # Run this till it shows Completed
Get-ComplianceSearch -Identity $compSearchName | Select Items # Show count of matching emails
Get-ComplianceSearch -Identity $compSearchName | fl # Show list of matching mailboxes
New-ComplianceSearchAction -SearchName $compSearchName -Purge -PurgeType HardDelete -Confirm:$True # Purge from mailboxes
Get-ComplianceSearchAction -Identity "$($compSearchName)_Purge" # Make sure it all purged fine