$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
# Import Complicance Search libraries
Import-PSSession $Session -DisableNameChecking
$compSearchName = "lonestar-pr@astonk-limousine.com"
ew-MailboxExportRequest -Mailbox <user> -FilePath \\<server FQDN>\<shared folder name>\<PST name>.pst