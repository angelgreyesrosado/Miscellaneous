Install-Module MSOnline
Connect-MsolService
Set-MsolUser -UserPrincipalName rberlingeri@guillemardlaw.com -PasswordNeverExpires $false
Get-MSOLUser | Select UserPrincipalName, PasswordNeverExpires