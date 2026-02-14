$adminUPN="aguillemard@guillemardlaw.com"
$orgName="guillemardlaw"
$userCredential = Get-Credential -UserName $adminUPN -Message "Andy8351!"
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential

' Allow specials characters as file names
Set-SPOTenant -SpecialCharactersStateInFileFolderNames allowed
