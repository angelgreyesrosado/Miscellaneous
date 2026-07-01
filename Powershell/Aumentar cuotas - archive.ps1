# One Time, install Module
Install-Module -Name PSWSMan

Set-ExecutionPolicy RemoteSigned
Install-Module -Name ExchangeOnlineManagement

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName aguillemard@guillemardlaw.com 

Get-Mailbox mguillemard | Select * quota
Set-Mailbox mguillemard -ProhibitSendQuota 100GB -ProhibitSendReceiveQuota 100GB -IssueWarningQuota 99GB

Enable-Mailbox mguillemard -AutoExpandingArchive