#https://www.interfacett.com/blogs/remotely-installing-simple-network-load-balance-server-2012-core-using-powershell/
# Description:
# Install roll Network Load Balancing
$session=New-PsSession -ComputerName ngdc01
Invoke-Command –session $session {Import-Module ServerManager}
Invoke-Command -session $session {Add-WindowsFeature NLB}