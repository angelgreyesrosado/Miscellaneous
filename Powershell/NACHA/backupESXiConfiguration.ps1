#powershell.exe -noe -c ". \"C:\Program Files (x86)\VMware\Infrastructure\PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1\" $true"
. "C:\Program Files (x86)\VMware\Infrastructure\PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1" $true
Connect-VIServer -Server 10.9.1.14
Get-VMHostFirmware -VMHost 10.9.1.14 -BackupConfiguration -DestinationPath e:\wpdata\backups\VMware
exit(0)