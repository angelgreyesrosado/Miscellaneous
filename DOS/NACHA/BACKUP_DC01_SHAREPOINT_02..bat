REM set PROGRAM_DIR="C:\Program Files\Windows Resource Kits\Tools"
REM SET DEST_FOLDER=Q:\NG_BACKUPS\H
REM SET SOURCE_FOLDER=H:
REM C:
REM CD %PROGRAM_DIR%
REM forfiles -pc:\backups\ipitomy -s -m*.* -c"cmd /c del @FILE" -d-31
REM ROBOCOPY.EXE c:\backups\ h:\backups\ /E /COPYALL /XO /R:2 /W:2 
REM ROBOCOPY.EXE %SOURCE_FOLDER%\ %DEST_FOLDER%\ /E /COPYALL /XO /R:2 /W:2 
REM SET DEST_FOLDER=z:\backups
REM SET SOURCE_FOLDER=e:\wpdata\backups
REM ROBOCOPY.EXE %SOURCE_FOLDER%\ %DEST_FOLDER%\ /E /XO /R:2 /W:2
REM ROBOCOPY.EXE %SOURCE_FOLDER%\ %DEST_FOLDER%\ /E /COPYALL /XO /R:2 /W:2 /MIR
REM System Drive
REM SET DEST_FOLDER=F:\NGDC01_BACKUPS\c-drive\users
REM SET SOURCE_FOLDER=c:\users
REM ROBOCOPY.EXE %SOURCE_FOLDER%\ %DEST_FOLDER%\ /E /COPYALL /XO /R:2 /W:2 REM Update New Windows Server
REM
rem Delete old ipitomy backups
rem forfiles -p"h:\backups\ipitomy" -c"cmd /c del @FILE" -d-15
rem forfiles -p"z:\backups\ipitomy" -c"cmd /c del @FILE" -d-15
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Clients\' -destDir 'z:\Clients\' 2> c:\logs\Sharepoint_copy_Crepoint_copy_clients.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Administration\alopez\' -destDir 'z:\Administration\alopez\' 2> c:\logs\Sharepoint_copy_alopez.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Administration\nrodriguez\' -destDir 'z:\Administration\nrodriguez\' 2> c:\logs\Sharepoint_copy_nrodriguez.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Administration\zvirella\' -destDir 'z:\Administration\zvirella\' 2> c:\logs\Sharepoint_copy_zvirella.txt
reM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Administration\mcouto\' -destDir 'z:\Administration\mcouto\' 2> c:\logs\Sharepoint_copy_mcouto.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Administration\rcatinchi\' -destDir 'z:\Administration\rcatinchi\' 2> c:\logs\Sharepoint_copy_rcatinchi.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Administration\mlopez\' -destDir 'z:\Administration\mlopez\' 2> c:\logs\Sharepoint_copy_mlopez.txt
REM
EXIT