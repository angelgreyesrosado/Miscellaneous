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
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Backups\' -destDir 'z:\Backups\' 2> c:\logs\Sharepoint_copy_Backups.txt
rem
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\Archivo\' -destDir 'z:\Archivo\' 2> c:\logs\Sharepoint_copy_archivo.txt
rem
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\ATTY\aguillemard\' -destDir 'z:\ATTY\aguillemard\' 2> c:\logs\Sharepoint_copy_aguillemard.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\ATTY\vgonzalezvega\' -destDir 'z:\ATTY\vgonzalezvega\' 2> c:\logs\Sharepoint_copy_vgonzalezvega.txt
rem
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\ATTY\mguillemard\' -destDir 'z:\ATTY\mguillemard\' 2> c:\logs\Sharepoint_copy_mguillemard.txt
rem
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\ATTY\gcolon\' -destDir 'z:\ATTY\gcolon\' 2> c:\logs\Sharepoint_copy_gcolon.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\ATTY\jpeters\' -destDir 'z:\ATTY\jpeters\' 2> c:\logs\Sharepoint_copy_jpeters.txt
REM
start cmd /c powershell c:\src\powershell\syncFilesV3.ps1 -sourceDir 'e:\wpdata\ATTY\gcastiel\' -destDir 'z:\ATTY\gcastiel\' 2> c:\logs\Sharepoint_copy_gcastiel.txt
rem
EXIT