###############################################################################
##script:           syncFilesV5.ps1
##Adapted by Angel G. Reyes Rosado, P.E.
##4/21/2016
##
##Description:      Syncs/copies contents of one dir to another. Uses MD5
#+                  checksums to verify the version of the files and if they
#+                  need to be synced.
##Created by:       Noam Wajnman
##Creation Date:    June 9, 2014
## https://scriptingblog.com/2014/06/10/synchronize-folderdirectory-contents/
###############################################################################

param (
    [string]$sourceDir = "",
	[string]$destDir = ""
)
#FUNCTIONS
function Get-FileMD5 {
    Param([string]$file)
    $md5 = [System.Security.Cryptography.HashAlgorithm]::Create("MD5")
    $IO = New-Object System.IO.FileStream($file, [System.IO.FileMode]::Open)
    $StringBuilder = New-Object System.Text.StringBuilder
    $md5.ComputeHash($IO) | % { [void] $StringBuilder.Append($_.ToString("x2")) }
    $hash = $StringBuilder.ToString() 
    $IO.Dispose()
    return $hash
}
#VARIABLES
$DebugPreference = "continue"
#parameters
#$sourceDir = 'h:\clients\'
#$destDir = 'z:\clients\'
$filesCopied = 0
$currDate=Get-Date
#SCRIPT MAIN
clear
Write-Debug "Starting SYNC at $currDate..."
$SourceFiles = GCI -Recurse $sourceDir | ? { $_.PSIsContainer -eq $false} #get the files in the source dir.
$SourceFiles | % { # loop through the source dir files
    $cpySrc = $false
    $cpyDest = $false
    $src = $_.FullName #current source dir file
    Write-Debug $src
    $dest = $src -replace $sourceDir.Replace('\','\\'),$destDir #current destination dir file
    if (test-path $dest) { #if file exists in destination folder check MD5 hash
        $srcMD5 = Get-FileMD5 -file $src
        #Write-Debug "Source file hash: $srcMD5"
        $destMD5 = Get-FileMD5 -file $dest
        #Write-Debug "Destination file hash: $destMD5"
        if ($srcMD5 -eq $destMD5) { #if the MD5 hashes match then the files are the same
            Write-Debug "File hashes match. File already exists in destination folder and will be skipped."
            $cpy = $false
        }
        else { #if the MD5 hashes are different then copy the file and overwrite the older version in the destination dir
			$srcLastModifiedDate = Get-Item $src | select LastWriteTime
			$destLastModifiedDate = Get-Item $dest | select LastWriteTime
            $cpy = $true
            if ((Get-Item $src).LastWriteTime -ge (Get-Item $dest).LastWriteTime) {
                Write-Debug "Source file is newer. File will be copied to destination folder.  Source File last Modified on $srcLastModifiedDate and Destination on $destLastModifiedDate"
                $cpySrc = $true
                $cpyDest = $false
            }
            else {
                Write-Debug "Destination file is newer. File will be copied to source folder.  Source File last Modified on $srcLastModifiedDate and Destination on $destLastModifiedDate"
                $cpySrc = $false
                $cpyDest = $true
            }
        }
    }
    else { #if the file doesn't in the destination dir it will be copied.
        Write-Debug "File doesn't exist in destination folder and will be copied."
        $cpy = $true
    }
    Write-Debug "Copy is $cpy"
    if (($cpy -eq $true) -and ($cpySrc = $false) -and ($cpyDest = $false)) { #copy the file if file version is newer or if it doesn't exist in the destination dir.
        Write-Debug "Copying $src to $dest"
        # Create dest directory if it doesn't exists...
        if (!(test-path $dest)) {
            New-Item -ItemType "File" -Path $dest -Force   
        }
    #    Copy-Item -Path $src -Destination $dest -Force
        $filesCopied = $filesCopied+1
    }
    if (($cpy -eq $true) -and ($cpySrc = $true) -and ($cpyDest = $false)) { #copy the file if file version is newer or if it doesn't exist in the destination dir.
        Write-Debug "Copying $src to existing $dest"
        # Create dest directory if it doesn't exists...
        if (!(test-path $dest)) {
            New-Item -ItemType "File" -Path $dest -Force   
        }
    #    Copy-Item -Path $src -Destination $dest -Force
        $filesCopied = $filesCopied+1
    }
}
Write-Debug "$filesCopied files copied..."
Write-Debug "Ending SYNC at $currDate..." 