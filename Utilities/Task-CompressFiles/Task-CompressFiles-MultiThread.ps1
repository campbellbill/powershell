### Task-CompressFiles-MultiThread.ps1 ###
##
##	Task - Compress Files - MultiThread
##	Script Name:	Task-CompressFiles-MultiThread.ps1
##	Written by:		Bill Campbell
##	Written:		June 27, 2013
##	Added Script:	27 June 2013
##	Last Modified:	2013.July.22
##
##	Version:		2013.07.22.03
##	Version Notes:	Version format is a date taking the following format: YYYY.MM.DD.RR		- where RR is the revision/save count for the day modified.
##	Version Exmpl:	If this is the 6th revision/save on January 13 2012 then RR would be "06" and the version number will be formatted as follows: 2012.01.23.06
##
##	Purpose:		This is the Backend Script for Multi-Threading purposes to Compress files.
##
Param(
	[string]$SrcDir					# Directory where the files to be compressed are located
	, [string]$DestDir				# Directory where the compressed files will be saved at
	, [string]$ArchiveType			# Specifies the type of archive. Valid values are: "7z", "zip", "xz", "bzip2, "gzip", "tar"
	, [string]$ArchiveBaseName		# Name that will be given to the Archive file, without the extension.
	, [string]$FilesToBeCompressed	# Files to include in the archive
	, [string]$WorkingDirectory
	, [string]$ScriptLogPath
	, [switch]$RecurseDirectory
)
#[string]$WorkingDirectory = ($MyInvocation.MyCommand.Definition | Split-Path)
# Recursive
# 32-bit installed version
# 7z.exe a TxtFiles.7z *.txt -r
#
# 32-bit stand alone version
# 7za.exe a TxtFiles.7z *.txt -r
#
# These commands add all *.txt files from current folder and its subfolders to archive "TxtFiles.7z"
#
# Path to the Installed 7-Zip compression executable.
#[string]$PathToExe = "C:\Program Files\7-Zip\7z.exe"
#
# Path to the stand alone 7-Zip compression executable.
[string]$PathToExe = "$($WorkingDirectory)\7za920\7za.exe"
[string]$FilesToArch = "$($SrcDir)\$($FilesToBeCompressed)"
[string]$DestArchive = "$($DestDir)\$($ArchiveBaseName).$($ArchiveType)"

If ($RecurseDirectory)
{
	[array]$Arguments = "a", "-t$($ArchiveType)", "$($DestArchive)", "$($FilesToArch)", "-r"
}
Else
{
	[array]$Arguments = "a", "-t$($ArchiveType)", "$($DestArchive)", "$($FilesToArch)"
}

Add-Content -Path $ScriptLogPath -Value "`t`tCommand being executed:`n`t`t`t $([char]34)$($PathToExe) $([string]::Join(" ", $Arguments))$([char]34)"

& $PathToExe $Arguments
