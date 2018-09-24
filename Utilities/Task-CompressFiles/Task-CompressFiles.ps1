<#
	Task - Compress Files
	Script Name:	Task-CompressFiles.ps1
	Written by:		Bill Campbell
	Written:		June 25, 2013
	Added Script:	25 June 2013
	Last Modified:	2013.July.22

	Version:		2013.07.22.03
	Version Notes:	Version format is a date taking the following format: YYYY.MM.DD.RR		- where RR is the revision/save count for the day modified.
	Version Exmpl:	If this is the 6th revision/save on January 13 2012 then RR would be "06" and the version number will be formatted as follows: 2012.01.23.06

	Purpose:		Query with Powershell for all Windows Servers from Active Directory For Licensing purposes and create a report to be sent via email.

	Notes:			Use the next line to change to your user profile "Desktop" folder:
					cd ${env:userprofile}\Desktop
					echo ${env:userprofile}
					echo ${env:computername}
					echo ${env:username}
					echo ${env:PSModulePath}
					cd C:\Pearson_Docs

	Compress Kaplan CSV files
	Source Directory:		\\stor22\datadump\kaplan\ER_Reports
	Destination Directory:	\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed
	Compression Format:		7-Zip
	
	Steps that need to happen:
	1. Get a list of all of the CSV files with-in the "\\stor22\datadump\kaplan\ER_Reports" directory.
	2. Loop over the list and verify a compressed file does NOT exist in the "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed" directory.
	3. If a compressed file does NOT exist then compress the CSV file and move it into the "CSV_Compressed" directory.

	.SYNOPSIS
		#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
		#! IMPORTANT NOTE:																							 !#
		#! 																											 !#
		#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#

	.PARAMETER SourceDirectory [string] (REQUIRED)
    	Source Directory where the files to be compressed are located.

	.PARAMETER DestinationDirectory [string] (REQUIRED)
    	Destination Directory where the compressed files will be saved to.

	.PARAMETER MaxAge [int] (OPTIONAL)
    	Maximum age of files to return (In Days). Defaults to 90 days.

	.PARAMETER MinAge [int] (OPTIONAL)
    	Minimum age of files to return (In Days). Defaults to 0 days.

	.PARAMETER CompressionFormat [string] (OPTIONAL)
		Compression Format for the new archives. Valid values are: "7z", "xz", "zip", "gzip", "bzip2, "tar". Defaults to "7z".

	.PARAMETER Extension [string] (OPTIONAL)
		Extension of the files you want to compress. Defaults to "*"

	.PARAMETER ScrDebug [switch] (OPTIONAL)
		The ScrDebug switch turns the debugging code in the script on or off ($true/$false). Defaults to FALSE.

	.EXAMPLES
		.\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed" -MaxAge 30 -CompressionFormat "zip" -Extension "csv" -ScrDebug
        .\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed\zip" -MaxAge 60 -CompressionFormat "zip" -Extension "csv" -ScrDebug
		. ${env:userprofile}\Dropbox\Scripting\PowerShell_Scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed" -MaxAge 30 -CompressionFormat "7z" -Extension "csv"
		. ${env:userprofile}\Dropbox\Scripting\PowerShell_Scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed"
		. ${env:userprofile}\Dropbox\Scripting\PowerShell_Scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed" -ScrDebug
		. ${env:userprofile}\Dropbox\Scripting\PowerShell_Scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed" -MaxAge 21 -MinAge 7 -CompressionFormat "7z"
		. C:\utilities\scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory "\\stor22\datadump\kaplan\ER_Reports" -DestinationDirectory "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed" -MaxAge 30 -CompressionFormat "zip" -Extension "csv"

	.EXAMPLES Scheduled Tasks
		Server 2003
		C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -Command "pushd C:\utilities\scripts\Task-CompressFiles; C:\utilities\scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory '\\stor22\datadump\kaplan\ER_Reports' -DestinationDirectory '\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed' -CompressionFormat '7z' -Extension 'csv' -MaxAge 3"
		Server 2008 and newer
		powershell.exe -c "pushd C:\utilities\scripts\Task-CompressFiles; C:\utilities\scripts\Task-CompressFiles\Task-CompressFiles.ps1 -SourceDirectory '\\stor22\datadump\kaplan\ER_Reports' -DestinationDirectory '\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed' -CompressionFormat '7z' -Extension 'csv' -MaxAge 3"

	-- Script Changes --
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
	#!   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE   !#
	#! RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. !#
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#

	Changes for 25.Jun.2013 - 28.Jun.2013
		- Initial Script Writing and Debugging.

	Changes for 22.July.2013
		- Changes to the output in the log file.
		- Simplified some logic for better performance.
#>

#region Script Initialization
	Param(
		[Parameter(
			Position			= 0
			, Mandatory			= $true
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Source Directory where the files to be compressed are located.'
		)][string]$SourceDirectory
		, [Parameter(
			Position			= 1
			, Mandatory			= $true
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Destination Directory where the compressed files will be saved to.'
		)][string]$DestinationDirectory
		, [Parameter(
			Position			= 2
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Maximum age of files to return (In Days). Defaults to 90 days.'
		)][int]$MaxAge = 90
		, [Parameter(
			Position			= 3
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Minimum age of files to return (In Days). Defaults to 0 days.'
		)][int]$MinAge = 0
		, [Parameter(
			Position			= 4
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Compression Format for the new archives. Valid values are: "7z", "zip", "xz", "bzip2, "gzip", "tar". Defaults to "7z".'
		)][string]$CompressionFormat = "7z"
		, [Parameter(
			Position			= 5
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Extension of the files you want to compress. Defaults to "*"'
		)][string]$Extension = "*"
		, [Parameter(
			Position			= 6
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'The ScrDebug switch turns the debugging code in the script on or off ($true/$false). Defaults to FALSE.'
		)][switch]$ScrDebug = $false
	)

	#region Assembly and PowerShell Module Initialization
		#If (!(Get-Command Write-Zip -ErrorAction SilentlyContinue))
		#{
		#	try
		#	{
		#		Import-Module Pscx
		#	}
		#	catch
		#	{
		#		If (!(Get-Command Write-Zip -ErrorAction SilentlyContinue))
		#		{
		#			Import-Module Pscx
		#		}
		#	}
		#	finally
		#	{
		#		If (!(Get-Command Write-Zip -ErrorAction SilentlyContinue))
		#		{
		#			throw "Cannot load the $([char]34)Pscx$([char]34) module!! Please make sure the $([char]34)PowerShell Community Extensions (PSCX)$([char]34) are installed in one of the following path: $([char]34)$(${env:PSModulePath})$([char]34)."
		#		}
		#	}
		#}
		#region Import Assemblies
			#[void][Reflection.Assembly]::LoadFrom("C:\Program Files\7-Zip\7z.dll")
		#endregion Import Assemblies
	#endregion Assembly and PowerShell Module Initialization

	######################################
	##  Script Variable Initialization  ##
	######################################
	# These are used primarliy by Functions but are used elsewhere in the script.
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path)
	#[string]$Script:FullScrPath	= ($MyInvocation.MyCommand.Definition)

	[string]$MultiThreadScript = "$($ScriptDirectory)\Task-CompressFiles-MultiThread.ps1"
	If (!(Test-Path -Path "$($MultiThreadScript)"))
	{
		throw "The Multi-Thread Script was not found. Please make sure this path is correct: $([char]34)$($MultiThreadScript)$([char]34)"
	}

	If ($ScrDebug)
	{
    	$CutOffDate = (Get-Date).AddDays(-90)
    	#$CutOffDate = (Get-Date).AddDays(-$MaxAge)
	}
	Else
	{
    	$CutOffDate = (Get-Date).AddDays(-$MaxAge)
	}

	If ($MinAge -ge 1)
	{
		$MinCutOffDate = (Get-Date).AddDays(-$MinAge)
	}
	Else
	{
		$MinCutOffDate = Get-Date
	}
    #$Script:OldestCreateDate = New-Object DateTime($CutOffDate.Year, $CutOffDate.Month, $CutOffDate.Day, $CutOffDate.Hour, $CutOffDate.Minute, $CutOffDate.Second)
	#$Script:NewestCreateDate = New-Object DateTime($MinCutOffDate.Year, $MinCutOffDate.Month, $MinCutOffDate.Day, $MinCutOffDate.Hour, $MinCutOffDate.Minute, $MinCutOffDate.Second)

	# Maximum number of threads to have active at one time.
	$MaxThreads = 25

	# Sleep Timer in Milliseconds
	$SleepTimer = 500

	##	Date and Time Formatting
	##	http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo%28VS.85%29.aspx
	##	Formatting chosen for the date string is as follows:
	##	Date Display  |  -Format
	##	------------  |  ------------
	##	2012.Jan.25   |  yyyy.MMM.dd
	##	2012.01.25    |  yyyy.MM.dd
	##	Mon			  |  ddd
	##	------------  |  ------------
	##	Time Display  |  -Format
	##	------------  |  ------------
	##	22:00         |  HH:mm
	#[string]$LogDirDateName = (Get-Date -Format 'yyyy.MM.dd.ddd').ToString()
	[string]$LogDirDateName = (Get-Date -Format 'yyyy.MMM').ToString()
	[string]$LogFileDate = (Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()
	[string]$PtlLogName = ($MyInvocation.MyCommand.Name).Replace(".ps1", "")
	[string]$Ext = "log"
	[string]$LogFileName = "$($LogFileDate)_$($PtlLogName).$($Ext)"

	If ($ScrDebug)
	{
		If (!(Test-Path -Path "$($ScriptDirectory)\Logs\ScrDebug\$($LogDirDateName)"))
		{
			New-Item -ItemType Directory "$($ScriptDirectory)\Logs\ScrDebug\$($LogDirDateName)" | Out-Null
		}
		$Script:FullScriptLogPath = "$($ScriptDirectory)\Logs\ScrDebug\$($LogDirDateName)\$($LogFileName)"
	}
	Else
	{
		If (!(Test-Path -Path "$($ScriptDirectory)\Logs\$($LogDirDateName)"))
		{
			New-Item -ItemType Directory "$($ScriptDirectory)\Logs\$($LogDirDateName)" | Out-Null
		}
		$Script:FullScriptLogPath = "$($ScriptDirectory)\Logs\$($LogDirDateName)\$($LogFileName)"
	}

	If (!(Test-Path -Path "$($DestinationDirectory)"))
	{
		New-Item -ItemType Directory "$($DestinationDirectory)" | Out-Null
	}
#endregion Script Initialization

#region User Defined Functions
	#==============================
	#	Functions Start Here...
	#==============================
	Function Create-CompressedArchive ()
	{
		Param(
			[string]$SrcDir					# Directory where the files to be compressed are located
			, [string]$DestDir				# Directory where the compressed files will be saved at
			, [string]$ArchiveType			# Specifies the type of archive. Valid values are: "7z", "xz", "zip", "gzip", "bzip2, "tar"
			, [string]$ArchiveBaseName		# Name that will be given to the Archive file, without the extension.
			, [string]$FilesToBeCompressed	# Files to include in the archive
			#, [string]$ScriptLogPath
			, [string]$WorkingDirectory
			, [switch]$RecurseDirectory
		)
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

		If ($ScrDebug)
		{
			#Write-Host -ForegroundColor Cyan "$([string]::Join(", ", $Arguments))"
			Write-Host -ForegroundColor Cyan "Command being executed:`n $([char]34)$($PathToExe) $([string]::Join(" ", $Arguments))$([char]34)"
		}
		Add-Content -Path $FullScriptLogPath -Value "`t`tCommand being executed:`n`t`t`t $([char]34)$($PathToExe) $([string]::Join(" ", $Arguments))$([char]34)"

	    & $PathToExe $Arguments
	}
#endregion User Defined Functions

Add-Content -Path $FullScriptLogPath -Value "============================================================================================================================================="
Add-Content -Path $FullScriptLogPath -Value "`tStart Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
Add-Content -Path $FullScriptLogPath -Value "============================================================================================================================================="
Add-Content -Path $FullScriptLogPath -Value ""
Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"
Add-Content -Path $FullScriptLogPath -Value "`tGathering a list of all $([char]34)$($Extension)$([char]34) files in $([char]34)$($SourceDirectory)$([char]34)"

$DirectoryFiles = @(Get-ChildItem -Path $SourceDirectory | Where-Object { ($_.Extension -like "*$($Extension)") -and ($_.CreationTime -gt $CutOffDate) } | Select-Object Name,CreationTime,BaseName,Length) #,FullName)

Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"
Add-Content -Path $FullScriptLogPath -Value "`tBegin - Checking to see if a compressed file exists for each file in the array..."
Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"
Add-Content -Path $FullScriptLogPath -Value ""

$FilesToCompress = @()

ForEach ($RawFile in $DirectoryFiles)
{
	If (($RawFile.CreationTime -gt $CutOffDate) -and ($RawFile.CreationTime -lt $MinCutOffDate))
	{
		If (!(Test-Path -Path "$($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"))
		{
			If ($ScrDebug)
			{
				Write-Host -ForegroundColor Magenta "File Does NOT Exist: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
				Write-Host -ForegroundColor Magenta "`tAdding file $([char]34)$($RawFile.Name)$([char]34) to the list of files to compress.`n"
			}
			Add-Content -Path $FullScriptLogPath -Value "File Does NOT Exist: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
			Add-Content -Path $FullScriptLogPath -Value "`tAdding file $([char]34)$($RawFile.Name)$([char]34) to the list of files to compress."
			$FilesToCompress += $RawFile
		}
		Else
		{
			If ($ScrDebug)
			{
				Write-Host -ForegroundColor Green "`tCompressed File Exists: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
				Write-Host -ForegroundColor Green "`t`tNo compression task needed for file: $([char]34)$($RawFile.Name)$([char]34).`n"
			}
			Add-Content -Path $FullScriptLogPath -Value "`tCompressed File Exists: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
			Add-Content -Path $FullScriptLogPath -Value "`t`tNo compression task needed for file: $([char]34)$($RawFile.Name)$([char]34)."
		}
	}
}

If ($FilesToCompress.Count -ge 1)
{
	Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"
	Add-Content -Path $FullScriptLogPath -Value "`tBegin - Compressing Files that have no corresponding compressed file..."
	Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"

	#region Kill any existing jobs in the multi-thread queue
		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Magenta "Killing any existing jobs . . ."
		}

		# Kill any existing jobs in the multi-thread queue
		If ((Get-Job).Count -ge 1)
		{
			Add-Content -Path $FullScriptLogPath -Value "Killing any existing jobs . . ."
			Get-Job | Remove-Job -Force
			Add-Content -Path $FullScriptLogPath -Value "Done killing any existing jobs."
		}

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Magenta "Done killing any existing jobs."
		}
	#endregion Kill any existing jobs in the multi-thread queue

	$TotalFilesToCompress = $FilesToCompress.Count
	$i = 0

	#ForEach ($File in $FilesToCompress)
	#{
	#	If ($ScrDebug)
	#	{
	#		Write-Host -ForegroundColor Blue "Compressing file: $([char]34)$($File.Name)$([char]34)"
	#	}
	#	Add-Content -Path $FullScriptLogPath -Value "`tCompressing file: $([char]34)$($File.Name)$([char]34)"
	#
	#	Create-CompressedArchive -SrcDir "$($SourceDirectory)" -DestDir "$($DestinationDirectory)" -ArchiveType "$($CompressionFormat)" -ArchiveBaseName "$($File.BaseName)" -FilesToBeCompressed "$($File.Name)"
	#	$FileCount++
	#}

	#region Multi-Threading
		ForEach ($File in $FilesToCompress)
		{
			$Arguments = @($SourceDirectory, $DestinationDirectory, $CompressionFormat, $File.BaseName, $File.Name, $ScriptDirectory, $FullScriptLogPath, $false)

			# Check to see if there are too many open threads
		    # If there are too many threads then wait here until some close
		    While ($(Get-Job -State Running).Count -ge $MaxThreads)
			{
		        Write-Progress -Activity "Compressing files in the List" -Status "Waiting for running threads to finish and close" -CurrentOperation "$($i) threads created - $($(Get-Job -State Running).Count) threads open" -PercentComplete ($i / $TotalFilesToCompress * 100)
		        Start-Sleep -Milliseconds $SleepTimer
		    }

			# Starting job
			$i++

			If ($ScrDebug)
			{
				Write-Host -ForegroundColor Yellow "Compressing file: $([char]34)$($File.Name)$([char]34)"
			}
			Add-Content -Path $FullScriptLogPath -Value "`tCompressing file: $([char]34)$($File.Name)$([char]34)"

			# Call Task-CompressFiles-MultiThread.ps1 script
			Start-Job -Name $File.BaseName -FilePath $MultiThreadScript -ArgumentList $Arguments | Out-Null

			# Show Progress Bar 
			Write-Progress -Activity "Compressing files..." -Status "Starting Threads" -CurrentOperation "$($i) threads created - $($(Get-Job -State Running).Count) threads open" -PercentComplete ($i / $TotalFilesToCompress * 100)
		}

		# Get-Job | Wait-Job
		While ($(Get-Job -State Running).Count -gt 0)
		{
			$StillRunning = ""
			ForEach ($System  in $(Get-Job -State Running))
			{
				$StillRunning += ", $($System.Name)"
			}

			$StillRunning = $StillRunning.SubString(2)
			Write-Progress -Activity "Compressing files......" -Status "$($(Get-Job -State Running).Count) threads remaining" -CurrentOperation "$StillRunning" -PercentComplete ($(Get-Job -State Completed).Count / $(Get-Job).Count * 100)

			Start-Sleep -Milliseconds $SleepTimer
		}

		$JobResults = Get-Job | Receive-Job

		Add-Content -Path $FullScriptLogPath -Value "*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*"
		ForEach ($Job in $JobResults)
		{
			Add-Content -Path $FullScriptLogPath -Value "Job data for $([char]34)$($Job)$([char]34)"
			Add-Content -Path $FullScriptLogPath -Value "$($Job)"
			#Add-Content -Path $FullScriptLogPath -Value "******************************************************************"
		}
		Add-Content -Path $FullScriptLogPath -Value "*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"
	#endregion Multi-Threading
}
Else
{
	Add-Content -Path $FullScriptLogPath -Value "*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*"
	Add-Content -Path $FullScriptLogPath -Value "`tThere were no $([char]34)$($Extension)$([char]34) files found in $([char]34)$($SourceDirectory)$([char]34) that need to be compressed!"
	Add-Content -Path $FullScriptLogPath -Value "*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*"
}

Add-Content -Path $FullScriptLogPath -Value "============================================================================================================================================="
Add-Content -Path $FullScriptLogPath -Value "`tEnd Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
Add-Content -Path $FullScriptLogPath -Value "============================================================================================================================================="

<#
	Get-ChildItem -Path \\stor22\datadump\kaplan\ER_Reports | Where-Object { $_.Name -eq "ThreadPostIncremental_kaplan_20130505.CSV" } | Select-Object * | Format-List
	---------------------------------------------------------------------------------------------------------------------------------------------
	The following output is from the preceeding command:
	---------------------------------------------------------------------------------------------------------------------------------------------
	PSPath            : Microsoft.PowerShell.Core\FileSystem::\\stor22\datadump\kaplan\ER_Reports\ThreadPostIncremental_kaplan_20130505.CSV
	PSParentPath      : Microsoft.PowerShell.Core\FileSystem::\\stor22\datadump\kaplan\ER_Reports
	PSChildName       : ThreadPostIncremental_kaplan_20130505.CSV
	PSProvider        : Microsoft.PowerShell.Core\FileSystem
	PSIsContainer     : False
	VersionInfo       : File:             \\stor22\datadump\kaplan\ER_Reports\ThreadPostIncremental_kaplan_20130505.CSV
	                    InternalName:
	                    OriginalFilename:
	                    FileVersion:
	                    FileDescription:
	                    Product:
	                    ProductVersion:
	                    Debug:            False
	                    Patched:          False
	                    PreRelease:       False
	                    PrivateBuild:     False
	                    SpecialBuild:     False
	                    Language:

	BaseName          : ThreadPostIncremental_kaplan_20130505
	Mode              : -a---
	Name              : ThreadPostIncremental_kaplan_20130505.CSV
	Length            : 123861878
	DirectoryName     : \\stor22\datadump\kaplan\ER_Reports
	Directory         : \\stor22\datadump\kaplan\ER_Reports
	IsReadOnly        : False
	Exists            : True
	FullName          : \\stor22\datadump\kaplan\ER_Reports\ThreadPostIncremental_kaplan_20130505.CSV
	Extension         : .CSV
	CreationTime      : 5/5/2013 3:07:03 AM
	CreationTimeUtc   : 5/5/2013 9:07:03 AM
	LastAccessTime    : 6/24/2013 11:11:42 AM
	LastAccessTimeUtc : 6/24/2013 5:11:42 PM
	LastWriteTime     : 5/5/2013 3:09:49 AM
	LastWriteTimeUtc  : 5/5/2013 9:09:49 AM
	Attributes        : Archive

	---------------------------------------------------------------------------------------------------------------------------------------------
	Compress each CSV file individually and move the compressed file to the "CSV_Compressed" folder.
	
	Get-ChildItem -Path \\stor22\datadump\kaplan\ER_Reports | Where-Object { $_.Attributes -ne "Directory" }
#>
