#region Assembly and PowerShell Module Initialization
	If (!(Get-Command Write-Zip -ErrorAction SilentlyContinue))
	{
		try
		{
			Import-Module Pscx
		}
		catch
		{
			If (!(Get-Command Write-Zip -ErrorAction SilentlyContinue))
			{
				Import-Module Pscx
			}
		}
		finally
		{
			If (!(Get-Command Write-Zip -ErrorAction SilentlyContinue))
			{
				throw "Cannot load the $([char]34)Pscx$([char]34) module!! Please make sure the $([char]34)PowerShell Community Extensions (PSCX)$([char]34) are installed in one of the following path: $([char]34)$(${env:PSModulePath})$([char]34)."
			}
		}
	}
	#region Import Assemblies
		#[void][Reflection.Assembly]::LoadFrom("C:\Program Files\7-Zip\7z.dll")
	#endregion Import Assemblies
#endregion Assembly and PowerShell Module Initialization
#Get-Command -Module Pscx -CommandType Cmdlet

#region Script Initialization
	[string]$LogDirDateName = (Get-Date -Format 'yyyy.MM.dd.ddd').ToString()
	[string]$LogFileDate = (Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()

	#[string]$ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path)
	[string]$PtlLogName = ($MyInvocation.MyCommand.Name).Replace(".ps1", "")
	[string]$Ext = "log"
	[string]$LogFileName = "$($LogFileDate)_$($PtlLogName).$($Ext)"

	#If ($ScrDebug)
	#{
		If (!(Test-Path -Path "$(${env:userprofile})\Desktop\Logs\$($LogDirDateName)"))
		{
			New-Item -ItemType Directory "$(${env:userprofile})\Desktop\Logs\$($LogDirDateName)" | Out-Null
		}
		$Script:FullScriptLogPath = "$(${env:userprofile})\Desktop\Logs\$($LogDirDateName)\$($LogFileName)"
	#}

	[string]$SourceDirectory = "\\stor22\datadump\kaplan\ER_Reports"
	[string]$DestinationDirectory = "\\stor22\datadump\kaplan\ER_Reports\CSV_Compressed"
	[int]$MaxAge = 60
	[int]$MinAge = 0
	[string]$CompressionFormat = "7z"
	[string]$Extension = "csv"

	If (!(Test-Path -Path "$($DestinationDirectory)"))
	{
		New-Item -ItemType Directory "$($DestinationDirectory)" | Out-Null
	}

	$CutOffDate = (Get-Date).AddDays(-$MaxAge)

	If ($MinAge -ge 1)
	{
		$MinCutOffDate = (Get-Date).AddDays(-$MinAge)
	}
	Else
	{
		$MinCutOffDate = Get-Date
	}
	#$MinCutOffDate = (Get-Date).AddDays(-$MinAge)
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
			, [string]$ArchiveType			# Specifies the type of archive. Valid values are: "7z", "zip", "xz", "bzip2, "gzip", "tar" and "wim"
			, [string]$ArchiveBaseName		# Name that will be given to the Archive file, without the extension.
			, [string]$FilesToBeCompressed	# Files to include in the archive
			, [string]$FullScriptLogPath
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
		# Path to the 7-Zip compression executable.
	    #[string]$PathToExe = "C:\Program Files\7-Zip\7z.exe"
	    [string]$PathToExe = ".\7za920\7za.exe"

		If ($RecurseDirectory)
		{
			[array]$Arguments = "a", "-t$($ArchiveType)", "$($DestDir)\$($ArchiveBaseName).$($ArchiveType)", "$($SrcDir)\$($FilesToBeCompressed)", "-r"
		}
		Else
		{
			[array]$Arguments = "a", "-t$($ArchiveType)", "$($DestDir)\$($ArchiveBaseName).$($ArchiveType)", "$($SrcDir)\$($FilesToBeCompressed)"
		}
		#Write-Host -ForegroundColor Cyan "$([string]::Join(", ", $Arguments))"
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
Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"

$FilesToCompress = @()
$DirectoryFiles = @(Get-ChildItem -Path $SourceDirectory | Where-Object { ($_.Extension -like "*$($Extension)") -and ($_.CreationTime -gt $CutOffDate) } | Select-Object Name,CreationTime,BaseName) #,FullName)

Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"
Add-Content -Path $FullScriptLogPath -Value "`tBegin - Checking to see if a compressed file exists for each file in the array..."
Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"

ForEach ($RawFile in $DirectoryFiles)
{
	If (($RawFile.CreationTime -gt $CutOffDate) -and ($RawFile.CreationTime -lt $MinCutOffDate))
	{
		If (!(Test-Path -Path "$($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"))
		{
			#Write-Host -ForegroundColor Magenta "File Does NOT Exist: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
			#Write-Host -ForegroundColor Magenta "Adding file $([char]34)$($RawFile.Name)$([char]34) to the list of files to compress."
			Add-Content -Path $FullScriptLogPath -Value "`tFile Does NOT Exist: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
			Add-Content -Path $FullScriptLogPath -Value "`t`tAdding file $([char]34)$($RawFile.Name)$([char]34) to the list of files to compress."
			$FilesToCompress += $RawFile
		}
		Else
		{
			#Write-Host -ForegroundColor Green "File Exists: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
			Add-Content -Path $FullScriptLogPath -Value "Compressed File Exists: $($DestinationDirectory)\$($RawFile.BaseName).$($CompressionFormat)"
			Add-Content -Path $FullScriptLogPath -Value "`tNo compression task needed for file: $([char]34)$($RawFile.Name)$([char]34)."
		}
	}
}

Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"
Add-Content -Path $FullScriptLogPath -Value "`tBegin Compressing Files that have no corresponding compressed file..."
Add-Content -Path $FullScriptLogPath -Value "---------------------------------------------------------------------------------------------------------------------------------------------"

ForEach ($File in $FilesToCompress)
{
	#Write-Host -ForegroundColor Blue "Compressing file: $([char]34)$($RawFile.Name)$([char]34)"
	Add-Content -Path $FullScriptLogPath -Value "`tCompressing file: $([char]34)$($RawFile.Name)$([char]34)"

	Create-CompressedArchive -SrcDir "$($SourceDirectory)" -DestDir "$($DestinationDirectory)" -ArchiveType "$($CompressionFormat)" -ArchiveBaseName "$($File.BaseName)" -FilesToBeCompressed "$($File.Name)"
}

#$FilesToCompress | Export-CSV -Path "${env:userprofile}\Desktop\Kaplan_Export_FilesToCompress_1.csv" -NoTypeInformation
Add-Content -Path $FullScriptLogPath -Value "============================================================================================================================================="
Add-Content -Path $FullScriptLogPath -Value "`tEnd Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
Add-Content -Path $FullScriptLogPath -Value "============================================================================================================================================="
