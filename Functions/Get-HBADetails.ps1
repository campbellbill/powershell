#Requires -Version 4.0
#Requires -RunAsAdministrator
####################################################################################################################
##
##	Script Name          : Get-HBADetails.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2015 Bill Campbell. All rights reserved.
##	Created On           : Jul 08, 2015
##	Added to Script Repo : 08 Jul 2015
##	Last Modified        : 2015.Jul.20
##
##	Version              : 2015.07.20.01
##	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
##	Version Example      : If this is the 6th revision/save on May 22 2013 then RR would be '06' and the version number will be formatted as follows: 2013.05.22.06
##
##	Purpose              : Execute a series of commands using Powershell for the purpose of 'X' and 'Y'.
##
##	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
##	#! IMPORTANT NOTICE:																						 !#
##	#!		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK FROM THE USE OF,	 !#
##	#!		OR THE RESULTS RETURNED FROM THE USE OF, THIS CODE REMAINS WITH THE USER.							 !#
##	#! 																											 !#
##	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
##	#! 																											 !#
##	#! 						Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force							 !#
##	#! 																											 !#
##	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
##
##	Notes                : Use this area for general purpose use notes for the script.
##							Gets a list of all current Environment Variables:			Get-ChildItem Env:
##							Change Directories to the Current Users 'Desktop' folder:	Set-Location ${Env:USERPROFILE}\Desktop
##							Some Commonly used Environment Variables:					Write-Output ${Env:USERPROFILE}; Write-Output ${Env:COMPUTERNAME}; Write-Output ${Env:USERNAME};
##							Some Commonly used commands:								[Guid]::NewGuid(); [Environment]::OSVersion.Version
##																						$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))
##
####################################################################################################################
[CmdletBinding(
	SupportsShouldProcess = $true,
	ConfirmImpact = 'Medium',
	DefaultParameterSetName = 'CSVFile'
)]
#region Script Variable Initialization
	#region Script Parameters
		Param(
			[Parameter(
				Position			= 0
				, ParameterSetName	= 'CSVFile'
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'The full path to a CSV file with the following Headers and info in it: "DNSName","ClusterName","NodeName". The DNSName column is required. It should be the FQDN of the server.'
			)][string]$ServerList
			#, [Parameter(
			#	Position			= 0
			#	, ParameterSetName	= 'Manual'
			#	, Mandatory			= $false
			#	, ValueFromPipeline	= $false
			#	, HelpMessage		= 'A comma separated list of FQDN Computer Name(s) to execute the script against.'
			#)][string[]]$ComputerNames
		)
	#endregion Script Parameters

	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	#[string]$Script:FullScrPath = ($MyInvocation.MyCommand.Definition)
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)

	#[string]$Script:ScriptNameNoExt = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	[string]$Script:ScriptNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

	#region Assembly and PowerShell Module Initialization
		######################################
		##  Import Modules & Add PSSnapins  ##
		######################################
		## Check for [ModuleName] Module, attempt to load
		## Check for [SnapinName] Snapin, attempt to load
		#if (!(Get-Command [CmdletNameFrom-ModuleOrSnappin] -ErrorAction SilentlyContinue))
		#{
		#	try
		#	{
		#		#Import-Module [ModuleName]
		#		#Add-PSSnapin [SnapinName]
		#	}
		#	catch
		#	{
		#		if (!(Get-Command [CmdletNameFrom-ModuleOrSnappin] -ErrorAction SilentlyContinue))
		#		{
		#			#Import-Module [ModuleName]
		#			#Add-PSSnapin [SnapinName]
		#		}
		#	}
		#	finally
		#	{
		#		if (!(Get-Command [CmdletNameFrom-ModuleOrSnappin] -ErrorAction SilentlyContinue))
		#		{
		#			#throw "Cannot load the $([char]34)[ModuleName]$([char]34) module!! Please correct any errors and try again."
		#			#throw "Cannot load the $([char]34)[SnapinName]$([char]34) Snapin!! Please correct any errors and try again."
		#		}
		#	}
		#}
		#$Error.Clear()
	#endregion Assembly and PowerShell Module Initialization
#endregion Script Variable Initialization

#region User Defined Functions
	function Get-HBAInfo
	{
		[CmdletBinding()]
		Param(
			[Parameter(
				Mandatory = $false
				, ValueFromPipeline = $true
				, Position = 0
			)][string]$ComputerName
		)

		begin 
		{
			$Namespace = 'root\WMI'
		}
		process
		{
			$port = Get-WmiObject -ErrorAction SilentlyContinue -Class MSFC_FibrePortHBAAttributes -Namespace $Namespace @PSBoundParameters
			$hbas = Get-WmiObject -ErrorAction SilentlyContinue -Class MSFC_FCAdapterHBAAttributes -Namespace $Namespace @PSBoundParameters

			$hbaProp = $hbas | Get-Member -MemberType Property, AliasProperty | Select-Object -ExpandProperty Name | Where-Object { $_ -notlike "__*" }
			$hbas = $hbas | Select-Object $hbaProp

			$hbas | ForEach-Object {
				$_.NodeWWN = ((($_.NodeWWN) | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()
			}

			foreach ($hba in $hbas)
			{
				Add-Member -MemberType NoteProperty -InputObject $hba -Name FabricName -Value (
					($port | Where-Object { $_.InstanceName -eq $hba.InstanceName }).Attributes | Select-Object @{Name = 'FabricName'; Expression = {(($_.FabricName | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()}}
				) #-PassThru

				Add-Member -MemberType NoteProperty -InputObject $hba -Name PortWWN -Value (
					($port | Where-Object { $_.InstanceName -eq $hba.InstanceName }).Attributes | Select-Object @{Name = 'PortWWN'; Expression = {(($_.PortWWN | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()}} 
				) -PassThru
			}
		}
	}
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	$ClustersHBAInfo = @()

################################################################################
## Figure out the parameter set that was used and then use that to make the array that gets looped over...
################################################################################
	## Has the following headers in it:
	##	'DNSName','ClusterName','NodeName'
	$SQLServers = @(Import-Csv -Path $ServerList)

	foreach ($SQLServer in $SQLServers)
	{
		$SrvDnsName = $SQLServer.DNSName

		if (!(Test-Connection -ComputerName $SrvDnsName -Count 1 -Quiet))
		{
			continue
		}

		$ClustersHBAInfo += Get-HBAInfo -ComputerName $SrvDnsName | Select-Object * -ExpandProperty 'FabricName' -ExcludeProperty 'FabricName' | Select-Object * -ExpandProperty 'PortWWN' -ExcludeProperty 'PortWWN'
	}

	$ClustersHBAInfo | Select-Object `
						@{Name = "ComputerName"; Expression = {$_.PSComputerName}}, `
						'Model', `
						@{Name = "HBA Bios"; Expression = {"$($_.OptionROMVersion)"}}, `
						'DriverName', `
						@{Name = "Stor Miniport Driver Version"; Expression = {$_.DriverVersion}}, `
						'FirmwareVersion', 'PortWWN', 'NodeWWN' `
					| Export-Csv -Path "$($ScriptDirectory)\$($ScriptNameNoExt)_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
	.SYNOPSIS
		<Brief overview of what the script does.>

	.DESCRIPTION
		The Get-HBADetails script does the following:

		<< IN DEPTH DESCRIPTION >>

	.PARAMETER ServerList
		CSV list of Computer Name(s) to execute the script against.

	.PARAMETER ComputerNames
		A comma separated list of FQDN Computer Name(s) to execute the script against.

	.EXAMPLE
		${Env:SystemDrive}\Path\To\Scripts\Directory\Get-HBADetails.ps1 -ServerList 'C:\Scripts\SQLSvr-List-Driver-Queries.csv'

		Description
		-----------
			Imports the CSV file 'C:\Scripts\SQLSvr-List-Driver-Queries.csv' and executes the script against the servers returning the requested information.

	.EXAMPLE
		${Env:SystemDrive}\Path\To\Scripts\Directory\Get-HBADetails.ps1 -ComputerNames 'Server1','Server2','Server3'

		Description
		-----------
			Uses the list of computer names provided in the 'ComputerNames' parameter and executes the script against those servers returning the requested information. 

	.INPUTS
		None
			This script does not accept any input.

	.OUTPUTS
		None
			This script does not return any output.

	.NOTES
		Author : Bill Campbell
		Version: 1.0.0.1
		Release: 2015-Jul-08

		REQUIREMENTS
			PowerShell Version 2.0

	.LINK
		about_Providers

	.LINK
		Get-ChildItem

	.LINK
		New-Item

	.LINK
		Rename-Item

	.LINK
		Set-Item
#>
#endregion Help

#region Script Change Log
###################################################################################################################
#
#	EXAMPLES
#		.\Get-HBADetails.ps1 -ServerList 'C:\Scripts\SQLSvr-List-Driver-Queries.csv'
#		.\Get-HBADetails.ps1 -ComputerNames 'Server1','Server2','Server3'
#		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\Get-HBADetails.ps1 -ServerList 'C:\Scripts\SQLSvr-List-Driver-Queries.csv'
#		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\Get-HBADetails.ps1 -ComputerNames 'Server1','Server2','Server3'
#		. ${Env:SystemDrive}\Path\To\Scripts\Directory\Get-HBADetails.ps1 -ServerList 'C:\Scripts\SQLSvr-List-Driver-Queries.csv'
#		. ${Env:SystemDrive}\Path\To\Scripts\Directory\Get-HBADetails.ps1 -ComputerNames 'Server1','Server2','Server3'
#
#	-- Script Change Log --
#	Changes for Jul-2015
#		- Initial Script/Module Writing and Debugging.
#
###################################################################################################################
#endregion Script Change Log
