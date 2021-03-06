#Requires -Version 2.0
####################################################################################################################
##
##	Script Name          : Create-GPOWithWMIFilter.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2016 Bill Campbell. All rights reserved.
##	Created On           : Jul 25, 2016
##	Added to Script Repo : 25 Jul 2016
##	Last Modified        : 2016.Jul.25
##
##	Version              : 2016.07.25.01
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
	ConfirmImpact = 'Medium'
	#, DefaultParameterSetName = '<ParameterSetName>'
)]
#region Script Variable Initialization
	#region Script Parameters
		Param(
			#[Parameter(
			#	Position			= 0
			#	, Mandatory			= $true
			#	, ValueFromPipeline	= $false
			#	, HelpMessage		= '<Help Message...> ("Value1", "Value2").'
			#)]
			#[ValidateSet('Value1', 'Value2')]
			#[string]$Parameter1Name
			#, [Parameter(
			#	Position			= 1
			#	, Mandatory			= $false
			#	, ValueFromPipeline	= $false
			#	, HelpMessage		= 'Enables/Disables custom debugging code embedded in the script ($true/$false). Defaults to FALSE.'
			#)][switch]$ScrDebug = $false
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
		## Check for GPWmiFilter Module, attempt to load
		if (!(Get-Command Get-GPWmiFilter -ErrorAction SilentlyContinue))
		{
			try
			{
				Import-Module GPWmiFilter
			}
			catch
			{
				if (!(Get-Command Get-GPWmiFilter -ErrorAction SilentlyContinue))
				{
					Import-Module GPWmiFilter
				}
			}
			finally
			{
				if (!(Get-Command Get-GPWmiFilter -ErrorAction SilentlyContinue))
				{
					throw "Cannot load the $([char]34)GPWmiFilter$([char]34) module!! Please correct any errors and try again."
				}
			}
		}
		$Error.Clear()
	#endregion Assembly and PowerShell Module Initialization
#endregion Script Variable Initialization

#region User Defined Functions
	#Function fn_FunctionName()
	#{
	#	<#
	#		.SYNOPSIS
	#			<Brief overview of what the function does.>
	#
	#		.DESCRIPTION
	#			The <fn_FunctionName> function does the following:
	#
	#			<< IN DEPTH DESCRIPTION. >>
	#
	#		.PARAMETER <FunctionVariable>
	#			<FunctionVariable Description.>
	#
	#		.EXAMPLE
	#			fn_FunctionName -FunctionVariable 'FunctionVariableValue'
	#
	#			Description
	#			-----------
	#				<< IN DEPTH DESCRIPTION OF WHAT THE EXAMPLE IS DOING. >>
	#	#>
	#	[CmdletBinding(
	#		SupportsShouldProcess = $true,
	#		ConfirmImpact = 'Medium'
	#	)]
	#	#region Function Parameters
	#		Param(
	#			[Parameter(
	#				Position			= 0
	#				, Mandatory			= $true
	#				, ValueFromPipeline	= $false
	#				, HelpMessage		= '<Help Message...>'
	#			)][string]$FunctionVariable
	#		)
	#	#endregion Function Parameters
	#
	#	Begin
	#	{}
	#	Process
	#	{}
	#	End
	#	{}
	#}
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	$filter = New-GPWmiFilter -Name 'Windows Servers Only' -Expression 'SELECT * FROM Win32_OperatingSystem WHERE ProductType != "1"' -Description 'Only Windows Server class operating systems'

	$gpo = New-GPO -Name "Test GPO"
	$gpo.WmiFilter = $filter

	Get-GPWmiFilter -Name 'Windows Servers Only'
	# Get-GPWmiFilter -Name 'Windows Servers Only' | Remove-GPWmiFilter

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
	.SYNOPSIS
		<Brief overview of what the script does.>

	.DESCRIPTION
		The Create-GPOWithWMIFilter script does the following:

		<< IN DEPTH DESCRIPTION >>

	.PARAMETER <Parameter1Name>
		<Parameter1 Description - Max width of (71? or 80?) characters. Then you need to add a new line.>

		Acceptable values:
			value1
			value2
			value3
			...

		Default value is 'value2'.

	.PARAMETER ScrDebug
		Enables/Disables custom debugging code embedded in the script ($true/$false). Defaults to FALSE.

	.EXAMPLE
		${Env:SystemDrive}\Path\To\Scripts\Directory\Create-GPOWithWMIFilter.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'

		Description
		-----------
			<< IN DEPTH DESCRIPTION OF WHAT THE EXAMPLE DOES. >>

	.INPUTS
		None
			This script does not accept any input.

	.OUTPUTS
		None
			This script does not return any output.

	.NOTES
		Author : Bill Campbell
		Version: 1.0.0.1
		Release: 2016-Jul-25

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
#		.\Create-GPOWithWMIFilter.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'
#		. ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Path\To\Proper\Directory\Create-GPOWithWMIFilter.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value' -ScrDebug
#		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\Create-GPOWithWMIFilter.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'
#		. ${Env:SystemDrive}\Path\To\Scripts\Directory\Create-GPOWithWMIFilter.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value' -ScrDebug
#
#	-- Script Change Log --
#	Changes for Jul-2016
#		- Initial Script/Module Writing and Debugging.
#
###################################################################################################################
#endregion Script Change Log
