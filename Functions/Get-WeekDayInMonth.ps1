#Requires -Version 2.0
####################################################################################################################
##
##	Script Name          : Get-WeekDayInMonth.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2015 Bill Campbell. All rights reserved. No part of this publication may be reproduced, stored in a retrieval system, or transmitted, in any form or by means electronic, mechanical, photocopying, or otherwise, without prior written permission of the publisher.
##	Created On           : Dec 10, 2015
##	Added to Script Repo : 10 Dec 2015
##	Last Modified        : 2015.Dec.10
##
##	Version              : 2015.12.10.01
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
	Function Get-WeekDayInMonth()
	{
		[CmdletBinding()]
		#region Function Parameters
			Param(
				[Parameter(
					Position			= 0
					, Mandatory			= $true
					, HelpMessage		= 'Specifies the month that is searched. Enter a value from 1 to 12.'
				)]
				[ValidateSet(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)]
				[int]$Month
				, [Parameter(
					Position			= 1
					, Mandatory			= $true
					, HelpMessage		= 'Specifies the year that is searched. Enter a value from 1 to 9999.'
				)][int]$Year
				, [Parameter(
					Position			= 2
					, Mandatory			= $true
					, HelpMessage		= 'The number of the week in the month to be searched.'
				)]
				[ValidateSet(1, 2, 3, 4, 5)]
				[int]$WeekNumber
				, [Parameter(
					Position			= 3
					, Mandatory			= $true
					, HelpMessage		= 'Represents the week day you are after. 0 = Sunday, 1 = Monday, 2 = Tuesday, 3 = Wednesday, 4 = Thursday, 5 = Friday, 6 = Saturday'
				)]
				[ValidateSet(1, 2, 3, 4, 5, 6, 0)]
				[int]$WeekDay
			)
		#endregion Function Parameters
		<#
			DisplayHint : DateTime
			DateTime    : Tuesday, December 1, 2015 12:00:00 AM
			Date        : 12/1/2015 12:00:00 AM
			Day         : 1
			DayOfWeek   : Tuesday
			DayOfYear   : 335
			Hour        : 0
			Kind        : Local
			Millisecond : 132
			Minute      : 0
			Month       : 12
			Second      : 0
			Ticks       : 635845248001328029
			TimeOfDay   : 00:00:00.1328029
			Year        : 2015
		#>
		$FirstDayOfMonth = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0

		## First week day of the month (i.e. first monday of the month)
		[int]$FirstDayofMonthDay = $FirstDayOfMonth.DayOfWeek
		$Difference = $WeekDay - $FirstDayofMonthDay

		if ($Difference -lt 0)
		{
			$DaysToAdd = 7 - ($FirstDayofMonthDay - $WeekDay)
		}
		elseif ($difference -eq 0)
		{
			$DaysToAdd = 0
		}
		else
		{
			$DaysToAdd = $Difference
		}

		$FirstWeekDayofMonth = $FirstDayOfMonth.AddDays($DaysToAdd)
		Remove-Variable DaysToAdd

		## Add Weeks
		$DaysToAdd = ($WeekNumber -1)*7
		$TheDay = $FirstWeekDayofMonth.AddDays($DaysToAdd)

		if (!($TheDay.Month -eq $Month -and $TheDay.Year -eq $Year))
		{
			$TheDay = $null
		}

		$TheDay
	}
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	<# PLACE YOUR MAIN SCRIPT CODE HERE!! #>
##Get-Date -Year 2015 -Month 10 -Day 1 -Hour 0 -Minute 0 -Second 0 | Select-Object * | Format-List
#$FirstDayOfMonth = Get-Date -Year 2015 -Month 10 -Day 1 -Hour 0 -Minute 0 -Second 0
#[int]$FirstDayofMonthDay = $FirstDayOfMonth.DayOfWeek
#$FirstDayOfMonth.DayOfWeek
#$FirstDayofMonthDay

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
	.SYNOPSIS
		<Brief overview of what the script does.>

	.DESCRIPTION
		The Get-WeekDayInMonth script does the following:

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
		${Env:SystemDrive}\Path\To\Scripts\Directory\Get-WeekDayInMonth.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'

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
		Release: 2015-Dec-10

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
#		.\Get-WeekDayInMonth.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'
#		. ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Path\To\Proper\Directory\Get-WeekDayInMonth.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value' -ScrDebug
#		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\Get-WeekDayInMonth.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'
#		. ${Env:SystemDrive}\Path\To\Scripts\Directory\Get-WeekDayInMonth.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value' -ScrDebug
#
#	-- Script Change Log --
#	Changes for Dec-2015
#		- Initial Script/Module Writing and Debugging.
#
###################################################################################################################
#endregion Script Change Log
