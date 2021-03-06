#Requires -Version 2.0
####################################################################################################################
##
##	Script Name          : Get-ServerUptime.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2016 Bill Campbell. All rights reserved.
##	Created On           : Feb 09, 2016
##	Added to Script Repo : 09 Feb 2016
##	Last Modified        : 2016.Feb.09
##
##	Version              : 2016.02.09.01
##	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
##	Version Example      : If this is the 6th revision/save on May 22 2013 then RR would be '06' and the version number will be formatted as follows: 2013.05.22.06
##
##	Purpose              : Execute a series of commands using Powershell to get the last boot time of the server and the number of days it has been up.
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
		Param()
	#endregion Script Parameters

	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	#[string]$Script:FullScrPath = ($MyInvocation.MyCommand.Definition)
	#[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)

	#[string]$Script:ScriptNameNoExt = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	#[string]$Script:ScriptNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
#endregion Script Variable Initialization

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	$Win32OpSys = Get-WmiObject -ComputerName ${Env:COMPUTERNAME} -Query 'SELECT * FROM Win32_OperatingSystem'

	[DateTime]$ServerLastBootUpTime = $Win32OpSys.ConvertToDateTime($Win32OpSys.LastBootUpTime)
	[TimeSpan]$ServerUpTime = New-TimeSpan -Start $ServerLastBootUpTime -End $(Get-Date)
	[string]$SvLastBootUpTime = $ServerLastBootUpTime.ToString()
	[int]$SvUpTime = $ServerUpTime.Days

	if ($SvUpTime -lt 90)
	{
		[string]$WriteColor = 'Green'
	}
	elseif ($SvUpTime -gt 180)
	{
		[string]$WriteColor = 'Red'
	}
	else
	{
		[string]$WriteColor = 'Yellow'
	}

	Write-Host -ForegroundColor $WriteColor "Last Boot Up Time: $($SvLastBootUpTime)"
	Write-Host -ForegroundColor $WriteColor "Number of days Since last reboot: $($SvUpTime)"

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
	.SYNOPSIS
		Gets the last boot time of the server and the number of days it has been up.

	.DESCRIPTION
		The Get-ServerUptime script does the following:

		Execute a series of commands using Powershell to get the last boot time of the server and the number of days it has been up.

	.EXAMPLE
		. ${Env:SystemDrive}\Scripts\Get-ServerUptime.ps1

	.INPUTS
		None
			This script does not accept any input.

	.OUTPUTS
		None
			This script does not return any output.

	.NOTES
		Author : Bill Campbell
		Version: 1.0.0.1
		Release: 2016-Feb-09

		REQUIREMENTS
			PowerShell Version 2.0
#>
#endregion Help
