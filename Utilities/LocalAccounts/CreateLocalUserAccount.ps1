<#
	Script Name          : CreateLocalUserAccount.ps1
	Written by           : Bill Campbell
	Written on           : Nov 26, 2013
	Added to Script Repo : '[Date Added to a Source Control System]'
	Last Modified        : 2013.Nov.26

	Version              : 2013.11.26.01
	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
	Version Example      : If this is the 6th revision/save on December 28 2012 then RR would be "06" and the version number will be formatted as follows: 2012.12.28.06

	Purpose              : Execute a series of commands using Powershell for the purpose of "X" and "Y".

	Notes                : Use this area for general purpose use notes for the script.
							Change Directories to the Current Users "Desktop" folder
								Set-Location ${Env:USERPROFILE}\Desktop
							Some Commonly used Environment Variables:
								Write-Output ${Env:USERPROFILE}; Write-Output ${Env:COMPUTERNAME}; Write-Output ${Env:USERNAME}
							Gets a list of all current Environment Variables:
								Get-ChildItem Env:

	External Links		 : http://www.petri.co.il/create-local-accounts-with-powershell.htm

	EXAMPLES
		.\CreateLocalUserAccount.ps1 -User 'xl_deploy' -Password 'Passw0rd^&!' -Computer 'localhost' -ScrDebug
		. ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Utilities\LocalAccounts\CreateLocalUserAccount.ps1 -User 'xl_deploy' -Password 'Passw0rd^&!' -Computer 'localhost'
		. ${Env:SystemDrive}\Projects\CreateLocalUserAccount.ps1 -User 'xl_deploy' -Password 'Passw0rd^&!' -Computer 'localhost'
		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\CreateLocalUserAccount.ps1 -User 'xl_deploy' -Password 'Passw0rd^&!' -Computer 'localhost'
		. ${Env:SystemDrive}\Path\To\Scripts\Directory\CreateLocalUserAccount.ps1 -User 'xl_deploy' -Password 'Passw0rd^&!' -Computer 'localhost'

	-- Script Change Log --
	Changes for Nov-2013
		- Initial Script Writing and Debugging.

	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
	#! IMPORTANT NOTE:																							 !#
	#!	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK FROM THE USE OF,		 !#
	#!	OR THE RESULTS RETURNED FROM THE USE OF, THIS CODE REMAINS WITH THE USER.								 !#
	#! 																											 !#
	#! 						Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force							 !#
	#! 																											 !#
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
	#! 																											 !#
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#>
#region Script Variable Initialization
	#region Script Parameters
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Username of the user to create.'
			)][string]$User
			, [Parameter(
				Position			= 1
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Password for the new user being created.'
			)][string]$Password
			, [Parameter(
				Position			= 2
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Specifies the name of the computer upon which to run the script.'
			)][string]$Computer = "localhost"
			, [Parameter(
				Position			= 3
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Prints the help information.'
			)][switch]$Help = $false
			, [Parameter(
				Position			= 4
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'The ScrDebug switch turns on some custom code to help with debugging the script ($true/$false). Defaults to FALSE.'
			)][switch]$ScrDebug = $false
		)
	#endregion Script Parameters

	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"

	#region Assembly and PowerShell Module Initialization
		######################################
		##  Import Modules & Add PSSnapins  ##
		######################################
		## Check for [ModuleName] Module, attempt to load
		## Check for [SanpinName] Sanpin, attempt to load
		#If (!(Get-Command [CmdletNameFrom-ModuleOrSnappin] -ErrorAction SilentlyContinue))
		#{
		#	Try
		#	{
		#		#Import-Module [ModuleName]
		#		#Add-PSSnapin [SanpinName]
		#	}
		#	Catch
		#	{
		#		If (!(Get-Command [CmdletNameFrom-ModuleOrSnappin] -ErrorAction SilentlyContinue))
		#		{
		#			#Import-Module [ModuleName]
		#			#Add-PSSnapin [SanpinName]
		#		}
		#	}
		#	Finally
		#	{
		#		If (!(Get-Command [CmdletNameFrom-ModuleOrSnappin] -ErrorAction SilentlyContinue))
		#		{
		#			#Throw "Cannot load the $([char]34)[ModuleName]$([char]34) module!! Please correct any errors and try again."
		#			#Throw "Cannot load the $([char]34)[SanpinName]$([char]34) Snapin!! Please correct any errors and try again."
		#		}
		#	}
		#}
		#$Error.Clear()
	#endregion Assembly and PowerShell Module Initialization
#endregion Script Variable Initialization

#region User Defined Functions
	
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	[ADSI]$Server = "WinNT://$($Computer)"
	$Server.Children | Where-Object {$_.class -eq "user"} | Format-Table Name,Description -AutoSize

	# Create the user account on the specified server
	$NewUserAcct = $Server.Create("User", "$($User)")
	$NewUserAcct.SetPassword("$($Password)")
	$NewUserAcct.SetInfo()

	# Sets the account password to never expire
	$UserFlags = $NewUserAcct.UserFlags.Value -bor 0x10000
	$NewUserAcct.Put("UserFlags", $UserFlags)
	# Sets the Description field on the new user account
	$NewUserAcct.Put("Description", "$($User) - Local Test Account")
	$NewUserAcct.SetInfo()

	$Server.Children | Where-Object {$_.class -eq "user"} | Format-Table Name,Description -AutoSize

	# Add new user to a local group
	[ADSI]$Group = "WinNT://$($Computer)/Power Users, Group"
	#$NewUserAcct.Path
	$Group.Add($NewUserAcct.Path)

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
.SYNOPSIS
Deletes the files from the location meeting the specific criteria.

.DESCRIPTION
The Remove-AgedFiles script deletes items from the specified Path that 
meet the following specified criteria:
        
    File Extention = specified extentions (full or partial)
    File Name = name(s) (full or partial) of files to remove
    Path of items to remove

If the log file will does not exist, a new will be created. If the log 
file exists, it will be appended. One log per month.  Log name will lo
ok like:

If settings file supplied:
    "Remove-AgedFiles_SettingsFileName_YYYY-mm.log"

If settings file not supplied:
    "Remove-AgedFiles_YYYY-mm.log"

.PARAMETER <ParameterName> ([string] [switch] [array] [int])
<Parameter Description - Max width of 71 characters. Then you need to ad
d a new line>

Acceptable values:
    value1
    value2
    value3
    ...

Default value is "value2".

.PARAMETER ScrDebug
The ScrDebug switch turns on some custom code to help with debugging the
script ($true/$false). Defaults to FALSE.

.PARAMETER WhatIf
Describes what would happen if you executed the command without actuall
y executing the command.  Items will not be removed. Output will only o
ccur on display. Redirection is not possible.

.INPUTS
    None
        This script does not accept any input.

.OUTPUTS
    None
        This script does not return any output.

.NOTES

.EXAMPLE
C:\PS> .\Remove-AgedFiles.ps1 -Path "D:\DeltaLocal\DeltaNetLocal\errorl
og\splunk" -FileName "crossover","imswebservices","liswebservices","osb
knewton","osbknewtonadmin" -FileExtensions ".log",".txt" -MaxFileAge 27
0 -ScriptLogPath "D:\ScriptLogs"

Description
-----------
This command deletes files whose name contains "crossover", "imswebserv
ices", "liswebservices", "osbknewton", or "osbknewtonadmin", extension 
is either ".log" or ".txt" the in the "D:\DeltaLocal\DeltaNetLocal\erro
rlog\splunk" directory whose last modified date is older than 270 days.
 This will not includes read-only and hidden files.  Additional (verbos
 e) messages will be displayed in console. The script log file is creat
 ed in the "D:\ScriptLogs" directory. 

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
