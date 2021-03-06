#Requires -Version 2.0
####################################################################################################################
##
##	Script Name          : ConvertImage-ToBase64String.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2015 Bill Campbell. All rights reserved. No part of this publication may be reproduced, stored in a retrieval system, or transmitted, in any form or by means electronic, mechanical, photocopying, or otherwise, without prior written permission of the publisher.
##	Created On           : Feb 16, 2015
##	Added to Script Repo : 16 Feb 2015
##	Last Modified        : 2018.Sept.21
##
##	Version              : 2018.09.21.02
##	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
##	Version Example      : If this is the 6th revision/save on May 22 2013 then RR would be '06' and the version number will be formatted as follows: 2013.05.22.06
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
####################################################################################################################
[CmdletBinding(
	SupportsShouldProcess = $true,
	ConfirmImpact = 'Medium'
	#, DefaultParameterSetName = '<ParameterSetName>'
)]
#region Script Variable Initialization
	#region Script Parameters
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'The full path and name to the image file.'
			)][string]$ImagePath
			, [Parameter(
				Position			= 1
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Enables/Disables custom debugging code embedded in the script ($true/$false). Defaults to FALSE.'
			)][switch]$ScrDebug = $false
		)
	#endregion Script Parameters

	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)

	#[string]$Script:ScriptNameNoExt = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	#[string]$Script:ScriptNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
#endregion Script Variable Initialization

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	#[Byte[]]$LogoToEncode = Get-Content -Path "$($ScriptDirectory)\PowerShell-icon_1280x1024.jpg" -Encoding Byte
	#$LogoToEncodeB64 = [System.Convert]::ToBase64String($LogoToEncode)
	#$LogoToEncodeB64 | Set-Content -Path "$($ScriptDirectory)\PowerShell-icon_1280x1024-base64.txt"

	[string]$ImageNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($ImagePath)
	[Byte[]]$LogoToEncode = Get-Content -Path "$($ImagePath)" -Encoding Byte
	$LogoToEncodeB64 = [System.Convert]::ToBase64String($LogoToEncode)
	$LogoToEncodeB64 | Set-Content -Path "$($ScriptDirectory)\$($ImageNameNoExt)-base64.txt"

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
	.SYNOPSIS
		Converts an image to a Base64 encoded string for embedding in HTML and other reports.

	.DESCRIPTION
		The ConvertImage-ToBase64String script does the following:

		1. Reads the image in to a variable.
		2. Converts the binary data of the image to a Base64 string.
		3. Writes the Base64 string to a text file.

	.PARAMETER <ImagePath>
		The full path and name to the image file.

	.PARAMETER ScrDebug
		Enables/Disables custom debugging code embedded in the script ($true/$false). Defaults to FALSE.

	.EXAMPLE
		${Env:SystemDrive}\Path\To\Scripts\Directory\ConvertImage-ToBase64String.ps1 -ImagePath 'ImagePathValue'

		Description
		-----------
			Takes the image and converts it to a Base64 string.

	.INPUTS
		None
			This script does not accept any input.

	.OUTPUTS
		None
			This script does not return any output.

	.NOTES
		Author : Bill Campbell
		Version: 1.0.0.1
		Release: 2015-Feb-16

		REQUIREMENTS
			PowerShell Version 2.0

	.LINK
		about_Providers

	.LINK
		Get-Content

	.LINK
		Set-Content
#>
#endregion Help

#region Script Change Log
###################################################################################################################
#
#	EXAMPLES
#		.\ConvertImage-ToBase64String.ps1 -ImagePath 'ImagePathValue'
#		. ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Path\To\Proper\Directory\ConvertImage-ToBase64String.ps1 -ImagePath 'ImagePathValue' -ScrDebug
#		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\ConvertImage-ToBase64String.ps1 -ImagePath 'ImagePathValue'
#		. ${Env:SystemDrive}\Path\To\Scripts\Directory\ConvertImage-ToBase64String.ps1 -ImagePath 'ImagePathValue' -ScrDebug
#
#	-- Script Change Log --
#	Changes for Feb-2015
#		- Initial Script/Module Writing and Debugging.
#	Changes for Sep-2018
#		- Added parameter for the image path.
#
###################################################################################################################
#endregion Script Change Log
