#Requires -Version 2.0
###################################################################################################################
#	Script Name          : Report-InfoServerLicensing-UserCredentialEncryption.ps1
#	Written by           : Bill Campbell
#	Copyright			 : © 2014. All rights reserved.
#	Written on           : Aug 24, 2015
#	Added to Script Repo : 24-Aug-2015
#	Last Modified        : 2018.Sept.21
#
#	Version              : 2018.09.21.02
#	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
#	Version Example      : If this is the 6th revision/save on December 28 2012 then RR would be "06" and the version number will be formatted as follows: 2012.12.28.06
#
#	Purpose              : Encrypt and Store User/Service Account Credentials for Re-Use.
#
#	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#	#! IMPORTANT NOTE:																							 !#
#	#!	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK FROM THE USE OF,		 !#
#	#!	OR THE RESULTS RETURNED FROM THE USE OF, THIS CODE REMAINS WITH THE USER.								 !#
#	#! 																											 !#
#	#! 						Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force							 !#
#	#! 																											 !#
#	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#	#! 																											 !#
#	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#
#	Usage Instructions:
#		1. YOU ONLY NEED TO FOLLOW THESE 6 STEPS IF YOU ARE USING THIS IN CONJUCTION WITH PUPPET OR CHEF!!!
#			1. Rename the script you need the credential for with a different extension.
#				EXAMPLE: Script 'Report-InfoServerLicensing.ps1' is renamed to 'Report-InfoServerLicensing.ps1.bak'
#
#			2. Rename this script to the original name of the script that will be using the encrypted credential.
#				This is CRITICAL as part of the ecrypting process uses the name of the script in some way. When I
#				tried it without performing this step, the password was not decrypted properly and caused failed
#				login attempts, locking the account.
#					EXAMPLE: Script 'Report-InfoServerLicensing-UserCredentialEncryption.ps1' is renamed to 'Report-InfoServerLicensing.ps1'
#
#			3. Change the information in the 'Main Script' region to match the user(s) that you want to create an encrypted password for and save this script.
#
#			4. Where ever the script is run from, is where the txt files storing the encrypted passwords will be saved.
#				EXAMPLE: C:\Scripts\Report-InfoServerLicensing
#
#			5. You may then take the contents of the encrypted password text files and put them in your configuration files or into the script directly.
#
#			6. Once you have generated your encrypted passwords, you need to reverse steps 1 and 2 so the scripts are back at their original names.
#
#			Example Usage based on the new script names as described in the instructions above:
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'Denver'
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'Boston'
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'Iowa City'
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'DCSX'
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'Denver' -ScrDebug
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'Boston' -ScrDebug
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'Iowa City' -ScrDebug
#				.\Report-InfoServerLicensing.ps1 -DataCenterCity 'DCSX' -ScrDebug
#
#		2. Example Usage based on the original script name:
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'Denver'
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'Boston'
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'Iowa City'
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'DCSX'
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'Denver' -ScrDebug
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'Boston' -ScrDebug
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'Iowa City' -ScrDebug
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'DCSX' -ScrDebug
#			.\Report-InfoServerLicensing-UserCredentialEncryption.ps1 -DataCenterCity 'ALL' -ScrDebug
#
###################################################################################################################
#region Script Variable Initialization
	#region Script Level Parameters
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'The city where the Data Centers are located for the domains to be queried. ("Denver", "Boston", "Iowa City", "DCSX", "All")'
			)]
			#[ValidateSet('Denver', 'Boston', 'Iowa City', 'DCSX', 'ALL')]
			[string]$DataCenterCity
			, [Parameter(
				Position			= 1
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'The ScrDebug switch turns on some custom code to help with debugging the script ($true/$false). Defaults to FALSE.'
			)][switch]$ScrDebug = $false
		)
	#endregion Script Level Parameters

	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"

	#[string]$Script:ScriptFullPath	= ($MyInvocation.MyCommand.Definition)
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)
	#[string]$Script:ScriptNameNoExt = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	[string]$Script:ScriptNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
	[Byte[]]$Script:Key = 211,209,124,250,220,103,226,241,248,133,153,152,129,183,215,121,147,238,253,125,236,154,144,165,190,146,228,227,170,191,218,163
	[string]$Script:FileDate = (Get-Date -Format 'yyyy.MM.dd.ddd-HHmm').ToString()
	[string]$Script:OutputXMLFile = "$($ScriptDirectory)\$($ScriptNameNoExt)_$($FileDate).txt"
	$Script:OutputXMLsb = New-Object System.Text.StringBuilder ''

	## Lookup the NETBIOS Domain name from the DNS Domain name for known Active Directory Domains.
	$Script:DNStoNBnames = @{
		## Denver
		'ecollege.net'				= 'ATHENS'			# PROD
		'ecollegeqa.net'			= 'ATHENSQA'		# Staging
		'eclgsecure.net'			= 'CAIRO'			# PROD
		'eclgsecuresc.net'			= 'CAIROSC'			# Staging
		'ecollege.int'				= 'CRETE'			# Secure Integration
		'example.com'				= 'EXAMPLE'			# Corp Office / Dev / Test
		## Boston
		'pad.examplecmg.com'		= 'PAD'				# PROD
		'wrk.pad.examplecmg.com'	= 'WRK'				# PROD & Staging
		## Iowa City
		'schoolnet.prd'				= 'SCHOOLNET'		# PROD
		'schoolnet.dct'				= 'SCHOOLNET'		# Staging
		## DCSX & PEROOT
		'dcsprod.dcsroot.local'		= 'DCSPROD'			# PROD
		'dcsutil.dcsroot.local'		= 'DCSUTIL'			# PROD
		'peroot.com'				= 'PEROOT'			# PROD
	}
#endregion Script Variable Initialization

#region User Defined Functions
	Function fn_GenerateStoredPSCredential
	{
		#region Function Parameters
			Param(
				[Parameter(
					Position			= 0
					, Mandatory			= $false
					, ValueFromPipeline	= $false
					, HelpMessage		= 'The AD DNS Domain Name of the user to be used.'
				)][string]$AdDmnName 	#= (Read-Host -Prompt "Enter the AD DNS Domain Name")
				, [Parameter(
					Position			= 1
					, Mandatory			= $false
					, ValueFromPipeline	= $false
					, HelpMessage		= 'The AD NETBIOS Domain Name of the user to be used.'
				)][string]$AdNetBiosName	# = (Read-Host -Prompt "Enter the AD NETBIOS Domain Name")
				, [Parameter(
					Position			= 2
					, Mandatory			= $false
					, ValueFromPipeline	= $false
					, HelpMessage		= 'The AD username.'
				)][string]$AdAdmUserName = (Read-Host -Prompt "Enter the AD username")
				, [Parameter(
					Position			= 3
					, Mandatory			= $false
					, ValueFromPipeline	= $false
					, HelpMessage		= 'File Name to store the encrypted password in for re-use.'
				)][string]$AdAdmCredsFileName
				, [Parameter(
					Position			= 4
					, Mandatory			= $false
					, ValueFromPipeline	= $false
					, HelpMessage		= 'The path to the location for storing the encrypted password file.'
				)][string]$AdAdmCredsFilePath
			)
		#endregion Function Parameters

		if ($AdDmnName -and !$AdNetBiosName)
		{
			#[string]$AdNetBiosName = fn_GetADDomainNetBIOSName -DomainFqdn $AdDmnName

			if ($DNStoNBnames[$AdDmnName])
			{
				[string]$AdNetBiosName = $DNStoNBnames[$AdDmnName]
			}
			else
			{
				throw "NETBIOS Domain name not found in the Hash Table 'DNStoNBnames'"
			}
		}

		[string]$AdNetBiosName = $AdNetBiosName.ToUpper()

		##STORED CREDENTIAL CODE
		if (!$AdAdmCredsFileName)
		{
			#[string]$AdAdmCredsFileName = "$($AdNetBiosName.ToLower())-$($AdAdmUserName.ToLower())_Encrypted-PowerShell-Password.txt"
			#[string]$AdAdmCredsFileName = "$($ScriptNameNoExt)_$($AdNetBiosName.ToLower())-$($AdAdmUserName.ToLower())_Encrypted.txt"
			#[string]$AdAdmCredsFileName = "$($ScriptNameNoExt)_$($AdDmnName.ToLower())-$($AdAdmUserName.ToLower())_Encrypted.txt"
			[string]$AdAdmCredsFileName = "$($AdAdmUserName.ToLower())@$($AdDmnName.ToLower())-$($ScriptNameNoExt)_Encrypted.txt"
		}

		if (!$AdAdmCredsFilePath)
		{
			## No path was specified... Using the same directory the script is in.
			[string]$CredsFileFullPath = "$($ScriptDirectory)\$($AdAdmCredsFileName)"
		}
		else
		{
			## A path was specified... Using it.
			[string]$CredsFileFullPath = "$($AdAdmCredsFilePath)\$($AdAdmCredsFileName)"
		}

		if (!(Test-Path -Path $CredsFileFullPath))
		{
			#Write-Host "Credential file not found. Enter the password for $($AdNetBiosName)\$($AdAdmUserName)" -ForegroundColor Red
			Read-Host -Prompt "Credential file not found. Enter the password for '$($AdDmnName.ToLower())' - $($AdNetBiosName)\$($AdAdmUserName)" -AsSecureString | ConvertFrom-SecureString -Key $Script:Key | Set-Content -Path $CredsFileFullPath
			$SecPasswd = Get-Content $CredsFileFullPath | ConvertTo-SecureString -Key $Script:Key #(1..32)
			$SecCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$($AdNetBiosName)\$($AdAdmUserName)", $SecPasswd
		}
		else
		{
				Write-Host 'Using the stored credential file...' -ForegroundColor Green
				$SecPasswd = Get-Content $CredsFileFullPath | ConvertTo-SecureString -Key $Script:Key #(1..32)
				$SecCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$($AdNetBiosName)\$($AdAdmUserName)", $SecPasswd
		}
		##END OF STORED CREDENTIAL CODE

		$EncPasswd = Get-Content $CredsFileFullPath

		#    <DomainUser>
		#      <DnsDomain>ecollege.net</DnsDomain>
		#      <NbDomain>ATHENS</NbDomain>
		#      <UserName>adminaccount</UserName>
		#      <EncPassword>76492d1116743f0423413b16050a5345MgB8AHUAWgA3AFgARgAyAFUARAB0AFAAWgBEAGMARgBhAFoANABzAFQANgB1AHcAPQA9AHwAMwBiADEANgA0ADEAYQBiADEAYQA1AGYANwBlAGUAOABhADAAMAA0ADgAOQAyAGEAMwAwAGEAOQBkAGIAYgBmAGQANgBlAGQAOQAzADEANABjADkAMwA0ADAANwAzADUAMQAyAGUAZABjADkAMQA2ADYAMQA2AGEANQAxAGUAMgA=</EncPassword>
		#      <IsCentralManagementCred>FALSE</IsCentralManagementCred>
		#    </DomainUser>

		#$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">')
		#$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine("			$($SvrCreatedOn)")
		$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine('    <DomainUser>')
		$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine("      <DnsDomain>$($AdDmnName)</DnsDomain>")
		$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine("      <NbDomain>$($AdNetBiosName)</NbDomain>")
		$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine("      <UserName>$($AdAdmUserName)</UserName>")
		$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine("      <EncPassword>$($EncPasswd)</EncPassword>")
		if ($AdNetBiosName -eq 'DCSUTIL')
		{
			$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine('      <IsCentralManagementCred>TRUE</IsCentralManagementCred>')
		}
		else
		{
			$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine('      <IsCentralManagementCred>FALSE</IsCentralManagementCred>')
		}
		$StringBuilderOutputXMLsbOutput = $OutputXMLsb.AppendLine('    </DomainUser>')
		return $SecCreds
	}
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	##$DomainCredential = fn_GenerateStoredPSCredential -AdDmnName 'example.com' -AdNetBiosName 'EXAMPLE' -AdAdmUserName 'adminaccount'

	if ($ScrDebug -and (($DataCenterCity -eq 'Denver') -or ($DataCenterCity -eq 'ALL')))
	{
		$Domain1Credential = fn_GenerateStoredPSCredential -AdDmnName 'ecollege.net' -AdAdmUserName 'adminaccount'
		$Domain2Credential = fn_GenerateStoredPSCredential -AdDmnName 'ecollegeqa.net' -AdAdmUserName 'adminaccount'
		$Domain3Credential = fn_GenerateStoredPSCredential -AdDmnName 'eclgsecure.net' -AdAdmUserName 'adminaccount'
		$Domain4Credential = fn_GenerateStoredPSCredential -AdDmnName 'eclgsecuresc.net' -AdAdmUserName 'adminaccount'
		$Domain5Credential = fn_GenerateStoredPSCredential -AdDmnName 'example.com' -AdAdmUserName 'adminaccount'
		$Domain11Credential = fn_GenerateStoredPSCredential -AdDmnName 'ecollege.int' -AdAdmUserName 'adminaccount'
	}
	elseif (!$ScrDebug -and ($DataCenterCity -eq 'Denver'))
	{
		$Domain1Credential = fn_GenerateStoredPSCredential -AdDmnName 'ecollege.net' -AdAdmUserName 'serviceaccount'
		$Domain2Credential = fn_GenerateStoredPSCredential -AdDmnName 'ecollegeqa.net' -AdAdmUserName 'serviceaccount'
		$Domain3Credential = fn_GenerateStoredPSCredential -AdDmnName 'eclgsecure.net' -AdAdmUserName 'serviceaccount'
		$Domain4Credential = fn_GenerateStoredPSCredential -AdDmnName 'eclgsecuresc.net' -AdAdmUserName 'serviceaccount'
		$Domain5Credential = fn_GenerateStoredPSCredential -AdDmnName 'example.com' -AdAdmUserName 'serviceaccount'
	}

	if ($ScrDebug -and (($DataCenterCity -eq 'Boston') -or ($DataCenterCity -eq 'ALL')))
	{
		$Domain6Credential = fn_GenerateStoredPSCredential -AdDmnName 'pad.examplecmg.com' -AdAdmUserName 'adminaccount'
		$Domain8Credential = fn_GenerateStoredPSCredential -AdDmnName 'wrk.pad.examplecmg.com' -AdAdmUserName 'adminaccount'
	}
	elseif (!$ScrDebug -and ($DataCenterCity -eq 'Boston'))
	{
		$Domain6Credential = fn_GenerateStoredPSCredential -AdDmnName 'pad.examplecmg.com' -AdAdmUserName 'serviceaccount'
		$Domain8Credential = fn_GenerateStoredPSCredential -AdDmnName 'wrk.pad.examplecmg.com' -AdAdmUserName 'serviceaccount'
	}

	if ($ScrDebug -and (($DataCenterCity -eq 'DCSX') -or ($DataCenterCity -eq 'ALL')))
	{
		$Domain12Credential = fn_GenerateStoredPSCredential -AdDmnName 'dcsprod.dcsroot.local' -AdAdmUserName 'adminaccount'
		$Domain13Credential = fn_GenerateStoredPSCredential -AdDmnName 'dcsutil.dcsroot.local' -AdAdmUserName 'adminaccount'
		$Domain13Credential = fn_GenerateStoredPSCredential -AdDmnName 'peroot.com' -AdAdmUserName 'adminaccount'
	}
	elseif (!$ScrDebug -and ($DataCenterCity -eq 'DCSX'))
	{
		$Domain12Credential = fn_GenerateStoredPSCredential -AdDmnName 'dcsprod.dcsroot.local' -AdAdmUserName 'serviceaccount'
		$Domain13Credential = fn_GenerateStoredPSCredential -AdDmnName 'dcsutil.dcsroot.local' -AdAdmUserName 'serviceaccount'
		$Domain13Credential = fn_GenerateStoredPSCredential -AdDmnName 'peroot.com' -AdAdmUserName 'serviceaccount'
	}

	if ($ScrDebug -and (($DataCenterCity -eq 'Iowa City') -or ($DataCenterCity -eq 'ALL')))
	{
		$Domain9Credential = fn_GenerateStoredPSCredential -AdDmnName 'schoolnet.prd' -AdAdmUserName 'adminaccount'
		$Domain10Credential = fn_GenerateStoredPSCredential -AdDmnName 'schoolnet.dct' -AdAdmUserName 'adminaccount'
	}
	elseif (!$ScrDebug -and ($DataCenterCity -eq 'Iowa City'))
	{
		$Domain9Credential = fn_GenerateStoredPSCredential -AdDmnName 'schoolnet.prd' -AdAdmUserName 'serviceaccount'
		$Domain10Credential = fn_GenerateStoredPSCredential -AdDmnName 'schoolnet.dct' -AdAdmUserName 'serviceaccount'
	}
	else
	{
		Write-Host -ForegroundColor Magenta "Unrecognized DataCenterCity - $($DataCenterCity)"

		[string]$NewAdDmnName = Read-Host -Prompt "Enter the Active Directory Fully Qualified Domain Name"
		[string]$NewAdNetBiosName = Read-Host -Prompt "Enter the Active Directory NETBIOS Domain Name"
		[string]$NewAdAdmUserName = Read-Host -Prompt "Enter an Admin level username for '$($NewAdNetBiosName.ToUpper())'"

		$NewDomainCredential = fn_GenerateStoredPSCredential -AdDmnName $NewAdDmnName -AdNetBiosName $NewAdNetBiosName -AdAdmUserName $NewAdAdmUserName

		Write-Host -ForegroundColor Green "Execute the script again to process additional domains..."
	}

	## Set the XML contents for the creds to a file.
	Set-Content -Path $OutputXMLFile -Value $OutputXMLsb.ToString()

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script
