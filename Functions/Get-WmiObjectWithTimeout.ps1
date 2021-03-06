#Requires -Version 2.0
####################################################################################################################
##
##	Script Name          : Get-WmiObjectWithTimeout.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2016 Bill Campbell. All rights reserved. No part of this publication may be reproduced, stored in a retrieval system, or transmitted, in any form or by means electronic, mechanical, photocopying, or otherwise, without prior written permission of the publisher.
##	Created On           : Apr 01, 2014
##	Added to Script Repo : 01 Apr 2016
##	Last Modified        : 2016.Apr.01
##
##	Version              : 2016.04.01.01
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

#region User Defined Functions
	Function Get-WmiObjectWithTimeout()
	{
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = 'Medium',
			DefaultParameterSetName = 'QueryBuilder'
		)]
		#region Function Parameters
			Param(
				[Parameter(
					Mandatory = $true
				)]
				[Alias('cn')]
				[ValidateNotNull()]
				[System.String]$ComputerName
				, [Parameter(
					Mandatory = $false
				)]
				[Alias('ns')]
				[ValidateNotNullOrEmpty()]
				[System.String]$Namespace = 'ROOT\CIMV2'
				, [Parameter(
					Mandatory = $false
					, ParameterSetName	= 'Query'
				)]
				[System.String]$Query
				, [Parameter(
					Mandatory = $true
					, ParameterSetName	= 'QueryBuilder'
				)]
				[ValidateNotNull()]
				[System.String]$Class
				, [Parameter(
					Mandatory = $false
					, ParameterSetName	= 'QueryBuilder'
				)]
				[ValidateNotNull()]
				[System.String[]]$Property = @('*')
				, [Parameter(
					Mandatory = $false
					, ParameterSetName	= 'QueryBuilder'
				)]
				[System.String]$Filter
				, [Parameter(
					Mandatory = $false
				)]
				[Alias('timeout')]
				[ValidateRange(1, 3600)]
				[Int]$TimeoutSeconds = 600
				, [Parameter(
					Mandatory = $false
				)]
				[ValidateNotNullOrEmpty()]
				[System.String]$WmiUserName
				, [Parameter(
					Mandatory = $false
				)]
				[ValidateNotNullOrEmpty()]
				[System.Security.SecureString]$WmiSecPasswd
				, [Parameter(
					Mandatory = $false
				)]
				[ValidateSet('Call', 'Connect', 'Default', 'None', 'Packet', 'PacketIntegrity', 'PacketPrivacy', 'Unchanged')]
				[System.String]$WmiAuthentication = 'Unchanged'
			)
		#endregion Function Parameters
		## Function Usage:
		##	Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class '' -Property '*' -Filter '' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd
		##	Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class '' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd
		##	Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Query "Full WQL Query String" -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd

		begin
		{
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tStart Function: $($MyInvocation.InvocationName)"

			if ($Class)
			{
				[string]$WmiErrorVariableName = $Class + '_Error'
			}
			else
			{
				[string]$WmiErrorVariableName = 'GWmiT_Query_Error'
			}

			$WmiSearcher = [WMISearcher]''

			if (!$Query)
			{
				$Query = 'SELECT ' + [String]::Join(', ', $Property) + ' FROM ' + $Class

				if ($Filter)
				{
					$Query = $Query + ' WHERE ' + $Filter
				}
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tRunning WMI Query         : $($Query)"

			## The shortest known NETBIOS Domain Name at Pearson is 3 characters, Plus the \ and at least one character for the username is 5 characters.
			## The shortest NETBIOS Domain Name could be 1 character, Plus the \ and at least one character for the username is 3 characters.
			if ($WmiUserName.Length -ge 3)
			{
				##	$WmiSearcher.Scope.Options | Select-Object *
				##
				##	Locale           :
				##	Username         :
				##	Password         :
				##	SecurePassword   :
				##	Authority        :
				##	Impersonation    : Impersonate
				##	Authentication   : Unchanged
				##	EnablePrivileges : False
				##	Context          : {}
				##	Timeout          : 10675199.02:48:05.4775807

				$WmiSearcher.Scope.Options.Username = $WmiUserName
				$WmiSearcher.Scope.Options.SecurePassword = $WmiSecPasswd

				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`tAs User           : $($WmiUserName)"
			}

			## Should be set whether a username and password are set or not.
			$WmiSearcher.Scope.Options.Authentication = $WmiAuthentication

			$WmiSearcher.Options.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
			$WmiSearcher.Options.ReturnImmediately = $true
			#$WmiSearcher.Options.Rewindable = $false
			$WmiSearcherPath = '\\' + $ComputerName + '\' + $Namespace

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`tWith a timeout of : $($TimeoutSeconds) Seconds"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`tAgainst Path      : $($WmiSearcherPath)"
		}
		process
		{
			$WmiSearcher.Scope.Path = $WmiSearcherPath
			$WmiSearcher.Query = $Query

			try
			{
				$WmiSearcherResults = $WmiSearcher.Get()
			}
			catch
			{
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name $WmiErrorVariableName -Value $Error[0].Exception.Message
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()): [ERROR] - [ERROR] - [ERROR] - [ERROR] : $($Error[0].Exception.Message)"
				$WmiSearcherResults = 'NoWmiResults'
			}
		}
		end
		{
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tEnd Function: $($MyInvocation.InvocationName)"
			$WmiSearcher.Dispose()
			return $WmiSearcherResults
		}
	}
#endregion User Defined Functions

$SQLServices = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd -Filter "Name LIKE 'mssql%' AND PathName LIKE '%sqlservr%'")
