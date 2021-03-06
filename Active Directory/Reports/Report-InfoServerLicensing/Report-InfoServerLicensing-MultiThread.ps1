#Requires -Version 2.0
####################################################################################################################
##
##	Old Script Name      : Report-ADServerCountForLicensing-MultiThread.ps1
##	Script Name          : Report-InfoServerLicensing-MultiThread.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2015 Bill Campbell. All rights reserved.
##	Created On           : Jun 15, 2015
##	Added to Script Repo : 15 Jun 2015
##	Last Modified        : 2018.Sept.21
##
##	Version              : 2018.09.21.02
##	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
##	Version Example      : If this is the 6th revision/save on May 22 2013 then RR would be '06' and the version number will be formatted as follows: 2013.05.22.06
##
##	Purpose              : This is the Backend Script for Multi-Threading purposes to Query with Powershell for all Windows Servers from Active Directory For Licensing purposes and create a report to be sent via email.
##	Notes                : 
##
####################################################################################################################
[CmdletBinding(
	SupportsShouldProcess = $true,
	ConfirmImpact = 'Medium'
)]
#region Script Variable Initialization
	#region Script Parameters
		Param(
			$PsoADSvr
			, [string]$WorkingDirectory
			, [string]$RptSvrName
			, $LogEntryDateFormat
			, $SubnetToDataCenter
			, $IncludeMSSQL
			, $IncludeHotFixDetails
			, $IncludeCustomAppInfo
			, $IncludeInfraDetails
			, $ScrDebug
		)
	#endregion Script Parameters

	[string]$ScrStartTime = (Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString()
	[string]$LogDirDateName = (Get-Date -Format 'yyyy.MM.dd.ddd').ToString()
	[string]$LogFileDate = (Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()
	[string]$Ext = 'log'

	if ($ScrDebug)
	{
		[string]$ScrLogDir = "$($WorkingDirectory)\Logs\ScrDebug\$($LogDirDateName)"
	}
	else
	{
		[string]$ScrLogDir = "$($WorkingDirectory)\Logs\$($LogDirDateName)"
	}

	if (!(Test-Path -Path "$($ScrLogDir)" -PathType Container))
	{
		New-Item -ItemType Directory "$($ScrLogDir)" | Out-Null
	}

	[string]$LogFileName = "$($LogFileDate)_$($PsoADSvr.SrvrDnsName).$($Ext)"
	[string]$Script:MultiThreadScriptLogFile = "$($ScrLogDir)\$($LogFileName)"

	if ($IncludeMSSQL)
	{
		$SQLServices = New-Object System.Collections.Arraylist
	}

	[string]$SvInstOsArch = 'Unknown'
	[int64]$PhysMemory = 0
	[bool]$ProcessWmi = $false
	[bool]$LocalWmi = $false
	[bool]$IPV6only = $false
	[bool]$IPV4only = $true
#endregion Script Variable Initialization

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

#region Main Multi-Threaded Execution
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t--------------------------------------------------------------------------------"
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t - Multi-Threaded Backend-Script -"
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Script Start Time : $([char]34)$($ScrStartTime)$([char]34)"
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t--------------------------------------------------------------------------------"

	$SrvName		= $PsoADSvr.SrvrName
	$SrvDnsName		= $PsoADSvr.SrvrDnsName
	$SrvOsName		= $PsoADSvr.SrvrOsName
	$SrvOsSpLvl		= $PsoADSvr.SrvrOsSpLvl
	$SrvChangedOn	= $PsoADSvr.SrvrChangedOn
	$SrvCreatedOn	= $PsoADSvr.SrvrCreatedOn
	$SrvDnsDomain	= $PsoADSvr.SrvrDnsDomain
	$SrvNBDomain	= $PsoADSvr.SrvrNBDomain
	$SrvSecUsrNam	= $PsoADSvr.SecUsrNam
	$SrvSecPasswd	= $PsoADSvr.SecPasswd
	$NbDmnUsrLgn	= $PsoADSvr.ConnectAcct

	if ($SrvDnsName -eq $RptSvrName)
	{
		[bool]$LocalWmi = $true
	}

	$UserCreds = New-Object System.Management.Automation.PSCredential ($NbDmnUsrLgn, $SrvSecPasswd)

	if ($ScrDebug)
	{
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Names and Values as passed into the backend script for processing..."
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t--------------------------------------------------------------------------------"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Name        : Value"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t--------------------- : --------------------------------------------------------"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvName              : $($SrvName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDnsName           : $($SrvDnsName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t RptSvrName           : $($RptSvrName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvOsName            : $($SrvOsName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvOsSpLvl           : $($SrvOsSpLvl)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvChangedOn         : $($SrvChangedOn)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvCreatedOn         : $($SrvCreatedOn)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvNBDomain          : $($SrvNBDomain)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvSecUsrNam         : $($SrvSecUsrNam)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvSecPasswd         : $($SrvSecPasswd)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t NbDmnUsrLgn          : $($NbDmnUsrLgn)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t--------------------------------------------------------------------------------"
	}

	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tChecking to see if server $([char]34)$($SrvDnsName)$([char]34) is responding..."

	if (Test-Connection -ComputerName $SrvDnsName -Count 1 -Quiet)
	{
		[string]$SvrResponse = 'Green'
	}
	else
	{
		[string]$SvrResponse = 'Red'
	}

	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tThe server response check returned $([char]34)$($SvrResponse)$([char]34)..."
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"

	if ($SvrResponse -eq 'Green')
	{
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tRetrieving IP Address info for $([char]34)$($SrvDnsName)$([char]34) from DNS..."

		$ProcessWmi = $true

		$SvIPv4Addr = @(
			([System.Net.Dns]::GetHostEntry($SrvDnsName)).AddressList | ForEach-Object {
				if ($IPV6only)
				{
					if ($_.AddressFamily -eq 'InterNetworkV6')
					{
						$_.IPAddressToString
					}
				}

				if ($IPV4only)
				{
					if ($_.AddressFamily -eq 'InterNetwork')
					{
						$_.IPAddressToString
					}
				}

				if (!($IPV6only -or $IPV4only))
				{
					$_.IPAddressToString
				}
			}
		)

		$IpAddress = $SvIPv4Addr[0]
		$SvIPv4Addr = [String]::Join(', ', $SvIPv4Addr)
	}
	else
	{
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tNOT Retrieving IP Address info for $([char]34)$($SrvDnsName)$([char]34)..."

		$ProcessWmi = $false
		## This is the HTML code for a space character.
		## They are added here for the HTML report and are changed in the main script for the CSV and XLSX reports.
		[string]$IpAddress = '&nbsp;'
		[string]$SvIPv4Addr = '&nbsp;'
	}

	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"

	if ($ProcessWmi)
	{
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`tStart gathering information for server: $([char]34)$($SrvDnsName)$([char]34)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"

		#region DataCenter
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tDetermining DataCenter Location Codes for server: $([char]34)$($SrvDnsName)$([char]34)"

			foreach ($DcSubnet in $SubnetToDataCenter)
			{
				#if ($IpAddress -match $DcSubnet.Subnet)	## Regex comparison - Returns more false positives. Returns the first one that similarly matches instead of the closest match.
				if ($IpAddress -like $DcSubnet.Subnet)		## Wildcard comparison - Returns more positive matches than false ones. Returns the closest match.
				{
					[string]$SvDataCenter = $DcSubnet.LocationCodes
					if ($DcSubnet.BMCLocationCode -eq 'DN3')
					{
						if ($DcSubnet.Subnet -eq '159.182.160.*')
						{
							[string]$SvBMCLocationCode = "$($DcSubnet.BMCLocationCode)_PROD"
						}
						elseif ($DcSubnet.Subnet -eq '10.200.*')
						{
							[string]$SvBMCLocationCode = "$($DcSubnet.BMCLocationCode)_PROD"
						}
						elseif ($DcSubnet.Subnet -eq '10.52.*')
						{
							[string]$SvBMCLocationCode = "$($DcSubnet.BMCLocationCode)_STAGE"
						}
						else
						{
							[string]$SvBMCLocationCode = $DcSubnet.BMCLocationCode
						}
					}
					else
					{
						[string]$SvBMCLocationCode = $DcSubnet.BMCLocationCode
					}
				}
			}

			if (!$SvDataCenter)
			{
				[string]$SvDataCenter = 'Unknown'
			}

			if ($SvBMCLocationCode.Length -lt 3)
			{
				[string]$SvBMCLocationCode = 'Unknown'
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion DataCenter

		#region Win32_BIOS
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_BIOS$([char]34)"

			#$Win32BiosQuery = 'SELECT * FROM Win32_BIOS'
			if ($LocalWmi)
			{
				#$Win32Bios = Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_BIOS_Error -Query $Win32BiosQuery
				$Win32Bios = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_BIOS' -Property '*' -TimeoutSeconds 300
			}
			else
			{
				#$Win32Bios = Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_BIOS_Error -Query $Win32BiosQuery
				$Win32Bios = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_BIOS' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd
			}

			if ($Win32_BIOS_Error)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_BIOS Query Error was: $([char]34)$($Win32_BIOS_Error[0].Exception.Message)$([char]34)"

				[string]$SvSerNum = 'Error Encountered'
				[string]$SvBiosVer = 'Error Encountered'
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_BIOS query did not return any Errors."

				[string]$SvSerNum = $Win32Bios.SerialNumber
				$SvSerNum = $SvSerNum.Trim()
				[string]$SvBiosVer = ($Win32Bios.SMBIOSBIOSVersion).Trim()
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Win32_BIOS

		#region Win32_ComputerSystem
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_ComputerSystem$([char]34)"

			#$Win32CompSysQuery = 'SELECT * FROM Win32_ComputerSystem'
			if ($LocalWmi)
			{
				#$Win32CompSys = Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_ComputerSystem_Error -Query $Win32CompSysQuery
				$Win32CompSys = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_ComputerSystem' -Property '*' -TimeoutSeconds 300
			}
			else
			{
				#$Win32CompSys = Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_ComputerSystem_Error -Query $Win32CompSysQuery
				$Win32CompSys = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_ComputerSystem' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd
			}

			if ($Win32_ComputerSystem_Error)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_ComputerSystem Query Error was: $([char]34)$($Win32_ComputerSystem_Error[0].Exception.Message)$([char]34)"

				[string]$SvMnfctr		= 'Error Encountered'
				[string]$SvModel		= 'Error Encountered'
				[string]$SvInstOsArch	= 'Error Encountered'
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_ComputerSystem query did not return any Errors."

				[string]$SvMnfctr = $Win32CompSys.Manufacturer
				[string]$SvModel = $Win32CompSys.Model

				if ($Win32CompSys.SystemType -eq 'x64-based PC')
				{
					[string]$SvInstOsArch = '64-bit'
				}
				elseif ($Win32CompSys.SystemType -eq 'X86-based PC')
				{
					[string]$SvInstOsArch = '32-bit'
				}
				else
				{
					[string]$SvInstOsArch = 'Unknown'
				}
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Win32_ComputerSystem

		#region Win32_NetworkAdapterConfiguration
#			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_NetworkAdapterConfiguration$([char]34)"
#
#			#$NetworkAdapterConfigs = Get-WmiObject -Class 'Win32_NetworkAdapterConfiguration' -ComputerName ${Env:COMPUTERNAME} | Where-Object { $_.IpEnabled -eq 'True' }
#
#			#$Win32NetworkAdapterConfigurationQuery = 'SELECT * FROM Win32_NetworkAdapterConfiguration'
#			if ($LocalWmi)
#			{
#				#$Win32NetworkAdapterConfiguration = Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_NetworkAdapterConfiguration_Error -Query $Win32NetworkAdapterConfigurationQuery
#				$Win32NetworkAdapterConfiguration = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_NetworkAdapterConfiguration' -Property '*' -TimeoutSeconds 300 -Filter "IpEnabled = 'True'"
#			}
#			else
#			{
#				#$Win32NetworkAdapterConfiguration = Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_NetworkAdapterConfiguration_Error -Query $Win32NetworkAdapterConfigurationQuery
#				$Win32NetworkAdapterConfiguration = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_NetworkAdapterConfiguration' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd -Filter "IpEnabled = 'True'"
#			}
#
#			if ($Win32_NetworkAdapterConfiguration_Error)
#			{
#				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_NetworkAdapterConfiguration Query Error was: $([char]34)$($Win32_NetworkAdapterConfiguration_Error[0].Exception.Message)$([char]34)"
#
#				[string]$SvMACAddress = 'Error Encountered'
#			}
#			else
#			{
#				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_NetworkAdapterConfiguration query did not return any Errors."
#
#				#[string]$SvMACAddress = $Win32NetworkAdapterConfiguration.MACAddress
#				[string]$SvMACAddress = [String]::Join(', ', $Win32NetworkAdapterConfiguration.MACAddress)
#				$SvMACAddress = $SvMACAddress.Trim()
#			}
#
#			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Win32_NetworkAdapterConfiguration

		#region Win32_OperatingSystem
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_OperatingSystem$([char]34)"

			#$Win32OpSysQuery = 'SELECT * FROM Win32_OperatingSystem'
			if ($LocalWmi)
			{
				#$Win32OpSys = Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_OperatingSystem_Error -Query $Win32OpSysQuery
				$Win32OpSys = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_OperatingSystem' -Property '*' -TimeoutSeconds 300
			}
			else
			{
				#$Win32OpSys = Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_OperatingSystem_Error -Query $Win32OpSysQuery
				$Win32OpSys = Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_OperatingSystem' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd
			}

			if ($Win32_OperatingSystem_Error)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_OperatingSystem Query Error was: $([char]34)$($Win32_OperatingSystem_Error[0].Exception.Message)$([char]34)"

				[string]$SvLastBootUpTime = 'Error Encountered'
				[string]$SvUpTime = 'Error Encountered'
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_OperatingSystem query did not return any Errors."

				[DateTime]$ServerLastBootUpTime = $Win32OpSys.ConvertToDateTime($Win32OpSys.LastBootUpTime)
				[TimeSpan]$ServerUpTime = New-TimeSpan -Start $ServerLastBootUpTime -End $(Get-Date)
				[string]$SvLastBootUpTime = $ServerLastBootUpTime.ToString()
				[int]$SvUpTime = $ServerUpTime.Days
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Win32_OperatingSystem

		if ($IncludeHotFixDetails)
		{
			#region Win32_QuickFixEngineering
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_QuickFixEngineering$([char]34) checking for the Uptime bug hotfixes."

				#$Win32QFEQuery = 'SELECT * FROM Win32_QuickFixEngineering'
				if ($LocalWmi)
				{
					#$Win32QFE = Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_QuickFixEngineering_Error -Query $Win32QFEQuery
					#$Win32QFE = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_QuickFixEngineering' -Property 'HotFixID' -TimeoutSeconds 300)
					$Win32QFE = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_QuickFixEngineering' -Property '*' -TimeoutSeconds 300)
				}
				else
				{
					#$Win32QFE = Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_QuickFixEngineering_Error -Query $Win32QFEQuery
					#$Win32QFE = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_QuickFixEngineering' -Property 'HotFixID' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd)
					$Win32QFE = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_QuickFixEngineering' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd)
				}

				if ($Win32_QuickFixEngineering_Error)
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_QuickFixEngineering Query Error was: $([char]34)$($Win32_QuickFixEngineering_Error[0].Exception.Message)$([char]34)"

					[string]$SvUpTimeKB2553549HotFix1Applied = 'Error Encountered'
					[string]$SvUpTimeKB2688338HotFix2Applied = 'Error Encountered'
					[string]$SvUpTimeTotalHotFixApplied = 'Error Encountered'
					[string]$SvKB3042553HotFixApplied = 'Error Encountered'
					[string]$SvLastPatchInstallDate = 'Error Encountered'
				}
				else
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_QuickFixEngineering query did not return any Errors."

					#[string]$SvLastPatchInstallDate = 'Not Available'	#$SrvrCreatedOn.ToShortDateString()
					## Set the following to 'No' initially or another default value, then they can be changed later.
					[string]$SvUpTimeKB2553549HotFix1Applied = 'No'
					[string]$SvUpTimeKB2688338HotFix2Applied = 'No'
					[string]$SvUpTimeTotalHotFixApplied = 'No'
					[string]$SvKB3042553HotFixApplied = 'No'

					#region Get Date of Last Patch Installed
						if ($SrvOsName -like "*2000*")
						{
							## The WMI class needed is not available on Windows 2000 operating systems.
							[string]$SvLastPatchInstallDate = 'Not Available'
						}
						else
						{
							#####################################################################
							## If this does not work, copy the original from the backup files. ##
							#####################################################################
							##
							##$InstalledUpdates = $Win32QFE | Select-Object HotFixID, InstalledOn | Sort-Object -Property InstalledOn -Descending
							#$LastInstalledUpdate = $Win32QFE | Select-Object HotFixID, InstalledOn | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1

							if (($SvrOsName -like "*2008 St*") -or ($SvrOsName -like "*2008 En*") -or ($SvrOsName -like "*Web*2008"))
							{
								$LastInstalledUpdate = $Win32QFE | Select-Object HotFixID, InstalledOn | % {
									if ($_.InstalledOn -like "*/*/*")
									{
										$DateIs = [DateTime]::ParseExact($_.InstalledOn, '%M/%d/yyyy', [Globalization.CultureInfo]::GetCultureInfo("en-US").DateTimeFormat)
									}
									elseif (($_.InstalledOn -eq $null) -or ($_.InstalledOn -eq ''))
									{
										## Creates a value formated as follows:
										##	Monday, January 1, 0001 12:00:00 AM
										$DateIs = [DateTime]1700
									}
									else
									{
										$DateIs = Get-Date ([DateTime][Convert]::ToInt64("$($_.InstalledOn)", 16)) -Format '%M/%d/yyyy'
									}

									$_.InstalledOn = $DateIs	#([DateTime][Convert]::ToInt64($_.InstalledOn, 16))	#.AddYears(1600).ToShortDateString()
								} | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1

								## For some unknown reason on Windows 2008 Non-R2 operating systems, this DateTime does not convert the way it does on other opperating systems.
								[DateTime]$tmpSvLastPatchInstallDate = ([DateTime][Convert]::ToInt64($LastInstalledUpdate.InstalledOn, 16))

								if ($tmpSvLastPatchInstallDate.Year -lt 2000)
								{
									## On some instances I have found that you have to add 1600 years to the date for it to display correctly.
									[string]$SvLastPatchInstallDate = $tmpSvLastPatchInstallDate.AddYears(1600).ToShortDateString()
								}
								else
								{
									[string]$SvLastPatchInstallDate = $tmpSvLastPatchInstallDate.ToShortDateString()
								}
							}
							else
							{
								$LastInstalledUpdate = $Win32QFE | Select-Object HotFixID, InstalledOn | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1
								[string]$SvLastPatchInstallDate = $LastInstalledUpdate.InstalledOn.ToShortDateString()
							}

							## 8 Characters is used because the shortest date string format returned should not be shorter than this:
							##		M/D/YYYY
							##		('M/D/YYYY').Length
							if ($SvLastPatchInstallDate.Length -lt 8)
							{
								[string]$SvLastPatchInstallDate = 'Not Available'
							}
						}
					#endregion Get Date of Last Patch Installed

					## Thinking of switching to putting the date installed in the field if it has been installed.
					foreach ($HFix in $Win32QFE)
					{
						if ($HFix.HotFixID -eq 'KB2553549')
						{
							## Looking for 'KB2553549'
							## Uptime bug HotFix1 has been applied...
							[string]$SvUpTimeKB2553549HotFix1Applied = 'Yes'
							#[string]$SvUpTimeKB2553549HotFix1AppliedDate = $HFix.InstalledOn.ToShortDateString()
						}

						if ($HFix.HotFixID -eq 'KB2688338')
						{
							## Looking for 'KB2688338' which includes 'KB2553549' for Server 2008 R2, but not for Server 2008.
							## Uptime bug HotFix2 has been applied...
							[string]$SvUpTimeKB2688338HotFix2Applied = 'Yes'
							#[string]$SvUpTimeKB2688338HotFix2AppliedDate = $HFix.InstalledOn.ToShortDateString()
						}

						if ($HFix.HotFixID -eq 'KB3042553')
						{
							## Looking for 'KB3042553' on Server 2008 R2 and above.
							## MS15-034 - Zero Day Vulnerability in HTTP.sys on Windows Server 2008 R2 and newer.
							[string]$SvKB3042553HotFixApplied = 'Yes'
							#[string]$SvKB3042553HotFixAppliedDate = $HFix.InstalledOn.ToShortDateString()
						}
					}

					#region Up Time Bug HotFix Check
						if ($SrvOsName -like "*2008*")
						{
							if (($SrvOsName -like "*2008 R2*") -and ($SvUpTimeKB2688338HotFix2Applied -eq 'Yes') -and ($SvUpTimeKB2553549HotFix1Applied -eq 'No'))
							{
								## Up Time Bug HotFix1 "KB2553549" is included in HotFix2 "KB2688338" for Windows Server 2008 R2 ONLY...
								[string]$SvUpTimeKB2553549HotFix1Applied = 'Yes'
								## Setting the HotFix1 applied date to match that of the HotFix2 applied date.
								#[string]$SvUpTimeKB2553549HotFix1AppliedDate = $SvUpTimeKB2688338HotFix2AppliedDate
							}

							if (($SvUpTimeKB2553549HotFix1Applied -eq 'Yes') -and ($SvUpTimeKB2688338HotFix2Applied -eq 'Yes'))
							{
								## Both HotFixes Applied... 'KB2553549' and 'KB2688338'
								[string]$SvUpTimeTotalHotFixApplied = 'Yes'
							}

							if (($SvUpTimeKB2553549HotFix1Applied -eq 'Yes') -and ($SvUpTimeKB2688338HotFix2Applied -eq 'No'))
							{
								## Only HotFix1 "KB2553549" Applied... Which technically fixes the Up Time bug on Server 2008 (Non R2).
								[string]$SvUpTimeTotalHotFixApplied = 'Maybe'
							}

							if (($SvUpTimeKB2553549HotFix1Applied -eq 'No') -and ($SvUpTimeKB2688338HotFix2Applied -eq 'Yes'))
							{
								## Only HotFix2 "KB2688338" Applied...
								[string]$SvUpTimeTotalHotFixApplied = 'No'
							}

							if (!$SvUpTimeTotalHotFixApplied)
							{
								## Neither of the HotFixes Applied...
								[string]$SvUpTimeTotalHotFixApplied = 'No'
							}
						}
						else
						{
							## Patches are not applicable...
							[string]$SvUpTimeKB2553549HotFix1Applied = 'N/A'
							#[string]$SvUpTimeKB2553549HotFix1AppliedDate = 'Not Applicable'
							[string]$SvUpTimeKB2688338HotFix2Applied = 'N/A'
							#[string]$SvUpTimeKB2688338HotFix2AppliedDate = 'Not Applicable'
							[string]$SvUpTimeTotalHotFixApplied = 'N/A'
							#[string]$SvKB3042553HotFixAppliedDate = 'Not Applicable'
						}
					#endregion Up Time Bug HotFix Check

					#region MS15-034 HotFix Check - KB3042553
						if (($SrvOsName -notlike "*2008 R2*") -and ($SrvOsName -notlike "*2012*"))
						{
							## Vulnerability only applies to Server 2008 R2 and above. Setting other server types to not applicable.
							[string]$SvKB3042553HotFixApplied = 'N/A'
						}
					#endregion MS15-034 HotFix Check - KB3042553
				}

				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
			#endregion Win32_QuickFixEngineering
		}

		#region MSCluster_Cluster
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)MSCluster_Cluster$([char]34)"

			[string]$MSClusterName = $null

			if ($LocalWmi)
			{
				#$MSCluster = Get-WmiObject -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -ComputerName $SrvDnsName -Authentication PacketPrivacy -ErrorVariable MSCluster_Cluster_Error
				$MSCluster = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -TimeoutSeconds 300 -WmiAuthentication 'PacketPrivacy')
			}
			else
			{
				#$MSCluster = Get-WmiObject -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -ComputerName $SrvDnsName -Credential $UserCreds -Authentication PacketPrivacy -ErrorVariable MSCluster_Cluster_Error
				$MSCluster = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd -WmiAuthentication 'PacketPrivacy')
			}

			if (((($MSCluster | Select-Object Name).Name).Length -lt 1) -or ($MSCluster_Cluster_Error -like '*Invalid namespace*') -or ($Error[0].Exception.Message -like '*no more endpoints available*'))	# -or ($MSCluster_Cluster_Error -like '*no more endpoints available*'))
			{
				[string]$MSClusterName = 'Not Clustered'
			}
			#else
			if ($MSCluster_Cluster_Error)
			{
				[string]$MSClusterName = 'Error Encountered'
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tMSCluster_Cluster MSCluster_Cluster_Error: $([char]34)$($MSCluster_Cluster_Error)$([char]34)"
			}
			elseif ($GWmiT_Query_Error)
			{
				[string]$MSClusterName = 'Error Encountered'
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tMSCluster_Cluster GWmiT_Query_Error: $([char]34)$($GWmiT_Query_Error)$([char]34)"
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tMSCluster_Cluster query did not return any Errors."
				[string]$MSClusterName = (($MSCluster | Select-Object Name).Name).ToLower()
			}

			if ($ScrDebug)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tMSCluster_Cluster_Error variable content: $([char]34)$($MSCluster_Cluster_Error)$([char]34)"
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tGWmiT_Query_Error variable content: $([char]34)$($GWmiT_Query_Error)$([char]34)"
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tMSClusterName variable contains: $([char]34)$($MSClusterName)$([char]34)"
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion MSCluster_Cluster

		#region StdRegProv-Remote_Registry_Query
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)StdRegProv-Remote_Registry_Query$([char]34)"
			if ($LocalWmi)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_Service$([char]34) Looking for MS SQL Server Services."
				#$SQLServices = Get-WmiObject -Namespace root\CIMV2 -ComputerName $SrvDnsName -Class Win32_Service -Filter {Name LIKE 'mssql%' AND PathName LIKE '%sqlservr%'} #-Property Name, PathName, StartMode, ProcessId, State, StartName, Description, DisplayName, Caption
				$SQLServices = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300 -Filter "Name LIKE 'mssql%' AND PathName LIKE '%sqlservr%'")

				#Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)ROOT\DEFAULT\StdRegProv$([char]34)"
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)ROOT\CIMV2\StdRegProv$([char]34)"
				##$StdReg = Get-WmiObject -List -Namespace ROOT\DEFAULT -ComputerName $SrvDnsName -ErrorVariable StdRegProv_Error | Where-Object { $_.Name -eq 'StdRegProv' }
				$StdReg = Get-WmiObject -List -Namespace ROOT\DEFAULT -ComputerName $SrvDnsName -ErrorVariable StdRegProv_Error -Class StdRegProv #-Authentication PacketPrivacy
				#$StdReg = Get-WmiObject -List -Namespace ROOT\CIMV2 -ComputerName $SrvDnsName -ErrorVariable StdRegProv_Error -Class StdRegProv #-Authentication PacketPrivacy
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_Service$([char]34) Looking for MS SQL Server Services."
				#$SQLServices = Get-WmiObject -Namespace ROOT\CIMV2 -ComputerName $SrvDnsName -Credential $UserCreds -Class Win32_Service -Filter {Name LIKE 'mssql%' AND PathName LIKE '%sqlservr%'} #-Property Name, PathName, StartMode, ProcessId, State, StartName, Description, DisplayName, Caption
				$SQLServices = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd -Filter "Name LIKE 'mssql%' AND PathName LIKE '%sqlservr%'")

				#Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)ROOT\DEFAULT\StdRegProv$([char]34)"
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)ROOT\CIMV2\StdRegProv$([char]34)"
				##$StdReg = Get-WmiObject -List -Namespace ROOT\DEFAULT -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable StdRegProv_Error | Where-Object { $_.Name -eq 'StdRegProv' }
				$StdReg = Get-WmiObject -List -Namespace ROOT\DEFAULT -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable StdRegProv_Error -Class StdRegProv #-Authentication PacketPrivacy
				#$StdReg = Get-WmiObject -List -Namespace ROOT\CIMV2 -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable StdRegProv_Error -Class StdRegProv #-Authentication PacketPrivacy
			}

			## http://tfl09.blogspot.com/2011/09/using-powershell-and-wmi-to-manage.html
			## The well known values for Registry Hives are as follows:
			##	HKEY_CLASSES_ROOT =		'2147483648'
			##	HKEY_CURRENT_USER =		'2147483649'
			##	HKEY_LOCAL_MACHINE =	'2147483650'
			##	HKEY_USERS =			'2147483651'
			##	HKEY_CURRENT_CONFIG =	'2147483653'
			##	HKEY_DYN_DATA =			'2147483654'
			$HKRegHive = 2147483650

			#region OS Name and SP Level
				if ($StdRegProv_Error)
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tStdRegProv Query Error was: $([char]34)$($StdRegProv_Error[0].Exception.Message)$([char]34)"
					#if ($ScrDebug)
					#{
					#	$WmiErrors += "`t`t`t`t`t`tStdRegProv Query Error was: $([char]34)$($StdRegProv_Error[0].Exception.Message)$([char]34)$([char]13)$([char]10)"
					#}

					[string]$SvRegOsName = 'Error Encountered'
					[string]$SvRegSpLvl = 'Error Encountered'
				}
				else
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tStdRegProv query did not return any Errors."
					#if ($ScrDebug)
					#{
					#	$WmiErrors += "`t`t`t`t`t`tStdRegProv query did not return any Errors.$([char]13)$([char]10)"
					#}

					$RegKey = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
					[string]$SvRegOsName = $StdReg.GetStringValue($HKRegHive, $RegKey, 'ProductName').sValue
					if (!$?)
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SvRegOsName$([char]34)."
					}

					[string]$SvRegSpLvl = $StdReg.GetStringValue($HKRegHive, $RegKey, 'CSDVersion').sValue
					if (!$?)
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SvRegSpLvl$([char]34)."
					}

					if ($SvRegOsName -like "*Windows*")
					{
						#$SvRegOsName = ($SvRegOsName).Replace(' (R)', '').Replace('(R)', '').Replace('®', '').Replace('?', '').Replace(',', '').Replace('Microsoft ', '').Replace('Windows ', '').Replace('Standard', 'Std').Replace('Enterprise', 'Ent').Replace('Datacenter', 'DC').Replace('Edition', '')	#.Replace('Server', 'Svr')
						$SvRegOsName = ($SvRegOsName).Replace(' (R)', '').Replace('(R)', '').Replace('®', '').Replace('?', '').Replace(',', '').Replace('Microsoft ', '').Replace('Windows ', '').Replace('Edition', '').Replace('Standard x64 ', 'Std x64').Replace('Standard ', 'Std').Replace('Standard', 'Std').Replace('Enterprise x64 ', 'Ent x64').Replace('Enterprise ', 'Ent').Replace('Enterprise', 'Ent').Replace('Datacenter ', 'DC').Replace('Datacenter', 'DC')
					}

					if ($SvRegOsName -eq '2000')
					{
						[string]$SvRegOsName = '2000 Server'
					}

					if ($SvRegSpLvl -like "Service Pack*")
					{
						$SvRegSpLvl = ($SvRegSpLvl).Replace('Service Pack ', 'SP')
					}
					else
					{
						$SvRegSpLvl = 'NONE'
					}
				}
			#endregion OS Name and SP Level

			if ($IncludeMSSQL)
			{
				#region SQL Server Inventory
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tIncluding Microsoft SQL Server information."
					## Setting it to false initailly.
					[bool]$SQLSvrIsInstalled = $false
					$SQLServerBaseKeys = @(
						'SOFTWARE\Microsoft\Microsoft SQL Server'
						, 'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server'
					)

					## If There is at least one service listed in the $SQLServices array and there were no errors durring the query processes, then do the following...
					if (($SQLServices.Count -ge 1) -and -not ($Win32_Service_Error -or $StdRegProv_Error))
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tMicrosoft SQL Server is installed."
						[bool]$SQLSvrIsInstalled = $true

						[string]$RegistryPath = $SQLServerBaseKeys[0]

						$SQLInstanceNamesArray = New-Object System.Collections.Arraylist		#$SQLInstanceNamesArray = @()
						$SQLInstanceNamesValues = New-Object System.Collections.Hashtable		#$SQLInstanceNamesValues = @{}
						$SQLSvcDisplayNamesArray = New-Object System.Collections.Arraylist		#$SQLSvcDisplayNamesArray = @()
						$SQLEditionsArray = New-Object System.Collections.Arraylist				#$SQLEditionsArray = @()
						$SQLVersionsArray = New-Object System.Collections.Arraylist				#$SQLVersionsArray = @()
						$SQLClusterNamesArray = New-Object System.Collections.Arraylist			#$SQLClusterNamesArray = @()
						$SQLClusterNodesArray = New-Object System.Collections.Arraylist			#$SQLClusterNodesArray = @()
						$SQLProductCaptionsArray = New-Object System.Collections.Arraylist		#$SQLProductCaptionsArray = @()

						$SQLInstalledInstances = $StdReg.GetMultiStringValue($HKRegHive, $RegistryPath, 'InstalledInstances')

						foreach ($SQLInstalledInstance in $SQLInstalledInstances)
						{
							## HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\InstalledInstances
							$SQLInstanceNamesArray += $SQLInstalledInstance.sValue
						}

						if ($SQLInstanceNamesArray.Count -ge 1)
						{
							foreach ($SQLObj in $SQLInstanceNamesArray)
							{
								$SQLObjValue = $StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\Instance Names\SQL", "$($SQLObj)")
								if (!$?)
								{
									Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SQLObjValue$([char]34)."
								}
								## HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL
								##	ValueName		= ValueValue
								##	"MSSQLSERVER"	= "MSSQL10_50.SQL550C"
								##	"I02"			= "MSSQL10_50.SQL551C"
								##	"I03"			= "MSSQL10_50.SQL552C"

								## Add each SQL Instance to the Hashtable
								$SQLInstanceNamesValues.Add($SQLObj, $SQLObjValue.sValue)
							}

							foreach ($SQLInstanceName in $SQLInstanceNamesArray)
							{
								#region SQLInstanceOptions, SQLClusterName
									## HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL.1\Cluster
									$SQLInstanceOptions = $StdReg.EnumKey($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])")

									if ($SQLInstanceOptions.sNames -contains 'Cluster')
									{
										[bool]$SQLIsClustered = $true

										#$tmpSQLClusterName = $StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])\Cluster", 'ClusterName')
										#$SQLClusterName = $tmpSQLClusterName.sValue
										[string]$SQLClusterName = ($StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])\Cluster", 'ClusterName')).sValue
										if (!$?)
										{
											Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SQLClusterName$([char]34)."
										}

										$SQLClusterNamesArray += $SQLClusterName.ToLower()

										$tmpSQLClusterNodes = $StdReg.EnumKey($HKRegHive, 'Cluster\Nodes')

										foreach ($tmpSQLClusterNode in $tmpSQLClusterNodes.sNames)
										{
											$tmpNodeName = $StdReg.GetStringValue($HKRegHive, "Cluster\Nodes\$($tmpSQLClusterNode)", 'NodeName')
											if (!$?)
											{
												Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)tmpNodeName$([char]34)."
											}
											else
											{
												$SQLClusterNodesArray.Add(($tmpNodeName.sValue).ToLower())
												#$SQLClusterNodesArray.Add(($StdReg.GetStringValue($HKRegHive, "Cluster\Nodes\$($tmpSQLClusterNode)", 'NodeName')).sValue)
											}
										}
									}
									else
									{
										[bool]$SQLIsClustered = $false
										[string]$SQLClusterName = 'IS NOT a member of a SQL Cluster!'
										$SQLClusterNamesArray += $SQLClusterName
										$SQLClusterNodesArray.Add('NONE')
									}
								#endregion SQLInstanceOptions, SQLClusterName

								#region MSClusterName - Commented out
								#	if ($SQLIsClustered)
								#	{
								#		if ($LocalWmi)
								#		{
								#			#$MSCluster = Get-WmiObject -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -ComputerName $SrvDnsName -Authentication PacketPrivacy #-TimeoutSeconds 300
								#			$MSCluster = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -TimeoutSeconds 300 -WmiAuthentication 'PacketPrivacy')
								#		}
								#		else
								#		{
								#			#$MSCluster = Get-WmiObject -ComputerName $SrvDnsName -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -Credential $UserCreds -Authentication PacketPrivacy
								#			$MSCluster = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\MSCluster' -Class 'MSCluster_Cluster' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd -WmiAuthentication 'PacketPrivacy')
								#		}
								#
								#		[string]$MSClusterName = (($MSCluster | Select-Object Name).Name).ToLower()
								#	}
								#	else
								#	{
								#		[string]$MSClusterName = 'Not Clustered'
								#	}
								#endregion MSClusterName - Commented out

								#region SQLEditionsArray
									#$tmpSQLEdition = $StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])\Setup", 'Edition')
									try
									{
										#[string]$SQLEdition = $tmpSQLEdition.sValue
										[string]$SQLEdition = ($StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])\Setup", 'Edition')).sValue
										if (!$?)
										{
											Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SQLEdition$([char]34)."
										}

										$SQLEditionsArray += $SQLEdition
									}
									catch
									{
										[string]$SQLEdition = 'Unknown Microsoft SQL Server Product Edition'
										$SQLEditionsArray += $SQLEdition
									}
								#endregion SQLEditionsArray

								#region SQLVersion
									#$tmpSQLVersion = $StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])\Setup", 'Version')
									try
									{
										#$SQLVersion = $tmpSQLVersion.sValue
										[string]$SQLVersion = ($StdReg.GetStringValue($HKRegHive, "$($RegistryPath)\$($SQLInstanceNamesValues[$SQLInstanceName])\Setup", 'Version')).sValue
										if (!$?)
										{
											Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SQLVersion$([char]34)."
										}

										$SQLVersionsArray += $SQLVersion

										$SQLFullVersion = $SQLVersion.Split('.')
										[int]$SQLMajorVersion = $SQLFullVersion[0]
										[int]$SQLMinorVersion = $SQLFullVersion[1]
										#[int]$SQLBuildVersion = $SQLFullVersion[2]
										#[int]$SQLRevisionVersion = $SQLFullVersion[3]
									}
									catch
									{
										[string]$SQLVersion = 'Unknown Microsoft SQL Server Product Version'
									}
								#endregion SQLVersion

								#region SQLProductCaptionsArray
									[string]$SQLSvcDisplayName = ($SQLServices | Where-Object { $_.DisplayName -match $SQLInstanceName } | Select-Object DisplayName).DisplayName
									$SQLSvcDisplayNamesArray += $SQLSvcDisplayName

									$SQLProductCaptionsArray += switch ($SQLMajorVersion)
									{
										12 { 'Microsoft SQL Server 2014' }

										11 { 'Microsoft SQL Server 2012' }

										10 {
											if ($SQLMinorVersion -ge 50)
											{
												'Microsoft SQL Server 2008 R2'
											}
											else
											{
												'Microsoft SQL Server 2008'
											}
										}

										9 { 'Microsoft SQL Server 2005' }

										8 { 'Microsoft SQL Server 2000' }

										7 { 'Microsoft SQL Server 7.0' }

										6 {
											if ($SQLMinorVersion -ge 50)
											{
												'Microsoft SQL Server 6.5'
											}
											else
											{
												'Microsoft SQL Server 6.0'
											}
										}

										default { 'Unrecognized Microsoft SQL Server Product' }
									}
								#endregion SQLProductCaptionsArray
							}

							#region Owner Role - Active - Passive - Not Clusterd
								if ($SQLIsClustered)
								{
									[bool]$IsActiveNodeForSQLClusterInstance = $false
									[string]$SQLClusterInstanceOwnerRole = 'Passive'

									foreach ($SQLClusterNetworkName in $SQLClusterNamesArray)
									{
										if ($IsActiveNodeForSQLClusterInstance)
										{
											## If the server is the Active Node for any instance, then there is no reason to check any of the other SQL Server instances.
											Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`t`t$([char]34)$($SrvName)$([char]34) is running as the Active node found for another SQL Server instance. No need to test the SQL Cluster Network Name: $([char]34)$($SQLClusterNetworkDnsName)$([char]34)"
											continue
										}

										[string]$SQLClusterNetworkDnsName = "$($SQLClusterNetworkName).$($SrvDnsDomain)"

										[string]$SQLClusterInstanceActiveNodeName = (Get-WmiObject -Class Win32_ComputerSystem -ComputerName $SQLClusterNetworkDnsName -Credential $UserCreds -ErrorVariable SQLClusterInstanceActiveNodeName_Error | Select-Object -ExpandProperty Name)
										#Get-WmiObjectWithTimeout -ComputerName $SQLClusterNetworkName -Namespace 'ROOT\CIMV2' -Class 'Win32_ComputerSystem' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd

										if ($SQLClusterInstanceActiveNodeName_Error)
										{
											Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tSQLClusterInstanceActiveNodeName_Error Win32_ComputerSystem Query Error was: $([char]34)$($SQLClusterInstanceActiveNodeName_Error[0].Exception.Message)$([char]34)"
											if ($ScrDebug)
											{
												Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`t`tTested SQL Cluster Network Name: $([char]34)$($SQLClusterNetworkDnsName)$([char]34)"
											}

											#$IsActiveNodeForSQLClusterInstance = $false
										}
										else
										{
											Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tSQLClusterInstanceActiveNodeName_Error Win32_ComputerSystem query did not return any Errors."
											if ($ScrDebug)
											{
												Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`t`tTested SQL Cluster Network Name: $([char]34)$($SQLClusterNetworkDnsName)$([char]34)"
											}

											if ($SQLClusterInstanceActiveNodeName -eq $SrvName)
											{
											    ## Server is the Active node for this instance...
												$IsActiveNodeForSQLClusterInstance = $true
												$SQLClusterInstanceOwnerRole = 'Active'
											}
											else
											{
											    ## Server is the Passive node for this instance...
												$IsActiveNodeForSQLClusterInstance = $false
												$SQLClusterInstanceOwnerRole = 'Passive'
											}
										}
									}
								}
								else
								{
									## Server is NOT in a SQL Cluster...
									[string]$SQLClusterInstanceOwnerRole = 'Not Clustered - Active'
								}
							#endregion Owner Role - Active - Passive - Not Clusterd

							## This may be more forgiving if errors are encountered.
							#[string]$SQLInstanceNames = ($SQLInstanceNamesArray | Select-Object -Unique) -join ', '
							#[string]$SQLProductCaptions = ($SQLProductCaptionsArray | Select-Object -Unique) -join ', '
							#[string]$SQLEditions = ($SQLEditionsArray | Select-Object -Unique) -join ', '
							#[string]$SQLVersions = ($SQLVersionsArray | Select-Object -Unique) -join ', '
							#[string]$SQLSvcDisplayNames = ($SQLSvcDisplayNamesArray | Select-Object -Unique) -join ', '
							#[string]$SQLClusterNames = ($SQLClusterNamesArray | Select-Object -Unique) -join ', '
							#[string]$SQLClusterNodes = ($SQLClusterNodesArray | Select-Object -Unique) -join ', '

							[string]$SQLInstanceNames = [String]::Join(', ', ($SQLInstanceNamesArray | Select-Object -Unique))
							[string]$SQLProductCaptions = [String]::Join(', ', ($SQLProductCaptionsArray | Select-Object -Unique))
							[string]$SQLEditions = [String]::Join(', ', ($SQLEditionsArray | Select-Object -Unique))
							[string]$SQLVersions = [String]::Join(', ', ($SQLVersionsArray | Select-Object -Unique))
							[string]$SQLSvcDisplayNames = [String]::Join(', ', ($SQLSvcDisplayNamesArray | Select-Object -Unique))
							[string]$SQLClusterNames = [String]::Join(', ', ($SQLClusterNamesArray | Select-Object -Unique))
							[string]$SQLClusterNodes = [String]::Join(', ', ($SQLClusterNodesArray | Select-Object -Unique))
						}
						else
						{
							[string]$SQLDetectionError = 'MS SQL Service(s) detected but could not retrieve any information about the instance(s).'
							[bool]$SQLSvrIsInstalled = $true
							[string]$SQLInstanceNames = $SQLDetectionError
							[string]$SQLProductCaptions = $SQLDetectionError
							[string]$SQLEditions = $SQLDetectionError
							[string]$SQLVersions = $SQLDetectionError
							[string]$SQLSvcDisplayNames = $SQLDetectionError
							[bool]$SQLIsClustered = $false
							[string]$SQLClusterNames = $SQLDetectionError
							[string]$SQLClusterNodes = $SQLDetectionError
							#[string]$MSClusterName = $SQLDetectionError
						}
					}
					elseif ($Win32_Service_Error)
					{
						[bool]$SQLSvrIsInstalled = $false
						[string]$SQLInstanceNames = 'Error Encountered'
						[string]$SQLProductCaptions = 'Error Encountered'
						[string]$SQLEditions = 'Error Encountered'
						[string]$SQLVersions = 'Error Encountered'
						[string]$SQLSvcDisplayNames = 'Error Encountered'
						[bool]$SQLIsClustered = $false
						[string]$SQLClusterNames = 'Error Encountered'
						[string]$SQLClusterNodes = 'Error Encountered'
						#[string]$MSClusterName = 'Error Encountered'
						[string]$SQLClusterInstanceOwnerRole = 'No'
					}
					else
					{
						[bool]$SQLSvrIsInstalled = $false
						[string]$SQLInstanceNames = 'NONE'
						[string]$SQLProductCaptions = 'NONE'
						[string]$SQLEditions = 'NONE'
						[string]$SQLVersions = 'NONE'
						[string]$SQLSvcDisplayNames = 'NONE'
						[bool]$SQLIsClustered = $false
						[string]$SQLClusterNames = 'NONE'
						[string]$SQLClusterNodes = 'NONE'
						#[string]$MSClusterName = 'NONE'
						[string]$SQLClusterInstanceOwnerRole = 'NONE'
					}
				#endregion SQL Server Inventory
			}

			if ($IncludeCustomAppInfo)
			{
				#region BrowserHawk Editor Software Install Check
					## Setting the initial value to 'No' and if the app is installed then it will be changed later.
					[string]$SvBrowserHawkInstall = 'No'

					##	[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\FC41734035064E24F94E6D8BFEB3DE1E]
					##	"049A69EFD5F6F634583C0F2C4C7176C3"="C?\\Program Files\\cyScape\\BrowserHawk\\BrowserHawk.exe"
					if ($StdRegProv_Error)
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tUnable to populate vairable for: $([char]34)BrowserHawk Software Check$([char]34)"
						[string]$SvAppBrowserHawk = 'Error Encountered'
					}
					else
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tStdRegProv query did not return any Errors. Checking for BrowserHawk Editor..."

						$RegKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\FC41734035064E24F94E6D8BFEB3DE1E'
						[string]$SvAppBrowserHawk = $StdReg.GetStringValue($HKRegHive, $RegKey, '049A69EFD5F6F634583C0F2C4C7176C3').sValue
						if (!$?)
						{
							Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SvAppBrowserHawk$([char]34)."
						}

						if ($SvAppBrowserHawk -like "*BrowserHawk.exe*")
						{
							[string]$SvBrowserHawkInstall = 'Yes'
						}
						else
						{
							[string]$SvBrowserHawkInstall = 'No'
						}
					}
				#endregion BrowserHawk Editor Software Install Check

				#region SoftArtisans FileUp Software Install Check
					$HKRegHive = 2147483648
					## Setting the initial value to 'No' and if the app is installed then it will be changed later.
					[string]$SvSoftArtisansFileUpInstall = 'No'

					if ($StdRegProv_Error)
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tUnable to populate vairable for: $([char]34)SoftArtisans FileUp Software Check$([char]34)"
						[string]$SvAppSoftArtisansFileUp = 'Error Encountered'
					}
					else
					{
						Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tStdRegProv query did not return any Errors. Checking for SoftArtisans FileUp..."

						#$RegKey = 'Installer\Products\388426FA47D1070428C39CA541B84A73'
						$RegKey = 'Installer\Products'

						$RegSubKeys = ($StdReg.EnumKey($HKRegHive, $RegKey)).sNames
						foreach ($RegSubKey in $RegSubKeys)
						{
							[string]$SvAppSoftArtisansFileUp = $StdReg.GetStringValue($HKRegHive, ("$($RegKey)\$($RegSubKey)"), 'ProductName').sValue
							if (!$?)
							{
								Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tCould not populate variable $([char]34)SvAppSoftArtisansFileUp$([char]34)."
							}

							if ($SvAppSoftArtisansFileUp -like "*SoftArtisans FileUp*")
							{
								[string]$SvSoftArtisansFileUpInstall = 'Yes'
							}
							#else
							#{
							#	## Variable is already set to 'No' above.
							#	[string]$SvSoftArtisansFileUpInstall = 'No'
							#}
						}
					}
				#endregion SoftArtisans FileUp Software Install Check
			}

			#region Pending Reboot
				if ($StdRegProv_Error)
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tUnable to populate vairable for: $([char]34)Pending Reboot$([char]34)"
					[string]$SvPendingReboot = 'Error Encountered'
				}
				else
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tStdRegProv query did not return any Errors. Checking for a Pending Reboot..."

					## Setting pending values to false to cut down on the number of else statements
					$SvPendingReboot = $false
					$tmpPendFileRename = $false
					$tmpWUAURebootReq = $false
					## Setting CBSRebootPend to $false since not all versions of Windows have this value
					$tmpCBSRebootPend = $false

					$RegistryPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\'
					$RegKeyName = 'RebootPending'

					if ($StdReg.GetStringValue($HKRegHive, $RegistryPath, $RegKeyName).sValue)
					{
						$tmpCBSRebootPend = $true
					}

					## Query WUAU from the registry
					$RegistryPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\'
					$RegKeyName = 'RebootRequired'

					if ($StdReg.GetStringValue($HKRegHive, $RegistryPath, $RegKeyName).sValue)
					{
						$tmpWUAURebootReq = $true
					}

					## Query PendingFileRenameOperations from the registry
					$RegistryPath = 'SYSTEM\CurrentControlSet\Control\Session Manager\'
					$RegKeyName = 'PendingFileRenameOperations'
					$RegValuePFRO = @($StdReg.GetMultiStringValue($HKRegHive, $RegistryPath, $RegKeyName))

					## If PendingFileRenameOperations has a value set $RegValuePFRO variable to $true
					if ($RegValuePFRO.sValue.Length -ge 1)
					{
						$tmpPendFileRename = $true
					}

					## If any of the variables are true, set $SvPendingReboot variable to $true
					if ($tmpCBSRebootPend -or $tmpWUAURebootReq -or $tmpPendFileRename)
					{
						$SvPendingReboot = $true
					}
				}
			#endregion Pending Reboot

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion StdRegProv-Remote_Registry_Query

		#region Win32_PhysicalMemory
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_PhysicalMemory$([char]34)"

			#$Win32PhysMemQuery = 'SELECT * FROM Win32_PhysicalMemory'
			if ($LocalWmi)
			{
				#$Win32PhysMem = @(Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_PhysicalMemory_Error -Query $Win32PhysMemQuery)
				$Win32PhysMem = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_PhysicalMemory' -Property '*' -TimeoutSeconds 300)
			}
			else
			{
				#$Win32PhysMem = @(Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_PhysicalMemory_Error -Query $Win32PhysMemQuery)
				$Win32PhysMem = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_PhysicalMemory' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd)
			}

			if ($Win32_PhysicalMemory_Error)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_PhysicalMemory Query Error was: $([char]34)$($Win32_PhysicalMemory_Error[0].Exception.Message)$([char]34)"

				$SvPhysMemoryGB = 'Error Encountered'
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_PhysicalMemory query did not return any Errors."

				foreach ($objDimm in $Win32PhysMem)
				{
					$PhysMemory += $objDimm.Capacity
				}

				## Byte Conversion Table
				## 1 * 1024 = 1024 = 1KB
				## 1 * 1024 * 1024 = 1048576 = 1MB
				## 1 * 1024 * 1024 * 1024 = 1073741824 = 1GB
				if ([Math]::Round(($PhysMemory / 1073741824), 0) -lt 1)
				{
					[float]$SvPhysMemoryGB = switch ([Math]::Round(($PhysMemory / 1048576), 0))
					{
						 64 { 0.0625 }
						128 { 0.125 }
						256 { 0.25 }
						384 { 0.375 }
						512 { 0.5 }
						640 { 0.625 }
						768 { 0.75 }
						896 { 0.875 }
					}
				}
				else
				{
					[float]$SvPhysMemoryGB = [Math]::Round(($PhysMemory / 1073741824), 0)
				}
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Win32_PhysicalMemory

		#region Win32_Processor
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_Processor$([char]34)"

			#$Win32ProcQuery = 'SELECT Manufacturer,NumberOfCores,NumberOfLogicalProcessors,SocketDesignation FROM Win32_Processor'
			#$Win32ProcQuery = 'SELECT * FROM Win32_Processor'
			if ($LocalWmi)
			{
				#$Win32Proc = @(Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_Processor_Error -Query $Win32ProcQuery)
				$Win32Proc = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Processor' -Property '*' -TimeoutSeconds 300)
			}
			else
			{
				#$Win32Proc = @(Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_Processor_Error -Query $Win32ProcQuery)
				$Win32Proc = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Processor' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd)
			}

			if ($Win32_Processor_Error)
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_Processor Query Error was: $([char]34)$($Win32_Processor_Error[0].Exception.Message)$([char]34)" # $([char]34)$($SrvDnsName)$([char]34) 

				[string]$SvProcMnftr		= 'Error Encountered'
				[string]$SvSocketCount		= 'Error Encountered'
				[string]$SvProcCoreCount	= 'Error Encountered'
				[string]$SvLogProcsCount	= 'Error Encountered'
				[string]$SvHyperT			= 'Error Encountered'

				if ($IncludeInfraDetails)
				{
					[string]$SvProcName			= 'Error Encountered'
				}
			}
			else
			{
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_Processor query did not return any Errors."

				[string]$SvProcMnftr = $Win32Proc[0].Manufacturer

				if ($IncludeInfraDetails)
				{
					[string]$SvProcName = $Win32Proc[0].Name

					if ($SvProcName.Length -ge 1)
					{
						[string]$SvProcName = ($SvProcName).Replace('           ', ' ').Replace('  ', ' ')
					}
				}

				## A better way to do this might be to check to see if the 'NumberOfCores' and 'NumberOfLogicalProcessors' exist or are null, instead of detecting the OS.
				## Server 2003 needs this HotFix installed for the 'NumberOfCores' and 'NumberOfLogicalProcessors' to show correctly.
				## http://support.microsoft.com/kb/932370
				if (($SrvOsName -like "*2008 R2*") -or ($SrvOsName -like "*2012*"))
				{
					$SvSocketCount = ($Win32Proc | Sort-Object SocketDesignation -Unique).Count
					$SvProcCoreCount = $($Win32Proc | Measure-Object NumberOfCores -Sum).Sum
					$SvLogProcsCount = $($Win32Proc | Measure-Object NumberOfLogicalProcessors -Sum).Sum
				}
				elseif ($SrvOsName -like "*2008*")
				{
					$SvSocketCount = ($Win32Proc | Sort-Object DeviceID -Unique).Count
					$SvProcCoreCount = $($Win32Proc | Measure-Object NumberOfCores -Sum).Sum
					$SvLogProcsCount = $($Win32Proc | Measure-Object NumberOfLogicalProcessors -Sum).Sum
				}
				else
				{
					$SvSocketCount = ($Win32Proc | Sort-Object SocketDesignation -Unique).Count
					$SvProcCoreCount = ($Win32Proc).Count
					$SvLogProcsCount = ($Win32Proc).Count
					## If the Total Logical Procs Divided by 'SvSocketCount' is greater than or equal to 8 the HT is Active for Intel.
					## Need to see about getting the following HotFix applied to ALL 2003 Server instances.
					##	http://support.microsoft.com/kb/932370
				}

				if ($SvSocketCount -le 1)
				{
					$SvSocketCount = 1
				}

				## Intel Hyperthreading -OR- AMD HyperTransport
				## If the number of logical processors divided by the number of cores equals 2, then HT is enabled.
				if (($SvLogProcsCount / $SvProcCoreCount) -eq 2)
				{
					[string]$SvHyperT = 'Active'
				}
				else
				{
					[string]$SvHyperT = 'Inactive'
				}
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Win32_Processor

		#region Critical Windows Services
			#region BMC/BSA/BladeLogic Server Automation RSCD Agent Service
#				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_Service$([char]34) Looking for BMC/BSA/BladeLogic Server Automation RSCD Agent Service."
#				<#
#				#	PSComputerName          : DN3WSCORUTLA01
#				#	Name                    : RSCDsvc
#				#	Status                  : OK
#				#	ExitCode                : 0
#				#	DesktopInteract         : False
#				#	ErrorControl            : Normal
#				#	PathName                : "C:\Program Files\BMC Software\BladeLogic\RSCD\RSCDsvc.exe"
#				#	ServiceType             : Own Process
#				#	StartMode               : Auto
#				#	__GENUS                 : 2
#				#	__CLASS                 : Win32_Service
#				#	__SUPERCLASS            : Win32_BaseService
#				#	__DYNASTY               : CIM_ManagedSystemElement
#				#	__RELPATH               : Win32_Service.Name="RSCDsvc"
#				#	__PROPERTY_COUNT        : 25
#				#	__DERIVATION            : {Win32_BaseService, CIM_Service, CIM_LogicalElement, CIM_ManagedSystemElement}
#				#	__SERVER                : DN3WSCORUTLA01
#				#	__NAMESPACE             : root\CIMV2
#				#	__PATH                  : \\DN3WSCORUTLA01\root\CIMV2:Win32_Service.Name="RSCDsvc"
#				#	AcceptPause             : True
#				#	AcceptStop              : True
#				#	Caption                 : BladeLogic Server Automation RSCD Agent
#				#	CheckPoint              : 0
#				#	CreationClassName       : Win32_Service
#				#	Description             : BladeLogic Server Automation Remote System Call Daemon
#				#	DisplayName             : BladeLogic Server Automation RSCD Agent
#				#	InstallDate             :
#				#	ProcessId               : 2324
#				#	ServiceSpecificExitCode : 0
#				#	Started                 : True
#				#	StartName               : LocalSystem
#				#	State                   : Running
#				#	SystemCreationClassName : Win32_ComputerSystem
#				#	SystemName              : DN3WSCORUTLA01
#				#	TagId                   : 0
#				#	WaitHint                : 0
#				#	Scope                   : System.Management.ManagementScope
#				#	Path                    : \\DN3WSCORUTLA01\root\CIMV2:Win32_Service.Name="RSCDsvc"
#				#	Options                 : System.Management.ObjectGetOptions
#				#	ClassPath               : \\DN3WSCORUTLA01\root\CIMV2:Win32_Service
#				#	Properties              : {AcceptPause, AcceptStop, Caption, CheckPoint...}
#				#	SystemProperties        : {__GENUS, __CLASS, __SUPERCLASS, __DYNASTY...}
#				#	Qualifiers              : {dynamic, Locale, provider, UUID}
#				#	Site                    :
#				#	Container               :
#				#>
#
#				if (Get-Variable -Name Win32_Service_Error -ErrorAction SilentlyContinue)
#				{
#					Clear-Variable -Name Win32_Service_Error -Confirm:$false
#				}
#
#				if ($LocalWmi)
#				{
#					#$BMCService = Get-WmiObject -Namespace root\CIMV2 -ComputerName $SrvDnsName -Class Win32_Service -Filter {Name LIKE 'RSCDsvc%' AND PathName LIKE '%RSCDsvc%'} #-Property Name, PathName, StartMode, ProcessId, State, StartName, Description, DisplayName, Caption
#					$BMCService = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300 -Filter "Name LIKE 'RSCDsvc%' AND PathName LIKE '%RSCDsvc%'")
#				}
#				else
#				{
#					#$BMCService = Get-WmiObject -Namespace root\CIMV2 -ComputerName $SrvDnsName -Credential $UserCreds -Class Win32_Service -Filter {Name LIKE 'RSCDsvc%' AND PathName LIKE '%RSCDsvc%'} #-Property Name, PathName, StartMode, ProcessId, State, StartName, Description, DisplayName, Caption
#					$BMCService = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd -Filter "Name LIKE 'RSCDsvc%' AND PathName LIKE '%RSCDsvc%'")
#				}
#
#				if ($Win32_Service_Error)
#				{
#					[string]$BMCAgentInstalled = 'Error Encountered'
#					[string]$BMCAgentStartMode = 'Error Encountered'
#					[string]$BMCAgentState = 'Error Encountered'
#				}
#				elseif ($BMCService.Status -eq 'OK')
#				{
#					[string]$BMCAgentInstalled = 'YES'
#					[string]$BMCAgentStartMode = $BMCService.StartMode
#					[string]$BMCAgentState = $BMCService.State
#				}
#				else
#				{
#					[string]$BMCAgentInstalled = 'NO'
#					[string]$BMCAgentStartMode = 'Not Installed'
#					[string]$BMCAgentState = 'Not Installed'
#				}
#
#				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
			#endregion BMC/BSA/BladeLogic Server Automation RSCD Agent Service

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_Service$([char]34) Looking for Critical Windows Services."
			<#
			#	PSComputerName          : DN3WSCORUTLA01
			#	Name                    : RSCDsvc
			#	Status                  : OK
			#	ExitCode                : 0
			#	DesktopInteract         : False
			#	ErrorControl            : Normal
			#	PathName                : "C:\Program Files\BMC Software\BladeLogic\RSCD\RSCDsvc.exe"
			#	ServiceType             : Own Process
			#	StartMode               : Auto
			#	__GENUS                 : 2
			#	__CLASS                 : Win32_Service
			#	__SUPERCLASS            : Win32_BaseService
			#	__DYNASTY               : CIM_ManagedSystemElement
			#	__RELPATH               : Win32_Service.Name="RSCDsvc"
			#	__PROPERTY_COUNT        : 25
			#	__DERIVATION            : {Win32_BaseService, CIM_Service, CIM_LogicalElement, CIM_ManagedSystemElement}
			#	__SERVER                : DN3WSCORUTLA01
			#	__NAMESPACE             : root\CIMV2
			#	__PATH                  : \\DN3WSCORUTLA01\root\CIMV2:Win32_Service.Name="RSCDsvc"
			#	AcceptPause             : True
			#	AcceptStop              : True
			#	Caption                 : BladeLogic Server Automation RSCD Agent
			#	CheckPoint              : 0
			#	CreationClassName       : Win32_Service
			#	Description             : BladeLogic Server Automation Remote System Call Daemon
			#	DisplayName             : BladeLogic Server Automation RSCD Agent
			#	InstallDate             :
			#	ProcessId               : 2324
			#	ServiceSpecificExitCode : 0
			#	Started                 : True
			#	StartName               : LocalSystem
			#	State                   : Running
			#	SystemCreationClassName : Win32_ComputerSystem
			#	SystemName              : DN3WSCORUTLA01
			#	TagId                   : 0
			#	WaitHint                : 0
			#	Scope                   : System.Management.ManagementScope
			#	Path                    : \\DN3WSCORUTLA01\root\CIMV2:Win32_Service.Name="RSCDsvc"
			#	Options                 : System.Management.ObjectGetOptions
			#	ClassPath               : \\DN3WSCORUTLA01\root\CIMV2:Win32_Service
			#	Properties              : {AcceptPause, AcceptStop, Caption, CheckPoint...}
			#	SystemProperties        : {__GENUS, __CLASS, __SUPERCLASS, __DYNASTY...}
			#	Qualifiers              : {dynamic, Locale, provider, UUID}
			#	Site                    :
			#	Container               :
			#>

			if (Get-Variable -Name Win32_Service_Error -ErrorAction SilentlyContinue)
			{
				Clear-Variable -Name Win32_Service_Error -Confirm:$false
			}

			#	a. BladeLogic Server Automation RSCD Agent - RSCDsvc - Automatic - Started
			#	b. Remote Registry - RemoteRegistry - Automatic - Started
			#	c. Windows Update - wuauserv - Automatic (Delayed Start) - Started
			#	d. Netlogon - Netlogon - Automatic - Started
			#	e. Windows Firewall - MpsSvc - Automatic - Started (Not needed for BMC, but critical to other OS functions.)

			if ($LocalWmi)
			{
				#$Win32Service = Get-WmiObject -Namespace root\CIMV2 -ComputerName $SrvDnsName -Class Win32_Service} #-Property Name, PathName, StartMode, ProcessId, State, StartName, Description, DisplayName, Caption
				$Win32Service = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300)
			}
			else
			{
				#$Win32Service = Get-WmiObject -Namespace root\CIMV2 -ComputerName $SrvDnsName -Credential $UserCreds -Class Win32_Service} #-Property Name, PathName, StartMode, ProcessId, State, StartName, Description, DisplayName, Caption
				$Win32Service = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_Service' -Property '*' -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd)
			}

			$BMCService = @($Win32Service | Where-Object { ($_.Name -like 'RSCDsvc*') -and ($_.PathName -like '*RSCDsvc*') } | Select-Object *)
			$SvcNetlogon = @($Win32Service | Where-Object { $_.Name -eq 'Netlogon' } | Select-Object *)
			$SvcRemRegistry = @($Win32Service | Where-Object { $_.Name -eq 'RemoteRegistry' } | Select-Object *)
			$SvcWinFirewall = @($Win32Service | Where-Object { $_.Name -eq 'MpsSvc' } | Select-Object *)
			$SvcWinUpdate = @($Win32Service | Where-Object { $_.Name -eq 'wuauserv' } | Select-Object *)

			if ($Win32_Service_Error)
			{
				## BladeLogic Agent
				[string]$BMCAgentInstalled = 'Error Encountered'
				[string]$BMCAgentStartMode = 'Error Encountered'
				[string]$BMCAgentState = 'Error Encountered'

				## Netlogon
				[string]$NetlogonStartMode = 'Error Encountered'
				[string]$NetlogonState = 'Error Encountered'

				## Remote Registry
				[string]$RemoteRegistryStartMode = 'Error Encountered'
				[string]$RemoteRegistryState = 'Error Encountered'

				## Windows Firewall
				[string]$WinFirewallStartMode = 'Error Encountered'
				[string]$WinFirewallState = 'Error Encountered'

				## Windows Update
				[string]$WinUpdateStartMode = 'Error Encountered'
				[string]$WinUpdateState = 'Error Encountered'
			}
			else
			{
				## BladeLogic Agent
				if ($BMCService.Status -eq 'OK')
				{
					[string]$BMCAgentInstalled = 'YES'
					[string]$BMCAgentStartMode = $BMCService.StartMode
					[string]$BMCAgentState = $BMCService.State
				}
				else
				{
					[string]$BMCAgentInstalled = 'NO'
					[string]$BMCAgentStartMode = 'Not Installed'
					[string]$BMCAgentState = 'Not Installed'
				}

				## Netlogon
				[string]$NetlogonStartMode = $SvcNetlogon.StartMode
				[string]$NetlogonState = $SvcNetlogon.State

				## Remote Registry
				[string]$RemoteRegistryStartMode = $SvcRemRegistry.StartMode
				[string]$RemoteRegistryState = $SvcRemRegistry.State

				## Windows Firewall
				if ($SvcWinFirewall.Status -eq 'OK')
				{
					[string]$WinFirewallStartMode = $SvcWinFirewall.StartMode
					[string]$WinFirewallState = $SvcWinFirewall.State
				}
				else
				{
					[string]$WinFirewallStartMode = 'Not Available'
					[string]$WinFirewallState = 'Not Available'
				}

				## Windows Update
				[string]$WinUpdateStartMode = $SvcWinUpdate.StartMode
				[string]$WinUpdateState = $SvcWinUpdate.State
			}

			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
		#endregion Critical Windows Services

		if ($IncludeInfraDetails)
		{
			#region Win32_LogicalDisk
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tRunning queries for $([char]34)Win32_LogicalDisk$([char]34)"

				#$Win32LogicalDiskQuery = 'SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3'
				if ($LocalWmi)
				{
					#$Win32LogicalDisk = @(Get-WmiObject -ComputerName $SrvDnsName -ErrorVariable Win32_LogicalDisk_Error -Query $Win32LogicalDiskQuery)
					$Win32LogicalDisk = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_LogicalDisk' -Property '*' -Filter "DriveType = 3" -TimeoutSeconds 300)
				}
				else
				{
					#$Win32LogicalDisk = @(Get-WmiObject -ComputerName $SrvDnsName -Credential $UserCreds -ErrorVariable Win32_LogicalDisk_Error -Query $Win32LogicalDiskQuery)
					$Win32LogicalDisk = @(Get-WmiObjectWithTimeout -ComputerName $SrvDnsName -Namespace 'ROOT\CIMV2' -Class 'Win32_LogicalDisk' -Property '*' -Filter "DriveType = 3" -TimeoutSeconds 300 -WmiUserName $NbDmnUsrLgn -WmiSecPasswd $SrvSecPasswd)
				}

				if ($Win32_LogicalDisk_Error)
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`t`tWin32_LogicalDisk Query Error was: $([char]34)$($Win32_LogicalDisk_Error[0].Exception.Message)$([char]34)"

					[string]$SvDiskTotalInfo = 'Error Encountered'
					[string]$SvDiskUsedInfo = 'Error Encountered'
					[string]$SvDiskFreeInfo = 'Error Encountered'
					[string]$SvDiskTotalAllocated = 'Error Encountered'
					[string]$SvDiskUsedAllocated = 'Error Encountered'
					[string]$SvDiskFreeAllocated = 'Error Encountered'
				}
				else
				{
					Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`t`tWin32_LogicalDisk query did not return any Errors."

					## Each Drive showing total disk space. 'C: 135.61; D: 1115.37; '
					$SvDiskTotalInfoArray = New-Object System.Collections.Arraylist
					## Each Drive showing used disk space. 'C: 36.49; D: 406.2; '
					$SvDiskUsedInfoArray = New-Object System.Collections.Arraylist
					## Each Drive showing free disk space. 'C: 99.12; D: 709.18; '
					$SvDiskFreeInfoArray = New-Object System.Collections.Arraylist
					## Sum of each drive showing total allocated disk space. '1250.98'
					[float]$SvDiskTotalAllocated = 0
					## Sum of each drive showing used allocated disk space. '442.69'
					[float]$SvDiskUsedAllocated = 0
					## Sum of each drive showing free allocated disk space. '808.29'
					[float]$SvDiskFreeAllocated = 0

					foreach ($SvrDisk in $Win32LogicalDisk)
					{
						$DiskID = $SvrDisk.DeviceID
						[float]$Size = $SvrDisk.Size
						[float]$FreeSpace = $SvrDisk.FreeSpace

						## Byte Conversion Table
						## 1 * 1024 = 1024 = 1KB
						## 1 * 1024 * 1024 = 1048576 = 1MB
						## 1 * 1024 * 1024 * 1024 = 1073741824 = 1GB
						$TtlGB = [Math]::Round(($Size / 1073741824), 2)
						$UsedGB = [Math]::Round((($Size - $FreeSpace) / 1073741824), 2)
						$FreeGB = [Math]::Round(($FreeSpace / 1073741824), 2)
						#$FreePcnt = [Math]::Round((($FreeSpace / $Size) * 100), 2)

						$SvDiskTotalInfoArray += "$($DiskID)\ $($TtlGB.ToString())"
						$SvDiskUsedInfoArray += "$($DiskID)\ $($UsedGB.ToString())"
						$SvDiskFreeInfoArray += "$($DiskID)\ $($FreeGB.ToString())"
						$SvDiskTotalAllocated += $TtlGB
						$SvDiskUsedAllocated += $UsedGB
						$SvDiskFreeAllocated += $FreeGB
					}

					[string]$SvDiskTotalInfo = [String]::Join('; ', $SvDiskTotalInfoArray)
					[string]$SvDiskUsedInfo = [String]::Join('; ', $SvDiskUsedInfoArray)
					[string]$SvDiskFreeInfo = [String]::Join('; ', $SvDiskFreeInfoArray)
				}
				Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------------"
			#endregion Win32_LogicalDisk
		}

		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`tFinished Processing WMI calls for server: $([char]34)$($SrvDnsName)$([char]34)"
	}
	else
	{
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`t`tNOT responding! There will be NO PROCESSING OF WMI calls for server: $([char]34)$($SrvDnsName)$([char]34)"

		[string]$SvDataCenter = 'No Response'
		[string]$SvBMCLocationCode = 'No Response'
		[string]$SvSerNum = 'No Response'
		[string]$SvBiosVer = 'No Response'
		[string]$SvMnfctr = 'No Response'
		[string]$SvModel = 'No Response'
		[string]$SvInstOsArch = 'No Response'
		[string]$SvLastBootUpTime = 'No Response'
		[string]$SvUpTime = 'No Response'

		if ($IncludeHotFixDetails)
		{
			[string]$SvLastPatchInstallDate = 'No Response'
			[string]$SvUpTimeKB2553549HotFix1Applied = 'No Response'
			[string]$SvUpTimeKB2688338HotFix2Applied = 'No Response'
			[string]$SvUpTimeTotalHotFixApplied = 'No Response'
			[string]$SvKB3042553HotFixApplied = 'No Response'
		}

		[string]$MSClusterName = 'No Response'
		[string]$SvRegOsName = 'No Response'
		[string]$SvRegSpLvl = 'No Response'

		if ($IncludeMSSQL)
		{
			[string]$SQLSvrIsInstalled = 'No Response'
			[string]$SQLInstanceNames = 'No Response'
			[string]$SQLProductCaptions = 'No Response'
			[string]$SQLEditions = 'No Response'
			[string]$SQLVersions = 'No Response'
			[string]$SQLSvcDisplayNames	= 'No Response'
			[string]$SQLIsClustered = 'No Response'
			[string]$SQLClusterNames = 'No Response'
			[string]$SQLClusterNodes = 'No Response'
			[string]$SQLClusterInstanceOwnerRole = 'No Response'
		}

		if ($IncludeCustomAppInfo)
		{
			[string]$SvBrowserHawkInstall = 'No Response'
			[string]$SvSoftArtisansFileUpInstall = 'No Response'
		}

		[string]$SvPendingReboot = 'No Response'
		[string]$SvPhysMemoryGB = 'No Response'
		[string]$SvProcMnftr = 'No Response'
		[string]$SvSocketCount = 'No Response'
		[string]$SvProcCoreCount = 'No Response'
		[string]$SvLogProcsCount = 'No Response'
		[string]$SvHyperT = 'No Response'
		## BladeLogic Agent
		[string]$BMCAgentInstalled = 'No Response'
		[string]$BMCAgentStartMode = 'No Response'
		[string]$BMCAgentState = 'No Response'
		## Netlogon
		[string]$NetlogonStartMode = 'No Response'
		[string]$NetlogonState = 'No Response'
		## Remote Registry
		[string]$RemoteRegistryStartMode = 'No Response'
		[string]$RemoteRegistryState = 'No Response'
		## Windows Firewall
		[string]$WinFirewallStartMode = 'No Response'
		[string]$WinFirewallState = 'No Response'
		## Windows Update
		[string]$WinUpdateStartMode = 'No Response'
		[string]$WinUpdateState = 'No Response'

		if ($IncludeInfraDetails)
		{
			[string]$SvProcName = 'No Response'
			[string]$SvDiskTotalInfo = 'No Response'
			[string]$SvDiskUsedInfo = 'No Response'
			[string]$SvDiskFreeInfo = 'No Response'
			[string]$SvDiskTotalAllocated = 'No Response'
			[string]$SvDiskUsedAllocated = 'No Response'
			[string]$SvDiskFreeAllocated = 'No Response'
		}
	}

	if ($ScrDebug)
	{
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Names and Values being returned to the main script for further processing..."
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-----------------------------------------------------------------------------------------"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Name           : Value"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------:-----------------------------------------------------------------"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvName                 : $($SrvName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDnsName              : $($SrvDnsName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvOsName               : $($SrvOsName)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvOsSpLvl              : $($SrvOsSpLvl)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvChangedOn            : $($SrvChangedOn)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvCreatedOn            : $($SrvCreatedOn)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvNBDomain             : $($SrvNBDomain)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t IpAddress               : $($IpAddress)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvIPv4Addr             : $($SvIPv4Addr)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDataCenter           : $($SvDataCenter)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvBMCLocationCode      : $($SvBMCLocationCode)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvSerNum               : $($SvSerNum)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvBiosVer              : $($SvBiosVer)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvMnfctr               : $($SvMnfctr)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvModel                : $($SvModel)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvInstOsArch           : $($SvInstOsArch)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvLastBootUpTime       : $($SvLastBootUpTime)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvUpTime               : $($SvUpTime)"

		if ($IncludeHotFixDetails)
		{
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvLastPatchInstallDate            : $($SvLastPatchInstallDate)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvUpTimeKB2553549HotFix1Applied   : $($SvUpTimeKB2553549HotFix1Applied)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvUpTimeKB2688338HotFix2Applied   : $($SvUpTimeKB2688338HotFix2Applied)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvUpTimeTotalHotFixApplied        : $($SvUpTimeTotalHotFixApplied)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvKB3042553HotFixApplied          : $($SvKB3042553HotFixApplied)"
		}

		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t MSClusterName           : $([char]34)$($MSClusterName)$([char]34)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvRegOsName            : $([char]34)$($SvRegOsName.Trim())$([char]34)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvRegSpLvl             : $([char]34)$($SvRegSpLvl.Trim())$([char]34)"

		if ($IncludeMSSQL)
		{
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLSvrIsInstalled       : $([char]34)$($SQLSvrIsInstalled.ToString())$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLInstanceNames        : $([char]34)$($SQLInstanceNames)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLProductCaptions      : $([char]34)$($SQLProductCaptions)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLEditions             : $([char]34)$($SQLEditions)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLVersions             : $([char]34)$($SQLVersions)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLSvcDisplayNames      : $([char]34)$($SQLSvcDisplayNames)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLIsClustered          : $([char]34)$($SQLIsClustered.ToString())$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t isSQLClusterNode        : $([char]34)$(($SQLClusterNodesArray -contains $SrvName).ToString())$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLClusterNames         : $([char]34)$($SQLClusterNames)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLClusterNodes         : $([char]34)$($SQLClusterNodes)$([char]34)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SQLClusterInstanceOwner : $([char]34)$($SQLClusterInstanceOwnerRole)$([char]34)"
		}

		if ($IncludeCustomAppInfo)
		{
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvBrowserHawkInstall   : $($SvBrowserHawkInstall)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvSoftArtisansFileUpInstall : $($SvSoftArtisansFileUpInstall)"
		}

		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvPendingReboot        : $([char]34)$($SvPendingReboot.ToString())$([char]34)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvPhysMemoryGB         : $($SvPhysMemoryGB)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvProcMnftr            : $($SvProcMnftr)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvSocketCount          : $($SvSocketCount)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvProcCoreCount        : $($SvProcCoreCount)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvLogProcsCount        : $($SvLogProcsCount)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvHyperT               : $($SvHyperT)"
		## BladeLogic Agent
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvBMCAgentInstalled    : $($BMCAgentInstalled)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvBMCAgentStartMode    : $($BMCAgentStartMode)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvBMCAgentState        : $($BMCAgentState)"
		## Netlogon
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvNetlogonStartMode    : $($NetlogonStartMode)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvNetlogonState        : $($NetlogonState)"
		## Remote Registry
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvRemoteRegistryStartMode        : $($RemoteRegistryStartMode)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvRemoteRegistryState  : $($RemoteRegistryState)"
		## Windows Firewall
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvWinFirewallStartMode : $($WinFirewallStartMode)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvWinFirewallState     : $($WinFirewallState)"
		## Windows Update
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvWinUpdateStartMode   : $($WinUpdateStartMode)"
		Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvWinUpdateState       : $($WinUpdateState)"

		if ($IncludeInfraDetails)
		{
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvProcName             : $($SvProcName)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDiskTotalInfo        : $($SvDiskTotalInfo)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDiskUsedInfo         : $($SvDiskUsedInfo)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDiskFreeInfo         : $($SvDiskFreeInfo)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDiskTotalAllocated   : $($SvDiskTotalAllocated)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDiskUsedAllocated    : $($SvDiskUsedAllocated)"
			Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvDiskFreeAllocated    : $($SvDiskFreeAllocated)"
		}
	}

	#region Return PSObject Array
		$PsObjSvrProps = New-Object PSObject #New-Object System.Object
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvName'			-Value $SrvName
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDnsName'			-Value $SrvDnsName
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvOsName'			-Value $SrvOsName
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvOsSpLvl'			-Value $SrvOsSpLvl
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvChangedOn'		-Value $SrvChangedOn
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvCreatedOn'		-Value $SrvCreatedOn
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvNBDomain'		-Value $SrvNBDomain
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDnsDomain'		-Value $SrvDnsDomain
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvIPv4Addr'		-Value $SvIPv4Addr
		#$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvIPv4Addr'		-Value $IpAddress
		#$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvAddlIPv4Addr'	-Value $SvIPv4Addr
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDataCenter'		-Value $SvDataCenter
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvBMCLocationCode'	-Value $SvBMCLocationCode
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvSerNum'			-Value $SvSerNum
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvBiosVer'			-Value $SvBiosVer
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvMnfctr'			-Value $SvMnfctr
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvModel'			-Value $SvModel.Trim()
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvInstOsArch'		-Value $SvInstOsArch
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvLastBootUpTime'	-Value $SvLastBootUpTime
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvUpTime'			-Value $SvUpTime

		if ($IncludeHotFixDetails)
		{
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvLastPatchInstallDate' -Value $SvLastPatchInstallDate
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvUpTimeKB2553549HotFix1Applied' -Value $SvUpTimeKB2553549HotFix1Applied
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvUpTimeKB2688338HotFix2Applied' -Value $SvUpTimeKB2688338HotFix2Applied
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvUpTimeTotalHotFixApplied' -Value $SvUpTimeTotalHotFixApplied
			## MS15-034 - Zero Day Vulnerability in HTTP.sys on Windows Server 2008 R2 and newer.
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvKB3042553HotFixApplied' -Value $SvKB3042553HotFixApplied
		}

		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'MSClusterName'		-Value $MSClusterName
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvRegOsName'		-Value $SvRegOsName.Trim()
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvRegSpLvl'		-Value $SvRegSpLvl.Trim()

		if ($IncludeMSSQL)
		{
			if ($SQLSvrIsInstalled)
			{
				$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLSvrName'			-Value $SrvName.ToLower()
			}
			else
			{
				$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLSvrName'			-Value ''
			}

			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLSvrIsInstalled'	-Value $SQLSvrIsInstalled	#.ToString()
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLProductCaptions'	-Value $SQLProductCaptions
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLEditions'		-Value $SQLEditions
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLVersions'		-Value $SQLVersions
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLSvcDisplayNames'	-Value $SQLSvcDisplayNames
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLInstanceNames'	-Value $SQLInstanceNames
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLIsClustered'		-Value $SQLIsClustered	#.ToString()
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'isSQLClusterNode'	-Value ($SQLClusterNodesArray -contains $SrvName)	#.ToString()
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLClusterNames'	-Value $SQLClusterNames
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLClusterNodes'	-Value $SQLClusterNodes
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SQLClusterInstanceOwner'	-Value $SQLClusterInstanceOwnerRole
		}

		if ($IncludeCustomAppInfo)
		{
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvBrowserHawkInstall'		-Value $SvBrowserHawkInstall
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvSoftArtisansFileUpInstall'	-Value $SvSoftArtisansFileUpInstall
		}

		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvPendingReboot'	-Value $SvPendingReboot	#.ToString()
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvPhysMemoryGB'	-Value $SvPhysMemoryGB
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvProcMnftr'		-Value $SvProcMnftr
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvSocketCount'		-Value $SvSocketCount
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvProcCoreCount'	-Value $SvProcCoreCount
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvLogProcsCount'	-Value $SvLogProcsCount
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvHyperT'			-Value $SvHyperT
		## BladeLogic Agent
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvBMCAgentInstalled' -Value $BMCAgentInstalled
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvBMCAgentStartMode' -Value $BMCAgentStartMode
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvBMCAgentState'	-Value $BMCAgentState
		## Netlogon
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvNetlogonStartMode' -Value $NetlogonStartMode
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvNetlogonState' -Value $NetlogonState
		## Remote Registry
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvRemoteRegistryStartMode' -Value $RemoteRegistryStartMode
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvRemoteRegistryState' -Value $RemoteRegistryState
		## Windows Firewall
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvWinFirewallStartMode' -Value $WinFirewallStartMode
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvWinFirewallState' -Value $WinFirewallState
		## Windows Update
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvWinUpdateStartMode' -Value $WinUpdateStartMode
		$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvWinUpdateState' -Value $WinUpdateState

		if ($IncludeInfraDetails)
		{
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvProcName'			-Value $SvProcName
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDiskTotalInfo'		-Value $SvDiskTotalInfo
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDiskUsedInfo'		-Value $SvDiskUsedInfo
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDiskFreeInfo'		-Value $SvDiskFreeInfo
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDiskTotalAllocated'	-Value $SvDiskTotalAllocated
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDiskUsedAllocated'	-Value $SvDiskUsedAllocated
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'SrvDiskFreeAllocated'	-Value $SvDiskFreeAllocated
		}
	#endregion Return PSObject Array

	$RtnServerProps	+= $PsObjSvrProps
	Add-Content -Path $MultiThreadScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Returning variables and values to the main script for further processing..."

	return $RtnServerProps
#endregion Main Multi-Threaded Execution
