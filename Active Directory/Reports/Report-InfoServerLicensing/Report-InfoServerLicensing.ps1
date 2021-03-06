#Requires -Version 2.0
####################################################################################################################
##
##	Old Script Name      : Report-ADServerCountForLicensing.ps1
##	Script Name          : Report-InfoServerLicensing.ps1
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
##	Purpose              : Execute a series of WMI and other queries using Powershell to list all Windows Servers from Active Directory For Licensing purposes and create a report to be sent via email.
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
	DefaultParameterSetName = 'SettingsFile'
)]
#region Script Variable Initialization
	#region Script Parameters
		Param(
			[Parameter(
				Position			= 0
				, ParameterSetName	= 'SettingsFile'
				, Mandatory			= $true
				#, ValueFromPipeline	= $false
				#, HelpMessage		= 'The name of the XML settings file. If the settings file is not located in the same directory as the script then provide the full path to the file.'
			)]
			[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
			[string]$SettingsFile
			, [Parameter(
				Position			= 1
				, ParameterSetName	= 'SettingsFile'
				, Mandatory			= $false
				#, ValueFromPipeline	= $false
				#, HelpMessage		= 'Attempt to schedule the command just executed to run at 2:30AM on the last Monday of each month. Specify the username here, SCHTASKS (under the hood) will ask for a password later.'
			)][string]$ScheduleAs
			, [Parameter(
				Position			= 2
				, ParameterSetName	= 'SettingsFile'
				, Mandatory			= $false
				#, ValueFromPipeline	= $false
				#, HelpMessage		= 'A name which will uniquely identify the scheduled task.'
			)][string]$SchTskName
		)
	#endregion Script Parameters

	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"

	#region Assembly and PowerShell Module Initialization
		######################################
		##  Import Modules & Add PSSnapins  ##
		######################################
		## Check for ActiveRoles Management Shell for Active Directory Sanpin, attempt to load.
		## WORKS IN A SINGLE OR MULTIPLE DOMAIN ENVIRONMENT!
		if (!(Get-Command Get-QADComputer -ErrorAction SilentlyContinue))
		{
			try
			{
				Add-PSSnapin Quest.ActiveRoles.ADManagement
			}
			catch
			{
				if (!(Get-Command Get-QADComputer -ErrorAction SilentlyContinue))
				{
					Add-PSSnapin Quest.ActiveRoles.ADManagement
				}
			}
			finally
			{
				if (!(Get-Command Get-QADComputer -ErrorAction SilentlyContinue))
				{
					#throw "Cannot load the PowerShell Snapin $([char]34)Quest.ActiveRoles.ADManagement$([char]34)!! Please make sure the $([char]34)ActiveRoles Management Shell for Active Directory$([char]34) is installed and try again."
					throw "The script '$($MyInvocation.MyCommand.Name)' cannot be run because the following PowerShell Snap-ins are missing: $([char]34)Quest.ActiveRoles.ADManagement$([char]34). Please make sure the $([char]34)ActiveRoles Management Shell for Active Directory$([char]34) is installed and try again."
				}

				$SnapinSettingsSizeLimit = Get-QADPSSnapinSettings -DefaultSizeLimit
				if ($SnapinSettingsSizeLimit -ne 0)
				{
					Set-QADPSSnapinSettings -DefaultSizeLimit 0
				}
			}
		}

		## Add the Excel Interop assembly
		#Add-Type -AssemblyName Microsoft.Office.Interop.Excel
		[Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")

		$Error.Clear()
	#endregion Assembly and PowerShell Module Initialization

	######################################
	##  Script Variable Initialization  ##
	######################################
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)
	[string]$Script:FullScrPath	= ($MyInvocation.MyCommand.Definition)

	#[string]$Script:ScriptNameNoExt = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	[string]$Script:ScriptNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

	#region Load Settings File
		#[xml]$XmlConfig = Get-Content -Path $SettingsFile
		$XmlConfig = New-Object -TypeName XML
		$XmlConfig.Load($SettingsFile)

		#region DataCenters and Subnets
			$Script:DataCenters = $XmlConfig.ScriptConfig.DataCenters.DataCenter

			#$Script:SubnetToDataCenter = @()
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SubnetToDataCenter -Value @()

			foreach ($DataCenter in $DataCenters)
			{
				foreach ($DCSubnet in $DataCenter.Subnets)
				{
					foreach ($objSubnetAddr in $DCSubnet.Subnet)
					{
						## Dynamically build an array...
						$tmpSubnet = New-Object -TypeName PSObject -Property @{
							'DataCenterCity' = $DataCenter.DataCenterCity
							'DataCenterCompany' = $DataCenter.DataCenterCompany
							'LocationCodes' = $DataCenter.LocationCodes
							'BMCLocationCode' = $DataCenter.BMCLocationCode
							'Subnet' = $objSubnetAddr
						}

						$SubnetToDataCenter += $tmpSubnet
					}
				}
			}
		#endregion DataCenters and Subnets

		#region DefaultSettings
			$Script:DefaultSettings = $XmlConfig.ScriptConfig.DefaultSettings

			if ($DefaultSettings.ScrDebug -eq 'TRUE')
			{
				#[switch]$ScrDebug = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name ScrDebug -Value $true
			}
			else
			{
				#[switch]$ScrDebug = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name ScrDebug -Value $false
			}

			#[string]$DataCenterCity = $DefaultSettings.DataCenterCity
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DataCenterCity -Value $DefaultSettings.DataCenterCity

			#[string]$MultiThreadScript = [System.IO.Path]::GetFullPath(($DefaultSettings.MultiThreadScript).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity))
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name MultiThreadScript -Value ([System.IO.Path]::GetFullPath(($DefaultSettings.MultiThreadScript).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)))

			#[Byte[]]$Key = 211,209,124,250,220,103,226,241,248,133,153,152,129,183,215,121,147,238,253,125,236,154,144,165,190,146,228,227,170,191,218,163
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FlatKey -Value $DefaultSettings.Key
			[Byte[]]$Script:Key = $FlatKey -split ',' | ForEach-Object { $_ }
			## The following needs a little more work. It doesn't work just yet. Trying to combine the above 2 lines in to one if possible.
			#New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name Key -Value $([Byte[]]($DefaultSettings.Key -split ',' | ForEach-Object { $_ }))

			if ($DefaultSettings.DisableOldADServerObjects -eq 'TRUE')
			{
				#[switch]$DisableOldADServerObjects = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DisableOldADServerObjects -Value $true
			}
			else
			{
				#[switch]$DisableOldADServerObjects = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DisableOldADServerObjects -Value $false
			}

			if ($ScrDebug)
			{
				#[int]$Script:DisableServersOlderThan = ($DisableServersOlderThan * 2)
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DisableServersOlderThan -Value ([int]($DefaultSettings.DisableServersOlderThan) * 2)
			}
			else
			{
				#[int]$Script:DisableServersOlderThan = $DefaultSettings.DisableServersOlderThan
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DisableServersOlderThan -Value ([int]($DefaultSettings.DisableServersOlderThan))
			}

			if ($DefaultSettings.CreateCsv -eq 'TRUE')
			{
				#[switch]$CreateCsv = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name CreateCsv -Value $true
			}
			else
			{
				#[switch]$CreateCsv = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name CreateCsv -Value $false
			}

			if ($DefaultSettings.CreateXlsx -eq 'TRUE')
			{
				#[switch]$CreateXlsx = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name CreateXlsx -Value $true
			}
			else
			{
				#[switch]$CreateXlsx = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name CreateXlsx -Value $false
			}

			if ($DefaultSettings.IncludeMSSQL -eq 'TRUE')
			{
				#[switch]$IncludeMSSQL = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeMSSQL -Value $true
			}
			else
			{
				#[switch]$IncludeMSSQL = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeMSSQL -Value $false
			}

			if ($DefaultSettings.IncludeExcessiveUpTime -eq 'TRUE')
			{
				#[switch]$IncludeExcessiveUpTime = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeExcessiveUpTime -Value $true
			}
			else
			{
				#[switch]$IncludeExcessiveUpTime = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeExcessiveUpTime -Value $false
			}

			if ($DefaultSettings.IncludeHotFixDetails -eq 'TRUE')
			{
				#[switch]$IncludeHotFixDetails = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeHotFixDetails -Value $true
			}
			else
			{
				#[switch]$IncludeHotFixDetails = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeHotFixDetails -Value $false
			}

			if ($DefaultSettings.IncludeCustomAppInfo -eq 'TRUE')
			{
				#[switch]$IncludeCustomAppInfo = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeCustomAppInfo -Value $true
			}
			else
			{
				#[switch]$IncludeCustomAppInfo = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeCustomAppInfo -Value $false
			}

			if ($DefaultSettings.IncludeInfraDetails -eq 'TRUE')
			{
				#[switch]$IncludeInfraDetails = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeInfraDetails -Value $true
			}
			else
			{
				#[switch]$IncludeInfraDetails = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name IncludeInfraDetails -Value $false
			}

			if ($DefaultSettings.ExportCsvForBMC -eq 'TRUE')
			{
				#[switch]$ExportCsvForBMC = $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name ExportCsvForBMC -Value $true
			}
			else
			{
				#[switch]$ExportCsvForBMC = $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name ExportCsvForBMC -Value $false
			}
		#endregion DefaultSettings

		#region EmailSettings
			$Script:EmailSettings = $XmlConfig.ScriptConfig.EmailSettings

			if ($EmailSettings.SendMail -eq 'TRUE')
			{
				#[switch]$SendMail = $true
				#[string]$FromAddr = ($EmailSettings.FromAddr).Replace('[DataCenterCity]', $DataCenterCity).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt)

				#[String[]]$ToAddr = @()
				#foreach ($Recipient in $EmailSettings.Recipients.ToAddr)
				#{
				#	$ToAddr += $Recipient.Trim()
				#}

				##[string]$SmtpServers = New-Object System.Collections.Hashtable
				##foreach ($SmtpServer in $SmtpServers)
				##{
				##	[string]$SmtpServers.Add($SmtpServer.DataCenterCity, $SmtpServer.SmtpSvrName)
				##}
				#[string]$SmtpSvr = ($EmailSettings.SmtpServers.SmtpSvr | Where-Object { $_.DataCenterCity -eq $DataCenterCity } | Select-Object SmtpSvrName).SmtpSvrName

				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SendMail -Value $true
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FromAddr -Value ($EmailSettings.FromAddr).Replace('[DataCenterCity]', $DataCenterCity).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace(' ', '')
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name ToAddr -Value @()

				foreach ($Recipient in $EmailSettings.Recipients.ToAddr)
				{
					$ToAddr += $Recipient.Trim()
				}

				#New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SmtpSvr -Value ($EmailSettings.SmtpServers.SmtpSvr | Where-Object { $_.DataCenterCity -eq $DataCenterCity } | Select-Object SmtpSvrName).SmtpSvrName
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SmtpSvr -Value @($EmailSettings.SmtpServers.SmtpSvr)
			}
			else
			{
				#[switch]$SendMail = $false
				#[string]$FromAddr = ''
				#[String[]]$ToAddr = ''
				#[string]$SmtpSvr = ''

				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SendMail -Value $false
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FromAddr -Value ''
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name ToAddr -Value ''
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SmtpSvr -Value ''
			}
		#endregion EmailSettings

		#region DomainUserInfo
			#$Script:DomainUserInfo = $XmlConfig.ScriptConfig.DomainUserInfo

			$Script:DomainUsers = $XmlConfig.ScriptConfig.DomainUserInfo.DomainUser

			#$DomainUserInfo = @()
			#$DomainUserInfo = New-Object System.Collections.Arraylist
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DomainUserInfo -Value @()

			foreach ($DomainUser in $DomainUsers)
			{
				#if ($DomainUser.IsCentralManagementCred -eq 'TRUE')
				#{
				#	[bool]$DUiCMC = $true
				#}
				#else
				#{
				#	[bool]$DUiCMC = $false
				#}

				$tmpADUser = New-Object -TypeName PSObject -Property @{
					'DnsDomain' = $DomainUser.DnsDomain
					'NbDomain' = $DomainUser.NbDomain
					'UserName' = $DomainUser.UserName
					'EncPassword' = $DomainUser.EncPassword
					#'IsCentralManagementCred' = $DUiCMC
					'IsCentralManagementCred' = $DomainUser.IsCentralManagementCred
				}

				#$tmpADUser = New-Object -TypeName PSObject
				#$tmpADUser | Add-Member -MemberType NoteProperty -Name 'DnsDomain' -Value $DomainUser.DnsDomain
				#$tmpADUser | Add-Member -MemberType NoteProperty -Name 'NbDomain' -Value $DomainUser.NbDomain
				#$tmpADUser | Add-Member -MemberType NoteProperty -Name 'UserName' -Value $DomainUser.UserName
				#$tmpADUser | Add-Member -MemberType NoteProperty -Name 'EncPassword' -Value $DomainUser.EncPassword

				$DomainUserInfo += $tmpADUser
			}
		#endregion DomainUserInfo

		#region InScopeDNSDomains
			#$Script:InScopeDNSDomains = $XmlConfig.ScriptConfig.InScopeDNSDomains

			#$Script:ActiveDirectoryDomains = @($XmlConfig.ScriptConfig.InScopeDNSDomains.ActiveDirectoryDomain)
			$Script:ActiveDirectoryDomains = $XmlConfig.ScriptConfig.InScopeDNSDomains.ActiveDirectoryDomain

			#$Script:InScopeDNSDomains = @()
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name InScopeDNSDomains -Value @()

			#[String[]]$Script:SplaDnsDomains = @()
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name SplaDnsDomains -Value @()

			#$Script:DNStoNBnames = @{}
			#$Script:DNStoNBnames = New-Object System.Collections.Hashtable
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name DNStoNBnames -Value @{}

			foreach ($ADDomain in $ActiveDirectoryDomains)
			{
				if ($DataCenterCity -eq 'All')
				{
					$tmpDNSName		= ($ADDomain | Select-Object DNSName).DNSName
					$tmpNetBiosName	= ($ADDomain | Select-Object NetBiosName).NetBiosName
				}
				else
				{
					$tmpDNSName		= ($ADDomain | Where-Object { $_.DataCenterCity -eq $DataCenterCity } | Select-Object DNSName).DNSName
					$tmpNetBiosName	= ($ADDomain | Where-Object { $_.DataCenterCity -eq $DataCenterCity } | Select-Object NetBiosName).NetBiosName
				}

				if ($tmpDNSName.Length -ge 3)
				{
					#if ($ADDomain.IsCentrallyManaged -eq 'TRUE')
					#{
					#	[bool]$ADDiCM = $true
					#}
					#else
					#{
					#	[bool]$ADDiCM = $false
					#}

					$tmpDNSDomain = New-Object -TypeName PSObject -Property @{
						'DNSName' = $ADDomain.DNSName.ToLower()
						'NetBiosName' = $ADDomain.NetBiosName.ToUpper()
						'Environment' = $ADDomain.Environment
						'DataCenterCity' = $ADDomain.DataCenterCity
						#'IsCentrallyManaged' = $ADDiCM
						'IsCentrallyManaged' = $ADDomain.IsCentrallyManaged
					}

					## Array
					$InScopeDNSDomains += $tmpDNSDomain
					## [String[]] String Array
					$SplaDnsDomains += $tmpDNSName.ToLower()
					## [Hashtable]
					$DNStoNBnames.Add($tmpDNSName.ToLower(), $tmpNetBiosName.ToUpper())
				}
			}
		#endregion InScopeDNSDomains

		#region Reports and Logs Variables
			$Script:ReportAndLogSettings = $XmlConfig.ScriptConfig.ReportAndLogSettings

			if ($CreateXlsx)
			{
				#$XlsxCreationEngine = $ReportAndLogSettings.XlsxCreationEngine
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name XlsxCreationEngine -Value $ReportAndLogSettings.XlsxCreationEngine
			}

			#$PartialReportAndLogName = ($ReportAndLogSettings.PartialReportAndLogName).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name PartialReportAndLogName -Value ($ReportAndLogSettings.PartialReportAndLogName).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)

			####################################################################################################
			##	Date and Time Formatting
			##	http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo%28VS.85%29.aspx
			##	Formatting chosen for the date string is as follows:
			##	Date Display  |  -Format
			##	------------  |  ------------
			##	2012.Jan.25   |  yyyy.MMM.dd
			##	2012.01.25    |  yyyy.MM.dd
			##	Mon			  |  ddd
			##	------------  |  ------------
			##	Time Display  |  -Format
			##	------------  |  ------------
			##	22:00         |  HH:mm
			####################################################################################################
			#[string]$LogFileDate = (Get-Date -Format "$($ReportAndLogSettings.LogFileDateFormat)").ToString()
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name LogFileDate -Value (Get-Date -Format "$($ReportAndLogSettings.LogFileDateFormat)").ToString()

			#[string]$RptFileDate = (Get-Date -Format "$($ReportAndLogSettings.ReportFileDateFormat)").ToString()
			New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name RptFileDate -Value (Get-Date -Format "$($ReportAndLogSettings.ReportFileDateFormat)").ToString()

			#region Log Variables
				#[string]$Script:FullLogDirPath = ($ReportAndLogSettings.LogDirPath).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name LogDirPath -Value ([System.IO.Path]::GetFullPath(($ReportAndLogSettings.LogDirPath).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)))

				if ($ScrDebug)
				{
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FullLogDirPath -Value "$($LogDirPath)\ScrDebug"
				}
				else
				{
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FullLogDirPath -Value "$($LogDirPath)"
				}

				if (!(Test-Path -Path "$($FullLogDirPath)" -PathType Container))
				{
					New-Item -Path "$($FullLogDirPath)" -ItemType Directory | Out-Null
				}

				#[string]$LogFileName = ($ReportAndLogSettings.BaseLogNameFormat).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[LogFileDateFormat]', $LogFileDate).Replace('[PartialReportAndLogName]', $PartialReportAndLogName).Replace('[DataCenterCity]', $DataCenterCity)
				#[string]$Script:FullPathMainScriptLogFile = "$($FullLogDirPath)\$($LogFileName)"
				#[string]$Script:LogEntryDateFormat = $ReportAndLogSettings.LogEntryDateFormat

				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name LogFileName -Value ($ReportAndLogSettings.BaseLogNameFormat).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[LogFileDateFormat]', $LogFileDate).Replace('[PartialReportAndLogName]', $PartialReportAndLogName).Replace('[DataCenterCity]', $DataCenterCity).Replace(' ', '')
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FullPathMainScriptLogFile -Value "$($FullLogDirPath)\$($LogFileName)"
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name LogEntryDateFormat -Value $ReportAndLogSettings.LogEntryDateFormat
			#endregion Log Variables

			#region Report Variables
				#[string]$Script:RptDirPath = ($ReportAndLogSettings.ReportDirPath).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name RptDirPath -Value ([System.IO.Path]::GetFullPath(($ReportAndLogSettings.ReportDirPath).Replace('[ScriptDirectory]', $ScriptDirectory).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[DataCenterCity]', $DataCenterCity)))

				if ($ScrDebug)
				{
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FullRptDirPath -Value "$($RptDirPath)\ScrDebug"
				}
				else
				{
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name FullRptDirPath -Value "$($RptDirPath)"
				}

				if (!(Test-Path -Path "$($FullRptDirPath)" -PathType Container))
				{
					New-Item -Path "$($FullRptDirPath)" -ItemType Directory | Out-Null
				}

				#[string]$RptFileName = ($ReportAndLogSettings.BaseReportNameFormat).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[ReportFileDateFormat]', $RptFileDate).Replace('[PartialReportAndLogName]', $PartialReportAndLogName).Replace('[DataCenterCity]', $DataCenterCity)
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name RptFileName -Value ($ReportAndLogSettings.BaseReportNameFormat).Replace('[ScriptNameNoExt]', $ScriptNameNoExt).Replace('[ReportFileDateFormat]', $RptFileDate).Replace('[PartialReportAndLogName]', $PartialReportAndLogName).Replace('[DataCenterCity]', $DataCenterCity).Replace(' ', '')

				#[string]$HtmlRptFile = "$($FullRptDirPath)\$($RptFileName)"
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name HtmlRptFile -Value "$($FullRptDirPath)\$($RptFileName)"

				if ($CreateCsv)
				{
					#[string]$CsvRptFile	= [System.IO.Path]::ChangeExtension($HtmlRptFile, 'csv')
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name CsvRptFile -Value ([System.IO.Path]::ChangeExtension($HtmlRptFile, 'csv'))
				}

				if ($CreateXlsx)
				{
					#[string]$XlsxRptFile = [System.IO.Path]::ChangeExtension($HtmlRptFile, 'xlsx')
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name XlsxRptFile -Value ([System.IO.Path]::ChangeExtension($HtmlRptFile, 'xlsx'))
				}

				if ($ExportCsvForBMC)
				{
					#[string]$BMCRptFile	= (($HtmlRptFile).Replace('.html', '_BMC.csv'))
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name BMCRptFile -Value (($HtmlRptFile).Replace('.html', '_BMC.csv'))
					#New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name BMCAgentInstallTaskScript -Value (($HtmlRptFile).Replace('.html', '_BMC.cmd'))
				}

				if ($ScrDebug)
				{
					#[string]$MS15034RptFile	= (($HtmlRptFile).Replace('.html', '_MS15-034.csv'))
					New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name MS15034RptFile -Value (($HtmlRptFile).Replace('.html', '_MS15-034.csv'))
				}
			#endregion Report Variables
		#endregion Reports and Logs Variables

		#region BMC-BladeLogic Relevant Export Sections
			#region LegacyPatchWindowCodes
				$Script:LegacyPatchWindowCodes = $XmlConfig.ScriptConfig.BladeLogicSettings.LegacyPatchWindowCodes.LegacyPatchWindow

				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name LegacyPatchWindows -Value @()

				foreach ($LegacyPatchWindow in $LegacyPatchWindowCodes)
				{
					$tmpLegacyPatchWindow = New-Object -TypeName PSObject -Property @{
						'Environment' = $LegacyPatchWindow.Environment
						'PatchCode' = $LegacyPatchWindow.PatchCode
						'Priority' = $LegacyPatchWindow.Priority
						'DataCenterCity' = $LegacyPatchWindow.DataCenterCity
						'Description' = $LegacyPatchWindow.Description
					}

					$LegacyPatchWindows += $tmpLegacyPatchWindow
				}
			#endregion LegacyPatchWindowCodes
			
			#region Environments
				$Script:Environments = $XmlConfig.ScriptConfig.BladeLogicSettings.Environments.Environment

				## Array
				#New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name BMCEnvironments -Value @()
				## [Hashtable]
				New-Variable -Scope Script -Option AllScope -Visibility Public -Force -Name BMCEnvironments -Value @{}

				foreach ($Environment in $Environments)
				{
					## Array
					#$tmpEnvironment = New-Object -TypeName PSObject -Property @{
					#	'EnvName' = $Environment.EnvName
					#	'EnvCode' = $Environment.EnvCode
					#}
					#
					#$BMCEnvironments += $tmpEnvironment

					## [Hashtable]
					$BMCEnvironments.Add($Environment.EnvName, $Environment.EnvCode)
				}
			#endregion Environments
			## May be more to come...
		#endregion BMC-BladeLogic Relevant Export Sections
	#endregion Load Settings File

	if ($XlsxCreationEngine -eq 'MSExcel')
	{
		## Requires Microsoft Excel 2010 or Newer...
		#New-Variable -Scope Script -Option Constant -Name Delimiter -Value "`n`r"
		#New-Variable -Scope Script -Option Constant -Name ScanErrorThreshold -Value 3
		#New-Variable -Scope Script -Option Constant -Name XlNumFmtDate -Value '[$-409]mm/dd/yyyy h:mm:ss AM/PM;@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtDate -Value '[$-409]mm/dd/yyyy hh:mm:ss;@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtTime -Value '[$-409]h:mm:ss AM/PM;@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtText -Value '@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtNumberGeneral -Value '0;@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtNumberS0 -Value '#,##0;@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtNumberS2 -Value '#,##0.00;@'
		New-Variable -Scope Script -Option Constant -Name XlNumFmtNumberS3 -Value '#,##0.000;@'
	}
	elseif ($XlsxCreationEngine -eq 'OpenXML')
	{
		## Open XML specific variables...
	}
	else
	{
		## Other App specific variables...
	}

	if ($ScheduleAs)
	{
		if (!$SchTskName)
		{
			throw 'The following parameter is required when setting the "ScheduleAs" parameter: "SchTskName"'
		}
	}

	$CutOffDate = (Get-Date).AddDays(-$DisableServersOlderThan)
	$Script:OldestCreateDate = New-Object DateTime($CutOffDate.Year, $CutOffDate.Month, $CutOffDate.Day, $CutOffDate.Hour, $CutOffDate.Minute, $CutOffDate.Second)

	## Resolve the FQDN of the server that this script is running on. It will be added to the end of the report so future users will know where to go to find it.
	#[string]$Script:RptSvrFqdn = ("$(${Env:COMPUTERNAME}).$(${Env:USERDNSDOMAIN})").ToLower()
	[string]$Script:RptSvrFqdn = ([System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName).ToLower()

	[int]$Script:ServerCount = 0

	## Maximum number of threads to create
	$MaxThreads = 20

	## Sleep Timer in Milliseconds:
	##		500		= 1/2 second
	##		1000	= 1 second
	##		2000	= 2 seconds
	$SleepTimer = 2000
#endregion Script Variable Initialization

#region User Defined Functions
	Function ConvertTo-ClearText()
	{
		<#
			.SYNOPSIS
				Decrypt a Secure String.

			.DESCRIPTION
				The ConvertTo-ClearText function does the following:

				Decrypt a secure string and return a clear text value.

			.PARAMETER SecureString
				A Secure String generated by the cmdlet 'ConvertTo-SecureString'.

			.EXAMPLE
				$EncString = '01000000d08c9ddf0115d1118c7a00c04fc297eb0100000079c701f6f261884c9a445ff1bb96702c0000000002000000000003660000c000000010000000ce7f213abf14167193743b4ae0a8258c0000000004800000a000000010000000006067c9c6d418dac3781c87fee685ad4000000082f355ac904c3f00f56290622bd539ad27f5ea705cb05ef7e879d4014cee100d5a7974a01e0ee1398dd52c049a14077b2cf77e6b46022187f09b84410f86b54d140000007f657ddd977514338e461f2d7c887b7c958f4e59'
				$SecString = $EncString | ConvertTo-SecureString
				ConvertTo-ClearText -SecureString $SecString
				Takes the encrypted string and converts it to a Secure String. It is then passed in to the function to decrypt the Secure String and return its clear text value. The value returned for this example is:
				My Sup3r $ecr3t S3cure $tr1ng

			.NOTES
				Author: Bill Campbell
				Version: 1.1.0.1
				Release: 2014-Dec-18
				Requirements: PowerShell Version 2.0

				An example of how to accomplish this was found at the link provided. It has been changed and updated for newer versions.
				The Original Name I gave to the function was 'fn_DecryptSecureString'.

			.LINK
				http://huddledmasses.org/jaykul/extending-strings-and-secure-strings-in-powershell/
		#>
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = 'Medium'
		)]
		#region Function Parameters
			Param(
				[Parameter(
					Position		= 0
					, Mandatory		= $true
					, ValueFromPipeline = $true
					#, HelpMessage	= 'The Secure String to Decrypt.'
				)]$SecureString
			)
		#endregion Function Parameters

		## Example from: http://huddledmasses.org/jaykul/extending-strings-and-secure-strings-in-powershell/
		#$SecureString = Read-Host -Prompt 'Enter your Secure String value' -AsSecureString
		$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
		$DecryptedSecureString = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
		[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

		return $DecryptedSecureString
	}

	Function Get-DateLastXDayOfMonth()
	{
		<#
			.SYNOPSIS
				Calculates the date of the last specified weekday of any given month, in any given year and returns the result to you as a date time type/object.

			.DESCRIPTION
				The Get-DateLastXDayOfMonth function does the following:

				Calculates the date of the last Monday/Tuesday/Wednesday..etc of any given month, in any given year and returns the result to you as a date time type/object.

			.PARAMETER DayName
				The name of the day of the week to be found.

			.PARAMETER Month
				The number which indicates the month of the year. If left blank it will use the current month.
				12 = December

			.PARAMETER Year
				The year in which you want to search. If left blank it will use the current year.

			.EXAMPLE
				Get-DateLastXDayOfMonth -DayName 'Monday'

				Description
				-----------
					Calculates the date of the last Monday in the current month of the current year.

			.EXAMPLE
				Get-DateLastXDayOfMonth -DayName 'Wednesday' -Month 5 -Year 2013

				Description
				-----------
					Calculates the date of the last Wednesday in the month of May in the current year 2013.
					It returns 'Wednesday, May 29, 2013 1:26:10 PM'
		#>
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = 'Medium'
		)]
		#region Function Parameters
			Param(
				[Parameter(
					Position = 0
					, Mandatory = $true
				)][string]$DayName
				, [int]$Month
				, [int]$Year
			)
		#endregion Function Parameters

		begin
		{
			if ($Month)
			{
				if ($Year)
				{
					$LastDayOfMonth = (Get-Date -Year $Year -Month $Month -Day 1).AddMonths(1).AddDays(-1)
				}
				else
				{
					$LastDayOfMonth = (Get-Date -Year (Get-Date).Year -Month $Month -Day 1).AddMonths(1).AddDays(-1)
				}
			}
			else
			{
				$LastDayOfMonth = (Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day 1).AddMonths(1).AddDays(-1)
			}

			$Answer = $null
		}
		process
		{
			if ($LastDayOfMonth.DayOfWeek -eq $DayName)
			{
				$Answer = $LastDayOfMonth
			}
			else
			{
				while($Answer -eq $null)
				{
					$LastDayOfMonth = $LastDayOfMonth.AddDays(-1)

					if ($LastDayOfMonth.DayOfWeek -eq $DayName)
					{
						$Answer = $LastDayOfMonth
					}
				}
			}
		}
		end
		{
			return $Answer
		}
	}

	Function Disable-OldComputerObjects()
	{
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = "Medium"
		)]
		#region Function Parameters
			Param(
				[Parameter(
					Position = 0
					, Mandatory = $true
					#, HelpMessage = 'Active Directory DNS Domain Name.'
				)][string]$ADDSDomain
				, [Parameter(
					Position = 1
					, Mandatory = $true
					#, HelpMessage = 'Sets the scope that the script will use to query computer objects from Active Directory ("AllWindows", "Server", "Workstation").'
				)]
				[ValidateSet('AllWindows', 'Server', 'Workstation')]
				[string]$ComputerScope
				, [Parameter(
					Position = 2
					, Mandatory = $false
					#, HelpMessage = 'Sets the number of days that a computer object may go with no change to its Active Directory object. Used in the query for computer objects from Active Directory (If left blank it will calculate the number of days back to January 1 of the current year or use 90 days, which ever is larger. Minimum of 90 days will be used regardless.).'
				)][int]$DaysNoChange
				, [Parameter(
					Position = 3
					, Mandatory = $true
					#, HelpMessage = 'Username of the user account performing the work.'
				)][string]$SecUserName
				, [Parameter(
					Position = 4
					, Mandatory = $true
					#, HelpMessage = 'Secure String Password of the user account performing the work.'
				)]$SecPassword
				, [Parameter(
					Position = 5
					, Mandatory = $false
					#, ValueFromPipeline	= $false
					#, HelpMessage = 'The fnDebug switch enables/disables custom debugging code in the Function ($true/$false). Defaults to FALSE.'
				)][switch]$fnDebug = $false
			)
		#endregion Function Parameters
		## Function Usage:
		##	Disable-OldComputerObjects -ADDSDomain 'ecollege.net' -ComputerScope 'Server' -DaysNoChange $DisableServersOlderThan -SecUserName $SecUsrNam -SecPassword $SecPasswd -fnDebug

		#region Logfile Header
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t----------------------------------------------------------------------------------------------------"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t -  $($MyInvocation.InvocationName) -"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Function Start Time   : $([char]34)$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())$([char]34)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Computers Scope       : $([char]34)$($ComputerScope)$([char]34)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t AD DS Domain Checked  : $([char]34)$($ADDSDomain)$([char]34)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Days Without Change   : $([char]34)$($DaysNoChange)$([char]34)"

			if ($fnDebug)
			{
				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t ScrDebug State        : $([char]34)$($ScrDebug.ToString())$([char]34) NO Changes will be applied"
			}
			else
			{
				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t ScrDebug State        : $([char]34)$($ScrDebug.ToString())$([char]34) Changes will be applied"
			}

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t----------------------------------------------------------------------------------------------------"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tWhen Created       `tLast Changed       `tAction                  `tResult  `tdNSHostName                                  `tDN                                                                                                                      `tErrors"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------`t-------------------`t------------------------`t--------`t---------------------------------------------`t------------------------------------------------------------------------------------------------------------------------`t--------"
		#endregion Logfile Header

		[string]$ComputersNodNSHostNameNames = [string]::Empty
		[int]$ComputersNodNSHostNameCount = 0
		[int]$ComputersDisabledCount = 0
		[int]$ComputersTotalCount = $OldComputers.Count

		if ($DNStoNBnames[$DNSDomain])
		{
			$NBName = $DNStoNBnames[$DNSDomain]
		}
		else
		{
			throw 'NETBIOS name not found in the Hash Table "DNStoNBnames"'
		}

		$ConnectAcct = "$($NBName)\$($SecUserName)"
		$GetQADActivity = "Retreiving a list of $([char]34)$($ComputerScope)$([char]34) Computer Objects from domain $([char]34)$($DNSDomain)$([char]34). . ."

		if ($ComputerScope -eq 'Workstation')
		{
			$OldComputers = @(Get-QADComputer -Activity $GetQADActivity -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPassword | Where-Object {
				((((Get-Date) - $_.whenChanged).Days) -gt $DaysNoChange) -and
				(
					#($_.operatingSystem -like "Windows*9*") -or			# "Windows 9x" Operating Systems
					($_.operatingSystem -like "Windows 2000 Pro*") -or	# "Windows 2000 Pro" Operating Systems
					($_.operatingSystem -like "Windows*XP*") -or		# "Windows XP" Operating Systems
					($_.operatingSystem -like "Windows*Vis*") -or		# "Windows Vista" Operating Systems
					($_.operatingSystem -like "Windows 7*") -or			# "Windows 7" Operating Systems
					($_.operatingSystem -like "Windows 8*") -or			# "Windows 8" Operating Systems
					($_.operatingSystem -like "Windows*Dev*")			# "Windows Developer Preview" Operating Systems
				) -and
				($_.AccountIsDisabled -eq $false)
			})
		}
		elseif ($ComputerScope -eq 'Server')
		{
			$OldComputers = @(Get-QADComputer -Activity $GetQADActivity -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPassword | Where-Object {
				((((Get-Date) - $_.whenChanged).Days) -gt $DaysNoChange) -and
				($_.operatingSystem -like "Windows*Server*") -and
				($_.AccountIsDisabled -eq $false)
			})
		}
		else	#if ($ComputerScope -eq 'AllWindows')
		{
			$OldComputers = @(Get-QADComputer -Activity $GetQADActivity -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPassword | Where-Object {
				((((Get-Date) - $_.whenChanged).Days) -gt $DaysNoChange) -and
				($_.operatingSystem -like "Windows*") -and
				## Some of the older NetApp devices show up as "Windows NT" and we don't want to disable them.
				($_.operatingSystem -ne "Windows NT") -and
				($_.AccountIsDisabled -eq $false)
			})
		}

		$ComputersTotalCount += $OldComputers.Count

		foreach ($Comp in $OldComputers)
		{
			if ($Comp.dNSHostName -eq $null)
			{
				$CompCanonicalName = $Comp.CanonicalName
				$CompDNSDomain = $CompCanonicalName.Substring(0, $CompCanonicalName.IndexOf('/')).ToLower()
				$CompName = $CompCanonicalName.Substring($CompCanonicalName.LastIndexOf('/') + 1).ToLower()
				$CompDNSName = "$($CompName).$($CompDNSDomain)"

				if ($ComputersNodNSHostNameNames.Length -lt 1)
				{
					$ComputersNodNSHostNameNames = "$([char]34)$($CompName)$([char]34)"
				}
				else
				{
					$ComputersNodNSHostNameNames += ", $([char]34)$($CompName)$([char]34)"
				}

				## .PadRight(IntTotalWidth, [StrPaddingChar])
				[string]$strPtlLogEntry = "$($Comp.whenCreated)`t$($Comp.whenChanged)`tDisable Computer Account`tActionResult`t$($CompDNSName.PadRight(45, ' '))`t$($Comp.DN.PadRight(120, ' '))"

				$ComputersNodNSHostNameCount++
			}
			else
			{
				[string]$strPtlLogEntry = "$($Comp.whenCreated)`t$($Comp.whenChanged)`tDisable Computer Account`tActionResult`t$($Comp.dNSHostName.ToLower().PadRight(45, ' '))`t$($Comp.DN.PadRight(120, ' '))"
			}

			if ($fnDebug)
			{
				Disable-QADComputer -Identity $Comp.DN -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPassword -Confirm:$false -WhatIf
			}
			else
			{
				Disable-QADComputer -Identity $Comp.DN -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPassword -Confirm:$false
			}

			if ($?)
			{
				if ($fnDebug)
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`twSuccess")
				}
				else
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`tSuccess ")
				}

				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t$($strPtlLogEntry)"

				$ComputersDisabledCount++
			}
			else
			{
				if ($fnDebug)
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`twFAILURE") + "`t$($Error[0].Exception.Message)"
				}
				else
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`tFAILURE ") + "`t$($Error[0].Exception.Message)"
				}

				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t$($strPtlLogEntry)"
			}
		}

		if ([string]::IsNullOrEmpty($ComputersNodNSHostNameNames))
		{
			[string]$ComputersNodNSHostNameNames = 'None found missing this attribute.'
		}

		#region Logfile Footer
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t------------------------------------:------------------------------------"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Total Computer Objects from AD     : $($ComputersTotalCount)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Total Computers No AD dNSHostName  : $($ComputersNodNSHostNameCount)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Computers No AD dNSHostName Names  : $($ComputersNodNSHostNameNames)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Total Computer Objects Disabled    : $($ComputersDisabledCount)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t------------------------------------:------------------------------------"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Function $($MyInvocation.InvocationName) End Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
		#endregion Logfile Footer
	}

	Function Get-ListOfWindowsServersFromActiveDirectory
	{
		[CmdletBinding(
			SupportsShouldProcess = $true
		)]
		#region Function Paramters
			Param()
		#endregion Function Paramters

		begin
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tStart Function: $($MyInvocation.InvocationName)"

			$ServerList = @()
			$SearchScope = 'Subtree'
			## userAccountControl=4098 means the computer account is disabled. We don't want those.
			## userAccountControl=4130 means the computer account is disabled and that it does not require a machine password. We don't want those either.
			## The following DOES NOT include the Windows based hypervisor 'Hyper-V Server'.
			#$LDAPFilter = "(&(ObjectClass=computer)(operatingSystem=Windows*Server*)(!userAccountControl=4098)(!userAccountControl=4130))"

			## The following DOES include 'Hyper-V Server'.
			#$LDAPFilter = "(&(ObjectClass=computer)(operatingSystem=*Server*)(!userAccountControl=4098)(!userAccountControl=4130))"

			## The following Builds on the filter above, but removes "Failover cluster virtual network name account" objects.
			## A.K.A. "Cluster Name Object (CNO)" or "Virtual Computer Objects (VCO)"
			## With some changes to the way Failover clusters are built in Server 2012 R2, this is now necessary. Otherwise the report counts these objects when it should not.
			$LDAPFilter = "(&(ObjectClass=computer)(operatingSystem=*Server*)(!servicePrincipalName=MSClusterVirtualServer*)(!userAccountControl=4098)(!userAccountControl=4130))"

			$SvrPropList = @('Name', 'dNSHostName', 'operatingSystem', 'operatingSystemServicePack', 'whenChanged', 'whenCreated', 'DN')
		}
		process
		{
			foreach ($DNSDomain in $SplaDnsDomains)
			{
				if (($InScopeDNSDomains | Where-Object { $_.DNSName -eq $DNSDomain } | Select-Object IsCentrallyManaged).IsCentrallyManaged -like 'TRUE')
				{
					[bool]$DomainIsCentrallyManaged = $true
				}
				else
				{
					[bool]$DomainIsCentrallyManaged = $false
				}

				#$DomainIsCentrallyManaged = ($InScopeDNSDomains | Where-Object { $_.DNSName -eq $DNSDomain } | Select-Object IsCentrallyManaged).IsCentrallyManaged

				if ($DNStoNBnames[$DNSDomain])
				{
					$NBName = $DNStoNBnames[$DNSDomain]
				}
				else
				{
					throw "NETBIOS name not found in the Hash Table 'DNStoNBnames'"
				}

				if ($DomainIsCentrallyManaged)
				{
					$CMNBName = ($DomainUserInfo | Where-Object { ($_.IsCentralManagementCred -like 'TRUE') } | Select-Object NbDomain).NbDomain
					$SecUsrNam = ($DomainUserInfo | Where-Object { ($_.IsCentralManagementCred -like 'TRUE') } | Select-Object UserName).UserName
					$EncPasswd = ($DomainUserInfo | Where-Object { ($_.IsCentralManagementCred -like 'TRUE') } | Select-Object EncPassword).EncPassword
					$SecPasswd = $EncPasswd | ConvertTo-SecureString -Key $Script:Key #(1..32)

					$ConnectAcct = "$($CMNBName)\$($SecUsrNam)"
				}
				else
				{
					#$SecUsrNam = ($DomainUserInfo | Where-Object { $_.NbDomain -eq $NBName } | Select-Object UserName).UserName
					#$EncPasswd = ($DomainUserInfo | Where-Object { $_.NbDomain -eq $NBName } | Select-Object EncPassword).EncPassword
					$SecUsrNam = ($DomainUserInfo | Where-Object { ($_.NbDomain -eq $NBName) -and ($_.DnsDomain -eq $DNSDomain) } | Select-Object UserName).UserName
					$EncPasswd = ($DomainUserInfo | Where-Object { ($_.NbDomain -eq $NBName) -and ($_.DnsDomain -eq $DNSDomain) } | Select-Object EncPassword).EncPassword
					$SecPasswd = $EncPasswd | ConvertTo-SecureString -Key $Script:Key #(1..32)
					
					#if ((!$DomainIsCentrallyManaged) -and ($SecUsrNam -like "*\*"))
					if ($SecUsrNam -like "*\*")
					{
						$ConnectAcct = "$($SecUsrNam)"
					}
					else
					{
						$ConnectAcct = "$($NBName)\$($SecUsrNam)"
					}
				}

				#### $ConnectAcct = "$($NBName)\$($SecUsrNam)"
				#if ($SecUsrNam -like "*\*")
				#{
				#	$ConnectAcct = "$($SecUsrNam)"
				#}
				#else
				#{
				#	$ConnectAcct = "$($NBName)\$($SecUsrNam)"
				#}

				if ($DisableOldADServerObjects)
				{
					## Gets the date of the last Monday of the current Month of the current year. This is when the report is due.
					$ReportDueDate = Get-DateLastXDayOfMonth -DayName 'Monday'
					## Gets todays date (Or the current days date).
					$ReportRunDate = Get-Date

					## Checks to see if the Day value of the day the report is due, equals the Day value of today.
					## We only need to run this once a month and not every time the script is tested for changes or additions.
					if ($ReportDueDate.Day -eq $ReportRunDate.Day)
					{
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Running Function 'Disable-OldComputerObjects'"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------"

						## This is only commented out until the bug with machine passwords on a domain with Server 2003 and Server 2012 domain controllers has been fixed by Microsoft and the patch applied in the environments. We also need to wait to role back the GPO that was pushed out.
						Disable-OldComputerObjects -ADDSDomain $DNSDomain -ComputerScope 'Server' -DaysNoChange $DisableServersOlderThan -SecUserName $SecUsrNam -SecPassword $SecPasswd #-fnDebug
					}
					else
					{
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Function 'Disable-OldComputerObjects' does not need to be run."
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------"
					}
				}

				if ($ScrDebug)
				{
					#Write-Host -ForegroundColor Yellow ''
					Write-Host -ForegroundColor Yellow " Variable Names and Values used by the Cmdlet $([char]34)Get-QADObject$([char]34)... "
					Write-Host -ForegroundColor Yellow '-------------------------------------------------------------------------------------'
					Write-Host -ForegroundColor Yellow ' Variable Name      : Value'
					Write-Host -ForegroundColor Yellow '------------------- : ---------------------------------------------------------------'
					Write-Host -ForegroundColor Yellow " DNSDomain          : $($DNSDomain)"
					Write-Host -ForegroundColor Yellow " SearchScope        : $($SearchScope)"
					Write-Host -ForegroundColor Yellow " LDAPFilter         : $($LDAPFilter)"
					Write-Host -ForegroundColor Yellow " SvrPropList        : $($SvrPropList)"
					Write-Host -ForegroundColor Yellow " NBName             : $($NBName)"
					Write-Host -ForegroundColor Yellow " SecUsrNam          : $($SecUsrNam)"
					Write-Host -ForegroundColor Yellow " SecPasswd          : $($SecPasswd)"
					Write-Host -ForegroundColor Yellow " ConnectAcct        : $($ConnectAcct)"
					Write-Host -ForegroundColor Yellow '-------------------------------------------------------------------------------------'
					Write-Host -ForegroundColor Yellow ''

					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Names and Values used by the Cmdlet $([char]34)Get-QADObject$([char]34)... "
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Name      : Value"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t------------------- : ---------------------------------------------------------------"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t DNSDomain          : $($DNSDomain)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SearchScope        : $($SearchScope)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t LDAPFilter         : $($LDAPFilter)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvrPropList        : $($SvrPropList)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t NBName             : $($NBName)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SecUsrNam          : $($SecUsrNam)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SecPasswd          : $($SecPasswd)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t ConnectAcct        : $($ConnectAcct)"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t-------------------------------------------------------------------------------------"
					Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
				}

				##Connect-QADService -Service $DNSDomain -ConnectionAccount "$($NBName)\$($SecUsrNam)" -ConnectionPassword $SecPasswd #-UseGlobalCatalog
				$SvrRslts = @(Get-QADObject -Service $DNSDomain -SearchScope $SearchScope -LdapFilter $LDAPFilter -IncludedProperties $SvrPropList -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPasswd)
				##Disconnect-QADService

				foreach ($SvrObject in $SvrRslts)
				{
					[string]$SrvrName		= $SvrObject.Name
					[string]$SrvrDnsName	= $SvrObject.dNSHostName
					if ($SrvrDnsName.Length -le $SrvrName.Length)
					{
						[string]$SrvrDnsName = "$($SrvrName).$($DNSDomain)"
					}
					$SrvrDnsName			= $SrvrDnsName.ToLower()

					[string]$SrvrOsName		= $SvrObject.operatingSystem #.ToString()
					if ($SrvrOsName -like "*Windows*")
					{
						#$SrvrOsName			= ($SrvrOsName).Replace(' (R)', '').Replace('(R)', '').Replace('®', '').Replace('?', '').Replace(',', '').Replace('Microsoft ', '').Replace('Windows ', '').Replace('Standard', 'Std').Replace('Enterprise', 'Ent').Replace('Datacenter', 'DC').Replace('Edition', '')	#.Replace('Server', 'Svr')
						$SrvrOsName			= ($SrvrOsName).Replace(' (R)', '').Replace('(R)', '').Replace('®', '').Replace('?', '').Replace(',', '').Replace('Microsoft ', '').Replace('Windows ', '').Replace('Edition', '').Replace('Standard x64 ', 'Std x64').Replace('Standard ', 'Std').Replace('Standard', 'Std').Replace('Enterprise x64 ', 'Ent x64').Replace('Enterprise ', 'Ent').Replace('Enterprise', 'Ent').Replace('Datacenter ', 'DC').Replace('Datacenter', 'DC')
					}

					[string]$SrvrOsSpLvl	= $SvrObject.operatingSystemServicePack #.ToString()
					if ($SrvrOsSpLvl -like "Service Pack*")
					{
						[string]$SrvrOsSpLvl = ($SrvrOsSpLvl).Replace('Service Pack ', 'SP')
					}
					else
					{
						[string]$SrvrOsSpLvl = 'NONE'
					}

					$SrvrChangedOn = $SvrObject.whenChanged
					$SrvrCreatedOn = $SvrObject.whenCreated

					$objSvrName = New-Object PSObject #New-Object System.Object
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrName'			-Value $SrvrName
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrDnsName'		-Value $SrvrDnsName
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrOsName'		-Value $SrvrOsName
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrOsSpLvl'		-Value $SrvrOsSpLvl
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrChangedOn'		-Value $SrvrChangedOn
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrCreatedOn'		-Value $SrvrCreatedOn
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrDnsDomain'		-Value $DNSDomain
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SrvrNBDomain'		-Value $NBName
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SecUsrNam'			-Value $SecUsrNam
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'SecPasswd'			-Value $SecPasswd
					$objSvrName | Add-Member -MemberType NoteProperty -Name 'ConnectAcct'		-Value $ConnectAcct
					$ServerList	+= $objSvrName

					if ($ScrDebug)
					{
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Names and Values to be passed on for further processing..."
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t----------------------------------------------------------------------------"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Variable Name      : Value"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t------------------- : ------------------------------------------------------"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrName           : $($SrvrName)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrDnsName        : $($SrvrDnsName)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrOsName         : $($SrvrOsName)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrOsSpLvl        : $($SrvrOsSpLvl)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrChangedOn      : $($SrvrChangedOn)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrCreatedOn      : $($SrvrCreatedOn)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrDnsDomain      : $($DNSDomain)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SrvrNBDomain       : $($NBName)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SecUsrNam          : $($SecUsrNam)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SecPasswd          : $($SecPasswd)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t ConnectAcct        : $($ConnectAcct)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t----------------------------------------------------------------------------"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
					}
					Clear-Variable -Name Srvr* -Confirm:$false
				}
				Clear-Variable -Name SvrRslts -Confirm:$false
			}
		}
		end
		{
			$SrvrList = $ServerList | Sort-Object -Property 'SrvrNBDomain', 'SrvrName'

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t End Function: $($MyInvocation.InvocationName)"
			return $SrvrList
		}
	}

	Function Export-ReportToExcel()
	{
		<#
			.SYNOPSIS
				Writes an Excel file containing the information from an Inventory object.

			.DESCRIPTION
				The Export-ReportToExcel function uses COM Interop to write an Excel file containing the Inventory information. Microsoft Excel 2007 or higher must be installed in order to write the Excel file.

			.PARAMETER Inventory
				A [System.Management.Automation.PSCustomObject] Inventory object.

			.PARAMETER Path
				Specifies the path where the Excel file will be written. This is a fully qualified path to a .xlsx file. If not specified then the file is named 'License Inventory - [Year][Month][Day][Hour][Minute].xlsx' and is written to your 'My Documents' folder.

			.PARAMETER ColorTheme
				An Office Theme Color to apply to each worksheet. If not specified or if an unknown theme color is provided the default 'Office' theme colors will be used.
				Office 2013 theme colors include: Aspect, Blue Green, Blue II, Blue Warm, Blue, Grayscale, Green Yellow, Green, Marquee, Median, Office, Office 2007 - 2010, Orange Red, Orange, Paper, Red Orange, Red Violet, Red, Slipstream, Violet II, Violet, Yellow Orange, Yellow
				Office 2010 theme colors include: Adjacency, Angles, Apex, Apothecary, Aspect, Austin, Black Tie, Civic, Clarity, Composite, Concourse, Couture, Elemental, Equity, Essential, Executive, Flow, Foundry, Grayscale, Grid, Hardcover, Horizon, Median, Metro, Module, Newsprint, Office, Opulent, Oriel, Origin, Paper, Perspective, Pushpin, Slipstream, Solstice, Technic, Thatch, Trek, Urban, Verve, Waveform
				Office 2007 theme colors include: Apex, Aspect, Civic, Concourse, Equity, Flow, Foundry, Grayscale, Median, Metro, Module, Office, Opulent, Oriel, Origin, Paper, Solstice, Technic, Trek, Urban, Verve

			.PARAMETER ColorScheme
				The color theme to apply to each worksheet. Valid values are 'Light', 'Medium', and 'Dark'. If not specified then 'Medium' is used as the default value.

			.PARAMETER ParentProgressId
				If the caller is using Write-Progress then all progress information will be written using ParentProgressId as the ParentID

			.PARAMETER RptTitleHeading
				The Report Title or Heading to be used in the output.

			.PARAMETER SplaAdDnsDomains
				The PowerShell Object array that holds all of the "In Scope" Active Directory DNS domain names.

			.PARAMETER SvrCountPhy
				The count of Physical servers.

			.PARAMETER SvrCountVm
				The count of Virtual Machine servers.

			.PARAMETER SvrCountEE
				The count of Windows servers that have encountered errors.

			.PARAMETER SvrCountNR
				The count of Windows servers that are not responding.

			.PARAMETER SvrCountTotal
				The total count of Windows servers.

			.PARAMETER SvrCountNew
				The count of newly built Windows servers.

			.EXAMPLE
				Export-ReportToExcel -Inventory $InventoryPSobject

				Description
				-----------
				Write an inventory using $InventoryPSobject.
				The Excel workbook will be written to your 'My Documents' folder.
				The Office color theme and Medium color scheme will be used by default.

			.EXAMPLE
				Export-ReportToExcel -Inventory $InventoryPSobject -Path 'C:\License Inventory.xlsx'

				Description
				-----------
				Write an inventory using $InventoryPSobject.
				The Excel workbook will be written to your 'C:\License Inventory.xlsx'.
				The Office color theme and Medium color scheme will be used by default.

			.EXAMPLE
				Export-ReportToExcel -Inventory $InventoryPSobject -ColorTheme Blue -ColorScheme Dark

				Description
				-----------
				Write an inventory using $InventoryPSobject.
				The Excel workbook will be written to your 'My Documents' folder.
				The Blue color theme and Dark color scheme will be used.

			.NOTES
				Blue and Green are nice looking Color Themes for Office 2013.
				Waveform is a nice looking Color Theme for Office 2010.
		#>
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = 'Medium'
		)]
		#region Function Paramters
			Param(
				[Parameter(
					Mandatory = $true
					, ValueFromPipeline = $true
					#, HelpMessage = 'The PowerShell Object array that holds all of the collected information from the "In Scope" Active Directory servers.'
				)]
				#[System.Management.Automation.PSCustomObject]
				$Inventory
				, [Parameter(
					Mandatory = $false
					#, HelpMessage = 'Specifies the path where the Excel file will be written. This is a fully qualified path to a .xlsx file. If not specified then the file is named "License Inventory - [Year]-[Month]-[Day]-[Hour]-[Minute].xlsx" and is written to your "My Documents" folder.'
				)]
				[Alias('File')]
				[ValidateNotNullOrEmpty()]
				[string]$Path = (Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)) -ChildPath ('License Inventory - ' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '.xlsx'))
				, [Parameter(
					Mandatory = $false
				)]
				[Alias('Theme')]
				[string]$ColorTheme = 'Office'
				, [Parameter(
					Mandatory = $false
				)]
				[ValidateSet('Light', 'Medium', 'Dark')]
				[string]$ColorScheme = 'Medium'
				, [Parameter(
					Mandatory = $false
				)]
				[ValidateNotNull()]
				[Int32]$ParentProgressId = -1
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The Report Title or Heading to be used in the output.'
				)][string]$RptTitleHeading
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The PowerShell Object array that holds all of the "In Scope" Active Directory DNS domain names.'
				)]
				[ValidateNotNullOrEmpty()]
				$SplaAdDnsDomains
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The count of Physical servers.'
				)][int]$SvrCountPhy
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The count of Virtual Machine servers.'
				)][int]$SvrCountVm
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The count of Windows servers that have encountered errors.'
				)][int]$SvrCountEE
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The count of Windows servers that are not responding.'
				)][int]$SvrCountNR
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The total count of Windows servers.'
				)][int]$SvrCountTotal
				, [Parameter(
					Mandatory = $true
					#, HelpMessage = 'The count of newly built Windows servers.'
				)][int]$SvrCountNew
			)
		#endregion Function Paramters

		begin
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tStart Function: $($MyInvocation.InvocationName)"

			## $Inventory - Is an Array with the following columns in it:
			##	'SrvName', 'SrvDnsName', 'SrvOsName', 'SrvOsSpLvl', 'SrvChangedOn', 'SrvCreatedOn', 'SrvNBDomain', 'SrvIPv4Addr', 'SrvAddlIPv4Addr', 'SrvDataCenter', 'SrvBMCLocationCode', 'SrvSerNum', 'SrvBiosVer', 'SrvMnfctr', 'SrvModel', 'SrvInstOsArch', 'SrvLastBootUpTime', 'SrvUpTime'
			##	if ($IncludeHotFixDetails)
			##		'SrvLastPatchInstallDate', 'SrvUpTimeKB2553549HotFix1Applied', 'SrvUpTimeKB2688338HotFix2Applied', 'SrvUpTimeTotalHotFixApplied', 'SrvKB3042553HotFixApplied'
			##	'MSClusterName', 'SrvRegOsName', 'SrvRegSpLvl'
			##	if ($IncludeMSSQL)
			##		'SQLSvrName', 'SQLSvrIsInstalled', 'SQLProductCaptions', 'SQLEditions', 'SQLVersions', 'SQLSvcDisplayNames', 'SQLInstanceNames', 'SQLIsClustered', 'isSQLClusterNode', 'SQLClusterNames', 'SQLClusterNodes'
			##	if ($IncludeCustomAppInfo)
			##		'SrvBrowserHawkInstall', 'SrvSoftArtisansFileUpInstall'
			##	'SrvPendingReboot', 'SrvPhysMemoryGB', 'SrvProcMnftr', 'SrvSocketCount', 'SrvProcCoreCount', 'SrvLogProcsCount', 'SrvHyperT', 'SrvBMCAgentInstalled', 'SrvBMCAgentStartMode', 'SrvBMCAgentState'
			##	'SrvNetlogonStartMode', 'SrvNetlogonState', 'SrvRemoteRegistryStartMode', 'SrvRemoteRegistryState', 'SrvWinFirewallStartMode', 'SrvWinFirewallState', 'SrvWinUpdateStartMode', 'SrvWinUpdateState'
			##	if ($IncludeInfraDetails)
			##		'SrvProcName', 'SrvDiskTotalInfo', 'SrvDiskUsedInfo', 'SrvDiskFreeInfo', 'SrvDiskTotalAllocated', 'SrvDiskUsedAllocated', 'SrvDiskFreeAllocated'

			#region Excel 2010 Enumerations
				## Excel 2010 Enumerations Reference: http://msdn.microsoft.com/en-us/library/ff838815.aspx
				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlsortorder(v=office.14).ASPX
				##	Member name		: Description
				##	----------------:-------------
				##	xlAscending		: default. Sorts the specified field in ascending order.
				##	xlDescending	: Sorts the specified field in descending order.
				$XlSortOrder = 'Microsoft.Office.Interop.Excel.XlSortOrder' -as [Type]

				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlyesnoguess(v=office.14).aspx
				##	Member name	: Description
				##	------------:-------------
				##	xlGuess		: Excel determines whether there’s a header, and to determine where it is, if there is one.
				##	xlNo		: default. (The entire range should be sorted).
				##	xlYes		: (The entire range should not be sorted).
				$XlYesNoGuess = 'Microsoft.Office.Interop.Excel.XlYesNoGuess' -as [Type]

				## http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.excel.xlhalign%28v=office.14%29.aspx
				##	Member name						: Description
				##	--------------------------------:------------------------------
				##	xlHAlignCenter					: Center
				##	xlHAlignCenterAcrossSelection	: Center across selection
				##	xlHAlignDistributed				: Distribute
				##	xlHAlignFill					: Fill
				##	xlHAlignGeneral					: Align according to data type
				##	xlHAlignJustify					: Justify
				##	xlHAlignLeft					: Left
				##	xlHAlignRight					: Right
				$XlHAlign = 'Microsoft.Office.Interop.Excel.XlHAlign' -as [Type]

				## http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.excel.xlvalign%28v=office.14%29.aspx
				##	Member name			: Description
				##	--------------------:-------------
				##	xlVAlignBottom		: Bottom
				##	xlVAlignCenter		: Center
				##	xlVAlignDistributed	: Distributed
				##	xlVAlignJustify		: Justify
				##	xlVAlignTop			: Top
				$XlVAlign = 'Microsoft.Office.Interop.Excel.XlVAlign' -as [Type]

				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xllistobjectsourcetype(v=office.14).aspx
				##	Member name		: Description
				##	----------------:-------------
				##	xlSrcExternal	: External data source (Microsoft Windows SharePoint Services site).
				##	xlSrcRange		: Microsoft Office Excel range.
				##	xlSrcXml		: XML.
				##	xlSrcQuery		: Query.
				$XlListObjectSourceType = 'Microsoft.Office.Interop.Excel.XlListObjectSourceType' -as [Type]

				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlthemecolor(v=office.14).aspx
				##	Member name						: Description			: Color
				##	--------------------------------:-----------------------:-----------
				##	xlThemeColorDark1				: Dark1					: White
				##	xlThemeColorLight1				: Light1				: Black
				##	xlThemeColorDark2				: Dark2					: Gray-25%
				##	xlThemeColorLight2				: Light2				: Blue-Gray
				##	xlThemeColorAccent1				: Accent1				: Blue (Sky Blue)
				##	xlThemeColorAccent2				: Accent2				: Orange
				##	xlThemeColorAccent3				: Accent3				: Gray-50%
				##	xlThemeColorAccent4				: Accent4				: Gold
				##	xlThemeColorAccent5				: Accent5				: Blue
				##	xlThemeColorAccent6				: Accent6				: Green (Shade of)
				##	xlThemeColorHyperlink			: Hyperlink				: Unvisited website link (Using system default clolor)
				##	xlThemeColorFollowedHyperlink	: Followed hyperlink	: Visited website link (Using system default clolor)
				$XlThemeColor = 'Microsoft.Office.Interop.Excel.XlThemeColor' -as [Type]

				#$RawDataTabColor = $XlThemeColor::xlThemeColorAccent3	# WAS xlThemeColorDark2
				$OverviewTabColor = $XlThemeColor::xlThemeColorDark1	#xlThemeColorLight2	# WAS xlThemeColorLight1
				$ActiveServersTabColor = $XlThemeColor::xlThemeColorAccent6
				$ErrorEncounteredServersTabColor = $XlThemeColor::xlThemeColorAccent4
				$NoResponseServersTabColor = $XlThemeColor::xlThemeColorLight1	# WAS xlThemeColorAccent3
				$CriticalWindowsServicesTabColor = $XlThemeColor::xlThemeColorAccent5
				$MSSQLServersTabColor = $XlThemeColor::xlThemeColorAccent2
				$HotFixDetailsTabColor = $XlThemeColor::xlThemeColorAccent1
				$CustomAppInfoTabColor = $XlThemeColor::xlThemeColorAccent3
				$InfraDetailsDetailsTabColor = $XlThemeColor::xlThemeColorDark2

				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xllinestyle(v=office.14).aspx
				##	Member name		: Description
				##	----------------:-------------
				##	xlContinuous	: Continuous line.
				##	xlDash			: Dashed line.
				##	xlDashDot		: Alternating dashes and dots.
				##	xlDashDotDot	: Dash followed by two dots.
				##	xlDot			: Dotted line.
				##	xlDouble		: Double line.
				##	xlSlantDashDot	: Slanted dashes.
				##	xlLineStyleNone	: No line.
				$XlLineStyle = 'Microsoft.Office.Interop.Excel.XlLineStyle' -as [Type]

				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlbordersindex(v=office.14).aspx
				##	Member name			: Description
				##	--------------------:-------------
				##	xlInsideHorizontal	: Horizontal borders for all cells in the range except borders on the outside of the range.
				##	xlInsideVertical	: Vertical borders for all the cells in the range except borders on the outside of the range.
				##	xlDiagonalDown		: Border running from the upper left-hand corner to the lower right of each cell in the range.
				##	xlDiagonalUp		: Border running from the lower left-hand corner to the upper right of each cell in the range.
				##	xlEdgeBottom		: Border at the bottom of the range.
				##	xlEdgeLeft			: Border at the left-hand edge of the range.
				##	xlEdgeRight			: Border at the right-hand edge of the range.
				##	xlEdgeTop			: Border at the top of the range.
				$XlBordersIndex = 'Microsoft.Office.Interop.Excel.XlBordersIndex' -as [Type]

				##	http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlborderweight(v=office.14).aspx
				##	Member name	: Description
				##	------------:-------------
				##	xlHairline	: Hairline (thinnest border).
				##	xlMedium	: Medium.
				##	xlThick		: Thick (widest border).
				##	xlThin		: Thin.
				$XlBorderWeight = 'Microsoft.Office.Interop.Excel.XlBorderWeight' -as [Type]
			#endregion Excel 2010 Enumerations

			$Workbook = $null
			$Worksheet = $null
			[int]$WorksheetNumber = 0
			$Range = $null
			$WorksheetData = $null
			$ColorThemePathPattern = $null
			$ColorThemePath = $null
			$MissingType = [System.Type]::Missing

			if ($IncludeExcessiveUpTime)
			{
				## Set this to whatever you consider to be the number of days that a server is online with out a reboot to be excessive.
				## I have chosen to set it to two days less than the number of days that will trigger the "Up Time" bug (497 days) in Windows Server 2008 and Windows Server 2008 R2.
				## It can be found in the KB Article 'KB2553549' or at the following internet link:
				##		http://support.microsoft.com/en-us/kb/2553549/en-us
				## Normally I would set it to 365.
				[int]$ExcessiveUpTimeDays = 495
				[int]$ExcessiveUpTimeColorIndex = 29
			}

			[int]$Row = 0
			[int]$RowCount = 0
			[int]$Col = 0
			##[int]$ColumnCount = 0
			## Instead of setting this for each worksheet, I have changed it to set it at the number of columns there are in the $Inventory array.
			## This way we never have to worry about forgetting to change it on each worksheet when a new column is added.
			## All it is doing is getting the first item in the array and sending that through the Get-Member Cmdlet to return only the Properties columns.
			## It is then doing the count on the number of Properties columns that are returned.
			[int]$ColumnCount = (($Inventory[0] | Get-Member -MemberType Properties) | Measure-Object).Count

			## Used to hold all of the formatting to be applied at the end of the function.
			$WorksheetFormat = @{}
			[int]$WorksheetCount = 5

			if ($IncludeMSSQL)
			{
				$WorksheetCount++
			}

			if ($IncludeHotFixDetails)
			{
				$WorksheetCount++
			}

			if ($IncludeCustomAppInfo)
			{
				$WorksheetCount++
			}

			if ($IncludeInfraDetails)
			{
				$WorksheetCount++
			}

			$BoldHeaderRows = @()
			$BlankRowCells = @()

			## The 2 variables for Worksheet 1 are intentionally set high here to prevent a couple of errors. ##
			## They are set to the corrected/actual used values later in the code.                            ##
			[int]$Worksheet1RowCount = 65536
			[int]$Worksheet1ColumnCount = 26
			[int]$Worksheet2RowCount = 0
			[int]$Worksheet2ColumnCount = 0
			[int]$Worksheet3RowCount = 0
			[int]$Worksheet3ColumnCount = 0
			[int]$Worksheet4RowCount = 0
			[int]$Worksheet4ColumnCount = 0
			[int]$Worksheet5RowCount = 0
			[int]$Worksheet5ColumnCount = 0

			if ($IncludeMSSQL)
			{
				[int]$Worksheet6RowCount = 0
				[int]$Worksheet6ColumnCount = 0
			}

			if ($IncludeHotFixDetails)
			{
				[int]$Worksheet7RowCount = 0
				[int]$Worksheet7ColumnCount = 0
			}

			if ($IncludeCustomAppInfo)
			{
				[int]$Worksheet8RowCount = 0
				[int]$Worksheet8ColumnCount = 0
			}

			if ($IncludeInfraDetails)
			{
				[int]$Worksheet9RowCount = 0
				[int]$Worksheet9ColumnCount = 0
			}
		}
		process
		{
			$ProgressId = Get-Random
			$ProgressActivity = 'Export-ReportToExcel'
			$ProgressStatus = 'Beginning output to Excel'

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
			Write-Progress -Activity $ProgressActivity -PercentComplete 0 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

			#region Write to Excel
				$Excel = New-Object -ComObject Excel.Application

				## Hide the Excel instance (is this necessary?)
				$Excel.Visible = $false

				## Turn off screen updating
				$Excel.ScreenUpdating = $false

				## Turn off automatic calculations
				##$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

				## Add a workbook
				$Workbook = $Excel.Workbooks.Add()
				##$Workbook.Title = 'Server Inventory - For Licensing'
				$Workbook.Title = $RptTitleHeading

				## Try to load the theme specified by $ColorTheme
				## The default theme is called 'Office'. If that's what was specified then skip over this stuff - it's already loaded
				if ($ColorTheme -ine 'Office')
				{
					$ColorThemePathPattern = [String]::Join([System.IO.Path]::DirectorySeparatorChar, @([System.IO.Path]::GetDirectoryName($Excel.Path), 'Document Themes *', 'Theme Colors', [System.IO.Path]::ChangeExtension($ColorTheme, 'xml')))
					$ColorThemePath = $null

					Get-ChildItem -Path $ColorThemePathPattern | ForEach-Object {
						$ColorThemePath = $_.FullName
					}

					if ($ColorThemePath)
					{
						$Workbook.Theme.ThemeColorScheme.Load($ColorThemePath)
					}
				}

				## Add enough worksheets to get us to 5
				$Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), $WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type) | Out-Null
				$WorksheetNumber = 1

				try
				{
					#region Worksheet 1: Overview - Summary of Inventory
						$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Overview"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
						Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

						$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
						$Worksheet.Name = 'Overview'
						$Worksheet.Tab.ThemeColor = $OverviewTabColor

						$RowCount = $Worksheet1RowCount

						$WorksheetData = New-Object -TypeName 'String[,]' -ArgumentList $RowCount, $ColumnCount

						#region Row 1 - Report Title
							$Col = 0
							$WorksheetData[0, $Col++] = $RptTitleHeading
							$WorksheetData[0, $Col++] = ''
							$WorksheetData[0, $Col++] = ''

							if ($Col -lt $Worksheet1ColumnCount)
							{
								$Worksheet1ColumnCount = $Col
							}

							$Row++
						#endregion Row 1 - Report Title

						<#
						#region Row - Blank or Empty 1
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
						#endregion Row - Blank or Empty 1
						#>

						#region Row - Headings - AD Domains - In Scope
							$Col = 0
							$WorksheetData[$Row, $Col++] = 'Active Directory Domains - In Scope'
							$WorksheetData[$Row, $Col++] = 'NETBIOS Name'
							$WorksheetData[$Row, $Col++] = 'DNS Domain (FQDN)'
							$Row++
							$BoldHeaderRows += $Row
						#endregion Row - Headings - AD Domains - In Scope

						#region Rows - Values - AD Domains - In Scope
							foreach ($SplaAdDnsDomain in $SplaAdDnsDomains)
							{
								if ($DNStoNBnames[$SplaAdDnsDomain])
								{
									$Col = 1
									$WorksheetData[$Row, $Col++] = $DNStoNBnames[$SplaAdDnsDomain]
									$WorksheetData[$Row, $Col++] = $SplaAdDnsDomain
									$Row++
								}
							}
						#endregion Row - Values - AD Domains - In Scope

						#region Row - Blank or Empty 2
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
							$BlankRowCells += $Row
						#endregion Row - Blank or Empty 2

						#region Row - Headings - New Systems created after MM/dd/yyyy HH:mm:ss
							$Col = 0
							$WorksheetData[$Row, $Col++] = 'Newly Built Systems for ALL "In Scope" Active Directory domains'
							$Row++
							$BoldHeaderRows += $Row
						#endregion Row - Headings - New Systems created after MM/dd/yyyy HH:mm:ss

						#region Row - Values - New Systems created after MM/dd/yyyy HH:mm:ss
							$Col = 1
							$WorksheetData[$Row, $Col++] = "Built after: $($OldestCreateDate.ToString())"
							$WorksheetData[$Row, $Col++] = $SvrCountNew
							$Row++
						#endregion Row - Values - New Systems created after MM/dd/yyyy HH:mm:ss

						#region Row - Blank or Empty 3
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
							$BlankRowCells += $Row
						#endregion Row - Blank or Empty 3

						#region Row - Headings - Server Counts - ALL In Scope
							$Col = 0
							$WorksheetData[$Row, $Col++] = 'Servers found in ALL "In Scope" Active Directory domains'
							$Row++
							$BoldHeaderRows += $Row
						#endregion Row - Headings - Server Counts - ALL In Scope

						#region Row - Values - Server Counts - In Scope
							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Physical Svr Count'
							$WorksheetData[$Row, $Col++] = $SvrCountPhy
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'VM Svr Count'
							$WorksheetData[$Row, $Col++] = $SvrCountVm
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'EE Svr Count'
							$WorksheetData[$Row, $Col++] = $SvrCountEE
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'NR Svr Count'
							$WorksheetData[$Row, $Col++] = $SvrCountNR
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Total Servers'
							$WorksheetData[$Row, $Col++] = $SvrCountTotal
							$Row++
						#endregion Row - Values - Server Counts - In Scope

						#region Row - Blank or Empty 4
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
							$BlankRowCells += $Row
						#endregion Row - Blank or Empty 4

						#region Row - Headings - Server Counts - By Data Center Location
							$Col = 0
							$WorksheetData[$Row, $Col++] = 'Server Counts by Data Center Location'
							$WorksheetData[$Row, $Col++] = 'Data Center'
							$WorksheetData[$Row, $Col++] = 'Quantity'
							$Row++
							$BoldHeaderRows += $Row
						#endregion Row - Headings - Server Counts - By Data Center Location

						#region Row - Values - Server Counts - By Data Center Location
							foreach ($PsnDataCenter in $DataCenters)
							{
								[int]$PsnDCServerCount = (($Inventory | Where-Object { $_.SrvDataCenter -eq $PsnDataCenter.LocationCodes }) | Measure-Object).Count

								if ($PsnDCServerCount -ge 1)
								{
									$Col = 1
									$WorksheetData[$Row, $Col++] = $PsnDataCenter.LocationCodes
									$WorksheetData[$Row, $Col++] = $PsnDCServerCount	#(($Inventory | Where-Object { $_.SrvDataCenter -eq $PsnDataCenter.LocationCodes }) | Measure-Object).Count
									$Row++
								}
							}
						#endregion Row - Values - Server Counts - By Data Center Location

						#region Row - Blank or Empty 5
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
							$BlankRowCells += $Row
						#endregion Row - Blank or Empty 5

						#region Row - Headings - Server Counts - By OS
							$Col = 0
							$WorksheetData[$Row, $Col++] = 'Server Counts by Operating System'
							$WorksheetData[$Row, $Col++] = 'Operating System'
							$WorksheetData[$Row, $Col++] = 'Quantity'
							$Row++
							$BoldHeaderRows += $Row
						#endregion Row - Headings - Server Counts - By OS

						#region Row - Values - Server Counts - By OS
							## Need to change/fix this so that it pulls a unique list of OS names from the array and loops over that to get these counts.
							$Col = 1
							$WorksheetData[$Row, $Col++] = '2000 Server (ALL)'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2000*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2003 (ALL)'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2003*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2008 Std'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2008 St*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2008 Ent'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2008 En*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Web Server 2008'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*Web*2008*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2008 R2 Std'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2008 R2 St*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2008 R2 Ent'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2008 R2 En*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Web Server 2008 R2'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*Web*2008 R2*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2012 Std'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2012 St*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Server 2012 R2 Std'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*2012 R2 St*" }) | Measure-Object).Count
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Hyper-V Server'
							$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvOsName -like "*Hyper-V*" }) | Measure-Object).Count
							$Row++
						#endregion Row - Values - Server Counts - By OS

						#region Row - Blank or Empty 6
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
							$BlankRowCells += $Row
						#endregion Row - Blank or Empty 6

						if ($IncludeCustomAppInfo)
						{
							#region Row - Headings - Application Counts
								$Col = 0
								$WorksheetData[$Row, $Col++] = "Total Counts by Application"
								$WorksheetData[$Row, $Col++] = 'Application'
								$WorksheetData[$Row, $Col++] = 'Quantity'
								$Row++
								$BoldHeaderRows += $Row
							#endregion Row - Headings - Application Counts

							#region Row - Values - Application Counts
								$Col = 1
								$WorksheetData[$Row, $Col++] = 'BrowserHawk Editor'
								$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvBrowserHawkInstall -eq 'Yes' }) | Measure-Object).Count
								$Row++

								$Col = 1
								$WorksheetData[$Row, $Col++] = 'SoftArtisans FileUp'
								$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { $_.SrvSoftArtisansFileUpInstall -eq 'Yes' }) | Measure-Object).Count
								$Row++
							#endregion Row - Values - Application Counts

							#region Row - Blank or Empty 7
								$Col = 0
								## Intentionally Left blank. Spacer Row.
								#$WorksheetData[$Row, $Col++] = ''
								$Row++
								$BlankRowCells += $Row
							#endregion Row - Blank or Empty 7
						}

						#region Rows - ForEach In Scope Domain
							foreach ($SplaAdDnsDomain in $SplaAdDnsDomains)
							{
								#region Row - Headings - Server Counts - ForEach In Scope Domain 1
									$Col = 0
									$WorksheetData[$Row, $Col++] = "Server Counts for $([char]34)$($DNStoNBnames[$SplaAdDnsDomain])$([char]34)"
									$Row++
									$BoldHeaderRows += $Row
								#endregion Row - Headings - Server Counts - ForEach In Scope Domain 1

								#region Row - Values - Server Counts - ForEach In Scope Domain 1
									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Physical Svr Count'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvModel -notlike "*Virtual*") -and ($_.SrvModel -ne 'Error Encountered') -and ($_.SrvModel -ne 'No Response') -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'VM Svr Count'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvModel -like "*Virtual*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'EE Svr Count'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvModel -eq 'Error Encountered') -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'NR Svr Count'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvModel -eq 'No Response') -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Total Domain Servers'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++
								#endregion Row - Values - Server Counts - ForEach In Scope Domain 1

								#region Row - Blank or Empty 8
									$Col = 0
									## Intentionally Left blank. Spacer Row.
									#$WorksheetData[$Row, $Col++] = ''
									$Row++
									$BlankRowCells += $Row
								#endregion Row - Blank or Empty 8

								#region Row - Headings - Server Counts - ForEach In Scope Domain 2
									$Col = 0
									$WorksheetData[$Row, $Col++] = "Server Counts by Operating System for $([char]34)$($DNStoNBnames[$SplaAdDnsDomain])$([char]34)"
									$WorksheetData[$Row, $Col++] = 'Operating System'
									$WorksheetData[$Row, $Col++] = 'Quantity'
									$Row++
									$BoldHeaderRows += $Row
								#endregion Row - Headings - Server Counts - ForEach In Scope Domain 2

								#region Row - Values - Server Counts - ForEach In Scope Domain 2
									$Col = 1
									$WorksheetData[$Row, $Col++] = '2000 Server (ALL)'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2000*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2003 (ALL)'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2003*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2008 Std'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2008 St*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2008 Ent'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2008 En*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Web Server 2008'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*Web*2008*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2008 R2 Std'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2008 R2 St*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2008 R2 Ent'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2008 R2 En*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Web Server 2008 R2'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*Web*2008 R2*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2012 Std'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2012 St*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Server 2012 R2 Std'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*2012 R2 St*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++

									$Col = 1
									$WorksheetData[$Row, $Col++] = 'Hyper-V Server'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvOsName -like "*Hyper-V*") -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++
								#endregion Row - Values - Server Counts - ForEach In Scope Domain 2

								#region Row - Blank or Empty 9
									$Col = 0
									## Intentionally Left blank. Spacer Row.
									#$WorksheetData[$Row, $Col++] = ''
									$Row++
									$BlankRowCells += $Row
								#endregion Row - Blank or Empty 9

								#region Row - Headings - Application Counts - ForEach In Scope Domain 1
									$Col = 0
									$WorksheetData[$Row, $Col++] = "Counts by Application for $([char]34)$($DNStoNBnames[$SplaAdDnsDomain])$([char]34)"
									$WorksheetData[$Row, $Col++] = 'Application'
									$WorksheetData[$Row, $Col++] = 'Quantity'
									$Row++
									$BoldHeaderRows += $Row
								#endregion Row - Headings - Application Counts - ForEach In Scope Domain 1

								#region Row - Values - Application Counts - ForEach In Scope Domain 1
									$Col = 1
									$WorksheetData[$Row, $Col++] = 'BrowserHawk Editor'
									$WorksheetData[$Row, $Col++] = (($Inventory | Where-Object { ($_.SrvBrowserHawkInstall -eq $true) -and ($_.SrvNBDomain -eq "$($DNStoNBnames[$SplaAdDnsDomain])") }) | Measure-Object).Count
									$Row++
								#endregion Row - Values - Application Counts - ForEach In Scope Domain 1

								#region Row - Blank or Empty 10
									$Col = 0
									## Intentionally Left blank. Spacer Row.
									#$WorksheetData[$Row, $Col++] = ''
									$Row++
									$BlankRowCells += $Row
								#endregion Row - Blank or Empty 10
							}
						#endregion Rows - ForEach In Scope Domain

						#region Row - Headings - Legend
							$Col = 0
							$WorksheetData[$Row, $Col++] = 'Legend'
							$WorksheetData[$Row, $Col++] = 'Abbreviation'
							$WorksheetData[$Row, $Col++] = 'Definition'
							$Row++
							$BoldHeaderRows += $Row
						#endregion Row - Headings - Legend

						#region Row - Values - Legend
							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Text/Font Color'
							$WorksheetData[$Row, $Col++] = 'Text Color for rows with Servers reporting Excessive Up Time'
							$Row++
							if ($IncludeExcessiveUpTime)
							{
									## Used to hold the rows where servers are showing as up for more than '$ExcessiveUpTimeDays' days.
									## We will them loop over it later to change the font color on the entire row.
									$ExcessiveUpTimeRows += $Row
							}

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'EE'
							$WorksheetData[$Row, $Col++] = 'Error Encountered'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'NR'
							$WorksheetData[$Row, $Col++] = 'No Response'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'VM'
							$WorksheetData[$Row, $Col++] = 'Virtual Machine'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Std'
							$WorksheetData[$Row, $Col++] = 'Standard'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Ent'
							$WorksheetData[$Row, $Col++] = 'Enterprise'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'MS SQL'
							$WorksheetData[$Row, $Col++] = 'Microsoft SQL Server'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Svr'
							$WorksheetData[$Row, $Col++] = 'Server'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Svc'
							$WorksheetData[$Row, $Col++] = 'Service'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'OS'
							$WorksheetData[$Row, $Col++] = 'Operating System'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'SP Lvl'
							$WorksheetData[$Row, $Col++] = 'Service Pack Level'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'HT'
							$WorksheetData[$Row, $Col++] = 'Intel Hyperthreading -OR- AMD HyperTransport'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'Arch'
							$WorksheetData[$Row, $Col++] = 'Architecture'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'FQDN'
							$WorksheetData[$Row, $Col++] = 'Fully Qualified Domain Name'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'NB'
							$WorksheetData[$Row, $Col++] = 'NETBIOS'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'DNS'
							$WorksheetData[$Row, $Col++] = 'Domain Name System'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'GB'
							$WorksheetData[$Row, $Col++] = 'Gigabyte'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'MB'
							$WorksheetData[$Row, $Col++] = 'Megabyte'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'KB'
							$WorksheetData[$Row, $Col++] = 'Kilobyte'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'SPLA'
							$WorksheetData[$Row, $Col++] = 'Service Provider License Agreement'
							$Row++

							$Col = 1
							$WorksheetData[$Row, $Col++] = 'EA'
							$WorksheetData[$Row, $Col++] = 'Enterprise License Agreement'
							$Row++
						#endregion Row - Values - Legend

						#region Row - Blank or Empty 11
							$Col = 0
							## Intentionally Left blank. Spacer Row.
							#$WorksheetData[$Row, $Col++] = ''
							$Row++
							$BlankRowCells += $Row
						#endregion Row - Blank or Empty 11

						## If the total used row count is not equal to the value initially set for $Worksheet1RowCount,
						## Then set $Worksheet1RowCount to the Actual used row count. Otherwise leave it alone.
						if ($Row -ne $Worksheet1RowCount)
						{
							$Worksheet1RowCount = $Row
						}

						$Row++
						$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($Row, $Worksheet1ColumnCount))
						$Range.Value2 = $WorksheetData
						#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $MissingType, $MissingType, $MissingType, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null

						$WorksheetFormat.Add($WorksheetNumber, @{
								BoldFirstRow = $true
								BoldFirstColumn = $false
								AutoFilter = $false
								ApplyCellBorders = $true
								FreezeAtCell = 'A2'
								ApplyTableFormatting = $false
								ColumnFormat = @(
									#@{ColumnNumber = 3; NumberFormat = $XlNumFmtText}
								)
								RowFormat = @()
								HeadingsFormat = @(
									@{
										RowNumber = 1
										StartColumnNumber = 1
										EndColumnNumber = $Worksheet1ColumnCount
										FontBold = $true
										FontSize = 20
										## http://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx
										## http://blog.softartisans.com/2013/05/13/kb-excels-color-palette-explained/
										##	Common Colors are:
										##	1 = Black;		2 = White;
										##	3 = Red;		4 = Green;		5 = Blue;		6 = Yellow;			7 = Magenta;		8 = Cyan
										##	9 = DarkRed;	10 = DarkGreen;	11 = DarkBlue;	12 = DarkYellow;	13 = DarkMagenta;	14 = DarkCyan;
										##	15 = Gray;	16 = DarkGray
										BackGroundColor = 1
										FontColor = 2
										VerticalAlignment = $XlVAlign::xlVAlignCenter
										HorizontalAlignment = $XlHAlign::xlHAlignCenter
									}
									foreach ($BoldHeaderRow in $BoldHeaderRows)
									{
										@{
											RowNumber = $BoldHeaderRow
											StartColumnNumber = 1
											EndColumnNumber = $Worksheet1ColumnCount
											FontBold = $true
											FontSize = 11
											BackGroundColor = 10
											FontColor = 2
											VerticalAlignment = $XlVAlign::xlVAlignTop
											HorizontalAlignment = $XlHAlign::xlHAlignLeft
										}
									}
								)
								BlankRowCellFormat = @(
									foreach ($BlankRowCell in $BlankRowCells)
									{
										@{
											RowNumber = $BlankRowCell
											StartColumnNumber = 1
											EndColumnNumber = $Worksheet1ColumnCount
											MergeCells = $true
											BackGroundColor = 2
											FontColor = 2
											FontSize = 2
											RowHeight = 4
											VerticalAlignment = $XlVAlign::xlVAlignCenter
											HorizontalAlignment = $XlHAlign::xlHAlignCenter
										}
									}
								)
								ExcessiveUpTimeRowFormat = @(
									if ($IncludeExcessiveUpTime)
									{
										foreach ($ExcessiveUpTimeRow in $ExcessiveUpTimeRows)
										{
											@{
												RowNumber = $ExcessiveUpTimeRow
												StartColumnNumber = 1
												EndColumnNumber = $Worksheet1ColumnCount
												FontBold = $true
												FontSize = 11
												FontColor = $ExcessiveUpTimeColorIndex
												#BackGroundColor = 10
												VerticalAlignment = $XlVAlign::xlVAlignTop
												HorizontalAlignment = $XlHAlign::xlHAlignLeft
											}
										}
									}
								)
							}
						)

						$WorksheetNumber++
					#endregion Worksheet 1: Overview - Summary of Inventory

					#region Worksheet 2: Active Servers
						$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Active Servers"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
						Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

						$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
						$Worksheet.Name = 'Active Servers'
						$Worksheet.Tab.ThemeColor = $ActiveServersTabColor

						## Used to hold the columns where Number or Date formatting needs to be applied.
						## We will them loop over them later to apply the proper formatting on the entire column.
						$FmtNumberGeneralCols = @()
						$FmtDateCols = @()
						$FmtNumberS2Cols = @()

						$RowCount = (($Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') }) | Measure-Object).Count + 1

						$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

						if ($IncludeExcessiveUpTime)
						{
							$ExcessiveUpTimeRows = @()
						}

						## If the total row count is not equal to the value initially set for $Worksheet2RowCount,
						## Then set $Worksheet2RowCount to the Actual used row count. Otherwise leave it alone.
						if ($RowCount -ne $Worksheet2RowCount)
						{
							$Worksheet2RowCount = $RowCount
						}

						#region Worksheet 2: Column Headers
							$Col = 0
							$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
							$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
							$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
							$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
							$WorksheetData[0, $Col++] = 'OS Name'				# Column Number 5
							$WorksheetData[0, $Col++] = 'SP Lvl'				# Column Number 6
							$WorksheetData[0, $Col++] = 'Installed OS Arch'		# Column Number 7
							$WorksheetData[0, $Col++] = 'Physical Memory (GB)'	# Column Number 8
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'Processor Sockets'		# Column Number 9
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'Processor Cores'		# Column Number 10
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'Logical Processors'	# Column Number 11
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'HT'					# Column Number 12
							$WorksheetData[0, $Col++] = 'Processor Manufacturer'# Column Number 13
							$WorksheetData[0, $Col++] = 'Created On'			# Column Number 14
							$FmtDateCols += $Col
							$WorksheetData[0, $Col++] = 'Changed On'			# Column Number 15
							$FmtDateCols += $Col
							$WorksheetData[0, $Col++] = 'Manufacturer'			# Column Number 16
							$WorksheetData[0, $Col++] = 'Model'					# Column Number 17
							$WorksheetData[0, $Col++] = 'Serial Number'			# Column Number 18
							$WorksheetData[0, $Col++] = 'Data Center Location'	# Column Number 19
							$WorksheetData[0, $Col++] = 'MS Cluster Name'		# Column Number 20
							#$WorksheetData[0, $Col++] = 'BMC Agent Installed'	# Column Number 21
							#$WorksheetData[0, $Col++] = 'BMC Agent State'		# Column Number 22
							$WorksheetData[0, $Col++] = 'Last Boot Up Time'		# Column Number 23
							$FmtDateCols += $Col
							$WorksheetData[0, $Col++] = 'Up Time (Days)'		# Column Number 24
							$FmtNumberGeneralCols += $Col
							#$WorksheetData[0, $Col++] = 'Pending Reboot'		# Column Number 25

							#if ($IncludeHotFixDetails)
							#{
							#	$WorksheetData[0, $Col++] = 'Last Patch Install Date'	# Column Number 26
							#	$FmtDateCols += $Col
							#	$WorksheetData[0, $Col++] = 'Up Time HotFix 1 Applied'	# Column Number 27
							#	$WorksheetData[0, $Col++] = 'Up Time HotFix 2 Applied'	# Column Number 28
							#	$WorksheetData[0, $Col++] = 'Up Time Total HotFix Applied'	# Column Number 29
							#	$WorksheetData[0, $Col++] = 'MS15-034 HotFix Applied'	# Column Number 30
							#}

							#if ($IncludeCustomAppInfo)
							#{
							#	$WorksheetData[0, $Col++] = 'BrowserHawk Installed'		# Column Number 31
							#	$WorksheetData[0, $Col++] = 'SoftArtisans FileUp Installed'
							#}

							#if ($IncludeInfraDetails)
							#{
							#	$WorksheetData[0, $Col++] = 'CPU_Type'					# Column Number 32
							#	$WorksheetData[0, $Col++] = 'Disk_Drives'				# Column Number 33
							#	$WorksheetData[0, $Col++] = 'Disks_Space_in-used'		# Column Number 34
							#	$WorksheetData[0, $Col++] = 'Drives_free_Space'			# Column Number 35
							#	$WorksheetData[0, $Col++] = 'Total_Disk_Allocated_GB'	# Column Number 36
							#	$FmtNumberS2Cols += $Col
							#	$WorksheetData[0, $Col++] = 'Total_Disks_in-used_GB'	# Column Number 37
							#	$FmtNumberS2Cols += $Col
							#	$WorksheetData[0, $Col++] = 'Total_Disks_free_Space_GB'	# Column Number 38
							#	$FmtNumberS2Cols += $Col
							#}

							$Worksheet2ColumnCount = $Col
						#endregion Worksheet 2: Column Headers

						#region Worksheet 2: Column Values
							$Row = 1
							$Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') } | ForEach-Object {
								$Col = 0

								$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

								if ($_.SrvIPv4Addr -eq '&nbsp;')
								{
									$WorksheetData[$Row, $Col++] = ''				#'IP Address'
								}
								else
								{
									$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
								}

								$WorksheetData[$Row, $Col++] = $_.SrvOsName			#'OS Name'
								$WorksheetData[$Row, $Col++] = $_.SrvOsSpLvl		#'SP Lvl'
								$WorksheetData[$Row, $Col++] = $_.SrvInstOsArch		#'Installed OS Arch'
								$WorksheetData[$Row, $Col++] = $_.SrvPhysMemoryGB	#'Physical Memory (GB)'
								$WorksheetData[$Row, $Col++] = $_.SrvSocketCount	#'Processor Sockets'
								$WorksheetData[$Row, $Col++] = $_.SrvProcCoreCount	#'Processor Cores'
								$WorksheetData[$Row, $Col++] = $_.SrvLogProcsCount	#'Logical Processors'
								$WorksheetData[$Row, $Col++] = $_.SrvHyperT			#'HT' - Intel Hyperthreading -OR- AMD HyperTransport
								$WorksheetData[$Row, $Col++] = $_.SrvProcMnftr		#'Processor Manufacturer'
								$WorksheetData[$Row, $Col++] = $_.SrvCreatedOn		#'Created On'
								$WorksheetData[$Row, $Col++] = $_.SrvChangedOn		#'Changed On'
								$WorksheetData[$Row, $Col++] = $_.SrvMnfctr			#'Manufacturer'
								$WorksheetData[$Row, $Col++] = $_.SrvModel			#'Model'
								$WorksheetData[$Row, $Col++] = $_.SrvSerNum			#'Serial Number'
								$WorksheetData[$Row, $Col++] = $_.SrvDataCenter		#'Data Center Location'
								$WorksheetData[$Row, $Col++] = $_.MSClusterName		#'MS Cluster Name'
								#$WorksheetData[$Row, $Col++] = $_.SrvBMCAgentInstalled	#'BMC Agent Installed'
								#$WorksheetData[$Row, $Col++] = $_.SrvBMCAgentState	#'BMC Agent State'
								$WorksheetData[$Row, $Col++] = $_.SrvLastBootUpTime	#'Last Boot Up Time'
								$WorksheetData[$Row, $Col++] = $_.SrvUpTime			#'Up Time (Days)'
								#$WorksheetData[$Row, $Col++] = $_.SrvPendingReboot	#'SrvPendingReboot'

								#if ($IncludeHotFixDetails)
								#{
								#	$WorksheetData[$Row, $Col++] = $_.SrvLastPatchInstallDate	#'Last Patch Install Date'
								#	$WorksheetData[$Row, $Col++] = $_.SrvUpTimeKB2553549HotFix1Applied	#'Up Time HotFix 1 Applied'
								#	$WorksheetData[$Row, $Col++] = $_.SrvUpTimeKB2688338HotFix2Applied	#'Up Time HotFix 2 Applied'
								#	$WorksheetData[$Row, $Col++] = $_.SrvUpTimeTotalHotFixApplied	#'Up Time Total HotFix Applied'
								#	$WorksheetData[$Row, $Col++] = $_.SrvKB3042553HotFixApplied	#'MS15-034 HotFix Applied'
								#}

								#if ($IncludeCustomAppInfo)
								#{
								#	$WorksheetData[$Row, $Col++] = $_.SrvBrowserHawkInstall		#'BrowserHawk Installed'
								#	$WorksheetData[$Row, $Col++] = $_.SrvSoftArtisansFileUpInstall	#'SoftArtisans FileUp Installed'
								#}

								#if ($IncludeInfraDetails)
								#{
								#	$WorksheetData[$Row, $Col++] = $_.SrvProcName			#'CPU_Type'	#'Processor Family'
								#	$WorksheetData[$Row, $Col++] = $_.SrvDiskTotalInfo		#'Disk_Drives'
								#	$WorksheetData[$Row, $Col++] = $_.SrvDiskUsedInfo		#'Disks_Space_in-used'
								#	$WorksheetData[$Row, $Col++] = $_.SrvDiskFreeInfo		#'Drives_free_Space'
								#	$WorksheetData[$Row, $Col++] = $_.SrvDiskTotalAllocated	#'Total_Disk_Allocated_GB'
								#	$WorksheetData[$Row, $Col++] = $_.SrvDiskUsedAllocated	#'Total_Disks_in-used_GB'
								#	$WorksheetData[$Row, $Col++] = $_.SrvDiskFreeAllocated	#'Total_Disks_free_Space_GB'
								#}

								$Row++

								if ($IncludeExcessiveUpTime)
								{
									if ([int]$_.SrvUpTime -gt $ExcessiveUpTimeDays)
									{
										## Used to hold the rows where servers are showing as up for more than '$ExcessiveUpTimeDays' days.
										## We will them loop over it later to change the font color on the entire row.
										$ExcessiveUpTimeRows += $Row
									}
								}
							}
						#endregion Worksheet 2: Column Values

						$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
						$Range.Value2 = $WorksheetData
						<#
							object Sort(
								object Key1,
								XlSortOrder Order1 = XlSortOrder.xlAscending,
								object Key2,
								object Type,
								XlSortOrder Order2 = XlSortOrder.xlAscending,
								object Key3,
								XlSortOrder Order3 = XlSortOrder.xlAscending,
								XlYesNoGuess Header = XlYesNoGuess.xlNo,
								object OrderCustom,
								object MatchCase,
								XlSortOrientation Orientation = XlSortOrientation.xlSortRows,
								XlSortMethod SortMethod = XlSortMethod.xlPinYin,
								XlSortDataOption DataOption1 = XlSortDataOption.xlSortNormal,
								XlSortDataOption DataOption2 = XlSortDataOption.xlSortNormal,
								XlSortDataOption DataOption3 = XlSortDataOption.xlSortNormal
							)
						#>
						#$Range.Sort(Key1,						Order1,						Key2,						Type,		 Order2,					Key3,			Order3,		Header) | Out-Null
						#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
						$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

						$WorksheetFormat.Add($WorksheetNumber, @{
								BoldFirstRow = $true
								BoldFirstColumn = $false
								AutoFilter = $true
								FreezeAtCell = 'A2'
								ApplyCellBorders = $false
								ApplyTableFormatting = $true
								ColumnFormat = @(
									#@{ColumnNumber = 7; NumberFormat = $XlNumFmtNumberGeneral},	#'Physical Memory (GB)'
									#@{ColumnNumber = 8; NumberFormat = $XlNumFmtNumberGeneral},	#'Processor Sockets'
									#@{ColumnNumber = 9; NumberFormat = $XlNumFmtNumberGeneral},	#'Processor Cores'
									#@{ColumnNumber = 10; NumberFormat = $XlNumFmtNumberGeneral},	#'Logical Processors'
									#@{ColumnNumber = 13; NumberFormat = $XlNumFmtDate},			#'Created On'
									#@{ColumnNumber = 14; NumberFormat = $XlNumFmtDate},			#'Changed On'
									#@{ColumnNumber = 22; NumberFormat = $XlNumFmtDate},			#'Last Boot Up Time'
									#@{ColumnNumber = 23; NumberFormat = $XlNumFmtNumberGeneral}	#'Up Time (Days)'

									#if ($IncludeInfraDetails)
									#{
									#	, @{ColumnNumber = 35; NumberFormat = $XlNumFmtNumberS2},	#'Total_Disk_Allocated_GB'
									#	@{ColumnNumber = 36; NumberFormat = $XlNumFmtNumberS2},		#'Total_Disks_in-used_GB'
									#	@{ColumnNumber = 37; NumberFormat = $XlNumFmtNumberS2}		#'Total_Disks_free_Space_GB'
									#}

									if ($FmtNumberGeneralCols.Length -ge 1)
									{
										foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
										{
											@{
												ColumnNumber = $FmtNumberGeneralCol
												NumberFormat = $XlNumFmtNumberGeneral
											}
										}
									}

									if ($FmtDateCols.Length -ge 1)
									{
										foreach ($FmtDateCol in $FmtDateCols)
										{
											@{
												ColumnNumber = $FmtDateCol
												NumberFormat = $XlNumFmtDate
											}
										}
									}

									if ($FmtNumberS2Cols.Length -ge 1)
									{
										foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
										{
											@{
												ColumnNumber = $FmtNumberS2Col
												NumberFormat = $XlNumFmtNumberS2
											}
										}
									}
								)
								RowFormat = @()
								HeadingsFormat = @()
								BlankRowCellFormat = @()
								ExcessiveUpTimeRowFormat = @(
									if ($IncludeExcessiveUpTime)
									{
										foreach ($ExcessiveUpTimeRow in $ExcessiveUpTimeRows)
										{
											@{
												RowNumber = $ExcessiveUpTimeRow
												StartColumnNumber = 1
												EndColumnNumber = $Worksheet2ColumnCount
												FontBold = $true
												FontSize = 11
												FontColor = $ExcessiveUpTimeColorIndex
												#BackGroundColor = 10
												VerticalAlignment = $XlVAlign::xlVAlignTop
												HorizontalAlignment = $XlHAlign::xlHAlignLeft
											}
										}
									}
								)
							}
						)

						$WorksheetNumber++
					#endregion Worksheet 2: Active Servers

					#region Worksheet 3: Error Encountered Servers
						$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Error Encountered Servers"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
						Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

						$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
						$Worksheet.Name = 'Error Encountered Servers'
						$Worksheet.Tab.ThemeColor = $ErrorEncounteredServersTabColor

						## Used to hold the columns where Number or Date formatting needs to be applied.
						## We will them loop over them later to apply the proper formatting on the entire column.
						$FmtNumberGeneralCols = @()
						$FmtDateCols = @()
						$FmtNumberS2Cols = @()

						$RowCount = (($Inventory | Where-Object { $_.SrvPhysMemoryGB -eq 'Error Encountered' }) | Measure-Object).Count + 1

						$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

						## If the total row count is not equal to the value initially set for $Worksheet3RowCount,
						## Then set $Worksheet3RowCount to the Actual used row count. Otherwise leave it alone.
						if ($RowCount -ne $Worksheet3RowCount)
						{
							$Worksheet3RowCount = $RowCount
						}

						#region Worksheet 3: Column Headers
							$Col = 0
							$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
							$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
							$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
							$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
							$WorksheetData[0, $Col++] = 'OS Name'				# Column Number 5
							$WorksheetData[0, $Col++] = 'SP Lvl'				# Column Number 6
							$WorksheetData[0, $Col++] = 'Installed OS Arch'		# Column Number 7
							$WorksheetData[0, $Col++] = 'Physical Memory (GB)'	# Column Number 8
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'Processor Sockets'		# Column Number 9
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'Processor Cores'		# Column Number 10
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'Logical Processors'	# Column Number 11
							$FmtNumberGeneralCols += $Col
							$WorksheetData[0, $Col++] = 'HT'					# Column Number 12
							$WorksheetData[0, $Col++] = 'Processor Manufacturer'# Column Number 13
							$WorksheetData[0, $Col++] = 'Created On'			# Column Number 14
							$FmtDateCols += $Col
							$WorksheetData[0, $Col++] = 'Changed On'			# Column Number 15
							$FmtDateCols += $Col
							$WorksheetData[0, $Col++] = 'Manufacturer'			# Column Number 16
							$WorksheetData[0, $Col++] = 'Model'					# Column Number 17
							$WorksheetData[0, $Col++] = 'Serial Number'			# Column Number 18
							$WorksheetData[0, $Col++] = 'Data Center Location'	# Column Number 19
							$WorksheetData[0, $Col++] = 'MS Cluster Name'		# Column Number 20
							$WorksheetData[0, $Col++] = 'BMC Agent Installed'	# Column Number 21
							$WorksheetData[0, $Col++] = 'BMC Agent State'		# Column Number 22
							$Worksheet3ColumnCount = $Col
						#endregion Worksheet 3: Column Headers

						#region Worksheet 3: Column Values
							$Row = 1
							$Inventory | Where-Object { $_.SrvPhysMemoryGB -eq 'Error Encountered' } | ForEach-Object {
								$Col = 0

								$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

								if ($_.SrvIPv4Addr -eq '&nbsp;')
								{
									$WorksheetData[$Row, $Col++] = ''				#'IP Address'
								}
								else
								{
									$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
								}

								$WorksheetData[$Row, $Col++] = $_.SrvOsName			#'OS Name'
								$WorksheetData[$Row, $Col++] = $_.SrvOsSpLvl		#'SP Lvl'
								$WorksheetData[$Row, $Col++] = $_.SrvInstOsArch		#'Installed OS Arch'
								$WorksheetData[$Row, $Col++] = $_.SrvPhysMemoryGB	#'Physical Memory (GB)'
								$WorksheetData[$Row, $Col++] = $_.SrvSocketCount	#'Processor Sockets'
								$WorksheetData[$Row, $Col++] = $_.SrvProcCoreCount	#'Processor Cores'
								$WorksheetData[$Row, $Col++] = $_.SrvLogProcsCount	#'Logical Processors'
								$WorksheetData[$Row, $Col++] = $_.SrvHyperT			#'HT' - Intel Hyperthreading -OR- AMD HyperTransport
								$WorksheetData[$Row, $Col++] = $_.SrvProcMnftr		#'Processor Manufacturer'
								$WorksheetData[$Row, $Col++] = $_.SrvCreatedOn		#'Created On'
								$WorksheetData[$Row, $Col++] = $_.SrvChangedOn		#'Changed On'
								$WorksheetData[$Row, $Col++] = $_.SrvMnfctr			#'Manufacturer'
								$WorksheetData[$Row, $Col++] = $_.SrvModel			#'Model'
								$WorksheetData[$Row, $Col++] = $_.SrvSerNum			#'Serial Number'
								$WorksheetData[$Row, $Col++] = $_.SrvDataCenter		#'Data Center Location'
								$WorksheetData[$Row, $Col++] = $_.MSClusterName		#'MS Cluster Name'
								$WorksheetData[$Row, $Col++] = $_.SrvBMCAgentInstalled	#'BMC Agent Installed'

								$Row++
							}
						#endregion Worksheet 3: Column Values

						$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
						$Range.Value2 = $WorksheetData
						#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
						$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

						$WorksheetFormat.Add($WorksheetNumber, @{
								BoldFirstRow = $true
								BoldFirstColumn = $false
								AutoFilter = $true
								FreezeAtCell = 'A2'
								ApplyCellBorders = $false
								ApplyTableFormatting = $true
								ColumnFormat = @(
									#@{ColumnNumber = 13; NumberFormat = $XlNumFmtDate},			#'Created On'
									#@{ColumnNumber = 14; NumberFormat = $XlNumFmtDate}				#'Changed On'

									if ($FmtNumberGeneralCols.Length -ge 1)
									{
										foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
										{
											@{
												ColumnNumber = $FmtNumberGeneralCol
												NumberFormat = $XlNumFmtNumberGeneral
											}
										}
									}

									if ($FmtDateCols.Length -ge 1)
									{
										foreach ($FmtDateCol in $FmtDateCols)
										{
											@{
												ColumnNumber = $FmtDateCol
												NumberFormat = $XlNumFmtDate
											}
										}
									}

									if ($FmtNumberS2Cols.Length -ge 1)
									{
										foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
										{
											@{
												ColumnNumber = $FmtNumberS2Col
												NumberFormat = $XlNumFmtNumberS2
											}
										}
									}
								)
								RowFormat = @()
								HeadingsFormat = @()
								BlankRowCellFormat = @()
								ExcessiveUpTimeRowFormat = @()
							}
						)

						$WorksheetNumber++
					#endregion Worksheet 3: Error Encountered Servers

					#region Worksheet 4: No Response Servers
						$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): No Response Servers"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
						Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

						$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
						$Worksheet.Name = 'No Response Servers'
						$Worksheet.Tab.ThemeColor = $NoResponseServersTabColor

						## Used to hold the columns where Number or Date formatting needs to be applied.
						## We will them loop over them later to apply the proper formatting on the entire column.
						$FmtNumberGeneralCols = @()
						$FmtDateCols = @()
						$FmtNumberS2Cols = @()

						$RowCount = (($Inventory | Where-Object { $_.SrvPhysMemoryGB -eq 'No Response' }) | Measure-Object).Count + 1

						$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

						## If the total row count is not equal to the value initially set for $Worksheet4RowCount,
						## Then set $Worksheet4RowCount to the Actual used row count. Otherwise leave it alone.
						if ($RowCount -ne $Worksheet4RowCount)
						{
							$Worksheet4RowCount = $RowCount
						}

						#region Worksheet 4: Column Headers
							$Col = 0
							$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
							$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
							$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
							$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
							$WorksheetData[0, $Col++] = 'OS Name'				# Column Number 5
							$WorksheetData[0, $Col++] = 'SP Lvl'				# Column Number 6
							$WorksheetData[0, $Col++] = 'Created On'			# Column Number 7
							$FmtDateCols += $Col
							$WorksheetData[0, $Col++] = 'Changed On'			# Column Number 8
							$FmtDateCols += $Col
							$Worksheet4ColumnCount = $Col
						#endregion Worksheet 4: Column Headers

						#region Worksheet 4: Column Values
							$Row = 1
							$Inventory | Where-Object { $_.SrvPhysMemoryGB -eq 'No Response' } | ForEach-Object {
								$Col = 0

								$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

								if ($_.SrvIPv4Addr -eq '&nbsp;')
								{
									$WorksheetData[$Row, $Col++] = ''				#'IP Address'
								}
								else
								{
									$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
								}

								$WorksheetData[$Row, $Col++] = $_.SrvOsName			#'OS Name'
								$WorksheetData[$Row, $Col++] = $_.SrvOsSpLvl		#'SP Lvl'
								$WorksheetData[$Row, $Col++] = $_.SrvCreatedOn		#'Created On'
								$WorksheetData[$Row, $Col++] = $_.SrvChangedOn		#'Changed On'

								$Row++
							}
						#endregion Worksheet 4: Column Values

						$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
						$Range.Value2 = $WorksheetData
						#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
						$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

						$WorksheetFormat.Add($WorksheetNumber, @{
								BoldFirstRow = $true
								BoldFirstColumn = $false
								AutoFilter = $true
								FreezeAtCell = 'A2'
								ApplyCellBorders = $false
								ApplyTableFormatting = $true
								ColumnFormat = @(
									#@{ColumnNumber = 6; NumberFormat = $XlNumFmtDate},				#'Created On'
									#@{ColumnNumber = 7; NumberFormat = $XlNumFmtDate}				#'Changed On'

									if ($FmtNumberGeneralCols.Length -ge 1)
									{
										foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
										{
											@{
												ColumnNumber = $FmtNumberGeneralCol
												NumberFormat = $XlNumFmtNumberGeneral
											}
										}
									}

									if ($FmtDateCols.Length -ge 1)
									{
										foreach ($FmtDateCol in $FmtDateCols)
										{
											@{
												ColumnNumber = $FmtDateCol
												NumberFormat = $XlNumFmtDate
											}
										}
									}

									if ($FmtNumberS2Cols.Length -ge 1)
									{
										foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
										{
											@{
												ColumnNumber = $FmtNumberS2Col
												NumberFormat = $XlNumFmtNumberS2
											}
										}
									}
								)
								RowFormat = @()
								HeadingsFormat = @()
								BlankRowCellFormat = @()
								ExcessiveUpTimeRowFormat = @()
							}
						)

						$WorksheetNumber++
					#endregion Worksheet 4: No Response Servers

					#region Worksheet 5: Critical Windows Services
						$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Critical Windows Services"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
						Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

						$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
						$Worksheet.Name = 'Critical Windows Services'
						$Worksheet.Tab.ThemeColor = $CriticalWindowsServicesTabColor

						## Used to hold the columns where Number or Date formatting needs to be applied.
						## We will them loop over them later to apply the proper formatting on the entire column.
						$FmtNumberGeneralCols = @()
						$FmtDateCols = @()
						$FmtNumberS2Cols = @()

						$RowCount = (($Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') }) | Measure-Object).Count + 1

						$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

						## If the total row count is not equal to the value initially set for $Worksheet5RowCount,
						## Then set $Worksheet5RowCount to the Actual used row count. Otherwise leave it alone.
						if ($RowCount -ne $Worksheet5RowCount)
						{
							$Worksheet5RowCount = $RowCount
						}

						#region Worksheet 5: Column Headers
							$Col = 0
							$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
							$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
							$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
							#$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
							$WorksheetData[0, $Col++] = 'OS Name'				# Column Number 5
							$WorksheetData[0, $Col++] = 'SP Lvl'				# Column Number 6
							$WorksheetData[0, $Col++] = 'Netlogon Start Mode'	# Column Number 7
							$WorksheetData[0, $Col++] = 'Netlogon State'		# Column Number 8
							$WorksheetData[0, $Col++] = 'Remote Registry Start Mode'	# Column Number 9
							$WorksheetData[0, $Col++] = 'Remote Registry State'	# Column Number 10
							$WorksheetData[0, $Col++] = 'Windows Firewall Start Mode'	# Column Number 11
							$WorksheetData[0, $Col++] = 'Windows Firewall State'	# Column Number 12
							$WorksheetData[0, $Col++] = 'Windows Update Start Mode'	# Column Number 13
							$WorksheetData[0, $Col++] = 'Windows Update State'	# Column Number 14
							$WorksheetData[0, $Col++] = 'BMC Agent Installed'	# Column Number 15
							$WorksheetData[0, $Col++] = 'BMC Agent Start Mode'	# Column Number 16
							$WorksheetData[0, $Col++] = 'BMC Agent State'		# Column Number 17
							$WorksheetData[0, $Col++] = 'Pending Reboot'		# Column Number 18
							$Worksheet5ColumnCount = $Col
						#endregion Worksheet 5: Column Headers

						#region Worksheet 5: Column Values
							$Row = 1
							$Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') } | ForEach-Object {
								$Col = 0

								$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
								$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

								#if ($_.SrvIPv4Addr -eq '&nbsp;')
								#{
								#	$WorksheetData[$Row, $Col++] = ''				#'IP Address'
								#}
								#else
								#{
								#	$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
								#}

								$WorksheetData[$Row, $Col++] = $_.SrvOsName			#'OS Name'
								$WorksheetData[$Row, $Col++] = $_.SrvOsSpLvl		#'SP Lvl'
								$WorksheetData[$Row, $Col++] = $_.SrvNetlogonStartMode		#'Netlogon Start Mode'
								$WorksheetData[$Row, $Col++] = $_.SrvNetlogonState		#'Netlogon State'
								$WorksheetData[$Row, $Col++] = $_.SrvRemoteRegistryStartMode	#'Remote Registry Start Mode'
								$WorksheetData[$Row, $Col++] = $_.SrvRemoteRegistryState	#'Remote Registry State'
								$WorksheetData[$Row, $Col++] = $_.SrvWinFirewallStartMode	#'Windows Firewall Start Mode'
								$WorksheetData[$Row, $Col++] = $_.SrvWinFirewallState	#'Windows Firewall State'
								$WorksheetData[$Row, $Col++] = $_.SrvWinUpdateStartMode	#'Windows Update Start Mode'
								$WorksheetData[$Row, $Col++] = $_.SrvWinUpdateState	#'Windows Update State'
								$WorksheetData[$Row, $Col++] = $_.SrvBMCAgentInstalled	#'BMC Agent Installed'
								$WorksheetData[$Row, $Col++] = $_.SrvBMCAgentStartMode	#'BMC Agent Start Mode'
								$WorksheetData[$Row, $Col++] = $_.SrvBMCAgentState	#'BMC Agent State'
								$WorksheetData[$Row, $Col++] = $_.SrvPendingReboot	#'Pending Reboot'

								$Row++
							}
						#endregion Worksheet 5: Column Values

						$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
						$Range.Value2 = $WorksheetData
						#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
						$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

						$WorksheetFormat.Add($WorksheetNumber, @{
								BoldFirstRow = $true
								BoldFirstColumn = $false
								AutoFilter = $true
								FreezeAtCell = 'A2'
								ApplyCellBorders = $false
								ApplyTableFormatting = $true
								ColumnFormat = @(
									if ($FmtNumberGeneralCols.Length -ge 1)
									{
										foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
										{
											@{
												ColumnNumber = $FmtNumberGeneralCol
												NumberFormat = $XlNumFmtNumberGeneral
											}
										}
									}

									if ($FmtDateCols.Length -ge 1)
									{
										foreach ($FmtDateCol in $FmtDateCols)
										{
											@{
												ColumnNumber = $FmtDateCol
												NumberFormat = $XlNumFmtDate
											}
										}
									}

									if ($FmtNumberS2Cols.Length -ge 1)
									{
										foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
										{
											@{
												ColumnNumber = $FmtNumberS2Col
												NumberFormat = $XlNumFmtNumberS2
											}
										}
									}
								)
								RowFormat = @()
								HeadingsFormat = @()
								BlankRowCellFormat = @()
								ExcessiveUpTimeRowFormat = @()
							}
						)

						$WorksheetNumber++
					#endregion Worksheet 5: Critical Windows Services

					#region Worksheet 6: MS SQL Servers
						if ($IncludeMSSQL)
						{
							$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): MS SQL Servers"
							Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
							Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

							$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
							$Worksheet.Name = 'MS SQL Servers'
							$Worksheet.Tab.ThemeColor = $MSSQLServersTabColor

							## Used to hold the columns where Number or Date formatting needs to be applied.
							## We will them loop over them later to apply the proper formatting on the entire column.
							$FmtNumberGeneralCols = @()
							$FmtDateCols = @()
							$FmtNumberS2Cols = @()

							$RowCount = (($Inventory | Where-Object { $_.SQLSvrIsInstalled -eq 'True' }) | Measure-Object).Count + 1

							$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

							if ($IncludeExcessiveUpTime)
							{
								$ExcessiveUpTimeRows = @()
							}

							## If the total row count is not equal to the value initially set for $Worksheet6RowCount,
							## Then set $Worksheet6RowCount to the Actual used row count. Otherwise leave it alone.
							if ($RowCount -ne $Worksheet6RowCount)
							{
								$Worksheet6RowCount = $RowCount
							}

							#region Worksheet 6: Column Headers
								$Col = 0
								$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
								$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
								$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
								$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
								$WorksheetData[0, $Col++] = 'Product Name'			# Column Number 5
								$WorksheetData[0, $Col++] = 'Product Edition'		# Column Number 6
								$WorksheetData[0, $Col++] = 'Product Version'		# Column Number 7
								$WorksheetData[0, $Col++] = 'SQL Svc Display Names'	# Column Number 8
								$WorksheetData[0, $Col++] = 'SQL Svr Instance Names'# Column Number 9
								$WorksheetData[0, $Col++] = 'SQL Svr Cluster Names'	# Column Number 10
								$WorksheetData[0, $Col++] = 'SQL Svr Cluster Nodes'	# Column Number 11
								$WorksheetData[0, $Col++] = 'SQL Svr Cluster Instance Owner' # Column Number 12
								$WorksheetData[0, $Col++] = 'MS Cluster Name'		# Column Number 13
								$WorksheetData[0, $Col++] = 'Physical Memory (GB)'	# Column Number 14
								$FmtNumberGeneralCols += $Col
								$WorksheetData[0, $Col++] = 'Processor Sockets'		# Column Number 15
								$FmtNumberGeneralCols += $Col
								$WorksheetData[0, $Col++] = 'Processor Cores'		# Column Number 16
								$FmtNumberGeneralCols += $Col
								$WorksheetData[0, $Col++] = 'Logical Processors'	# Column Number 17
								$FmtNumberGeneralCols += $Col
								$WorksheetData[0, $Col++] = 'HT'					# Column Number 18
								$WorksheetData[0, $Col++] = 'Processor Manufacturer'# Column Number 19
								$WorksheetData[0, $Col++] = 'OS Name'				# Column Number 20
								$WorksheetData[0, $Col++] = 'SP Lvl'				# Column Number 21
								$WorksheetData[0, $Col++] = 'Installed OS Arch'		# Column Number 22
								$WorksheetData[0, $Col++] = 'Manufacturer'			# Column Number 23
								$WorksheetData[0, $Col++] = 'Model'					# Column Number 24
								$WorksheetData[0, $Col++] = 'Serial Number'			# Column Number 25
								$WorksheetData[0, $Col++] = 'Data Center Location'	# Column Number 26
								$WorksheetData[0, $Col++] = 'Last Boot Up Time'		# Column Number 27
								$FmtDateCols += $Col
								$WorksheetData[0, $Col++] = 'Up Time (Days)'		# Column Number 28
								$FmtNumberGeneralCols += $Col

								if ($IncludeHotFixDetails)
								{
									$WorksheetData[0, $Col++] = 'Last Patch Install Date'	# Column Number 29
									$FmtDateCols += $Col
									$WorksheetData[0, $Col++] = 'Up Time HotFix 1 Applied'	# Column Number 30
									$WorksheetData[0, $Col++] = 'Up Time HotFix 2 Applied'	# Column Number 31
									$WorksheetData[0, $Col++] = 'Up Time Total HotFix Applied'	# Column Number 32
									$WorksheetData[0, $Col++] = 'MS15-034 HotFix Applied'	# Column Number 33
								}

								$Worksheet6ColumnCount = $Col
							#endregion Worksheet 6: Column Headers

							#region Worksheet 6: Column Values
								$Row = 1
								$Inventory | Where-Object { $_.SQLSvrIsInstalled -eq 'True' } | ForEach-Object {
									$Col = 0

									$WorksheetData[$Row, $Col++] = $_.SrvNBDomain			#'Server NB Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain			#'Server DNS Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsName			#'Server DNS Name'

									if ($_.SrvIPv4Addr -eq '&nbsp;')
									{
										$WorksheetData[$Row, $Col++] = ''					#'IP Address'
									}
									else
									{
										$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr		#'IP Address'
									}

									$WorksheetData[$Row, $Col++] = $_.SQLProductCaptions	#'Product Name'
									$WorksheetData[$Row, $Col++] = $_.SQLEditions			#'Product Edition'
									$WorksheetData[$Row, $Col++] = $_.SQLVersions			#'Product Version'
									$WorksheetData[$Row, $Col++] = $_.SQLSvcDisplayNames	#'SQL Svc Display Names'
									$WorksheetData[$Row, $Col++] = $_.SQLInstanceNames		#'SQL Svr Instance Names'
									$WorksheetData[$Row, $Col++] = $_.SQLClusterNames		#'SQL Svr Cluster Names'
									$WorksheetData[$Row, $Col++] = $_.SQLClusterNodes		#'SQL Svr Cluster Nodes'
									$WorksheetData[$Row, $Col++] = $_.SQLClusterInstanceOwner #'SQL Svr Cluster Instance Owner'
									$WorksheetData[$Row, $Col++] = $_.MSClusterName			#'MS Cluster Name'
									$WorksheetData[$Row, $Col++] = $_.SrvPhysMemoryGB		#'Physical Memory (GB)'
									$WorksheetData[$Row, $Col++] = $_.SrvSocketCount		#'Processor Sockets'
									$WorksheetData[$Row, $Col++] = $_.SrvProcCoreCount		#'Processor Cores'
									$WorksheetData[$Row, $Col++] = $_.SrvLogProcsCount		#'Logical Processors'
									$WorksheetData[$Row, $Col++] = $_.SrvHyperT				#'HT' - Intel Hyperthreading -OR- AMD HyperTransport
									$WorksheetData[$Row, $Col++] = $_.SrvProcMnftr			#'Processor Manufacturer'
									$WorksheetData[$Row, $Col++] = $_.SrvOsName				#'OS Name'
									$WorksheetData[$Row, $Col++] = $_.SrvOsSpLvl			#'SP Lvl'
									$WorksheetData[$Row, $Col++] = $_.SrvInstOsArch			#'Installed OS Arch'
									$WorksheetData[$Row, $Col++] = $_.SrvMnfctr				#'Manufacturer'
									$WorksheetData[$Row, $Col++] = $_.SrvModel				#'Model'
									$WorksheetData[$Row, $Col++] = $_.SrvSerNum				#'Serial Number'
									$WorksheetData[$Row, $Col++] = $_.SrvDataCenter			#'Data Center Location'
									$WorksheetData[$Row, $Col++] = $_.SrvLastBootUpTime		#'Last Boot Up Time'
									$WorksheetData[$Row, $Col++] = $_.SrvUpTime				#'Up Time (Days)'

									if ($IncludeHotFixDetails)
									{
										$WorksheetData[$Row, $Col++] = $_.SrvLastPatchInstallDate			#'Last Patch Install Date'
										$WorksheetData[$Row, $Col++] = $_.SrvUpTimeKB2553549HotFix1Applied	#'Up Time HotFix 1 Applied'
										$WorksheetData[$Row, $Col++] = $_.SrvUpTimeKB2688338HotFix2Applied	#'Up Time HotFix 2 Applied'
										$WorksheetData[$Row, $Col++] = $_.SrvUpTimeTotalHotFixApplied		#'Up Time Total HotFix Applied
										$WorksheetData[$Row, $Col++] = $_.SrvKB3042553HotFixApplied			#'MS15-034 HotFix Applied'
									}

									$Row++

									if ($IncludeExcessiveUpTime)
									{
										if ([int]$_.SrvUpTime -gt $ExcessiveUpTimeDays)
										{
											## Used to hold the rows where servers are showing as up for more than '$ExcessiveUpTimeDays' days.
											## We will them loop over it later to change the font color on the entire row.
											$ExcessiveUpTimeRows += $Row
										}
									}
								}
							#endregion Worksheet 6: Column Values

							$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
							$Range.Value2 = $WorksheetData
							#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
							$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

							$WorksheetFormat.Add($WorksheetNumber, @{
									BoldFirstRow = $true
									BoldFirstColumn = $false
									AutoFilter = $true
									FreezeAtCell = 'A2'
									ApplyCellBorders = $false
									ApplyTableFormatting = $true
									ColumnFormat = @(
										#@{ColumnNumber = 12; NumberFormat = $XlNumFmtNumberGeneral},	#'Physical Memory (GB)'
										#@{ColumnNumber = 13; NumberFormat = $XlNumFmtNumberGeneral},	#'Processor Sockets'
										#@{ColumnNumber = 14; NumberFormat = $XlNumFmtNumberGeneral},	#'Processor Cores'
										#@{ColumnNumber = 15; NumberFormat = $XlNumFmtNumberGeneral},	#'Logical Processors'
										#@{ColumnNumber = 25; NumberFormat = $XlNumFmtDate},				#'Last Boot Up Time'
										#@{ColumnNumber = 26; NumberFormat = $XlNumFmtNumberGeneral}		#'Up Time (Days)'

										if ($FmtNumberGeneralCols.Length -ge 1)
										{
											foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
											{
												@{
													ColumnNumber = $FmtNumberGeneralCol
													NumberFormat = $XlNumFmtNumberGeneral
												}
											}
										}

										if ($FmtDateCols.Length -ge 1)
										{
											foreach ($FmtDateCol in $FmtDateCols)
											{
												@{
													ColumnNumber = $FmtDateCol
													NumberFormat = $XlNumFmtDate
												}
											}
										}

										if ($FmtNumberS2Cols.Length -ge 1)
										{
											foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
											{
												@{
													ColumnNumber = $FmtNumberS2Col
													NumberFormat = $XlNumFmtNumberS2
												}
											}
										}
									)
									RowFormat = @()
									HeadingsFormat = @()
									BlankRowCellFormat = @()
									ExcessiveUpTimeRowFormat = @(
										if ($IncludeExcessiveUpTime)
										{
											foreach ($ExcessiveUpTimeRow in $ExcessiveUpTimeRows)
											{
												@{
													RowNumber = $ExcessiveUpTimeRow
													StartColumnNumber = 1
													EndColumnNumber = $Worksheet6ColumnCount
													FontBold = $true
													FontSize = 11
													FontColor = $ExcessiveUpTimeColorIndex
													#BackGroundColor = 10
													VerticalAlignment = $XlVAlign::xlVAlignTop
													HorizontalAlignment = $XlHAlign::xlHAlignLeft
												}
											}
										}
									)
								}
							)

							$WorksheetNumber++
						}
					#endregion Worksheet 6: MS SQL Servers

					#region Worksheet 7: HotFix Details
						if ($IncludeHotFixDetails)
						{
							$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): HotFix Details"
							Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
							Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

							$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
							$Worksheet.Name = 'HotFix Details'
							$Worksheet.Tab.ThemeColor = $HotFixDetailsTabColor

							## Used to hold the columns where Number or Date formatting needs to be applied.
							## We will them loop over them later to apply the proper formatting on the entire column.
							$FmtNumberGeneralCols = @()
							$FmtDateCols = @()
							$FmtNumberS2Cols = @()

							$RowCount = (($Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') }) | Measure-Object).Count + 1

							$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

							## If the total row count is not equal to the value initially set for $Worksheet7RowCount,
							## Then set $Worksheet7RowCount to the Actual used row count. Otherwise leave it alone.
							if ($RowCount -ne $Worksheet7RowCount)
							{
								$Worksheet7RowCount = $RowCount
							}

							#region Worksheet 7: Column Headers
								$Col = 0
								$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
								$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
								$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
								#$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
								$WorksheetData[0, $Col++] = 'OS Name'				# Column Number 5
								$WorksheetData[0, $Col++] = 'SP Lvl'				# Column Number 6
								$WorksheetData[0, $Col++] = 'Last Patch Install Date'	# Column Number 7
								$FmtDateCols += $Col
								$WorksheetData[0, $Col++] = 'Up Time HotFix 1 Applied'	# Column Number 8
								$WorksheetData[0, $Col++] = 'Up Time HotFix 2 Applied'	# Column Number 9
								$WorksheetData[0, $Col++] = 'Up Time Total HotFix Applied'	# Column Number 10
								$WorksheetData[0, $Col++] = 'MS15-034 HotFix Applied'	# Column Number 11
								$Worksheet7ColumnCount = $Col
							#endregion Worksheet 7: Column Headers

							#region Worksheet 7: Column Values
								$Row = 1
								$Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') } | ForEach-Object {
									$Col = 0

									$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

									#if ($_.SrvIPv4Addr -eq '&nbsp;')
									#{
									#	$WorksheetData[$Row, $Col++] = ''				#'IP Address'
									#}
									#else
									#{
									#	$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
									#}

									$WorksheetData[$Row, $Col++] = $_.SrvOsName			#'OS Name'
									$WorksheetData[$Row, $Col++] = $_.SrvOsSpLvl		#'SP Lvl'
									$WorksheetData[$Row, $Col++] = $_.SrvLastPatchInstallDate	#'Last Patch Install Date'
									$WorksheetData[$Row, $Col++] = $_.SrvUpTimeKB2553549HotFix1Applied	#'Up Time HotFix 1 Applied'
									$WorksheetData[$Row, $Col++] = $_.SrvUpTimeKB2688338HotFix2Applied	#'Up Time HotFix 2 Applied'
									$WorksheetData[$Row, $Col++] = $_.SrvUpTimeTotalHotFixApplied	#'Up Time Total HotFix Applied'
									$WorksheetData[$Row, $Col++] = $_.SrvKB3042553HotFixApplied	#'MS15-034 HotFix Applied'

									$Row++
								}
							#endregion Worksheet 7: Column Values

							$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
							$Range.Value2 = $WorksheetData
							#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
							$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

							$WorksheetFormat.Add($WorksheetNumber, @{
									BoldFirstRow = $true
									BoldFirstColumn = $false
									AutoFilter = $true
									FreezeAtCell = 'A2'
									ApplyCellBorders = $false
									ApplyTableFormatting = $true
									ColumnFormat = @(
										if ($FmtNumberGeneralCols.Length -ge 1)
										{
											foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
											{
												@{
													ColumnNumber = $FmtNumberGeneralCol
													NumberFormat = $XlNumFmtNumberGeneral
												}
											}
										}

										if ($FmtDateCols.Length -ge 1)
										{
											foreach ($FmtDateCol in $FmtDateCols)
											{
												@{
													ColumnNumber = $FmtDateCol
													NumberFormat = $XlNumFmtDate
												}
											}
										}

										if ($FmtNumberS2Cols.Length -ge 1)
										{
											foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
											{
												@{
													ColumnNumber = $FmtNumberS2Col
													NumberFormat = $XlNumFmtNumberS2
												}
											}
										}
									)
									RowFormat = @()
									HeadingsFormat = @()
									BlankRowCellFormat = @()
									ExcessiveUpTimeRowFormat = @()
								}
							)

							$WorksheetNumber++
						}
					#endregion Worksheet 7: HotFix Details

					#region Worksheet 8: Custom App Info
						if ($IncludeCustomAppInfo)
						{
							$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Custom App Info"
							Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
							Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

							$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
							$Worksheet.Name = 'Custom App Info'
							$Worksheet.Tab.ThemeColor = $CustomAppInfoTabColor

							## Used to hold the columns where Number or Date formatting needs to be applied.
							## We will them loop over them later to apply the proper formatting on the entire column.
							$FmtNumberGeneralCols = @()
							$FmtDateCols = @()
							$FmtNumberS2Cols = @()

							$RowCount = (($Inventory | Where-Object { ($_.SrvBrowserHawkInstall -eq 'Yes') -or ($_.SrvSoftArtisansFileUpInstall -eq 'Yes') }) | Measure-Object).Count + 1

							$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

							## If the total row count is not equal to the value initially set for $Worksheet8RowCount,
							## Then set $Worksheet8RowCount to the Actual used row count. Otherwise leave it alone.
							if ($RowCount -ne $Worksheet8RowCount)
							{
								$Worksheet8RowCount = $RowCount
							}

							#region Worksheet 8: Column Headers
								$Col = 0
								$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
								$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
								$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
								#$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4
								$WorksheetData[0, $Col++] = 'BrowserHawk Installed'		# Column Number 5
								$WorksheetData[0, $Col++] = 'SoftArtisans FileUp Installed'	# Column Number 6
								$Worksheet8ColumnCount = $Col
							#endregion Worksheet 8: Column Headers

							#region Worksheet 8: Column Values
								$Row = 1
								$Inventory | Where-Object { ($_.SrvBrowserHawkInstall -eq 'Yes') -or ($_.SrvSoftArtisansFileUpInstall -eq 'Yes') } | ForEach-Object {
									$Col = 0

									$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

									#if ($_.SrvIPv4Addr -eq '&nbsp;')
									#{
									#	$WorksheetData[$Row, $Col++] = ''				#'IP Address'
									#}
									#else
									#{
									#	$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
									#}

									$WorksheetData[$Row, $Col++] = $_.SrvBrowserHawkInstall		#'BrowserHawk Installed'
									$WorksheetData[$Row, $Col++] = $_.SrvSoftArtisansFileUpInstall	#'SoftArtisans FileUp Installed'

									$Row++
								}
							#endregion Worksheet 8: Column Values

							$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
							$Range.Value2 = $WorksheetData
							#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
							$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

							$WorksheetFormat.Add($WorksheetNumber, @{
									BoldFirstRow = $true
									BoldFirstColumn = $false
									AutoFilter = $true
									FreezeAtCell = 'A2'
									ApplyCellBorders = $false
									ApplyTableFormatting = $true
									ColumnFormat = @(
										if ($FmtNumberGeneralCols.Length -ge 1)
										{
											foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
											{
												@{
													ColumnNumber = $FmtNumberGeneralCol
													NumberFormat = $XlNumFmtNumberGeneral
												}
											}
										}

										if ($FmtDateCols.Length -ge 1)
										{
											foreach ($FmtDateCol in $FmtDateCols)
											{
												@{
													ColumnNumber = $FmtDateCol
													NumberFormat = $XlNumFmtDate
												}
											}
										}

										if ($FmtNumberS2Cols.Length -ge 1)
										{
											foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
											{
												@{
													ColumnNumber = $FmtNumberS2Col
													NumberFormat = $XlNumFmtNumberS2
												}
											}
										}
									)
									RowFormat = @()
									HeadingsFormat = @()
									BlankRowCellFormat = @()
									ExcessiveUpTimeRowFormat = @()
								}
							)

							$WorksheetNumber++
						}
					#endregion Worksheet 8: Custom App Info

					#region Worksheet 9: Infrastructure Details
						if ($IncludeInfraDetails)
						{
							$ProgressStatus = "Writing Worksheet #$($WorksheetNumber): Infrastructure Details"
							Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
							Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

							$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)
							$Worksheet.Name = 'Infrastructure Details'
							$Worksheet.Tab.ThemeColor = $InfraDetailsDetailsTabColor

							## Used to hold the columns where Number or Date formatting needs to be applied.
							## We will them loop over them later to apply the proper formatting on the entire column.
							$FmtNumberGeneralCols = @()
							$FmtDateCols = @()
							$FmtNumberS2Cols = @()

							$RowCount = (($Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') }) | Measure-Object).Count + 1

							$WorksheetData = New-Object -TypeName 'string[,]' -ArgumentList $RowCount, $ColumnCount

							## If the total row count is not equal to the value initially set for $Worksheet9RowCount,
							## Then set $Worksheet9RowCount to the Actual used row count. Otherwise leave it alone.
							if ($RowCount -ne $Worksheet9RowCount)
							{
								$Worksheet9RowCount = $RowCount
							}

							#region Worksheet 9: Column Headers
								$Col = 0
								$WorksheetData[0, $Col++] = 'Server NB Domain'		# Column Number 1
								$WorksheetData[0, $Col++] = 'Server DNS Domain'		# Column Number 2
								$WorksheetData[0, $Col++] = 'Server DNS Name'		# Column Number 3
								#$WorksheetData[0, $Col++] = 'IP Address'			# Column Number 4

								$WorksheetData[0, $Col++] = 'CPU_Type'					# Column Number 5
								$WorksheetData[0, $Col++] = 'Disk_Drives'				# Column Number 6
								$WorksheetData[0, $Col++] = 'Disks_Space_in-used'		# Column Number 7
								$WorksheetData[0, $Col++] = 'Drives_free_Space'			# Column Number 8
								$WorksheetData[0, $Col++] = 'Total_Disk_Allocated_GB'	# Column Number 9
								$FmtNumberS2Cols += $Col
								$WorksheetData[0, $Col++] = 'Total_Disks_in-used_GB'	# Column Number 10
								$FmtNumberS2Cols += $Col
								$WorksheetData[0, $Col++] = 'Total_Disks_free_Space_GB'	# Column Number 11
								$FmtNumberS2Cols += $Col
								$Worksheet9ColumnCount = $Col
							#endregion Worksheet 9: Column Headers

							#region Worksheet 9: Column Values
								$Row = 1
								$Inventory | Where-Object { ($_.SrvPhysMemoryGB -ne 'No Response') -and ($_.SrvPhysMemoryGB -ne 'Error Encountered') } | ForEach-Object {
									$Col = 0

									$WorksheetData[$Row, $Col++] = $_.SrvNBDomain		#'Server NB Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsDomain		#'Server DNS Domain'
									$WorksheetData[$Row, $Col++] = $_.SrvDnsName		#'Server DNS Name'

									#if ($_.SrvIPv4Addr -eq '&nbsp;')
									#{
									#	$WorksheetData[$Row, $Col++] = ''				#'IP Address'
									#}
									#else
									#{
									#	$WorksheetData[$Row, $Col++] = $_.SrvIPv4Addr	#'IP Address'
									#}

									$WorksheetData[$Row, $Col++] = $_.SrvProcName			#'CPU_Type'	#'Processor Family'
									$WorksheetData[$Row, $Col++] = $_.SrvDiskTotalInfo		#'Disk_Drives'
									$WorksheetData[$Row, $Col++] = $_.SrvDiskUsedInfo		#'Disks_Space_in-used'
									$WorksheetData[$Row, $Col++] = $_.SrvDiskFreeInfo		#'Drives_free_Space'
									$WorksheetData[$Row, $Col++] = $_.SrvDiskTotalAllocated	#'Total_Disk_Allocated_GB'
									$WorksheetData[$Row, $Col++] = $_.SrvDiskUsedAllocated	#'Total_Disks_in-used_GB'
									$WorksheetData[$Row, $Col++] = $_.SrvDiskFreeAllocated	#'Total_Disks_free_Space_GB'

									$Row++
								}
							#endregion Worksheet 9: Column Values

							$Range = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($RowCount, $ColumnCount))
							$Range.Value2 = $WorksheetData
							#$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $MissingType, $MissingType, $XlYesNoGuess::xlYes) | Out-Null
							$Range.Sort($Worksheet.Columns.Item(1), $XlSortOrder::xlAscending, $Worksheet.Columns.Item(2), $MissingType, $XlSortOrder::xlAscending, $Worksheet.Columns.Item(3), $XlSortOrder::xlAscending, $XlYesNoGuess::xlYes) | Out-Null

							$WorksheetFormat.Add($WorksheetNumber, @{
									BoldFirstRow = $true
									BoldFirstColumn = $false
									AutoFilter = $true
									FreezeAtCell = 'A2'
									ApplyCellBorders = $false
									ApplyTableFormatting = $true
									ColumnFormat = @(
										if ($FmtNumberGeneralCols.Length -ge 1)
										{
											foreach ($FmtNumberGeneralCol in $FmtNumberGeneralCols)
											{
												@{
													ColumnNumber = $FmtNumberGeneralCol
													NumberFormat = $XlNumFmtNumberGeneral
												}
											}
										}

										if ($FmtDateCols.Length -ge 1)
										{
											foreach ($FmtDateCol in $FmtDateCols)
											{
												@{
													ColumnNumber = $FmtDateCol
													NumberFormat = $XlNumFmtDate
												}
											}
										}

										if ($FmtNumberS2Cols.Length -ge 1)
										{
											foreach ($FmtNumberS2Col in $FmtNumberS2Cols)
											{
												@{
													ColumnNumber = $FmtNumberS2Col
													NumberFormat = $XlNumFmtNumberS2
												}
											}
										}
									)
									RowFormat = @()
									HeadingsFormat = @()
									BlankRowCellFormat = @()
									ExcessiveUpTimeRowFormat = @()
								}
							)

							$WorksheetNumber++
						}
					#endregion Worksheet 9: Infrastructure Details

					#region Apply formatting to every worksheet
						## Work backwards so that the first sheet is active when the workbook is saved.
						$ProgressStatus = 'Applying formatting to all worksheets...'
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
						Write-Progress -Activity $ProgressActivity -PercentComplete (($WorksheetNumber / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

						for ( $WorksheetNumber = $WorksheetCount; $WorksheetNumber -ge 1; $WorksheetNumber-- )
						{
							$ProgressStatus = "Applying formatting to Worksheet #$($WorksheetNumber)"
							Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
							Write-Progress -Activity $ProgressActivity -PercentComplete (((($WorksheetCount * 2) - $WorksheetNumber + 1) / ($WorksheetCount * 2)) * 100) -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId

							if ($ScrDebug)
							{
								Write-Host -ForegroundColor Green $ProgressStatus
							}

							$Worksheet = $Excel.Worksheets.Item($WorksheetNumber)

							## Switch to the worksheet
							$Worksheet.Activate() | Out-Null

							switch ($WorksheetNumber)
							{
								1 {
									$WorksheetRowCount = $Worksheet1RowCount
									$WorksheetColumnCount = $Worksheet1ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight2' }
										'medium' { 'TableStyleMedium9' }
										'dark' { 'TableStyleDark2' }
									}
								}
								2 {
									$WorksheetRowCount = $Worksheet2RowCount
									$WorksheetColumnCount = $Worksheet2ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight14' }
										'medium' { 'TableStyleMedium14' }
										'dark' { 'TableStyleDark7' }
									}
								}
								3 {
									$WorksheetRowCount = $Worksheet3RowCount
									$WorksheetColumnCount = $Worksheet3ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight12' }
										'medium' { 'TableStyleMedium12' }
										'dark' { 'TableStyleDark5' }
									}
								}
								4 {
									$WorksheetRowCount = $Worksheet4RowCount
									$WorksheetColumnCount = $Worksheet4ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight8' }
										'medium' { 'TableStyleMedium15' }
										'dark' { 'TableStyleDark1' }
									}
								}
								5 {
									$WorksheetRowCount = $Worksheet5RowCount
									$WorksheetColumnCount = $Worksheet5ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight9' }
										'medium' { 'TableStyleMedium9' }
										'dark' { 'TableStyleDark2' }
									}
								}
								6 {
									$WorksheetRowCount = $Worksheet6RowCount
									$WorksheetColumnCount = $Worksheet6ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight10' }
										'medium' { 'TableStyleMedium10' }
										'dark' { 'TableStyleDark3' }
									}
								}
								7 {
									$WorksheetRowCount = $Worksheet7RowCount
									$WorksheetColumnCount = $Worksheet7ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight20' }
										'medium' { 'TableStyleMedium13' }
										'dark' { 'TableStyleDark2' }
									}
								}
								8 {
									$WorksheetRowCount = $Worksheet8RowCount
									$WorksheetColumnCount = $Worksheet8ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight1' }
										'medium' { 'TableStyleMedium4' }
										'dark' { 'TableStyleDark8' }
									}
								}
								9 {
									$WorksheetRowCount = $Worksheet9RowCount
									$WorksheetColumnCount = $Worksheet9ColumnCount
									$TableStyle = switch ($ColorScheme)
									{
										'light' { 'TableStyleLight15' }
										'medium' { 'TableStyleMedium15' }
										'dark' { 'TableStyleDark1' }
									}
								}
							}

							## Bold the header row
							$Worksheet.Rows.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstRow

							## Bold the 1st column
							$Worksheet.Columns.Item(1).Font.Bold = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn

							## AutoFilter
							if ($WorksheetFormat[$WorksheetNumber].AutoFilter)
							{
								$WorksheetAutoFilter = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item(1, $WorksheetColumnCount)).AutoFilter()
							}

							## Freeze View
							$Worksheet.Range($WorksheetFormat[$WorksheetNumber].FreezeAtCell).Select() | Out-Null
							$Worksheet.Application.ActiveWindow.FreezePanes = $true

							## Apply Column formatting
							$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
								$Worksheet.Columns.Item($_.ColumnNumber).NumberFormat = $_.NumberFormat
							}

							## Apply Row formatting
							$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
								$Worksheet.Rows.Item($_.RowNumber).NumberFormat = $_.NumberFormat
							}

							## Update worksheet values so row and column formatting apply
							try
							{
								$Worksheet.UsedRange.Value2 = $Worksheet.UsedRange.Value2
							}
							catch
							{
								## Sometimes trying to set the entire worksheet's value to itself will result in the following exception:
								## 	"Not enough storage is available to complete this operation. 0x8007000E (E_OUTOFMEMORY))"
								## See http://support.microsoft.com/kb/313275 for more information
								## When this happens the workaround is to try doing the work in smaller chunks
								## ...so we'll try to update the column\row values that have specific formatting one at a time instead of the entire worksheet at once
								$WorksheetFormat[$WorksheetNumber].ColumnFormat | ForEach-Object {
									$Worksheet.Columns.Item($_.ColumnNumber).Value2 = $Worksheet.Columns.Item($_.ColumnNumber).Value2
								}

								$WorksheetFormat[$WorksheetNumber].RowFormat | ForEach-Object {
									$Worksheet.Rows.Item($_.RowNumber).Value2 = $Worksheet.Rows.Item($_.RowNumber).Value2
								}
							}

							## Apply table formatting
							#if ($WorksheetNumber -ne 1)
							if ($WorksheetFormat[$WorksheetNumber].ApplyTableFormatting)
							{
								$ListObject = $Worksheet.ListObjects.Add($XlListObjectSourceType::xlSrcRange, $Worksheet.UsedRange, $null, $XlYesNoGuess::xlYes, $null) 
								$ListObject.Name = "Table $($WorksheetNumber)"
								$ListObject.TableStyle = $TableStyle
								## Put a background color behind the 1st column
								$ListObject.ShowTableStyleFirstColumn = $WorksheetFormat[$WorksheetNumber].BoldFirstColumn
								$ListObject.ShowAutoFilter = $WorksheetFormat[$WorksheetNumber].AutoFilter
							}

							## Zoom back to 80%
							#$Worksheet.Application.ActiveWindow.Zoom = 80

							## Adjust the column widths to 250 before autofitting contents
							## This allows longer lines of text to remain on one line
							$Worksheet.UsedRange.EntireColumn.ColumnWidth = 250

							## Wrap text
							$Worksheet.UsedRange.WrapText = $true

							## Left align contents
							$Worksheet.UsedRange.EntireColumn.HorizontalAlignment = $XlHAlign::xlHAlignLeft

							## Vertical align contents
							$Worksheet.UsedRange.EntireColumn.VerticalAlignment = $XlVAlign::xlVAlignTop

							## Apply Headings Formatting
							$WorksheetFormat[$WorksheetNumber].HeadingsFormat | ForEach-Object {
								$HeadingRange = $Worksheet.Range($Worksheet.Cells.Item($_.RowNumber, $_.StartColumnNumber), $Worksheet.Cells.Item($_.RowNumber, $_.EndColumnNumber))

								if (($WorksheetNumber -eq 1) -and ($_.RowNumber -eq 1))
								{
									$HeadingRange.Merge() | Out-Null
								}

								$HeadingRange.Font.Bold = $_.FontBold
								$HeadingRange.Font.Size = $_.FontSize
								$HeadingRange.Font.ColorIndex = $_.FontColor
								$HeadingRange.Interior.ColorIndex = $_.BackGroundColor
								$HeadingRange.HorizontalAlignment = $_.HorizontalAlignment
								$HeadingRange.VerticalAlignment = $_.VerticalAlignment
							}

							## Apply Blank Row Cell Formatting
							$WorksheetFormat[$WorksheetNumber].BlankRowCellFormat | ForEach-Object {
								$BlankRowCellRange = $Worksheet.Range($Worksheet.Cells.Item($_.RowNumber, $_.StartColumnNumber), $Worksheet.Cells.Item($_.RowNumber, $_.EndColumnNumber))

								if ($_.MergeCells)
								{
									$BlankRowCellRange.Merge() | Out-Null
								}

								$BlankRowCellRange.Font.Size = $_.FontSize
								$BlankRowCellRange.Font.ColorIndex = $_.FontColor
								$BlankRowCellRange.Interior.ColorIndex = $_.BackGroundColor
								$BlankRowCellRange.RowHeight = $_.RowHeight
								$BlankRowCellRange.HorizontalAlignment = $_.HorizontalAlignment
								$BlankRowCellRange.VerticalAlignment = $_.VerticalAlignment
							}

							## Apply Excessive Up Time Row Formatting
							if ($IncludeExcessiveUpTime)
							{
								$WorksheetFormat[$WorksheetNumber].ExcessiveUpTimeRowFormat | ForEach-Object {
									$ExcessiveUpTimeRowCellRange = $Worksheet.Range($Worksheet.Cells.Item($_.RowNumber, $_.StartColumnNumber), $Worksheet.Cells.Item($_.RowNumber, $_.EndColumnNumber))

									$ExcessiveUpTimeRowCellRange.Font.Bold = $_.FontBold
									$ExcessiveUpTimeRowCellRange.Font.Size = $_.FontSize
									$ExcessiveUpTimeRowCellRange.Font.ColorIndex = $_.FontColor
									#$ExcessiveUpTimeRowCellRange.Interior.ColorIndex = $_.BackGroundColor
									$ExcessiveUpTimeRowCellRange.HorizontalAlignment = $_.HorizontalAlignment
									$ExcessiveUpTimeRowCellRange.VerticalAlignment = $_.VerticalAlignment
								}
							}

							## Apply Cell Borders
							if ($WorksheetFormat[$WorksheetNumber].ApplyCellBorders)
							{
								##	Microsoft.Office.Interop.Excel.XlBordersIndex
								##	xlDiagonalDown = 5
								##	xlDiagonalUp = 6
								##	xlEdgeLeft = 7
								##	xlEdgeTop = 8
								##	xlEdgeBottom = 9
								##	xlEdgeRight = 10
								##	xlInsideVertical = 11
								##	xlInsideHorizontal = 12
								##
								##	Microsoft.Office.Interop.Excel.XlBorderWeight
								##	xlHairline = 1
								##	xlMedium = -4138
								##	xlThick = 4
								##	xlThin = 2

								#$BordersRange = $Worksheet.Range($Worksheet.Cells.Item(2, 1), $Worksheet.Cells.Item($WorksheetRowCount, $WorksheetColumnCount))
								#7..12 | ForEach-Object {
								#	$BordersRange.Borders.Item($_).LineStyle = $XlLineStyle::xlContinuous	#1
								#	$BordersRange.Borders.Item($_).Weight = $XlBorderWeight::xlThin	#2
								#}
								$BordersRange = $Worksheet.UsedRange
								$BordersRange.Borders.LineStyle = $XlLineStyle::xlContinuous
								$BordersRange.Borders.Weight = $XlBorderWeight::xlMedium	#xlThin
							}

							## Autofit column and row contents
							$Worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
							#$Worksheet.UsedRange.EntireRow.AutoFit() | Out-Null

							## Put the selection back to the upper left cell
							$Worksheet.Range('A1').Select() | Out-Null
						}
					#endregion Apply formatting to every worksheet
				}
				catch
				{
					## Something more sophisticated needs to be built for this section!! ##
					throw
				}
				finally
				{
					## Save Workbook and Quit Excel.
					##$Worksheet.Application.DisplayAlerts = $false
					$Workbook.SaveAs($Path)
					$Workbook.Saved = $true

					## Turn on screen updating
					$Excel.ScreenUpdating = $true

					## Turn on automatic calculations
					##$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationAutomatic

					$Excel.Quit()
				}
			#endregion Write to Excel

			$ProgressStatus = 'Output to Excel complete'
			Write-Progress -Activity $ProgressActivity -PercentComplete 100 -Status $ProgressStatus -Id $ProgressId -ParentId $ParentProgressId -Completed
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t $($ProgressStatus)"
		}
		end
		{
			Remove-Variable -Name Range -Confirm:$false
			Remove-Variable -Name Row -Confirm:$false
			Remove-Variable -Name RowCount -Confirm:$false
			Remove-Variable -Name Col -Confirm:$false
			Remove-Variable -Name ColumnCount -Confirm:$false
			Remove-Variable -Name MissingType -Confirm:$false
			Remove-Variable -Name ColorThemePathPattern -Confirm:$false
			Remove-Variable -Name ColorThemePath -Confirm:$false
			Remove-Variable -Name XlSortOrder -Confirm:$false
			Remove-Variable -Name XlYesNoGuess -Confirm:$false
			Remove-Variable -Name XlListObjectSourceType -Confirm:$false
			Remove-Variable -Name XlThemeColor -Confirm:$false
			Remove-Variable -Name OverviewTabColor -Confirm:$false
			Remove-Variable -Name TableStyle -Confirm:$false
			Remove-Variable -Name WorksheetData -Confirm:$false
			Remove-Variable -Name WorksheetFormat -Confirm:$false
			Remove-Variable -Name WorksheetNumber -Confirm:$false
			Remove-Variable -Name Worksheet -Confirm:$false
			Remove-Variable -Name Workbook -Confirm:$false
			Remove-Variable -Name Excel -Confirm:$false

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t End Function: $($MyInvocation.InvocationName)"
		}
	}

	Function Send-EmailWithReports()
	{
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = 'Medium'
		)]
		#region Function Parameters
			Param(
				[Parameter(
					Position			= 0
					, Mandatory			= $true
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'HTML for the Email Message Body.'
				)][string]$MsgBody
				, [Parameter(
					Position			= 1
					, Mandatory			= $true
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Email Address that the email will be sent from: "From".'
				)][string]$AddrFrom
				, [Parameter(
					Position			= 2
					, Mandatory			= $true
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'A comma separated list of Email Addresses that the email will be sent to: "To".'
				)][array]$AddrTo
				, [Parameter(
					Position			= 3
					, Mandatory			= $true
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Outbound SMTP server address. Passed to Send-MailMessage as "-SmtpServer".'
				)][string]$SmtpHost
				, [Parameter(
					Position			= 4
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Attachments to be sent.'
				)]$Attachments
			)
		#endregion Function Parameters

		begin
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tStart Function: $($MyInvocation.InvocationName)"

			## Setup splatting hashtable to hold the values to be passed to the 'Send-MailMessage' cmdlet.
			$SendMailMsg = @{}

			[string]$RptDate = (Get-Date -Format 'ddd MMMM dd, yyyy').ToString()		## Example: Fri September 13, 2013
			[string]$MsgSubject	= "[REPORT - $($DataCenterCity.ToUpper())] Windows Server Inventory for Licensing - $($DataCenterCity) Data Centers - Report for $($RptDate)"
		}
		process
		{
			$SendMailMsg.Add('SmtpServer', $SmtpHost)
			$SendMailMsg.Add('From', $AddrFrom)
			$SendMailMsg.Add('BodyAsHtml', $true)
			$SendMailMsg.Add('Body', $MsgBody.ToString())
			$SendMailMsg.Add('To', $AddrTo)
			$SendMailMsg.Add('Subject', $MsgSubject)
			$SendMailMsg.Add('Encoding', 'UTF8')

			if ($Attachments)
			{
				$SendMailMsg.Add('Attachments', $Attachments)
			}
		}
		end
		{
			Send-MailMessage @SendMailMsg

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t End Function: $($MyInvocation.InvocationName)"
		}
	}

	Function New-ScheduledJobTask()
	{
		[CmdletBinding(
			SupportsShouldProcess = $true,
			ConfirmImpact = 'Medium',
			DefaultParameterSetName = 'SettingsFile'
		)]
		#region Function Parameters
			Param(
				[Parameter(
					Position			= 0
					, ParameterSetName	= 'SettingsFile'
					, Mandatory			= $true
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'The name of the XML settings file. If the settings file is not located in the same directory as the script then provide the full path to the file.'
				)]
				[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
				[string]$SettingsFile
				, [Parameter(
					Position			= 0
					, ParameterSetName	= 'Manual'
					, Mandatory			= $true
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Partial file name used to create the full report(s) and log filenames. Defaults to the script name without its file extension.'
				)][string]$PtlRptName
				, [Parameter(
					Position			= 1
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Send Email ($true/$false)'
				)][switch]$SendEmail = $true
				, [Parameter(
					Position			= 2
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Mail From Address'
				)][string]$AddrFrom
				, [Parameter(
					Position			= 3
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Mail To Address'
				)][array]$AddrTo
				, [Parameter(
					Position			= 4
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Mail Server Address'
				)][string]$SmtpHost
				, [Parameter(
					Position			= 1
					, ParameterSetName	= 'SettingsFile'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Schedule as user'
				)]
				[Parameter(
					Position			= 5
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Schedule as user'
				)]
				[string]$ScheduleAsUser
				, [Parameter(
					Position			= 2
					, ParameterSetName	= 'SettingsFile'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Schedule as user Password'
				)]
				[Parameter(
					Position			= 6
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'Schedule as user Password'
				)]
				[string]$ScheduleAsUserPassword
				, [Parameter(
					Position			= 3
					, ParameterSetName	= 'SettingsFile'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'A name which will uniquely identify the scheduled task.'
				)]
				[Parameter(
					Position			= 7
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline	= $false
					#, HelpMessage		= 'A name which will uniquely identify the scheduled task.'
				)]
				[string]$ScheduledTaskName
				, [Parameter(
					Position			= 4
					, ParameterSetName	= 'SettingsFile'
					, Mandatory			= $false
					#, ValueFromPipeline = $false
					#, HelpMessage		= 'The full path to the directory where the script is located.'
				)]
				[Parameter(
					Position			= 8
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline = $false
					#, HelpMessage		= 'The full path to the directory where the script is located.'
				)]
				[string]$WorkingDir
				, [Parameter(
					Position			= 5
					, ParameterSetName	= 'SettingsFile'
					, Mandatory			= $false
					#, ValueFromPipeline = $false
					#, HelpMessage		= 'The full path to the directory where the script is located with the full script name and extension.'
				)]
				[Parameter(
					Position			= 9
					, ParameterSetName	= 'Manual'
					, Mandatory			= $false
					#, ValueFromPipeline = $false
					#, HelpMessage		= 'The full path to the directory where the script is located with the full script name and extension.'
				)]
				[string]$FullScriptName
			)
		#endregion Function Parameters
		## Function Usage:
		##	New-ScheduledJobTask -SettingsFile "$([System.IO.Path]::GetFullPath($SettingsFile))" -ScheduleAsUser "Domain\User" -ScheduleAsUserPassword 'DecryptedPassword' -ScheduledTaskName "Monthly Licensing Reports - $($ScriptNameNoExt)" -WorkingDir "$([System.IO.Path]::GetFullPath($ScriptDirectory))" -FullScriptName "$([System.IO.Path]::GetFullPath($MyInvocation.MyCommand.Definition))"

		begin
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tStart Function: $($MyInvocation.InvocationName)"

			if ($PSCmdlet.ParameterSetName -eq 'SettingsFile')
			{
				$Parameters = "-SettingsFile '$($SettingsFile)'"
			}
			else
			{
				$Parameters = "-PartialReportAndLogName $($PtlRptName)"
				if ($SendEmail)
				{
					if (!$SmtpHost)
					{
						$SmtpHost = 'smtp.ecollege.net'
					}

					if (!$AddrFrom)
					{
						$AddrFrom = "ServerCountForLicensing@example.com"
					}

					if (!$AddrTo)
					{
						$AddrTo = 'winadmins@examplecmg.com'
					}

					$Parameters += ' -SendMail:$true'
					$Parameters += " -FromAddr \$([char]34)$AddrFrom\$([char]34)"
					$AddressTo = [string]::Join(', ', $AddrTo)
					$Parameters += " -ToAddr \$([char]34)$AddressTo\$([char]34)"
					$Parameters += " -SmtpSvr \$([char]34)$SmtpHost\$([char]34)"
				}
			}
		}
		process
		{
			## These are passed in from Script Level Variables:
			## $WorkingDir		= $ScriptDirectory
			## $FullScriptName	= $FullScrPath
			$SchedTask = "powershell.exe -c \$([char]34)pushd $($WorkingDir); . $($FullScriptName) $($Parameters)\$([char]34)"

			if ($ScrDebug)
			{
				Write-Host -ForegroundColor Yellow "Parameters: $($Parameters)"
				Write-Host -ForegroundColor Yellow "Attempting to schedule task as: $([char]34)$($ScheduleAsUser)$([char]34)"
				Write-Host -ForegroundColor Yellow "Task to schedule: $([char]34)$($SchedTask)$([char]34)"
				Write-Host -ForegroundColor Yellow 'Task will be scheduled to run on the LAST Monday of every month at 02:30 (2:30 AM).'
			}
			else
			{
				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tParameters: $($Parameters)"
				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tAttempting to schedule task as: $([char]34)$($ScheduleAsUser)$([char]34)"
				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tTask to schedule: $([char]34)$($SchedTask)$([char]34)"
				Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t`tTask will be scheduled to run on the LAST Monday of every month at 02:30 (2:30 AM)."
			}

			if ($ScheduleAsUserPassword)
			{
				SCHTASKS /Create /RU $ScheduleAsUser /RP $ScheduleAsUserPassword /SC MONTHLY /MO LAST /D MON /ST 02:30 /TN $ScheduledTaskName /TR $SchedTask
			}
			else
			{
				SCHTASKS /Create /RU $ScheduleAsUser /RP /SC MONTHLY /MO LAST /D MON /ST 02:30 /TN $ScheduledTaskName /TR $SchedTask
			}
		}
		end
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t End Function: $($MyInvocation.InvocationName)"
		}
	}

	Function Remove-ExternalComObject
	{
		<#
			.SYNOPSIS
				Releases all <__ComObject> objects in the caller scope.

			.DESCRIPTION
				Releases all <__ComObject> objects in the caller scope, except for those that are Read-Only or Constant.

			.EXAMPLE
				Remove-ExternalComObject -Verbose
				Releases <__ComObject> objects in the caller scope and displays the released COM objects' variable names.

			.INPUTS
				None

			.OUTPUTS
				None

			.LINK
				http://gallery.technet.microsoft.com/scriptcenter/d16d0c29-78a0-4d8d-9014-d66d57f51f63

			.NOTES
				The Original Name I gave to the function was 'fn_Remove-ComObject'.
		#>
		[CmdletBinding(
			SupportsShouldProcess = $true
		)]
		#region Function Paramters
			Param()
		#endregion Function Paramters

		begin
		{
			Start-Sleep -Milliseconds 500
			[Management.Automation.ScopedItemOptions]$ScopedOpt = 'ReadOnly, Constant'
		}
		process
		{
			Get-Variable -Scope 1 | Where-Object {
				$_.Value.PSTypeNames -contains 'System.__ComObject' -and -not ($ScopedOpt -band $_.Options)
			} | ForEach-Object {
				$_ | Remove-Variable -Scope 1 -Verbose:([Bool]$PSBoundParameters['Verbose'].IsPresent)
			}
		}
		end
		{
			[System.GC]::Collect()
		}
	}
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Script Processing Start Time: $($StartTime)"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"

	if ($ScheduleAs)
	{
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Setting up a Scheduled Task."

		$objSchTsk = @{}

		## Initially sets $NBName to the NETBIOS Domain name of the system the
		## script is being executed on, so it can choose the correct password
		## for the user to schedule the task as if the $ScheduleAs variable
		## does not contain '\'.
		[string]$NBName = ${Env:USERDOMAIN}

		if ($ScheduleAs -like "*\*")
		{
			$SchedUserInfo = $ScheduleAs.Split('\')
			[string]$NBName = $SchedUserInfo[0]
			[string]$SecUsrNam = $SchedUserInfo[1]
			[string]$EncPasswd = ($DomainUserInfo | Where-Object { $_.NbDomain -ieq $NBName } | Select-Object EncPassword).EncPassword
			[string]$ScheduleTaskAccount = $ScheduleAs
		}
		elseif ($ScheduleAs -notlike "*\*")
		{
			[string]$NBName = ${Env:USERDOMAIN}

			if ($DomainUserInfo.UserName.Contains($ScheduleAs) | Where-Object { $DomainUserInfo.NbDomain -ieq $NBName })
			{
				[string]$SecUsrNam = ($DomainUserInfo | Where-Object { ($_.UserName -ieq $ScheduleAs) -and ($_.NbDomain -ieq $NBName) } | Select-Object UserName).UserName
				[string]$EncPasswd = ($DomainUserInfo | Where-Object { ($_.UserName -ieq $ScheduleAs) -and ($_.NbDomain -ieq $NBName) } | Select-Object EncPassword).EncPassword
				[string]$ScheduleTaskAccount = "$($NBName)\$($SecUsrNam)"
			}
		}
		else
		{
			[string]$SecUsrNam = ($DomainUserInfo | Where-Object { $_.NbDomain -ieq $NBName } | Select-Object UserName).UserName
			[string]$EncPasswd = ($DomainUserInfo | Where-Object { $_.NbDomain -ieq $NBName } | Select-Object EncPassword).EncPassword
			[string]$ScheduleTaskAccount = "$($NBName)\$($SecUsrNam)"
		}

		if ($EncPasswd.Length -gt 255)
		{
			$SecPasswd = $EncPasswd | ConvertTo-SecureString -Key $Script:Key #(1..32)
			$DecPasswd = ConvertTo-ClearText -SecureString $SecPasswd
		}
		else
		{
			$SecPasswd = Read-Host -Prompt "Enter the password for the $([char]34)$($ScheduleTaskAccount)$([char]34) account" -AsSecureString
			$DecPasswd = ConvertTo-ClearText -SecureString $SecPasswd
		}

		$objSchTsk.Add('SettingsFile', "$([System.IO.Path]::GetFullPath($SettingsFile))")
		$objSchTsk.Add('ScheduleAsUser', $ScheduleTaskAccount)
		$objSchTsk.Add('ScheduleAsUserPassword', $DecPasswd)
		$objSchTsk.Add('ScheduledTaskName', $SchTskName)
		$objSchTsk.Add('WorkingDir', "$([System.IO.Path]::GetFullPath($ScriptDirectory))")
		$objSchTsk.Add('FullScriptName', "$([System.IO.Path]::GetFullPath($FullScrPath))")

		New-ScheduledJobTask @objSchTsk

		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Scheduled Task Setup Complete."
		Read-Host -Prompt 'Press Enter to Continue script execution...'
	}

	$RawSvrList = @()
	$SvrList = @()

	if ($CreateCsv)
	{
		$CsvSvrList = @()
	}

	if ($CreateXlsx)
	{
		$XlsxSvrList = @()
	}

	if ($ExportCsvForBMC)
	{
		$BMCSvrList = @()
	}

	## The following array returns with the following columns in it:
	## 'SrvrName', 'SrvrDnsName', 'SrvrOsName', 'SrvrOsSpLvl', 'SrvrChangedOn', 'SrvrCreatedOn', 
	## 'SrvrDnsDomain', 'SrvrNBDomain', 'SecUsrNam', 'SecPasswd', 'ConnectAcct'
	$AllADSvrsList = Get-ListOfWindowsServersFromActiveDirectory

	#region Kill any existing jobs in the multi-thread queue
		if ((Get-Job).Count -ge 1)
		{
			if ($ScrDebug)
			{
				Write-Host -ForegroundColor Magenta 'Killing existing jobs . . .'
			}
			Write-Debug -Message 'Killing existing jobs . . .'

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tKilling existing jobs . . ."

			Get-Job | Remove-Job -Force

			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tDone killing existing jobs."

			if ($ScrDebug)
			{
				Write-Host -ForegroundColor Magenta 'Done killing existing jobs.'
			}
			Write-Debug -Message 'Done killing existing jobs.'
		}
	#endregion Kill any existing jobs in the multi-thread queue

	## Total Servers in the array returned from Active Directory
	[int]$TotalNumServers = $AllADSvrsList.Count

	## Used to count the Total number of Servers
	[int]$ServerCount = 0

	## Used to count the number of Virtual Servers
	[int]$VmSvrCount = 0

	## Used to count the number of Servers that have errors
	[int]$EESvrCount = 0

	## Used to count the number of Servers that are not responding
	[int]$NRSvrCount = 0

	## Used to count the number of Physical Servers
	[int]$PhySvrCount = 0

	## Used to count the number of Newly Built Servers
	[int]$NewSvrCount = 0

	## Used to count the Total number of Threads/Jobs for the progress bar
	[int]$JobCnt = 0

	foreach ($ADSvr in $AllADSvrsList)
	{
		$Arguments = ($ADSvr, $ScriptDirectory, $RptSvrFqdn, $LogEntryDateFormat, $SubnetToDataCenter, $IncludeMSSQL, $IncludeHotFixDetails, $IncludeCustomAppInfo, $IncludeInfraDetails, $ScrDebug)

		## Check to see how many open threads exist and if there are too many open threads then wait here until some close.
		while ((Get-Job -State Running).Count -ge $MaxThreads)
		{
			#Write-Progress -Activity 'Querying Servers in the List' -Status 'Waiting for running threads to finish and close' -CurrentOperation "$($JobCnt) threads created - $($(Get-Job -State Running).Count) threads open" -PercentComplete ($JobCnt / $TotalNumServers * 100)
			Write-Progress -Activity 'Querying Servers in Server List Array' -Status 'Waiting for running threads to finish and close' -CurrentOperation "$($JobCnt) threads created -> $((Get-Job -State Running).Count) threads running -> $($TotalNumServers - $JobCnt) threads remaining" -PercentComplete ($JobCnt / $TotalNumServers * 100)

			Start-Sleep -Milliseconds $SleepTimer
		}

		## Starting job
		$JobCnt++

		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Starting Multi-Thread Job with the Server Name as the Job Name: $([char]34)$($ADSvr.SrvrDnsName)$([char]34)"

		## Call Report-ADServerCountForLicensing-MultiThread.ps1 script to do individual server queries
		Start-Job -Name "$($ADSvr.SrvrDnsName)" -FilePath $MultiThreadScript -ArgumentList $Arguments | Out-Null

		## Show Progress Bar 
		Write-Progress -Activity 'Querying Servers in Server List Array' -Status 'Starting Threads' -CurrentOperation "$($JobCnt) threads created -> $((Get-Job -State Running).Count) threads running -> $($TotalNumServers - $JobCnt) threads remaining" -PercentComplete ($JobCnt / $TotalNumServers * 100)
	}

	## Get-Job | Wait-Job
	while ((Get-Job -State Running).Count -gt 0)
	{
		$ThreadsStillRunning = ''

		foreach ($JobThread in (Get-Job -State Running))
		{
			$ThreadsStillRunning += ", $($JobThread.Name)"
		}

		$ThreadsStillRunning = $ThreadsStillRunning.SubString(2)
		Write-Progress -Activity 'Creating Server List' -Status "$((Get-Job -State Running).Count) thread(s) remaining" -CurrentOperation "$($ThreadsStillRunning)" -PercentComplete ((Get-Job -State Completed).Count / (Get-Job).Count * 100)

		Start-Sleep -Milliseconds $SleepTimer
	}

	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Collecting Results from ALL Jobs..."

	## Get the results from all jobs. This will be an Array with the following columns in it:
	##	'SrvName', 'SrvDnsName', 'SrvOsName', 'SrvOsSpLvl', 'SrvChangedOn', 'SrvCreatedOn', 'SrvNBDomain', 'SrvIPv4Addr', 'SrvAddlIPv4Addr', 'SrvDataCenter', 'SrvBMCLocationCode', 'SrvSerNum', 'SrvBiosVer', 'SrvMnfctr', 'SrvModel', 'SrvInstOsArch', 'SrvLastBootUpTime', 'SrvUpTime'
	##	if ($IncludeHotFixDetails)
	##		'SrvLastPatchInstallDate', 'SrvUpTimeKB2553549HotFix1Applied', 'SrvUpTimeKB2688338HotFix2Applied', 'SrvUpTimeTotalHotFixApplied', 'SrvKB3042553HotFixApplied'
	##	'MSClusterName', 'SrvRegOsName', 'SrvRegSpLvl'
	##	if ($IncludeMSSQL)
	##		'SQLSvrName', 'SQLSvrIsInstalled', 'SQLProductCaptions', 'SQLEditions', 'SQLVersions', 'SQLSvcDisplayNames', 'SQLInstanceNames', 'SQLIsClustered', 'isSQLClusterNode', 'SQLClusterNames', 'SQLClusterNodes'
	##	if ($IncludeCustomAppInfo)
	##		'SrvBrowserHawkInstall', 'SrvSoftArtisansFileUpInstall'
	##	'SrvPendingReboot', 'SrvPhysMemoryGB', 'SrvProcMnftr', 'SrvSocketCount', 'SrvProcCoreCount', 'SrvLogProcsCount', 'SrvHyperT', 'SrvBMCAgentInstalled', 'SrvBMCAgentStartMode', 'SrvBMCAgentState'
	##	'SrvNetlogonStartMode', 'SrvNetlogonState', 'SrvRemoteRegistryStartMode', 'SrvRemoteRegistryState', 'SrvWinFirewallStartMode', 'SrvWinFirewallState', 'SrvWinUpdateStartMode', 'SrvWinUpdateState'
	##	if ($IncludeInfraDetails)
	##		'SrvProcName', 'SrvDiskTotalInfo', 'SrvDiskUsedInfo', 'SrvDiskFreeInfo', 'SrvDiskTotalAllocated', 'SrvDiskUsedAllocated', 'SrvDiskFreeAllocated'

	if ($ScrDebug)
	{
		$AllJobs = Get-Job
	
		foreach ($SvrJob in $AllJobs)
		{
			Write-Host -ForegroundColor Green "Receiving Job: $($SvrJob.Name)"
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`tReceiving Job: $($SvrJob.Name)"
			$RawSvrList += Get-Job -Name $SvrJob.Name | Receive-Job
			Write-Host -ForegroundColor Cyan "Job Received: $($SvrJob.Name)"
		}
	}
	else
	{
		$RawSvrList = Get-Job | Receive-Job
	}

	$SvrList = $RawSvrList | Where-Object { $_.SrvNBDomain.Length -gt 2 } | Sort-Object -Property 'SrvCreatedOn' -Descending
	if ($CreateXlsx)
	{
		$XlsxSvrList = $RawSvrList | Where-Object { $_.SrvNBDomain.Length -gt 2 } | Sort-Object -Property 'SrvNBDomain', 'SrvDnsName'
		#$XlsxSvrList = $RawSvrList | Where-Object { $_.SrvNBDomain.Length -gt 2 } | Sort-Object -Property 'SrvNBDomain', 'SrvDnsDomain', 'SrvDnsName'
	}

	if ($ScrDebug)
	{
		[string]$ExportSvrListCsvFile	= "$($FullRptDirPath)\$([System.IO.Path]::ChangeExtension($LogFileName, 'csv'))"
		$ExportSvrList = $RawSvrList | Where-Object { $_.SrvNBDomain.Length -gt 2 } | Sort-Object -Property 'SrvNBDomain', 'SrvDnsName'
		#$ExportSvrList = $RawSvrList | Where-Object { $_.SrvNBDomain.Length -gt 2 } | Sort-Object -Property 'SrvNBDomain', 'SrvDnsDomain', 'SrvDnsName'
		$ExportSvrList | Export-Csv -Path $ExportSvrListCsvFile -NoTypeInformation
		Remove-Variable -Name ExportSvr*
	}

	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*"

	$ProgressActivity = 'Collecting Results from ALL Jobs'
	$ProgressStatus = 'Results Collected from ALL Jobs'
	Write-Progress -Activity $ProgressActivity -Status $ProgressStatus -PercentComplete 100 -Completed

	$ReportTitleHeading = "Windows Server Inventory for Licensing - $($DataCenterCity) Data Centers"
	## For the HTML output to render correctly, the ' - ' is replaced with '<br />' so that it will appear as follows:
	##					"Windows Servers from Active Directory<br />$($DataCenterCity) Data Centers"

	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Begin Building HTML Report"

	#region Build HTML Report
		#region Begin HTML Code
			## StringBuilder construct for the HTML report.
			$HtmlReport = New-Object System.Text.StringBuilder ''
			##	$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')		$StringBuilderHtmlReportOutput = $HtmlReport.Append('')
			##	$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("")		$StringBuilderHtmlReportOutput = $HtmlReport.Append("")

			## Standard head for a normal web based HTML document. ##
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<html xmlns="http://www.w3.org/1999/xhtml">')
			## Generic head for an HTML document. Used for testing purposes. ##
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<html>')

			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<head>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<meta http-equiv="X-UA-Compatible" content="IE=edge" />')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("	<title>Pearson - $($ReportTitleHeading)</title>")
			#region CSS Styles
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<style type="text/css">')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<!--')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			body {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-family: Tahoma;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-size: 11px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				margin-top: 0px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				margin-bottom: 0px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			table {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				width: 895px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				border: 2px solid #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			tr.ColumnLabels {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #008b5e;') # Green from developer.pearson.com
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			tr.AltRowOn {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #DDDDDD;')	# Light Gray Background
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			tr.AltRowOff {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #FFFFFF;')	# White Background
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			tr.LegendRow {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #FFFFFF;')	# White Background
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-size: 9px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			th {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-top: 0px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-bottom: 0px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-right: 5px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-left: 5px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				white-space: nowrap;')	# Does not wrap the text in the table cell.
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				border: 0px hidden #008b5e;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			th.ThSpacers {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #008b5e;') # Green from developer.pearson.com
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				height: 2px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-top: 0px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-bottom: 0px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-right: 5px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				padding-left: 5px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				white-space: nowrap;')	# Does not wrap the text in the table cell.
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				vertical-align: middle;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.ReportHeaderLeft {')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #3366CC;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				vertical-align: middle;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: center;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				width: 160px;')	# 200px;'
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.ReportHeaderRight {')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #3366CC;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				vertical-align: middle;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: center;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				width: 160px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.ReportTitle {')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #3366CC;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				vertical-align: middle;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.HdrReportDate {')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #3366CC;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				vertical-align: bottom;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.HdrRowSpacers {')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #3366CC;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				height: 2px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.RowSpacers {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #313431;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				height: 2px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.SvrHeading {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #008b5e;') # Green from developer.pearson.com
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: center;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-size: 14px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-weight: bold;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.RedCell {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #FF0000;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.RedText {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FF0000;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.YellowCell {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #FFFF00;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.DarkYellowText {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #cc6a00;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.GreenCell {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #00FF00;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.GreenText {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #006600;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.BlueCell {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #0000FF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.BlueText {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #0000FF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.CenterCell {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: center;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.TextAlignRight {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: right;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.IpAddrCell {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				white-space: normal;')	# Allows the text in the table cell to wrap to the next line.
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.CellBostonDC {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #00AA00;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.CellDenverDC {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #FF9900;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.CellCentennialDC {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #003399;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.CellSriLankaDC {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #993300;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.RptFooter {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #FFFFFF;') # White
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-size: 9px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #949494;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			td.RptFooterSpacers {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				background-color: #949494;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				height: 1px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			div.ReportTitle {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: center;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-size: 22px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-weight: bold;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			div.ReportDate {')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				color: #FFFFFF;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				text-align: center;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('				font-size: 9px;')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			}')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		-->')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</style>')
			#endregion CSS Styles
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('</head>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<body>')
			#region Header Table
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<table border="0" cellpadding="0" cellspacing="0" align="center">')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td rowspan="2" class="ReportHeaderLeft">')
				## Pearson Logo from 2015
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			<img width="114" height="18" alt="Pearson Logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHIAAAASCAYAAACHKYonAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAABJlJREFUeNrEWc112kAQlnkuQCXIFUSuwKICxDUXUAWGCgwVCCoALrkaKgBXAB2YXHNSDrnk4uy8N+uMP+/PSJCXfW+ejbQ7Ozs/38ysbn7/+jFLkuQpcY+zoZOhzebbbuuaMPo6eDN/Dol/5IaGZn1oDvEZ0x/728zvJ5Fh1hTmzzPLaPdKDTX8jGgX25t5ZawH4pnxY1q3M7Q2PBoFDzrDgHmkggfJsTQ8zpH1e/FzauafAnMnvBeNXS8iGx2oJGXRJoZSz7wiQGmiGwO5zuxVKtY0zB/3ss/osCR3rTDAq6GxMKI9F609sqG9DmWI1q9YXynwIDleY3KwI9qzrAL6TvDcPYgm8oA+UwXv7KFCY21oLmjLPJqIIlNWABo2OBwea/ddw54Ts8csEIk16IB4LASP1HcGdoK9cIAG9HAGOY4BA53AqLUyCA636OEAQ2uz6Yo9lcbY/J4HIGKjgTHHKD3PqjZMzN4zoeA5KJgcY+bZJxWppG9hVPBYuqCVoX0lHpHjVjB3xjBYCwPRmqHiSKTvnS+tydFTMJs64PbaY+DwyJS9vdNgZ5uDh/sg6r0mkEbg//tsINdYQTRVLoObZwvQY6lMHQlDbHaxITVJ/pIBsLphWFLDqyKHtplToPPQ+T3RiPl0GtIVG1Mi2aPyDCkXdJcZ0oHn1zZsCdC0A89NL+D9IPOIZ84WzkQRUCv2HUAka1LKEpwmC3QLC4kmsUJJA601CHwKzaWigmlvKeYrFpoIDjkfSMWOO0Z6ydWiS4kIwZg+JlyplpEKM+YkoWImlKZID1OYPwnJ0wspgo0gFVkpKq0n0Y9ZCvVuhYBVGSVoaI3xMnaeI8DROlQwmHdrLj4aUPIzF3u+1syO78q83bYQHDrQIlMZknuiN1ZEIeDU19Q3YNSDg7SwmjiMSrCSKw9uHSOX7ZSRu1Iomfa/d8g7DhizSz3QtmCrNPny1pFY0TAb7c0GJ/w2XvcI3nZmGVIH/J6UPLfCQbIW66zi+tAuWGO+cORK3ViH+aLcIm9rfHIws/dCpAmbL3+GDJlzDrxpsVfW0Ttzxy2Kt59y5LFQu1QIh9hzpLVR3sLI10B78QAVtTRk0bLNagu1c0CaCTpo7wqI0bWvfIRiIQTJqbbvcsBRrrga8+XNU+Ccuw49bwm3YFpZGj5T44vuaxjy0raDKrS+i0CRozZwBIryVnxcE4yVVSbuIfvCp1AO5CvCDKKsjWOdQqj0XwzJiksdRc4n+aGnbBP9U1C0r+Ib8bsP7xn6S08Euqp4Wuv8sMA5V35hmse+hARQYusrdoor2qjm3PJJBigUZHS9RAqXGnLlTAtHRpaKc6Ss+O6h/RkL3lTUHBx5r3EpkHIc77EScEdfObaiJRlBJK7lnXCHoqji99m/jMg8cX/Kyjy9YxPp784heI21JVxMeG9IxH1sA0UXGnHoiyBHD5qyU9h+OoNIDLVCheARy5cf6pQew4898KHDldhBQWcQ1j5fKq+13vk4jBfrV+XntAP2pRwddwzFW3YcqxNaex+rLtkZ75K/n/4akI90e6eIRE3vbR10Kuf+EWAA32NJlyzDu3YAAAAASUVORK5CYII=" />')
				## New Pearson Brand Logo from 2016
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			<img width="144" height="60" alt="Pearson Logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAUAAAACECAYAAADhnvK8AAAaG0lEQVR42u2dB5gUVbbHW1ddn3ndfaY1rctT3kMFna7qNIPzjOu6PCO6hvdkRQETKCB5ZnqAFUEyksMw3T1DFJQkCAKSg0hSQBBQEIGRnNNw956qGqiprthd3dM8/7/vu1/PdHflrn+dc8+553o8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQIRjV3jyojX569884eirvIV5+8BTEB0mveZHCzz5sdc8+cWPe/KK7/GER1+NkwYAOIdFj4tYfuxRLnr5XOhG8LaMt+1c8E7xV6bTyvlnZfz1a97GesKRjvz/JzwFI27AyQQAZD5tSq/loteMi9dMLl5HDITOWZMFc5GnIFLoCZdWx0kGAGQW+dGHuEhFeTvqiuiZC+LnXGDrecKzLsCJBwBUofDFnuQu7oKUi55+28qFN0/qXwQAgLQRjtzJBWhyFQmftu2UAigAAJByCqLtM0T4tG0RF+b7cYEAAO7TLvInJULLMrrlR1rjYgEA3CM/UictAQ732nhPeOAluHAAgOQIx945h4RP3TYhoRoAkDgFsQ7nqPgpLbbZ07r497iQAACnll/+uS1+FXmDkS3cErwIFxQAYNPyizRKTYBCClIwT56m0Xv5KRXCpbioAABr8qOPuSY8bYYzz7sDmefNj5inYW/meYO/vt2PeZr0Z553BsiN/qb33lB9h95vNUwWxrBrIjgMFxcAYEyb6PVSUYKE3c2YLHokYo36ME/rYeyOnuPZU6WzWKvPl7G+i9exkas3s8nrf2Kff/+z1Cau28pKV21mfRatY22mf82eG/Ulq/nRBHYhre+tvnw9XBRbDJXXHU46ReYRXGQAgIHrG/02IdGj1mwQ87zWi4vVEFaneDobxAVtwy/7WKLsPHCEjf/mB1Z/7Dx2KW2nQS/ZmqzYXmIieBAXGQCg5/rmOSxKwDyFMdldbdiLXduxlLWcvISt3bmXuc3uw8dY9y9XsVs+GMVFtifzvDeYedqXJCqC3XCxAQAqyy/idSQiJD7kltbvwW7pNJL1m7+GHTt5iqWact66zl7FLmhdJLvGhQm6xTSqBQAAJMLRdY7cXS4+57ccytp+tpSdPH2apZvvdh9gtbqPkwRYiiI7dYnDsQguOgCASlq9bFv82hVLolO770S2gYtQVfPSiNlyvyPtl1MRRII0AIC7v7stxaKwRE5Lqd+d1Rs9x7ZAnTp1ii3/aRcbtXIT68Rd17cnLWZPcdG6b+g0VrP/JFaTC+m9/PWx6BesycTFrO+CNWzJjzsdiWDDTxYyzyvdE0mZaYqLD8Cv2/VtYC1+3LLi7i65vV3mrLYUpL1Hj7OiJd+xurGZ7PpOI5mn+WDmebOvnBbTSMnza0y5gBV5gAPkzykHkL7TdCC7rcsY1uSTBWxNmb0o8jNkCVKUuDDmJIizKpFTlp2d/bvsgLCTN8bbSd5+5m2dQTugfE/b9vG2KtsvjMgOCu1ygt7HcwXhOvwgAUir9RfdaSuvj7u93ReuNRWhA8dPsjZTlrIr84plMXu9jyx+5J46cU1peS6Cnld7SmkvzSYusiWCN3YeLSdcOxHBBIIhjz5a7bc5Ae97XMAm8va9gcBVtK1c4D6v1ALCfIPlTvA2PTsg1sUPE4BM6fvjQtRs2jJT8ek59xv2e/oupaiQ6LmRtExC1kqONN/xwSj2rUWf4+Jtu7ll2V8WbNvbiTRK9jQGg/fcwoXtKz0BDPmFp8yXFYRQ0Ns7blm+vtpBoTZ+pACkzv1dYWn9cYuK+ugM3V1u9T04cIrcB0dJyoUxN4euyftA6+Tu7UVcDDftP2zuCn88Tw6K2A+IRN04ldnZYk1dCzAovGFneW5Rioo7rV2+MX6oALjv+lazFIe28jjeFTv26IrN5A3b2DUUeCDBqRAr66FoiRU9oLxD7lbf0GkUO3Gq3FAAl2/fI/cr2nW7qVKMC/h8viu4YJ1KVAAlEfULD+q70d7n8YMFwFX3N9LaUhy4kNzaZ4Ku0PRevE4OaFARA7MRGSSKlDBNgQ/qE6RgR0VARLIYHYzmoO/W68aeGjPX1Aqs1udTucCCbSsz+bmGc7Oy/sDF6ngyAiiLoHemngjm5ta6Cj9aANyzAOdYCkOzQey6zqPjBOb9WStlq++9IcYCVuEGcxf6Ci6QdUfMZh2/XM2GLNvABn21nrWdsZxV6zGOi+xAZ0GLNkWS6K7YvsdQAN+atEQamuegQEKd5PsBg9coQYykBDDkF57Td6XFMH60ALhBi6GX8xv/pC1xaDqQ1ft4Hlu/az/7tmwf+/vIL+V0E8oJNBMusvze6seeic1kew4f0xWqk+Wn2Q3vj5SLKDgJjNTvwV6dsNBQAItXbJSjweRu2xPANzNFAHNy7vl3XVc6IKzFDxcANwgXP2g7ACGVpOon9weS9UXuKwmLldXWfDC7loubFY2nfSULqpP+wMb9WXXu5hox98edsqja7geMhjNFACU3OCBs0xHAclEUMXIFgOQFMNrGcRCCig+QANqN8HKhfDA6w1IAeyxcI/cNOgmMcNf70o4j2IFjJ/THCXNrVRbsYXYjzR9lmACu1XODg0GhluWyUjRafDInKLyUExAfyfH5/sutn00oFLo8WxRvrx3w3hMIZFWnY3Zr3VlZWRcafUbbqx0U7w+FhBq2+2QF4bqcgPBATtD7d34engiFRG+NGjUqTYsQ8nkfsnNOjfaJUpwoQEXryc4Wbkv2HND+mZ0H+Rrcc0Mw6L2TzgUeiIn3/41I+Twcb/dj9xVNsyeAFCBxIoAth7KL2pewHQeP6K5zy/5DsvVKOYT2Jk4qyiwB9K7RE0BKldEXPeG2bL/QgX/nqEEy9k+hgNglFPLe4XRfpPScoLezgVUqueZ8230pgm21LjmBXFigjJxZxttyGi0TCgib+WsZ307HSscVFBrz9zdqtjePxNdwf4Pex6URNvr7upvv5yASPSXtiGm3aX6NxQAXvahR0rt0HEExnOv332p+fcUn6XiV416ujBjaxI/3F3q/diDrLu12+fs9eVvN22nVNo/m+MXJSJx3QsvRV/Kbflk6BDA4dKqlABZQQIUsQCfr5u711R+MYkdP6Jff2rz3oOymtxpmd52DM8wC1B1dUtvn+494gRLeV23rF9pets+XRTe5HFDxjq98k4rd7R2PUIt/f4Zq2Y2hoPhPvvwLkuUTFJrx95Zq9nG6mVVC1g0JsXIj64nIPFnQpaGG80xG1xzSRsWVKPwU5fMyEk9KMidLjaxA/t5UfdES61tae/y88+9+rFpuAp2DCkuM//+XUFDoV/lhJXyotTjVDxW+/Fij4yOLlb4XCASuNvtepeYXMN2DPeuv+G5+w+9KuQA27s/8gz+z7gOcskQeNudk3U3M+wCXb98tj11uOzxtc4W4JYC5ubkXKJaA9ke+V+0e1a1b9zfcgvhC9fkMQ7fJLzbSrGuiqZtFLl2lm8uba7zuOItoE3/7PBsi30rnGHtK65QtI/r/My7gk6wSzOm88P+3VCxjuM2g0E5nPQ3MLWDxEep/VQnbfYZCKQg38e+sVwnZBi5ifzRzeVXHqurqCF7Dl/u3Sr8Dv3dmTlCIKUJ+2CBf9E0InKUARv7Kb/jTKRfAJgPYvQMmWwpgo4mLnAkgubYNe7FnTSrSfLJ2i1xsIS9i1wXunykCmJ2ddbOBpfKJ6mvnq908uolsuLLTNOvU7fcMBO6tpvkejU75i8mqz+Mu2EGNNdLL5vnSWD7C+7z1UP5vJe23X2xhYPX0VInpaPk98aBlH5p22KGJAOYExccq7V9QeM36AVbrKrJQVce0y6zIBRfJPprjOhoKiTn8dXHFg03r8pOVzY91jM45Ofbw3XdfCpEzDYBE6qVlLt53BrA7+060FMB64+fLSdFOyvA36MWKl280XGf72ascRpYjnTJFACl4oe8WCQ+obviPK9+Y3set11v5ZpaaLys7ToD9wiCd7R/Pza1xmUmf5XjtMlYd9LVF8U8GFXLonLWzsBTJ2nlF+lwUb1e9N8lmF8MyKwGUxndXvo62Rwwp3QPqfV3jQADpHPxkZc0qx1FmdF6AEfnRlukRwIHcTZ1gKYAvjJ1rvw+QrD9uWV7cvoQdNSm/Hxg2TZ5Nzr4AtsgYC1BvJAh/z0TIyuxGRdWunNLm6NxUc3QFx+fLMt5nKQDjqAiE0q+ml/A9V/09smj4+hdqLMyFFZYefzDkqz6bYutc+Hw3Wll1UlCi8jUosHsNJfc1PiD1kYFl3t/Awl1mGZWX+1O1yw2EyJlbgL3TIoDvDmS39/7EUgDrcldWigLbsfzIpf1HNzZ8hbH19/OBI3L/H6Xt2J+s/cVMEECq/qLzg95D6Scq62JL5c9F26X9+fd3aNfv8/mu1UQoIzr7cPrBrKwrjQVQbBpnsfq97yYggAcoEVxXKPxCPQos8OP/h2Z/1a7gMbs1FSmgo3a1NRZcAzvWsqnFLUVntX1799xiVwC118Wg/1Vn1JB3PETOtA8wGk2LADYdxG7tMc5SAB8rmSknV9ux/up1Y3VGzDJdX/MZy+U6gk4q0uRFc1IlgBSAsHfDeEPavrRQQPhafdPEBSecFltQddAbWWqKS7lfI7JPWqy3ic5N3CoBAZyVQMR8imYd31MupM1l5+QEvc/oCOAWbT+o07HY/KHQVqcft7sdAaS+UHvbiC+eoekrBjoCOCwtAthsELv+w7Gs3GLCpIciM8zd1UJlHpL/+5D5+00yXdeOQ0fZeVR4ocUQhyW3Sv+QKgGkFAyjZZS8uAeU6J72h9zFjlVhlhOnc8Ov03E5W+oeC3+f+uLs5A4q+XpuCODsBARwlL4LyS1Zf5bP6fpCIZ9fZ33fOl0Pv+5P66xnmz0B1M/3jN9XKVii3cY4iFwmCGDzweyaLmOk8b5m5AydJleUMRI/KnDaoBerG/2ClVtYk/cXT5cLshY6mi94pRun1dAFDog/cJFbUtHkwqlSovMenfyvDfyp3lWbCEtQykRctFV2eVZShWlKMDZqFZ9bRVMThVuRz9oRVhsCOMfxtgNCc4scucWU7EwjSewFMMRw/ENM+NTpftUOCEG9/dE+sPQEMBj0PQwBTFkfYCySLgE0S1auwDtgsixy2vHHVNPvle7spo6l0oRKVvReup67vj3sF0A4GwAZlFoB1G3HlUjffGkUBfU5mQQZ5PVLicksybadtx9VbVuyBVeVm3BCVQkgJQvbP37xB3rAmFm1JHY6Lulw5/uVVd2gQviz1n2A4t+S6DOGAJpbgJEeaRHA94awy/45QpogyQxKlZHL2BfJ9QEpfeWNvqxa59Gs4/Sv2dFy6zmHZ28pk61IKs/ldGrM/GjDVApgjl943Y31002hOzpEEG5KYrXnO/z+eSTUiss71VRs0iSASjT0BccPA7/QQT8qy630eEEa4Pz3oEmjOfN7EN+2dIGD4mMQwNS5wE3TIoAthrCLO5SyXw4fNRQukjYpsvtSF2kej+rdx7H6Y+ayid/+aHtKzAnrf5InUCKLsX2J8/1sF70jlQLoVl6WrpuZhkKpNDJF7s/yFuvknS2XxDDo7VaVAigLifhEXOqKRaORNNqHgN44YhrmloBl+kfNuN2KoFUzCGCVCmDshbQIYMuh7EKTggUEBUgiKzaysbxtTmCS9UFLvpPr/iUqfi6VwzcVwATGAjvoVGduVCExEL6LpUDI2aTcM6IXCnjfovG6Z0XD+3xVC+AZC1VKnNYvJqFrkQWFmCag85UbFiAVRNB3gStnBcAFTr8APsxv/vKUC6BSiOCHvYeY25QdOcZeLJ0lV30mtzcR8XOpCELaBDAg3KfvyjmPdNp0t7XVX1aFQr7/NojENskQAVT3mQrcwmutFHQ4bSaCwWDWn1XHMkUvouxWHyB3vetYW4DWI3sggImSN7wGv/F3pFwAWw+T0lfWJ2DZGUGTIfWe9w27QpqIqac8f3BhLPF9zI8+ds4IYE7W9QZjhF9wVfz83lyd7Sw3F8zME8BK+5eddbMypni9VREB1VhkdfvM+QNLKbcVZ7Fn3QwBrErCoy/zFEQWp14Ai5inbRFb+8u+pIVv7qbtrPnkxex6Kp9PQRKq9lxYkuw+HvCEZ11wrgigIjTr3eifMkKqqhKXBC2ctBqVUJUCqIxDbmX7+3HjdKXWySLYtNX5g0SooyO0a3SEEgJYBYGQ1I8GoVJU3A02m7woPipympUdOsqWbP2FFS3bwBqOX8Bu6TJGnjipQvikCddjbsw3HHHzlKZDALlLV6jzg9/p1vqVYqLaFJAFNoS5CgVQXtahi1+38r56O2seAsfigk0+342OtqGTHJ7j97axI4Ahv/d/IICpJB0FEUgAmw9hi7buMtS7Ddw9fmjIVBYaOIXV7DeR3dR1LLuQxK35YHl8MFWJIdEjV9cN0as8tjh4rgmgXAZJN7fsxcQEVewS8osvn/n/7PhYdfvY+mb3/m8VJkJ/qbXibArn2YowfrGppr/1Q6fFHXT6bIdr16E3ntrAAnwaAphSAYw8knIBpOFrzQaz+VvKDAVw6oZtzPNyV7kaDEVy+fel4IndyYwSbxvcPqWKOB2P76PzNnRzO1RVRK9CslmxAl3hOCN2Zyddj6uVp6rSbGFRDdDJd2uRFgswKM49s00H85+ozyONw9Z0BVyis28lDq/Tds2DsLFBX6FeFLguBDCVtCr5HbeATqVcAJsOZNM3bjfO4Vu3VU6CzouwtKTmnG3Pun1KpQmDdKaztFNy3QlSegrNbRH/w19lVrOvstXkfUcZc9pfc9Pq1d47Rts0cfXeqKgY40I1mERc4HlOCsOqlquoqbjbo1PBmtxQzb6dUFflsbD+HtAsO8EkWKJnAT4DAUx9P+DSlIoMua3cqpv83U+GAjhy9WZn8/e60iK7U3E6aaypwfCrUte3Fci6y6BM/BZ+4z5qtBxZiTR3hFHx0Np+/38a5Mp9atDR31VJ7RikroKs9Kt1MzsGJVFYu/+fJyOAShtl/RCpcdnZfj7jRHUS8UTGBPPj+E7Vh7rEohuie6KJ0DTTXbKW6q+XcPSD1ApgRBqeRuXpjWhHlZsb9kq39dfU7VNJuXjK2FqDEQfet9y3OKUxuDsNqkcvoKFe1DdIfVfkhvMbfYidOUEMSq1LRUip8IBUsVpOkC5TROEldTBCXaiVhJomK6KpHOP333uHG5O/q11gdeKyUWl8JZ1ombLvI22IWTOjiLF13593Eo2kseg+GJNon66OpZmyVKL/hxZgiZhysWnQi3WY+42u+K3btZ9dRoESp6Wrkk59CZ+f7Kmj4WdyNV4pBeNnm6MOymVrxTukQjSSFl5p9jTdAqZmw7+aW6z2PMOJiCqL7C6qeGJmyajaVMV9v4Dc71BQGGwyhecmqU/R7y2wUxSUf3dRRTVozfjkcio0Ic2PzC0qEhWyjlSC3tfhA049Q95efvztydqmc0CvytShFd0A29TBpfjrJtZU8g3nGF4nvxDVFk44I3xB7+P8HBaZnO+l1J9LEzpB6Mzd4B2pLohwVWGMzdm844zwrdm5l3X9cjW7hCzEJv2TGcWRQOJzpLUr/XBySfW9yo9tv5yfR2WpjJrwjWIxnVRuElfdFOoLUqZlXKcz6uG4NLwrKIb1LDFDC1OeSW5e5RJc0t/zjSxacnuVWc6OSOeFu4J0I5Jrreq//EYJFm2T90vVZMtsj+Ke7tMrC6bjAtMcGuvP7jeNmRZL9UuHyVN32nUxdcT2SWUCpnKDLojv6AFDJfFNLTd5WON+pS2TSv6rm+RCiwdpvhCD89xZOb7v45aVpyrdK59/MQyRM3WDU1wen8SNUlq4lXdbz/Hs5u7j5ARpSm9JbghbImkvh9yw/jId6lvL8WfdS1YLzVtLLmiy1q4UsBDF2+1Gmikizq23K9JxvDTiwmj+Xelc0IgMX1Y2ibDVjHFOoOFzNGqGrCxpG6qx0eCcEcDS6ikXHhqqRtYeVX2u101OeSHXt7AkvX1/+dGnccEBAFo3+LOUW4Ekflz07h8+nYWGTpX7/dLr/s7AhQYA6FiB0WBKrb+3+7HqH45lq3acHRI3/4ed7LoOpXIOYFoswcg1uNAAACMr8OtUFUW9nLu/B3Tm8P3hwGF2IU1fqZTNSmHgoz4uMADAxAqMZKdEfBr2Zq9PWWqYB/jcuPlyQCRlAhjDFIEAAFuu8Gz38wB7s67z1xgKYHjWSnkO39QI4EpcVACATQGM3ei6CL3ehz09Zq7JNJYWcwIn3g57Wo6+EhcVAGAfmiHNTSFqM1zKAxy7enOc+PVftFaOBOeloPJLuKQWLiYAwDkF0cnuRYFLmKfFUC50A1jusGmsYOZKVsDd3vuLPpcjwBQAKXS1zl85F9QALiIAIDHe7v1bV4fISYnQxfK8vRTwaNRb/psqwLgrfoel8c0AAJAU7Yb/mQvKiTRXaUmmbZf2GQAAXCEvdpdsVWW4+IUjM5HoDABwH2mscGR3xopffrQAFwkAkEIRjNzsCUdXZJj4bXRzTl8AADCnINY5A4SvXLL6fgVlrQAAGWcNFj/IRWh11fT1RWP8tRouAgCgasmPPs/FaFkaRG8ff+0n9UUCAEBmCWGkDhepSVykTrsrfLEFnvzYa/z1CpxkAEBm03bETVysXvEUxIq4gH0r9dXZF7wTyjLFUumqguLbcUIBAOcmdUf/xpM3vIYnv/hxLmpvcVHswFsv/vcwqYUjXbnVGOaC+YanIPJXT7voHdIyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAODXwr8A592ojMy1aa4AAAAASUVORK5CYII=" />')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="ReportTitle">')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			<div class=$([char]34)ReportTitle$([char]34)>$(($ReportTitleHeading).Replace(' - ', '<br />'))</div>")
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			<div class=$([char]34)ReportDate$([char]34)>Generated on: $([char]34)$((Get-Date -Format 'MMMM dd, yyyy HH:mm').ToString())$([char]34)</div>")
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td rowspan="2" class="ReportHeaderRight">')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('			<img width="137" height="8" alt="Pearson Tag Line Logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAIkAAAAICAQAAABnwRQ/AAADcUlEQVR4AY2VbWjV5R+HL7d5qKF/x0wylLkoCWy29aQvfKoIqaaoyx6IkhNWOFPnRNOEwMAQgkrURCJ7IAu1bBgxi1I3FS17oXOZYi+aOh9iKic1+TvduYLzu/nxGzvKruvN9z7n5tyHm8/NhwgbXS+xZWYcIMEfXRymtBnvCXOjSyW43m0SXOBfDhNxmh2OEvFBW8x60snSw+9dIglrzQSrJfaZbv8Ix+Z2HPNdUxJc4d54zwAzlvU4odp2KyS42xfC1MeZtpr1qi0uQMR77bLToRIsV0skOMN2i0Rs8YqrRRzhtXj/YK/aZWX88994wGIrvewsybnLDyz1MZ+XHja5TBLO9w/LcxZL7C9mu13dI+odPup550pwpbo9XFGJWt7jhKlqu8PC6qDpMK3xnC9a6iCfclUBAPVsYT/15GcTKaYA4xjOq6T5H1DLVtqJmMN+GlhChMwANvIdn7OOiL4Uk2EHG+kNnbTlvEKAMVSylDpuIckZdnKEMmJoZChfUciNaWcfPzGIJKN4nWo2cIEOGpmHeLtXHeMkL1uaNyX4jjvFzX5kgX9aZz8vOkFyFnvep33I696VeHgdNlkkwfGe91dHS69S0uHKnEMk+K0fm/KUs7qlZLlb/NcHEin5zDLbXXeTlLSZ8md/s38iJW/aKrEUAHP5Pw8znD7MIT/rGMdEalhNlrXM5iWO00zEy/TnTsZyjoUEOMFhmrhOgF2M4Ah7mEZv6CKTs4uIu5nKNWZzjIUUEcNAxtLAAZKc4Ame5e2bZrCGPjSQIkAqPukgYgHFzKKZKirZzjyKycdJtrKZ3bQCnzCUFawmooD5NFNBFc2kGUx+UvxNmvdYRG84y7KcZ4mo5yi3UsVxSplODLWM5nHq6M7vVLOIudyYizxJGRviB9bMfVQBMI7xUESaS0yjCyjiKK+wiogySoBOTgOwhprwzT98wXN8ScQUbuN+LgOwl/ksIR+f8iGHGcgl8lFCOQCn6QQgFdYXuAgMJM1kdgDwFovZhARo4zW+ZjutJNnHdBq4GR1MZA9DiGhiI9uo4wdgOOB+50hwpocsjEo4uDtRloVhqvD9xKfLJTjJNvvlrdY3PGnWFivylnAmONLuJVwrYp37JFjiGSfEJUzO9e6ybyjhtYmWzF/ChyQ40lNxCRda7xGzdnrMVf8BpQQyjhXd9T4AAAAASUVORK5CYII=" />')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="HdrRowSpacers"></td>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="3" class="ReportTitle">')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			<div class=$([char]34)ReportTitle$([char]34)>$(($ReportTitleHeading).Replace(' - ', '<br />'))</div>")
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			<div class=$([char]34)ReportDate$([char]34)>Generated on: $([char]34)$((Get-Date -Format 'MMMM dd, yyyy HH:mm').ToString())$([char]34)</div>")
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="3" class="HdrRowSpacers"></td>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
				$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('</table>')
			#endregion Header Table
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('<table border="0" cellpadding="0" cellspacing="0" align="center">')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th colspan="6" class="ThSpacers"></th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr class="ColumnLabels">')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th>DNS Host Name</th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th>Operating System</th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th>SP Lvl</th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th>AD Svr Add Date</th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th>AD Svr Chg Date</th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th>IP Address</th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<th colspan="6" class="ThSpacers"></th>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
		#endregion Begin HTML Code

		$AlternateRow = 0

		#region Row HTML Server Level foreach Loop
			foreach ($Svr in $SvrList)
			{
				## Increment the global server count for use in the generated report.
				$ServerCount++

				$SvrDnsName = $Svr.SrvDnsName

				$AdOsName = $Svr.SrvOsName
				$RegOsName = $Svr.SrvRegOsName
				if (($RegOsName -eq 'No Response') -or ($RegOsName -eq 'Error Encountered'))
				{
					$SvrOsName = $AdOsName
				}
				elseif ($RegOsName.Length -lt 6) ## The word 'Server' has a length of 6 characters and all of the OS name variables should contain it.
				{
					$SvrOsName = $AdOsName
				}
				else
				{
					$SvrOsName = $RegOsName
				}

				$AdSpLvl = $Svr.SrvOsSpLvl
				$RegSpLvl = $Svr.SrvRegSpLvl
				if (($RegSpLvl -eq 'No Response') -or ($RegSpLvl -eq 'Error Encountered'))
				{
					$SvrSpLvl = $AdSpLvl
				}
				elseif ($RegSpLvl.Length -lt 3) ## The value should be 'SP*' or 'NONE'. The shortest length is 3 characters.
				{
					$SvrSpLvl = $AdSpLvl
				}
				else
				{
					$SvrSpLvl = $RegSpLvl
				}

				$SvrIPAddr = $Svr.SrvIPv4Addr
				$SvrCreatedOn = $Svr.SrvCreatedOn
				$SvrChangedOn = $Svr.SrvChangedOn
				$SvrNBDomain = $Svr.SrvNBDomain
				$SvrDataCenter = $Svr.SrvDataCenter

				if ($SvrSpLvl.Length -lt 3)
				{
					$SvrSpLvl = 'NONE'
				}

				#region TR.ForEach.Server
					if ($AlternateRow)
					{
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr class="AltRowOn">')
						$AlternateRow = 0
					}
					else
					{
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr class="AltRowOff">')
						$AlternateRow = 1
					}

					#region TD.SvrDnsName
						if ($SvrDataCenter -like "*Boston*")
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="CellBostonDC">')
						}
						elseif (($SvrDataCenter -like "*Arapahoe*") -or ($SvrDataCenter -like "Cornell*"))
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="CellDenverDC">')
						}
						elseif ($SvrDataCenter -like "*Centennial*")
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="CellCentennialDC">')
						}
						elseif ($SvrDataCenter -like "*Sri Lanka*")
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="CellSriLankaDC">')
						}
						else
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td>')
						}
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			$($SvrDnsName)")
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
					#endregion TD.SvrDnsName

					#region TD.SvrOsName
						if ($SvrOsName -like "*2012 R2 *")
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText">')
						}
						elseif ($SvrOsName -like "*2012 St*")
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText">')
						}
						elseif ($SvrOsName -like "*2008 R2 St*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="BlueText">')
						}
						elseif ($SvrOsName -like "*2008 R2 En*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="BlueText">')
						}
						elseif ($SvrOsName -like "*Web*2008 R2*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="BlueText">')
						}
						elseif ($SvrOsName -like "*2008 St*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="BlueText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="DarkYellowText">')
						}
						elseif ($SvrOsName -like "*2008 En*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="BlueText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="DarkYellowText">')
						}
						elseif ($SvrOsName -like "*Web*2008*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="BlueText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="DarkYellowText">')
						}
						elseif ($SvrOsName -like "*2003*")
						{
							#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="DarkYellowText">')
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedText">')
						}
						elseif ($SvrOsName -like "*2000*")
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedText">')
						}
						else
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td>')
						}
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			$($SvrOsName)")
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
					#endregion TD.SvrOsName

					#region TD.SvrOsSpLv
						if ($SvrSpLvl -eq 'NONE')
						{
							if ($SvrOsName -like "*2012*")
							{
								$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td>')
							}
							else
							{
								$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedText">')
							}
						}
						else
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td>')
						}
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			$($SvrSpLvl)")
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
					#endregion TD.SvrOsSpLv

					#region TD.SvrCreatedOn
						if ($SvrCreatedOn -gt $OldestCreateDate)
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenCell">')
							$NewSvrCount++
							#$NewServer = $true
						}
						else
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td>')
							#$NewServer = $false
						}
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			$($SvrCreatedOn)")
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
					#endregion TD.SvrCreatedOn

					#region TD.SvrChangedOn
						## If $SvrChangedOn is more than 30 days old then a red cell is needed, else a normal cell is needed.
						if ($SvrChangedOn -lt $OldestCreateDate)
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedCell">')
						}
						else
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td>')
						}
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			$($SvrChangedOn)")
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
					#endregion TD.SvrChangedOn

					#region TD.SvrIPAddr
						if ($SvrIPAddr -eq '&nbsp;')
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedCell">')
						}
						else
						{
							$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="IpAddrCell">')
						}
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("			$($SvrIPAddr)")
						$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		</td>')
					#endregion TD.SvrIPAddr

					$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
				#endregion TR.ForEach.Server

				if ($Svr.SrvModel -like "*Virtual*")
				{
					$VmSvrCount++
				}
				elseif ($Svr.SrvModel -eq 'Error Encountered')
				{
					$EESvrCount++
				}
				elseif ($Svr.SrvModel -eq 'No Response')
				{
					$NRSvrCount++
				}
				else
				{
					$PhySvrCount++
				}

				if ($CreateCsv)
				{
					if ($ScrDebug)
					{
						$objSvrInfo = New-Object PSObject #New-Object System.Object
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrNBDomain'	-Value $SvrNBDomain	#$Svr.SrvNBDomain
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrDnsName'	-Value $SvrDnsName

						if ($SvrIPAddr -eq '&nbsp;')
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrIPAddr'	-Value ''
						}
						else
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrIPAddr'	-Value $SvrIPAddr
						}

						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrOsName'			-Value $SvrOsName
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrSpLvl'			-Value $SvrSpLvl
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvInstOsArch'		-Value $Svr.SrvInstOsArch
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvPhysMemoryGB'	-Value $Svr.SrvPhysMemoryGB
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvSocketCount'	-Value $Svr.SrvSocketCount
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvProcCoreCount'	-Value $Svr.SrvProcCoreCount
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvLogProcsCount'	-Value $Svr.SrvLogProcsCount
						## Intel Hyperthreading -OR- AMD HyperTransport
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvHyperT'			-Value $Svr.SrvHyperT
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvProcMnftr'		-Value $Svr.SrvProcMnftr
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrCreatedOn'		-Value $SvrCreatedOn
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SvrChangedOn'		-Value $SvrChangedOn
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvMnfctr'			-Value $Svr.SrvMnfctr
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvModel'			-Value $Svr.SrvModel
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvSerNum'			-Value $Svr.SrvSerNum

						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDataCenter'		-Value $SvrDataCenter
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvBMCLocationCode'	-Value $Svr.SrvBMCLocationCode
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvLastBootUpTime'	-Value $Svr.SrvLastBootUpTime
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvUpTime'			-Value $Svr.SrvUpTime
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvBMCAgentInstalled'	-Value $Svr.SrvBMCAgentInstalled
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'MSClusterName'		-Value $Svr.MSClusterName

						if ($IncludeHotFixDetails)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvLastPatchInstallDate' -Value $Svr.SrvLastPatchInstallDate
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvUpTimeKB2553549HotFix1Applied' -Value $Svr.SrvUpTimeKB2553549HotFix1Applied
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvUpTimeKB2688338HotFix2Applied' -Value $Svr.SrvUpTimeKB2688338HotFix2Applied
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvUpTimeTotalHotFixApplied' -Value $Svr.SrvUpTimeTotalHotFixApplied
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvKB3042553HotFixApplied' -Value $Svr.SrvKB3042553HotFixApplied
						}

						if ($IncludeMSSQL)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLSvrName'		-Value $Svr.SQLSvrName
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLSvrIsInstalled'	-Value $Svr.SQLSvrIsInstalled
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLProductCaptions'-Value $Svr.SQLProductCaptions
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLEditions'		-Value $Svr.SQLEditions
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLVersions'		-Value $Svr.SQLVersions
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLSvcDisplayNames'-Value $Svr.SQLSvcDisplayNames
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLInstanceNames'	-Value $Svr.SQLInstanceNames
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLIsClustered'	-Value $Svr.SQLIsClustered
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'isSQLClusterNode'	-Value $Svr.isSQLClusterNode
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLClusterNames'	-Value $Svr.SQLClusterNames
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQLClusterNodes'	-Value $Svr.SQLClusterNodes
						}

						if ($IncludeCustomAppInfo)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvBrowserHawkInstall'	-Value $Svr.SrvBrowserHawkInstall
						}

						if ($IncludeInfraDetails)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvProcName'			-Value $Svr.SrvProcName
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDiskTotalInfo'		-Value $Svr.SrvDiskTotalInfo
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDiskUsedInfo'		-Value $Svr.SrvDiskUsedInfo
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDiskFreeInfo'		-Value $Svr.SrvDiskFreeInfo
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDiskTotalAllocated'	-Value $Svr.SrvDiskTotalAllocated
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDiskUsedAllocated'	-Value $Svr.SrvDiskUsedAllocated
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SrvDiskFreeAllocated'	-Value $Svr.SrvDiskFreeAllocated
						}
					}
					else
					{
						$objSvrInfo = New-Object PSObject #New-Object System.Object
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Server NB Domain'	-Value $SvrNBDomain	#$Svr.SrvNBDomain
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Server DNS Name'	-Value $SvrDnsName

						if ($SvrIPAddr -eq '&nbsp;')
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'IP Address'	-Value ''
						}
						else
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'IP Address'	-Value $SvrIPAddr
						}

						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'OS Name'					-Value $SvrOsName
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SP Lvl'					-Value $SvrSpLvl
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Installed OS Arch'			-Value $Svr.SrvInstOsArch
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Physical Memory (GB)'		-Value $Svr.SrvPhysMemoryGB
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Socket Count'				-Value $Svr.SrvSocketCount
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Processor Core Count'		-Value $Svr.SrvProcCoreCount
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Logical Processor Count'	-Value $Svr.SrvLogProcsCount
						## Intel Hyperthreading -OR- AMD HyperTransport
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'HT'						-Value $Svr.SrvHyperT
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Processor Manufacturer'	-Value $Svr.SrvProcMnftr
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Created On'				-Value $SvrCreatedOn
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Changed On'				-Value $SvrChangedOn
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Manufacturer'				-Value $Svr.SrvMnfctr
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Model'						-Value $Svr.SrvModel
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Serial Num'				-Value $Svr.SrvSerNum

						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Data Center Location'		-Value $SvrDataCenter
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'BMC Location Code'			-Value $Svr.SrvBMCLocationCode
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Last Boot Up Time'			-Value $Svr.SrvLastBootUpTime
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Up Time'					-Value $Svr.SrvUpTime
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'BMC Agent Installed'		-Value $Svr.SrvBMCAgentInstalled
						$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'MS Cluster Name'			-Value $Svr.MSClusterName

						if ($IncludeHotFixDetails)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Last Patch Install Date' -Value $Svr.SrvLastPatchInstallDate
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Up Time HotFix 1 Applied' -Value $Svr.SrvUpTimeKB2553549HotFix1Applied
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Up Time HotFix 2 Applied' -Value $Svr.SrvUpTimeKB2688338HotFix2Applied
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Up Time Total HotFix Applied' -Value $Svr.SrvUpTimeTotalHotFixApplied
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'MS15-034 KB3042553 HotFix Applied' -Value $Svr.SrvKB3042553HotFixApplied
						}

						if ($IncludeMSSQL)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Name'				-Value $Svr.SQLSvrName
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Installed'			-Value $Svr.SQLSvrIsInstalled
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Product Captions'	-Value $Svr.SQLProductCaptions
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Editions'			-Value $Svr.SQLEditions
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Versions'			-Value $Svr.SQLVersions
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Svc DisplayNames'	-Value $Svr.SQLSvcDisplayNames
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Instance Names'	-Value $Svr.SQLInstanceNames
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Is Clustered'		-Value $Svr.SQLIsClustered
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr is Cluster Node'	-Value $Svr.isSQLClusterNode
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Cluster Names'		-Value $Svr.SQLClusterNames
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'SQL Svr Cluster Nodes'		-Value $Svr.SQLClusterNodes
						}

						if ($IncludeCustomAppInfo)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'BrowserHawk Install'		-Value $Svr.SrvBrowserHawkInstall
						}

						if ($IncludeInfraDetails)
						{
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Processor Family'			-Value $Svr.SrvProcName
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Disk_Drives'				-Value $Svr.SrvDiskTotalInfo
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Disks_Space_in-used'		-Value $Svr.SrvDiskUsedInfo
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Drives_free_Space'			-Value $Svr.SrvDiskFreeInfo
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Total_Disk_Allocated_GB'	-Value $Svr.SrvDiskTotalAllocated
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Total_Disks_in-used_GB'	-Value $Svr.SrvDiskUsedAllocated
							$objSvrInfo | Add-Member -MemberType NoteProperty -Name 'Total_Disks_free_Space_GB'	-Value $Svr.SrvDiskFreeAllocated
						}
					}

					$CsvSvrList += $objSvrInfo
				}

				#if ($ExportCsvForBMC -and ($Svr.SrvModel -ne 'No Response'))
				if ($ExportCsvForBMC -and ($NewServer -or ($Svr.SrvBMCAgentInstalled -eq 'NO')))
				{
					if ($SvrIPAddr -ne '&nbsp;')
					{
						##	BCM Server Property				Description
						##	-------------------------------:-----------
						##	NAME							Server Name
						##	CLM_DOMAIN						NETBIOS Domain Name
						##	CLM_OS							Chart below
						##	CLM_LOCATION					Chart below
						##	CLM_ENV							Chart below
						##	CLM_BU							Line of Biz (Should be HED for everything.)
						##	CLM_MANAGED						HED or BUSINESS_UNIT_MANAGED (Should be HED for everything.)
						##	CLM_HIPAA_COMPLIANCE			'TRUE' or 'FALSE'
						##	CLM_PCI_COMPLIANCE				'TRUE' or 'FALSE'
						##	CLM_SOX_COMPLIANCE				'TRUE' or 'FALSE'
						##	CLM_CIS_COMPLIANCE				'TRUE' or 'FALSE'
						##	CLM_DISA_COMPLIANCE				'TRUE' or 'FALSE'
						##
						##	CLM_PROGRAM						Project/Program/Biz Service Name
						##	CLM_TASK						Chart below
						##	LEGACY_PATCH_WINDOW				Chart below
						##	CLM_DISASTER_RECOVERY			BRONZE or None
						##
						##	PRIMARY_TECHNICAL_CONTACT		SN Queue or Person
						##	SECONDARY_TECHNICAL_CONTACT		If no queue above
						##	CONTACT_APP_PRIMARY				SN Queue or Person
						##	CONTACT_APP_SECONDARY			If no queue above
						##	CLM_OWNER						Person
						##	DBA1							SN Queue or Person
						##	DBA2							If no queue above
						##	Cost Center						Used if/when we ever do chargeback
						##	CLM_MAINTENANCE_WINDOW			Chart below (NOT USED)

						$SvrDomainEnvironment = ($InScopeDNSDomains | Where-Object { $_.NetBiosName -eq $SvrNBDomain } | Select-Object Environment).Environment

						if ($BMCEnvironments[$SvrDomainEnvironment])
						{
							$SvrEnvironment = $BMCEnvironments[$SvrDomainEnvironment]
						}
						else
						{
							$SvrEnvironment = 'CHANGE TO CORRECT ENVIRONMENT'
						}

						if ($SvrDomainEnvironment)
						{
							#if (($LegacyPatchWindows | Where-Object { $_.Environment -eq $SvrDomainEnvironment }).Count -ge 2)
							#{
							#	## There are two or more Patch Code options available for the environment. Rotating between the first two.
							#	if ($AlternateRow)
							#	{
							#		$SvrLegacyPatchWindowCode = ($LegacyPatchWindows | Where-Object { ($_.Environment -eq $SvrDomainEnvironment) -and ($_.DataCenterCity -eq $DataCenterCity) -and ($_.Priority -eq 102) } | Select-Object PatchCode).PatchCode
							#	}
							#	else
							#	{
							#		$SvrLegacyPatchWindowCode = ($LegacyPatchWindows | Where-Object { ($_.Environment -eq $SvrDomainEnvironment) -and ($_.DataCenterCity -eq $DataCenterCity) -and ($_.Priority -eq 101) } | Select-Object PatchCode).PatchCode
							#	}

							#	## 9 is one less than the length of the shortest code used by the GTO domains for the "LEGACY_PATCH_WINDOW" code in BMC.
							#	if ($SvrLegacyPatchWindowCode.Length -lt 9)
							#	{
							#		$SvrLegacyPatchWindowCode = ($LegacyPatchWindows | Where-Object { ($_.Environment -eq $SvrDomainEnvironment) -and ($_.DataCenterCity -eq $DataCenterCity) } | Select-Object PatchCode).PatchCode
							#	}
							#}
							#else
							#{
								## There is only one Patch Code option available for the environment.
								$SvrLegacyPatchWindowCode = ($LegacyPatchWindows | Where-Object { ($_.Environment -eq $SvrDomainEnvironment) -and ($_.DataCenterCity -eq $DataCenterCity) } | Select-Object PatchCode).PatchCode
							#}
						}
						else
						{
							$SvrLegacyPatchWindowCode = 'Legacy_Patch_Window_Code'
						}

						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvrDnsName               ->: $([char]34)$($SvrDnsName)$([char]34)"
						Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t SvrLegacyPatchWindowCode ->: $([char]34)$($SvrLegacyPatchWindowCode)$([char]34)"

						#$SvrIPAddresses = $SvrIPAddr.Split(', ')

						$SvrBMCInfo = New-Object PSObject #New-Object System.Object
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'NAME'					-Value $SvrDnsName
						#$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'IP'					-Value $SvrIPAddresses[0]
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_DOMAIN'			-Value $SvrNBDomain
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_OS'				-Value 'W'
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_LOCATION'			-Value $Svr.SrvBMCLocationCode	# This is the code used to select the correct repeater in BMC.
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_ENV'				-Value $SvrEnvironment			#'CHANGE TO CORRECT ENVIRONMENT'
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'Legacy_Patch_Window'	-Value $SvrLegacyPatchWindowCode		#'Legacy_Patch_Window_Code'
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_BU'				-Value 'HED'
						$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_MANAGED'			-Value 'HED'
						#$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_HIPAA_COMPLIANCE'	-Value 'FALSE'
						#$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_PCI_COMPLIANCE'	-Value 'FALSE'
						#$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_SOX_COMPLIANCE'	-Value 'FALSE'
						#$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_CIS_COMPLIANCE'	-Value 'FALSE'
						#$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'CLM_DISA_COMPLIANCE'	-Value 'FALSE'

						if ($Svr.SrvModel -like "*Virtual*")
						{
							$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'IS_VIRTUAL' -Value 'TRUE'
						}
						else
						{
							$SvrBMCInfo | Add-Member -MemberType NoteProperty -Name 'IS_VIRTUAL' -Value 'FALSE'
						}

						$BMCSvrList += $SvrBMCInfo
					}
				}

				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="CenterCell"></td>')
				#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')

				Clear-Variable -Name Svr* -Confirm:$false
			}
		#endregion Row HTML Server Level foreach Loop

		#region End HTML Code
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="TextAlignRight">Physical Svr Count: </td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("		<td class=$([char]34)CenterCell$([char]34)>$($PhySvrCount)</td>")
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="TextAlignRight">VM Svr Count: </td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("		<td class=$([char]34)CenterCell$([char]34)>$($VmSvrCount)</td>")
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="TextAlignRight">EE-NR Svr Count: </td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("		<td class=$([char]34)CenterCell$([char]34)>$($EESvrCount + $NRSvrCount)</td>")
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr class="AltRowOff">')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="CenterCell"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="TextAlignRight" colspan="4">Total Number of Servers found in ALL "In Scope" Active Directory domains:</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("		<td class=$([char]34)CenterCell$([char]34)>$($ServerCount)</td>")
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr class="LegendRow">')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="CenterCell">Legend/Key:</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td align="center"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine("		<td colspan=$([char]34)2$([char]34) class=$([char]34)GreenCell$([char]34) align=$([char]34)center$([char]34)>New Systems created after:<br />$([char]34)$($OldestCreateDate)$([char]34)</td>")
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="2" class="RedCell" align="center">Old Non-Responding Server<br />-OR-<br />NO Valid Data Returned</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="3" class="CellBostonDC" align="center">Servers in Boston Data Centers</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="3" class="CellDenverDC" align="center">Servers in Denver Data Centers</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="3" class="CellCentennialDC" align="center">Servers in the Centennial office Data Center</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="3" class="CellSriLankaDC" align="center">Servers in the Sri Lanka office Data Center</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr class="LegendRow">')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText" align="center">OS for new server builds<br />Windows Server 2008 R2<br />-OR-<br />A Newer Windows Server OS</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="GreenText" align="center">OS for new server builds<br />Windows Server 2012 R2<br />-OR-<br />A Newer Windows Server OS</td>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="2" class="BlueText" align="center">OS should not be used<br />without valid reasons<br />Windows Server 2008.</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="2" class="BlueText" align="center">OS should be used sparingly<br />when a newer OS cannot be used.<br />Windows Server 2008 R2</td>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="2" class="DarkYellowText" align="center">OS should not be used,<br />they went EOL in July 2014<br />Windows Server 2003<br />-OR-<br />Windows Server 2003 R2.</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="2" class="DarkYellowText" align="center">OS should not be used<br />without VERY valid reasons<br />Windows Server 2008.</td>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedText" align="center">OS should not be used,<br />the OS went EOL -OR- There is NO SERVICE PACK installed!<br />-OR-<br />Server should be retired.</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td class="RedText" align="center">OS should not be used as it went E.O.L. <br />-OR-<br /> There is NO SERVICE PACK installed!<br />-OR-<br />Server should be retired.</td>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	<tr>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('		<td colspan="6" class="RowSpacers"></td>')
			#$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('	</tr>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('</table>')
			$StringBuilderHtmlReportOutput = $HtmlReport.AppendLine('</body>')
			$StringBuilderHtmlReportOutput = $HtmlReport.Append('</html>')
		#endregion End HTML Code
	#endregion Build HTML Report

	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t End Building HTML Report"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!"

	## Write report out to the HTML file...
	#$HtmlReport | Out-File -FilePath $HtmlRptFile
	Set-Content -Path $HtmlRptFile -Value $HtmlReport.ToString()

	if ($CreateCsv)
	{
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"

		## Write data out to the CSV file...
		if ($ScrDebug)
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Sorting the Array by $([char]34)SvrNBDomain$([char]34) then $([char]34)SvrDnsName$([char]34) before exporting it to the CSV file..."

			$CsvSvrList = $CsvSvrList | Sort-Object -Property 'SvrNBDomain', 'SvrDnsName'
			$CsvSvrList | Export-Csv -Path $CsvRptFile -NoTypeInformation
		}
		else
		{
			Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Sorting the Array by $([char]34)Server NB Domain$([char]34) then $([char]34)Server DNS Name$([char]34) before exporting it to the CSV file..."

			$CsvSvrList = $CsvSvrList | Sort-Object -Property 'Server NB Domain', 'Server DNS Name'
			$CsvSvrList | Export-Csv -Path $CsvRptFile -NoTypeInformation
		}
	}

	if ($ExportCsvForBMC)
	{
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t----------------------------------------------------------------------------"
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Exporting to a CSV formatted file, for import into BMC-BSA-BladeLogic."

		## Write data out to the CSV file...
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Sorting the Array by $([char]34)CLM_DOMAIN$([char]34) then $([char]34)NAME$([char]34) before exporting it to the CSV file..."

		$BMCSvrList = $BMCSvrList | Sort-Object -Property 'CLM_DOMAIN', 'NAME'
		$BMCSvrList | Export-Csv -Path $BMCRptFile -NoTypeInformation
	}

	if ($CreateXlsx)
	{
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t----------------------------------------------------------------------------"
		Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t Exporting to an Excel formatted file."

		$objExportInventoryToExcel = @{}
		$objExportInventoryToExcel.Add('Inventory', $XlsxSvrList)
		$objExportInventoryToExcel.Add('Path', $XlsxRptFile)
		$objExportInventoryToExcel.Add('RptTitleHeading', $ReportTitleHeading)
		$objExportInventoryToExcel.Add('SplaAdDnsDomains', $SplaDnsDomains)
		$objExportInventoryToExcel.Add('SvrCountPhy', $PhySvrCount)
		$objExportInventoryToExcel.Add('SvrCountVm', $VmSvrCount)
		$objExportInventoryToExcel.Add('SvrCountEE', $EESvrCount)
		$objExportInventoryToExcel.Add('SvrCountNR', $NRSvrCount)
		$objExportInventoryToExcel.Add('SvrCountTotal', $ServerCount)
		$objExportInventoryToExcel.Add('SvrCountNew', $NewSvrCount)

		if ($XlsxCreationEngine -eq 'MSExcel')
		{
			Export-ReportToExcel @objExportInventoryToExcel
		}
		elseif ($XlsxCreationEngine -eq 'OpenXML')
		{
			## Open XML Export Function when ready...
		}
		else
		{
			## Other App Export Function if needed or when ready...
		}
	}

	if ($SendMail)
	{
		$objSendMail = @{}

		## Use a StringBuilder for better performace with larger text objects.
		$EmailMsgBody = New-Object System.Text.StringBuilder ''

		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Good Day,')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine("	Attached please find the report: $([char]34)$($ReportTitleHeading)$([char]34).")
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')

		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine("	This report includes any server that has been used on the network in the last $([char]34)$($DisableServersOlderThan.ToString())$([char]34) days. ")
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine("	If the Active Directory computer object is older than $([char]34)$($DisableServersOlderThan.ToString())$([char]34) days and has not had its machine password updated, ")
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	then the row for that computer object has been changed to indicate it with red in one or more of the columns in the HTML version of the report.')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')

		if ($IncludeMSSQL)
		{
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	The flag "IncludeMSSQL" has been selected.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Information about "Microsoft SQL Server" in the environment(s) is also included.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		}

		if ($IncludeHotFixDetails)
		{
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	The flag "IncludeHotFixDetails" has been selected.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Several columns that indicate whether certain "CRITICAL" or "HIGH" priority patches have been applied to a server if applicable to the Operating System have been included.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')

			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	A column has also been included with the last date a "Microsoft" update/patch was applied to each server. This does not mean that the server is current with all necessary updates/patches.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		}

		if ($IncludeCustomAppInfo)
		{
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	The flag "IncludeCustomAppInfo" has been selected.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Column(s) that indicate whether certain applications are installed on the server have been included.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		}

		if ($IncludeInfraDetails)
		{
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	The flag "IncludeInfraDetails" has been selected.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Some additional information requested by the platform architecture teams is also included. It is usually only included in the report when they request it.')
			$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		}

		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Some of the computers in this report may still be active but may not have been online at the time of report generation.')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	If you have any questions, please let us know.')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('<p>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Thank you,')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	Windows Systems Admins')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<br />')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('	<a href="mailto:winadmins@examplecmg.com">winadmins@examplecmg.com</a>')
		$StringBuilderEmailMsgBodyOutput = $EmailMsgBody.AppendLine('</p>')

		$objSendMail.Add('MsgBody', $EmailMsgBody.ToString())

		if ($PSCmdlet.ParameterSetName -eq 'SettingsFile')
		{
			$objSendMail.Add('AddrFrom', $FromAddr)
		}
		else
		{
			$objSendMail.Add('AddrFrom', "$($DataCenterCity).$($FromAddr)")
		}

		$objSendMail.Add('AddrTo', $ToAddr)

		if ($DataCenterCity -eq 'All')
		{
			$RptSrvIPAddr = @(
				([System.Net.Dns]::GetHostEntry(${Env:COMPUTERNAME})).AddressList | ForEach-Object {
					if ($_.AddressFamily -eq 'InterNetwork')
					{
						$_.IPAddressToString
					}
				}
			)

			$RptSrvIPAddr = [string]::Join(', ', $RptSrvIPAddr)

			foreach ($ObjSubnet in $SubnetToDataCenter)
			{
				#if ($RptSrvIPAddr -match $ObjSubnet.Subnet)	## Regex comparison - Returns more false positives. Returns the first one that similarly matches instead of the closest match.
				if ($RptSrvIPAddr -like $ObjSubnet.Subnet)		## Wildcard comparison - Returns more positive matches than false ones. Returns the closest match.
				{
					[string]$RptSrvDataCenterCity = $ObjSubnet.DataCenterCity
				}
			}

			$SmtpServer = ($SmtpSvr | Where-Object { $_.DataCenterCity -eq $RptSrvDataCenterCity } | Select-Object SmtpSvrName).SmtpSvrName
		}
		else
		{
			$SmtpServer = ($SmtpSvr | Where-Object { $_.DataCenterCity -eq $DataCenterCity } | Select-Object SmtpSvrName).SmtpSvrName
		}

		$objSendMail.Add('SmtpHost', $SmtpServer)

		$FileAttachments = @($HtmlRptFile)

		if ($CreateCsv)
		{
			$FileAttachments += $CsvRptFile
		}

		if ($CreateXlsx)
		{
			$FileAttachments += $XlsxRptFile
		}

		$objSendMail.Add('Attachments', $FileAttachments)

		Send-EmailWithReports @objSendMail
	}

	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t End Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	Add-Content -Path $FullPathMainScriptLogFile -Value "$((Get-Date -Format $LogEntryDateFormat).ToString()):`t!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!+!"

	## Release all lingering COM objects.
	Remove-ExternalComObject

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script

#region Help
<#
	.SYNOPSIS
		<Brief overview of what the script does.>

	.DESCRIPTION
		The Report-InfoServerLicensing script does the following:

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
		${Env:SystemDrive}\Path\To\Scripts\Directory\Report-InfoServerLicensing.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'

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
		Release: 2015-Jun-15

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
#		.\Report-InfoServerLicensing.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'
#		. ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Path\To\Proper\Directory\Report-InfoServerLicensing.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value' -ScrDebug
#		. ${Env:USERPROFILE}\Documents\WindowsPowerShell\Path\To\Proper\Directory\Report-InfoServerLicensing.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value'
#		. ${Env:SystemDrive}\Path\To\Scripts\Directory\Report-InfoServerLicensing.ps1 -Parameter1Name 'Parameter1Value' -Parameter2Name 'Parameter2Value' -Parameter3Name 'Parameter3Value' -ScrDebug
#
#	-- Script Change Log --
#	Changes for Jun-2015
#		- Initial Script/Module Writing and Debugging.
#
###################################################################################################################
#endregion Script Change Log
