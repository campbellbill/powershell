<#
	Task - SFTP Tranfer Script
	Script Name:	Task-sftpTranferScript.ps1
	Written by:		Bill Campbell
	Written:		July 23, 2013
	Added Script:	23 July 2013
	Last Modified:	2014.Mar.31

	Version:		2014.03.31.02
	Version Notes:	Version format is a date taking the following format: YYYY.MM.DD.RR
					where RR is the revision/save count for the day modified.
	Version Exmpl:	If this is the 6th revision/save on January 13 2012 then RR would be '06'
					and the version number will be formatted as follows: 2012.01.23.06

	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
	#! IMPORTANT NOTE:																							 !#
	#!					THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE				 !#
	#!					RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.			 !#
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#

	Example Found:	http://winscp.net/eng/docs/after_installation?ver=5.1.3&lang=en&utm_source=winscp&utm_medium=setup&utm_campaign=5.1.3&prevver=
					This script was created using the information that was pulled from the URL above as an example.

	Purpose:		Transfer files from SFTP server to a Secure Windows File Server on a scheduled basis.

	Notes:			Use the next line to change to your user profile 'Desktop' folder:
					Gets a list of all current Environment Variables:			Get-ChildItem Env:
					Change Directories to the Current Users 'Desktop' folder:	Set-Location ${Env:USERPROFILE}\Desktop
					Some Commonly used Environment Variables:					Write-Output ${Env:USERPROFILE}; Write-Output ${Env:COMPUTERNAME}; Write-Output ${Env:USERNAME};

	.SYNOPSIS
	.PARAMETER PtlLogName
    	Partial report name used to create the full report filename

	.PARAMETER SendMail
		Send Mail after completion. Set to $true to enable. If enabled, -FromAddr, -ToAddr, -MailServer are mandatory

	.PARAMETER FromAddr
		Email address to send from. Passed to Send-MailMessage as -From

	.PARAMETER ToAddr
		Email address to send to. Passed to Send-MailMessage as -To

	.PARAMETER MailServer
		SMTP Mail server to attempt to send through. Passed to Send-MailMessage as -SmtpServer

	.PARAMETER ScrDebug
		Turn on debugging code embedded in the script.

	.EXAMPLES
		.\Task-sftpTranferScript.ps1
		cd "${env:userprofile}\SkyDrive\Scripts\PowerShell\Task-sftpTranferScript"
		.\Task-sftpTranferScript.ps1 -ScrDebug
		.\Task-sftpTranferScript.ps1 -SendMail -FromAddr 'Task-sftpTranferScript@pearson.com' -ToAddr 'bill.campbell@pearson.com' -MailServer 'mail.ecollege.net' -ScrDebug
		. ${env:userprofile}\SkyDrive\Scripts\PowerShell\Task-sftpTranferScript\Task-sftpTranferScript.ps1 -ScrDebug
		. ${env:userprofile}\SkyDrive\Scripts\PowerShell\Task-sftpTranferScript\Task-sftpTranferScript.ps1 -SendMail -FromAddr 'Task-sftpTranferScript@pearson.com' -ToAddr 'bill.campbell@pearson.com' -MailServer 'mail.ecollege.net' -ScrDebug
		C:\Utilities\Scripts\Task-sftpTranferScript\Task-sftpTranferScript.ps1 -SendMail -FromAddr 'Task-sftpTranferScript@pearson.com' -ToAddr 'bill.campbell@pearson.com' -MailServer 'mail.ecollege.net' -ScrDebug

	#!#!#!#!#!#!#!#!#!#!#!#
	#!  Script Changes:  !#
	#!#!#!#!#!#!#!#!#!#!#!#
	Script Outline:
		1. Collect variables to connect to the SFTP Server.
		2. Connect to the SFTP Server.
		3. Transfer files from the SFTP Server to the Secure File Server.
		4. Transfer any files to go out, from the Secure File Server to the SFTP server.

	Changes for 23.July.2013 - 29.July.2013
		- Initial Script Development. (Adapted from script written for SquareTwo Financial with the same name.)

	Changes for 23.Aug.2013
		- Change the '$session.GetFiles()' parameter to 'Remove Files on Transfer' from '$false' to '$true'.

	Changes for 31.Mar.2014
		- Corrected some content in the HTML for the error message email.
#>

#region Script Initialization
	#region Script Parameters
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Partial report name used to create the full report file name and log file name'
			)][string]$PtlLogName
			, [Parameter(
				Position			= 1
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Send Email ($true/$false). Default is FALSE.'
			)][switch]$SendMail = $false
			, [Parameter(
				Position			= 2
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Email From Address (Originator)'
			)][string]$FromAddr
			, [Parameter(
				Position			= 3
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Email To Address (Recipient)'
			)][string]$ToAddr
			, [Parameter(
				Position			= 4
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'SMTP Server Address'
			)][string]$MailServer
			, [Parameter(
				Position			= 5
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'The ScrDebug switch turns the debugging code in the script on or off ($true/$false). Default is FALSE.'
			)][switch]$ScrDebug = $false
		)
	#endregion Script Parameters

#############################################################################################################
#	#	Load WinSCP .NET assembly																			#
#############################################################################################################
	[void][Reflection.Assembly]::LoadFrom('C:\Program Files (x86)\WinSCP\WinSCPnet.dll')
	[string]$Script:WinScpIni = 'C:\Program Files (x86)\WinSCP\WinSCP.ini'

#############################################################################################################
#	#	Setup Global SFTP Connection variables																#
#############################################################################################################
#	 SshHostKeyFingerprint is Mandatory for SFTP and SCP protocols.
#	 http://winscp.net/eng/docs/faq_script_hostkey
	#If ($ScrDebug)
	#{
	#	$Script:SftpHostName = 'sftp.mycmsc.com'
	#	$Script:SftpSshHostKeyFingerprint = 'ssh-rsa 1024 55:a0:60:e3:55:4d:b1:77:77:95:24:77:fb:71:e3:6d'
	#	[int]$Script:SftpPortNumber = 222
	#}
	#Else
	#{
		$Script:SftpHostName = 'sftp.mycmsc.com'
		$Script:SftpSshHostKeyFingerprint = 'ssh-rsa 1024 55:a0:60:e3:55:4d:b1:77:77:95:24:77:fb:71:e3:6d'
	#}
	$Script:SftpUserName = 'kct.pearson.eresource'
	$Script:SftpPassword = 'mzc9_8kMC6EU'

	# Resolve the FQDN of the server that this script is running on. It will be added to the end of the report so future users will know where to go to find it.
	$Script:RptSvrFqdn = ([System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName).ToLower()

#############################################################################################################
#	#	Global Script Variables																				#
#############################################################################################################
#	[string]$Script:FullScrPath = ($MyInvocation.MyCommand.Definition)
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path)
	If ($ScrDebug)
	{
		[string]$Script:LocalRootDirectory = '\\stor22\datadump\kctcs\API\DEV'
	}
	Else
	{
		[string]$Script:LocalRootDirectory = '\\stor22\datadump\kctcs\API'
	}
	[string]$Script:SftpRootDirectory = '.'

	If (!$FromAddr)
	{
		[string]$Script:FromAddr = 'Task-sftpTranferScript@pearson.com'
	}

	If (!$ToAddr)
	{
		If ($ScrDebug)
		{
			$Script:ToAddr = @('sysadmin@ecollege.com')
		}
		Else
		{
			#$Script:ToAddr = @('winadmins@pearsoncmg.com','sysadmin@ecollege.com')
			$Script:ToAddr = @('winadmins@pearsoncmg.com')
		}
	}

	If (!$MailServer)
	{
		[string]$Script:MailServer = 'mail.ecollege.net'
	}

	If ($SendMail)
	{
		If ((!$FromAddr) -or (!$ToAddr) -or (!$MailServer))
		{
			throw 'The following parameters are required when setting the "SendMail" parameter: "FromAddr", "ToAddr", "MailServer"'
		}
	}

#############################################################################################################
#	#	Script Variable Initialization																		#
#############################################################################################################
#	#	Date and Time Formatting																			#
#	#	http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo%28VS.85%29.aspx		#
#	#	Formatting chosen for the date string is as follows:												#
#	#	Date Display  |  -Format																			#
#	#	------------  |  ------------																		#
#	#	2012.Jan.25   |  yyyy.MMM.dd																		#
#	#	2012.01.25    |  yyyy.MM.dd																			#
#	#	Mon			  |  ddd																				#
#	#	------------  |  ------------																		#
#	#	Time Display  |  -Format																			#
#	#	------------  |  ------------																		#
#	#	22:00         |  HH:mm																				#
#############################################################################################################
	[string]$LogFileDate = (Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()	# Example Output: 2012.04.25.Wed.1400	- 1400 being the Military time designator for 2:00 PM.

	If (!$PtlLogName)
	{
		$PtlLogName = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	}

	[string]$BaseLogName	= "$($PtlLogName)_$($LogFileDate)"					# Example Output: Task-sftpTranferScript_2012.04.25.Wed.1400
	[string]$Ext1			= 'log'
	[string]$LogFileName	= "$($BaseLogName).$($Ext1)"						# Example Output: Task-sftpTranferScript_2012.04.25.Wed.1400.log

	If ($ScrDebug)
	{
		If (!(Test-Path -Path "$($ScriptDirectory)\Logs\Debug"))
		{
			New-Item -ItemType Directory "$($ScriptDirectory)\Logs\Debug" | Out-Null
		}

		$Script:FullScriptLogPath = "$($ScriptDirectory)\Logs\Debug\$($LogFileName)"
		$Script:FullDebugLogPath = "$($ScriptDirectory)\Logs\Debug\DLL_DebugLog_$($LogFileName)"
		$Script:FullSessionLogPath = "$($ScriptDirectory)\Logs\Debug\DLL_SessionLog_$($LogFileName)"
	}
	Else
	{
		If (!(Test-Path -Path "$($ScriptDirectory)\Logs"))
		{
			New-Item -ItemType Directory "$($ScriptDirectory)\Logs" | Out-Null
		}

		$Script:FullScriptLogPath = "$($ScriptDirectory)\Logs\$($LogFileName)"
	}
#endregion Script Initialization

#region User Functions
	Function fn_WriteToLog
	{
	    [CmdletBinding()]
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Entry to write to the log file.'
			)]$LogEntry
		)
		Add-Content -Path $FullScriptLogPath -Value $LogEntry
	}

	Function fn_GetYearWeekOfYear
	{
		## http://technet.microsoft.com/en-us/library/dd347647.aspx
		[string]$Year = (Get-Date -Format 'yyyy').ToString()
		[string]$WkOfYrNum = (Get-Date -UFormat '%V').ToString()

		If ($WkOfYrNum.length -lt 2)
		{
			$WkOfYrNum = "0$($WkOfYrNum)"
		}

		[string]$YrWkOfYr = "$($Year)\Week_$($WkOfYrNum)"

		return $YrWkOfYr
	}

	Function fn_SendEmail
	{
	    [CmdletBinding()]
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'HTML for the Email Message Body'
			)][string]$MsgBody
			, [Parameter(
				Position			= 1
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Email Address that the email will be sent from: "From".'
			)][string]$AddrFrom
			, [Parameter(
				Position			= 2
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Email Address that the email will be sent to: "To".'
			)]$AddrTo
			#)][array]$AddrTo
			, [Parameter(
				Position			= 3
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Outgoing "SMTP" Server Address'
			)][string]$SmtpHost
			, [Parameter(
				Position			= 4
				, Mandatory			= $false
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Attachments to be sent'
			)]$Attachments
		)

		#	USAGE:
		#		fn_SendEmail -MsgBody $MsgBody -AddrFrom $AddrFrom -AddrTo $AddrTo -SmtpHost $SmtpHost -Attachments $Attachments

		# Setup splatting hashtable to hold the values to be passed to the 'Send-MailMessage' cmdlet.
		$SendMailMsg = @{}

		$SendMailMsg.Add('SmtpServer', $SmtpHost)
		$SendMailMsg.Add('From', $AddrFrom)
		$SendMailMsg.Add('BodyAsHtml', $true)
		$SendMailMsg.Add('Body', $MsgBody)
		$SendMailMsg.Add('To', $AddrTo)

		[string]$RptDate = (Get-Date -Format 'ddd dd-MMM-yyyy HH:mm').ToString()	# Example Output: Fri 27-Apr-2012

		If ($ScrDebug)
		{
			[string]$MsgSubject	= "[DEV - ERROR] - An SFTP Sync Issue Occurred - $($RptDate)"
		}
		Else
		{
			[string]$MsgSubject	= "[ERROR] - An SFTP Sync Issue Occurred - $($RptDate)"
		}
		$SendMailMsg.Add('Subject', $MsgSubject)

		If ($Attachments)
		{
			$SendMailMsg.Add('Attachments', $Attachments)
		}

		Send-MailMessage @SendMailMsg
		fn_WriteToLog -LogEntry "`t`t`t`tError Message Email Sent!"
	}

	Function fn_BuildErrorMsgEmail
	{
	    [CmdletBinding()]
		Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline	= $false
				, HelpMessage		= 'Error Message for use in the Email Message Body.'
			)][string]$ErrMsg
		)

		#	USAGE:
		#		fn_BuildErrorMsgEmail -ErrMsg $ErrMsg

		[string]$HtmlReport = "<!DOCTYPE html PUBLIC $([char]34)-//W3C//DTD XHTML 1.0 Transitional//EN$([char]34) $([char]34)http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)<html xmlns=$([char]34)http://www.w3.org/1999/xhtml$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)<head>"
		$HtmlReport += "$([char]13)$([char]10)	<meta http-equiv=$([char]34)Content-Type$([char]34) content=$([char]34)text/html; charset=iso-8859-1$([char]34) />"
		$HtmlReport += "$([char]13)$([char]10)	<title>Task-sftpTranferScript - Error Report</title>"
		#region CSS Styles
			$HtmlReport += "$([char]13)$([char]10)	<style type=$([char]34)text/css$([char]34)>"
			#$HtmlReport += "$([char]13)$([char]10)	<!--"
			$HtmlReport += "$([char]13)$([char]10)		body {"
			$HtmlReport += "$([char]13)$([char]10)			font-family: Tahoma;"
			$HtmlReport += "$([char]13)$([char]10)			font-size: 11px;"
			$HtmlReport += "$([char]13)$([char]10)			margin-left: 0px;"
			$HtmlReport += "$([char]13)$([char]10)			margin-top: 0px;"
			$HtmlReport += "$([char]13)$([char]10)			margin-right: 0px;"
			$HtmlReport += "$([char]13)$([char]10)			margin-bottom: 0px;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		table {"
			$HtmlReport += "$([char]13)$([char]10)			width: 895px;"
			$HtmlReport += "$([char]13)$([char]10)			border-width: 0;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td {"
			$HtmlReport += "$([char]13)$([char]10)			padding-top: 0px;"
			$HtmlReport += "$([char]13)$([char]10)			padding-bottom: 0px;"
			$HtmlReport += "$([char]13)$([char]10)			padding-right: 2px;"
			$HtmlReport += "$([char]13)$([char]10)			padding-left: 2px;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.ReportHeaderLeft {"
			$HtmlReport += "$([char]13)$([char]10)			background-color: #313431;"
			$HtmlReport += "$([char]13)$([char]10)			vertical-align: middle;"
			$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
			$HtmlReport += "$([char]13)$([char]10)			width: 160px;"	# 200px;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.ReportHeaderRight {"
			$HtmlReport += "$([char]13)$([char]10)			background-color: #313431;"
			$HtmlReport += "$([char]13)$([char]10)			vertical-align: middle;"
			$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
			$HtmlReport += "$([char]13)$([char]10)			width: 160px;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.ReportTitle {"
			$HtmlReport += "$([char]13)$([char]10)			background-color: #313431;"
			$HtmlReport += "$([char]13)$([char]10)			vertical-align: middle;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.HdrReportDate {"
			$HtmlReport += "$([char]13)$([char]10)			background-color: #313431;"
			$HtmlReport += "$([char]13)$([char]10)			vertical-align: bottom;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.RowSpacers {"
			#$HtmlReport += "$([char]13)$([char]10)			color: #FFFFFF;"
			$HtmlReport += "$([char]13)$([char]10)			background-color: #313431;"
			$HtmlReport += "$([char]13)$([char]10)			height: 2px;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.RedCell {"
			$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
			$HtmlReport += "$([char]13)$([char]10)			background-color: #df202e;"	# Should be the Red from the SquareTwo Financial Logo.
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		td.CenterCell {"
			$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		div.ReportTitle {"
			$HtmlReport += "$([char]13)$([char]10)			width: 550px;"
			$HtmlReport += "$([char]13)$([char]10)			color: #FFFFFF;"
			$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
			$HtmlReport += "$([char]13)$([char]10)			font-size: 22px;"
			$HtmlReport += "$([char]13)$([char]10)			font-weight: bold;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			$HtmlReport += "$([char]13)$([char]10)"
			$HtmlReport += "$([char]13)$([char]10)		div.ReportDate {"
			$HtmlReport += "$([char]13)$([char]10)			width: 550px;"
			$HtmlReport += "$([char]13)$([char]10)			color: #FFFFFF;"
			$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
			$HtmlReport += "$([char]13)$([char]10)			font-size: 9px;"
			$HtmlReport += "$([char]13)$([char]10)		}"
			#$HtmlReport += "$([char]13)$([char]10)	-->"
			$HtmlReport += "$([char]13)$([char]10)	</style>"
		#endregion CSS Styles
		$HtmlReport += "$([char]13)$([char]10)</head>"
		$HtmlReport += "$([char]13)$([char]10)<body>"
		$HtmlReport += "$([char]13)$([char]10)<table border=$([char]34)0$([char]34) cellpadding=$([char]34)0$([char]34) cellspacing=$([char]34)0$([char]34) align=$([char]34)center$([char]34)>"
		#region Table Header Section
			$HtmlReport += "$([char]13)$([char]10)	<tr>"
			$HtmlReport += "$([char]13)$([char]10)		<td rowspan=$([char]34)2$([char]34) class=$([char]34)ReportHeaderLeft$([char]34)>"
			$HtmlReport += "$([char]13)$([char]10)			<img src=$([char]34)http://www.pearson.com/etc/designs/pearson-corporate/images/2010/hc/logo.gif$([char]34) alt=$([char]34)Pearson Logo$([char]34) />"
			$HtmlReport += "$([char]13)$([char]10)		</td>"
			$HtmlReport += "$([char]13)$([char]10)		<td class=$([char]34)ReportTitle$([char]34)>"
			$HtmlReport += "$([char]13)$([char]10)			<div class=$([char]34)ReportTitle$([char]34)>Task-sftpTranferScript - Error Report</div>"
			$HtmlReport += "$([char]13)$([char]10)		</td>"
			$HtmlReport += "$([char]13)$([char]10)		<td rowspan=$([char]34)2$([char]34) class=$([char]34)ReportHeaderRight$([char]34)>"
			$HtmlReport += "$([char]13)$([char]10)			<img src=$([char]34)http://www.pearson.com/etc/designs/pearson-corporate/images/2010/hc/tagLine.gif$([char]34) alt=$([char]34)Pearson Tag Line Logo$([char]34) />"
			$HtmlReport += "$([char]13)$([char]10)		</td>"
			$HtmlReport += "$([char]13)$([char]10)	</tr>"
			$HtmlReport += "$([char]13)$([char]10)	<tr>"
			$HtmlReport += "$([char]13)$([char]10)		<td class=$([char]34)HdrReportDate$([char]34)>"
			$HtmlReport += "$([char]13)$([char]10)			<div class=$([char]34)ReportDate$([char]34)>Generated on: $([char]34)$((Get-Date -Format 'MMMM dd, yyyy HH:mm').ToString())$([char]34)</div>"
			$HtmlReport += "$([char]13)$([char]10)		</td>"
			$HtmlReport += "$([char]13)$([char]10)	</tr>"
			$HtmlReport += "$([char]13)$([char]10)	<tr>"
			$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)3$([char]34) class=$([char]34)RowSpacers$([char]34)></td>"
			$HtmlReport += "$([char]13)$([char]10)	</tr>"
		#endregion Table Header Section
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)3$([char]34) class=$([char]34)CenterCell$([char]34)>&nbsp;</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)3$([char]34) class=$([char]34)RedCell$([char]34)>$($ErrMsg)</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)3$([char]34) class=$([char]34)CenterCell$([char]34)>&nbsp;</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)3$([char]34) class=$([char]34)CenterCell$([char]34)>Please see the log file on $([char]34)$($RptSvrFqdn)$([char]34) in the following location:$([char]13)$([char]10)$([char]34)$($FullScriptLogPath)$([char]34)</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)3$([char]34) class=$([char]34)CenterCell$([char]34)>&nbsp;</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)</table>"
		$HtmlReport += "$([char]13)$([char]10)</body>"
		$HtmlReport += "$([char]13)$([char]10)</html>"
		$HtmlReport += "$([char]13)$([char]10)"

		return $HtmlReport
	}

##############################################################################################################
#	#	Contents of each Object in the Array returned from the SFTP connection.
##############################################################################################################
#	#	Name            : CallLogs
#	#	FileType        : d
#	#	Length          : 0
#	#	LastWriteTime   : 2/5/2013 4:42:00 PM
#	#	FilePermissions : rwxr-xr-x
#	#	IsDirectory     : True
##############################################################################################################
#	#	Contents of each Directory Object in the Array returned from the Secure File Server.
##############################################################################################################
#	#	PSPath            : Microsoft.PowerShell.Core\FileSystem::D:\SecureShares\CallLogs\FromPearsonEcollege\70
#	#	PSParentPath      : Microsoft.PowerShell.Core\FileSystem::D:\SecureShares\CallLogs\FromPearsonEcollege
#	#	PSChildName       : 70
#	#	PSDrive           : D
#	#	PSProvider        : Microsoft.PowerShell.Core\FileSystem
#	#	PSIsContainer     : True
#	#	BaseName          : 70
#	#	Mode              : d----
#	#	Name              : 70
#	#	Parent            : FromPearsonEcollege
#	#	Exists            : True
#	#	Root              : D:\
#	#	FullName          : D:\SecureShares\CallLogs\FromPearsonEcollege\70
#	#	Extension         :
#	#	CreationTime      : 2/7/2013 12:51:27 PM
#	#	CreationTimeUtc   : 2/7/2013 7:51:27 PM
#	#	LastAccessTime    : 2/7/2013 12:52:47 PM
#	#	LastAccessTimeUtc : 2/7/2013 7:52:47 PM
#	#	LastWriteTime     : 2/7/2013 12:52:47 PM
#	#	LastWriteTimeUtc  : 2/7/2013 7:52:47 PM
#	#	Attributes        : Directory
##############################################################################################################
#	#	Contents of each NON-Directory Object in the Array returned from the Secure File Server.
##############################################################################################################
#	#	PSPath            : Microsoft.PowerShell.Core\FileSystem::D:\SecureShares\CallLogs\FromPearsonEcollege\Volume Shadow Copy (VSS) Settings Changes for NetBackup 7.txt
#	#	PSParentPath      : Microsoft.PowerShell.Core\FileSystem::D:\SecureShares\CallLogs\FromPearsonEcollege
#	#	PSChildName       : Volume Shadow Copy (VSS) Settings Changes for NetBackup 7.txt
#	#	PSDrive           : D
#	#	PSProvider        : Microsoft.PowerShell.Core\FileSystem
#	#	PSIsContainer     : False
#	#	VersionInfo       : File:             D:\SecureShares\CallLogs\FromPearsonEcollege\Volume Shadow Copy (VSS) Settings Changesfor NetBackup 7.txt
#	#	                    InternalName:
#	#	                    OriginalFilename:
#	#	                    FileVersion:
#	#	                    FileDescription:
#	#	                    Product:
#	#	                    ProductVersion:
#	#	                    Debug:            False
#	#	                    Patched:          False
#	#	                    PreRelease:       False
#	#	                    PrivateBuild:     False
#	#	                    SpecialBuild:     False
#	#	                    Language:
#	#	BaseName          : Volume Shadow Copy (VSS) Settings Changes for NetBackup 7
#	#	Mode              : -a---
#	#	Name              : Volume Shadow Copy (VSS) Settings Changes for NetBackup 7.txt
#	#	Length            : 591
#	#	DirectoryName     : D:\SecureShares\CallLogs\FromPearsonEcollege
#	#	Directory         : D:\SecureShares\CallLogs\FromPearsonEcollege
#	#	IsReadOnly        : False
#	#	Exists            : True
#	#	FullName          : D:\SecureShares\CallLogs\FromPearsonEcollege\Volume Shadow Copy (VSS) Settings Changes for NetBackup 7.txt
#	#	Extension         : .txt
#	#	CreationTime      : 2/8/2013 8:23:04 AM
#	#	CreationTimeUtc   : 2/8/2013 3:23:04 PM
#	#	LastAccessTime    : 2/8/2013 8:23:04 AM
#	#	LastAccessTimeUtc : 2/8/2013 3:23:04 PM
#	#	LastWriteTime     : 10/19/2011 8:45:00 AM
#	#	LastWriteTimeUtc  : 10/19/2011 2:45:00 PM
#	#	Attributes        : Archive
##############################################################################################################

	Function fn_GetSftpDirListing
	{
	    [CmdletBinding()]
	    Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'SFTP Server Directory Path'
			)][string]$SftpDirectoryPath
			, [Parameter(
				Position			= 1
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Returns only the type indicated. Valid values are: "File" or "Directory"'
			)]
			[ValidateSet('File', 'Directory')]
			[string]$FileOrDir
		)
		# USAGE:
		#	fn_GetSftpDirListing -SftpDirectoryPath $SftpDirectoryPath -FileOrDir 'Directory'
		#	fn_GetSftpDirListing -SftpDirectoryPath $SftpDirectoryPath -FileOrDir 'File'

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Yellow ("Parameters passed into the Function $([char]34)fn_GetSftpDirListing$([char]34) are: $([char]34){0}$([char]34) and $([char]34){1}$([char]34)" -f $SftpDirectoryPath, $FileOrDir)
		}
		fn_WriteToLog -LogEntry "`t`t`t`tParameters passed into the Function $([char]34)fn_GetSftpDirListing$([char]34) are: $([char]34)$($SftpDirectoryPath)$([char]34) and $([char]34)$($FileOrDir)$([char]34)"

		$ArrSftpSvrResults = @()
		$SftpDirListing = $session.ListDirectory($SftpDirectoryPath)

		If ($FileOrDir -eq 'Directory')
		{
			# Returns a listing of only Directories in the path specified.
			ForEach ($ObjInfo in $SftpDirListing.Files)
			{
				[string]$ObjName = $ObjInfo.Name
				[string]$ObjFileType = $ObjInfo.FileType
				[string]$ObjLastWriteTime = $ObjInfo.LastWriteTime
				[string]$ObjFilePermissions = $ObjInfo.FilePermissions
				[string]$ObjIsDirectory = $ObjInfo.IsDirectory

				# Add to the array if the object HAS the 'IsDirectory' attribute set to true.
				If ($ObjIsDirectory -eq 'True')
				{
					If ($ObjName -ne '.')
					{
						If ($ObjName -ne '..')
						{
							$ObjSftpInfo = New-Object System.Object
							$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'Name' -Value $ObjName
							$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'FileType' -Value $ObjFileType
							$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'LastWriteTime' -Value $ObjLastWriteTime
							$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'FilePermissions' -Value $ObjFilePermissions
							$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'IsDirectory' -Value $ObjIsDirectory
							$ArrSftpSvrResults += $ObjSftpInfo
						}
					}
				}
			}
		}
		Else
		{
			# Returns a listing of everything EXCEPT Directories in the path specified.
			ForEach ($ObjInfo in $SftpDirListing.Files)
			{
				[string]$ObjName = $ObjInfo.Name
				[string]$ObjFileType = $ObjInfo.FileType
				[string]$ObjLastWriteTime = $ObjInfo.LastWriteTime
				[string]$ObjFilePermissions = $ObjInfo.FilePermissions
				[string]$ObjIsDirectory = $ObjInfo.IsDirectory

				# Add to the array if the object DOES NOT HAVE the 'IsDirectory' attribute set to true.
				If ($ObjIsDirectory -ne 'True')
				{
					$ObjSftpInfo = New-Object System.Object
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'Name' -Value $ObjName
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'FileType' -Value $ObjFileType
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'LastWriteTime' -Value $ObjLastWriteTime
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'FilePermissions' -Value $ObjFilePermissions
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'IsDirectory' -Value $ObjIsDirectory
					$ArrSftpSvrResults += $ObjSftpInfo
				}
			}
		}

		#If ($ScrDebug)
		#{
		#	Write-Host -ForegroundColor Red 'Array ArrSftpSvrResults contains the following:'
		#	$ArrSftpSvrResults | Format-Table
		#}

		#return $ArrSftpSvrResults
		$ArrSftpSvrResults
	}

	Function fn_GetSecFileSvrDirListing
	{
	    [CmdletBinding()]
	    Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Full Secure File Server Directory Path. EX: D:\SecureFiles\Outbound\565'
			)][string]$SecFileSvrDirPath
			, [Parameter(
				Position			= 1
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Returns only the type indicated. Valid values are: "File" or "Directory"'
			)]
			[ValidateSet('File', 'Directory')]
			[string]$FileOrDir
		)
		# USAGE:
		#	fn_GetSecFileSvrDirListing -SecFileSvrDirPath $SecFileSvrDirPath -FileOrDir 'Directory'
		#	fn_GetSecFileSvrDirListing -SecFileSvrDirPath $SecFileSvrDirPath -FileOrDir 'File'

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Yellow ("Parameters passed into the Function $([char]34)fn_GetSecFileSvrDirListing$([char]34) are: $([char]34){0}$([char]34) and $([char]34){1}$([char]34)" -f $SecFileSvrDirPath, $FileOrDir)
		}
		fn_WriteToLog -LogEntry "`t`t`t`tParameters passed into the Function $([char]34)fn_GetSecFileSvrDirListing$([char]34) are: $([char]34)$($SecFileSvrDirPath)$([char]34) and $([char]34)$($FileOrDir)$([char]34)"

		$ArrSecFileSvrResults = @()

		If ($FileOrDir -eq 'Directory')
		{
			# Returns a listing of only Directories in the path specified.
			[array]$SecFileSvrDirListing = Get-ChildItem -Path $SecFileSvrDirPath | Where-Object { $_.Attributes -eq 'Directory' }

			ForEach ($ObjInfo in $SecFileSvrDirListing)
			{
				[string]$ObjName = $ObjInfo.Name
				[string]$ObjFullName = $ObjInfo.FullName
				[string]$ObjLastWriteTime = $ObjInfo.LastWriteTime
				[string]$ObjAttributes = $ObjInfo.Attributes

				If ($ObjAttributes -eq 'Directory')
				{
					[string]$ObjIsDirectory = 'True'
				}
				Else
				{
					[string]$ObjIsDirectory = 'False'
				}

				If ($ObjAttributes -eq 'Directory')
				{
					$ObjSftpInfo = New-Object System.Object
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'Name' -Value $ObjName
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'FullName' -Value $ObjFullName
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'LastWriteTime' -Value $ObjLastWriteTime
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'Attributes' -Value $ObjAttributes
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'IsDirectory' -Value $ObjIsDirectory
					$ArrSecFileSvrResults += $ObjSftpInfo
				}
			}
		}
		Else
		{
			# Returns a listing of everything EXCEPT Directories in the path specified.
			[array]$SecFileSvrDirListing = Get-ChildItem -Path $SecFileSvrDirPath | Where-Object { $_.Attributes -ne 'Directory' }

			ForEach ($ObjInfo in $SecFileSvrDirListing)
			{
				[string]$ObjName = $ObjInfo.Name
				[string]$ObjFullName = $ObjInfo.FullName
				[string]$ObjLastWriteTime = $ObjInfo.LastWriteTime
				[string]$ObjAttributes = $ObjInfo.Attributes

				If ($ObjAttributes -eq 'Directory')
				{
					[string]$ObjIsDirectory = 'True'
				}
				Else
				{
					[string]$ObjIsDirectory = 'False'
				}

				If ($ObjAttributes -ne 'Directory')
				{
					$ObjSftpInfo = New-Object System.Object
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'Name' -Value $ObjName
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'FullName' -Value $ObjFullName
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'LastWriteTime' -Value $ObjLastWriteTime
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'Attributes' -Value $ObjAttributes
					$ObjSftpInfo | Add-Member -MemberType NoteProperty -Name 'IsDirectory' -Value $ObjIsDirectory
					$ArrSecFileSvrResults += $ObjSftpInfo
				}
			}
		}

		return $ArrSecFileSvrResults
	}

	Function fn_PutFilesOnSftp
	{
	    [CmdletBinding()]
	    Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Full SFTP Server Directory Path. EX: ./Franchises/<CollectionFolderName>/FromPearsonEcollege/555'
			)][string]$SftpDirectoryPath
			, [Parameter(
				Position			= 1
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Full Secure File Server Directory Path. EX: D:\SecureShares\Franchises\<CollectionFolderName>\FromPearsonEcollege\555'
			)][string]$SecFileSvrDirSrcPath
			, [Parameter(
				Position			= 2
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Directory Path of the "TransferComplete" directory to move files to for successful uploads.'
			)][string]$TransferCompleteDirPath
		)
		# USAGE:
		#	fn_PutFilesOnSftp -SftpDirectoryPath $SftpDirectoryPath -SecFileSvrDirSrcPath $SecFileSvrDirSrcPath -TransferCompleteDirPath $TransferCompleteDirPath
		#	fn_PutFilesOnSftp -SftpDirectoryPath $SftpDirectoryPath -SecFileSvrDirSrcPath $SecFileSvrDirSrcPath -TransferCompleteDirPath $TransferCompleteDirPath

		If (!($session.FileExists($SftpDirectoryPath)))
		{
			$session.CreateDirectory($SftpDirectoryPath)
			If ($ScrDebug)
			{
				Write-Host -ForegroundColor Yellow "[Outbound] - [Directory Created]:`t$([char]34)$($SftpDirectoryPath)$([char]34)"
			}
			fn_WriteToLog -LogEntry "[Outbound] - [Directory Created]:`t$([char]34)$($SftpDirectoryPath)$([char]34)"
		}

		##############################################################################################################################
		## Get a directory listing of the $LocalRootDirectory\<CollectionFolderName>\FromPearsonEcollege							##
		##############################################################################################################################
		$FilesOutbound = fn_GetSecFileSvrDirListing -SecFileSvrDirPath $SecFileSvrDirSrcPath -FileOrDir 'File'
		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Yellow "[Outbound] -`tThe contents of the $([char]34)FilesOutbound$([char]34) array is:"
			$FilesOutbound | Format-Table
			Write-Host -ForegroundColor Yellow "[Outbound] -`tTransfer Complete directory is: $($TransferCompleteDirPath)"

			Write-Host -ForegroundColor Cyan ("[Outbound] - Secure File Server Out Going Directory Name: $([char]34){0}$([char]34)" -f $SecFileSvrDirSrcPath)
			Write-Host -ForegroundColor DarkCyan ("[Outbound] - SFTP Out Going Directory Name: $([char]34){0}$([char]34)" -f $SftpDirectoryPath)
			Write-Host -ForegroundColor Black ("[Outbound] - Transfer Complete Directory Name is: $([char]34){0}$([char]34)" -f $TransferCompleteDirPath)
		}
		fn_WriteToLog -LogEntry "[Outbound] -`tSecure File Server Out Going Directory Name: $([char]34)$SecFileSvrDirSrcPath$([char]34)"
		fn_WriteToLog -LogEntry "[Outbound] -`tSFTP Out Going Directory Name: $([char]34)$SftpDirectoryPath$([char]34)"
		fn_WriteToLog -LogEntry "[Outbound] -`tTransfer Complete Directory Name is: $([char]34)$TransferCompleteDirPath$([char]34)"

		ForEach ($FromFile in $FilesOutbound)
		{
			$LocalFileName = "$($SecFileSvrDirSrcPath)\$($FromFile.Name)"
			$RemoteFileName = "$($SftpDirectoryPath)/$($FromFile.Name)"

			# Checks to see if the variable is a directory.
			[string]$LocalFileNameLastChar = $LocalFileName.SubString($LocalFileName.Length - 1)
			If ($LocalFileNameLastChar -eq '\')
			{
				If ($ScrDebug)
				{
					Write-Host -ForegroundColor DarkRed ("[Outbound] - Local File is a Directory: $([char]34){0}$([char]34)" -f $LocalFileName)
				}
				fn_WriteToLog -LogEntry "[Outbound] -`tLocal File is a Directory: $([char]34)$($LocalFileName)$([char]34)"
			}
			Else
			{
				If (!(Test-Path -Path $TransferCompleteDirPath))
				{
					New-Item -ItemType Directory $TransferCompleteDirPath | Out-Null
					If ($ScrDebug)
					{
						Write-Host -ForegroundColor Yellow "[Outbound] - [Directory Created] -`t$([char]34)$($TransferCompleteDirPath)$([char]34)"
					}
					fn_WriteToLog -LogEntry "[Outbound] - [Directory Created] -`t$([char]34)$($TransferCompleteDirPath)$([char]34)"
				}

				If ($ScrDebug)
				{
					Write-Host -ForegroundColor Green ("[Outbound] - LocalFileName: $([char]34){0}$([char]34)" -f $LocalFileName)
					Write-Host -ForegroundColor DarkGreen ("[Outbound] - RemoteFileName: $([char]34){0}$([char]34)" -f $RemoteFileName)
				}
				fn_WriteToLog -LogEntry "[Outbound]`t- LocalFileName: $([char]34)$($LocalFileName)$([char]34)"
				fn_WriteToLog -LogEntry "[Outbound]`t- RemoteFileName: $([char]34)$($RemoteFileName)$([char]34)"

				$session.PutFiles(
					$LocalFileName
					, $RemoteFileName
					, $false		#, $true	# Remove Files on Transfer
					, $transferOptions
				).Check()

				Start-Sleep -Seconds 5

				Try
				{
					If (Test-Path -Path $LocalFileName)
					{
						fn_WriteToLog -LogEntry "`t`t`t`t1. - File Still Exists, trying to move the file: $($LocalFileName)"
						Move-Item -Path $LocalFileName -Destination $TransferCompleteDirPath -Force
					}
					Start-Sleep -Seconds 3

					If (Test-Path -Path $LocalFileName)
					{
						fn_WriteToLog -LogEntry "`t`t`t`t2. - File Still Exists, trying to move the file again: $($LocalFileName)"
						Move-Item -Path $LocalFileName -Destination $TransferCompleteDirPath -Force
					}
					Start-Sleep -Seconds 3

					If (Test-Path -Path $LocalFileName)
					{
						fn_WriteToLog -LogEntry "`t`t`t`t3. - File Still Exists, trying to move the file again: $($LocalFileName)"
						Move-Item -Path $LocalFileName -Destination $TransferCompleteDirPath -Force
					}
					Start-Sleep -Seconds 3
				}
				Catch [Exception]
				{
					fn_WriteToLog -LogEntry "$([char]13)$([char]10)`t`t`t[EXCEPTION ERROR MOVING FILE] - $([char]34)$($_.Exception.Message)$([char]34)"
				}
				Finally
				{
					If (Test-Path -Path $LocalFileName)
					{
						fn_WriteToLog -LogEntry "$([char]13)$([char]10)`t`t`t[ERROR MOVING FILE] - 4. - File Still Exists, trying to move the file one last time: $($LocalFileName)"
						Move-Item -Path $LocalFileName -Destination $TransferCompleteDirPath -Force
					}

					If (Test-Path -Path $LocalFileName)
					{
						[string]$FileMvErrorMsg = "[ERROR MOVING FILE] - A FILE WAS NOT MOVED AND STILL EXISTS AT: $($LocalFileName)"
						fn_WriteToLog -LogEntry "$([char]13)$([char]10)`t`t`t$($FileMvErrorMsg)"

						$EmlMsg = fn_BuildErrorMsgEmail -ErrMsg $FileMvErrorMsg
						fn_SendEmail -MsgBody $EmlMsg -AddrFrom $FromAddr -AddrTo $ToAddr -SmtpHost $MailServer
						fn_WriteToLog -LogEntry "$([char]13)$([char]10)[Outbound] - [UPLOAD - MOVE ERROR]`t- Uploaded Local File $([char]34)$($LocalFileName)$([char]34) to $([char]34)$($RemoteFileName)$([char]34) but was not able to move the local copy to: $([char]34)$($TransferCompleteDirPath)$([char]34)"
					}
					Else
					{
						fn_WriteToLog -LogEntry "[Outbound] - [UPLOAD]`t- Uploaded Local File $([char]34)$($LocalFileName)$([char]34) to $([char]34)$($RemoteFileName)$([char]34) and Moved Local copy to: $([char]34)$($TransferCompleteDirPath)$([char]34)"
					}

					If ($ScrDebug)
					{
						$TransCompDirListing = fn_GetSecFileSvrDirListing -SecFileSvrDirPath $TransferCompleteDirPath -FileOrDir 'File'
						Write-Host -ForegroundColor Yellow "[Outbound] - The contents of the $([char]34)TransCompDirListing$([char]34) array is:"
						$TransCompDirListing | Format-Table
						Write-Host -ForegroundColor Yellow '---------------------------------------------------------------------------------------'
					}
				}
			}
		}

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Yellow "[Outbound] - Finished with Directory: $([char]34)$($SecFileSvrDirSrcPath)$([char]34)"
		}
		fn_WriteToLog -LogEntry "[Outbound] -`tFinished with Directory: $([char]34)$($SecFileSvrDirSrcPath)$([char]34)"
	}

	Function fn_GetFilesFromSftp
	{
	    [CmdletBinding()]
	    Param(
			[Parameter(
				Position			= 0
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Full SFTP Server Directory Path EX: ./Franchises/<CollectionFolderName>/ToPearsonEcollege/555'
			)][string]$SftpDirPath
			, [Parameter(
				Position			= 1
				, Mandatory			= $true
				, ValueFromPipeline = $false
				, HelpMessage		= 'Full Secure File Server Directory Path. EX: D:\SecureShares\Franchises\<CollectionFolderName>\FromPearsonEcollege\555'
			)][string]$SfsDirSrcPath
		)
		# USAGE:
		#	fn_GetFilesFromSftp -SftpDirPath $SftpDirPath -SfsDirSrcPath $SfsDirSrcPath
		#	fn_GetFilesFromSftp -SftpDirPath $SftpDirPath -SfsDirSrcPath $SfsDirSrcPath

		##################################################################################################################
		## Get a directory listing of the $SftpRootDirectory/<CollectionFolderName>/ToPearsonEcollege					##
		##################################################################################################################
		$FilesInbound = @(fn_GetSftpDirListing -SftpDirectoryPath $SftpDirPath -FileOrDir 'File')

		If ($FilesInbound.Length -ge 1)
		{
			If ($ScrDebug)
			{
				Write-Host -ForegroundColor Magenta "[Inbound] - Array $([char]34)FilesInbound$([char]34) has a length greater than or equal to 1!"
			}
			fn_WriteToLog -LogEntry "[Inbound] -`t`tArray $([char]34)FilesInbound$([char]34) has a length greater than or equal to 1!"

			If (!(Test-Path -Path $SfsDirSrcPath))
			{
				New-Item -ItemType Directory $SfsDirSrcPath | Out-Null

				If ($ScrDebug)
				{
					Write-Host -ForegroundColor DarkRed ("[Inbound] - [Directory Created]: $([char]34){0}$([char]34)" -f $SfsDirSrcPath)
				}
				fn_WriteToLog -LogEntry "[Inbound] -`t`t[Directory Created]: $([char]34)$($SfsDirSrcPath)$([char]34)"
			}
		}
		Else
		{
			If ($ScrDebug)
			{
				Write-Host -ForegroundColor Red "[Inbound] - Array $([char]34)FilesInbound$([char]34) DOES NOT have a length greater than or equal to 1!"
			}
			fn_WriteToLog -LogEntry "[Inbound] -`t`tArray $([char]34)FilesInbound$([char]34) DOES NOT have a length greater than or equal to 1!"
		}

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Yellow "[Inbound] - The contents of the $([char]34)FilesInbound$([char]34) array is:"
			If ($FilesInbound.Length -ge 1)
			{
				$FilesInbound | Format-Table
			}
			Else
			{
				Write-Host -ForegroundColor DarkCyan "                The $([char]34)FilesInbound$([char]34) array is EMPTY!"
			}
			Write-Host -ForegroundColor Cyan ("[Inbound] - Secure File Server Incoming Directory Name: $([char]34){0}$([char]34)" -f $SfsDirSrcPath)
			Write-Host -ForegroundColor DarkCyan ("[Inbound] - SFTP Incoming Directory Name: $([char]34){0}$([char]34)" -f $SftpDirPath)
		}
		fn_WriteToLog -LogEntry "[Inbound] -`t`tSecure File Server Incoming Directory Name: $([char]34)$SfsDirSrcPath$([char]34)"
		fn_WriteToLog -LogEntry "[Inbound] -`t`tSFTP Incoming Directory Name: $([char]34)$SftpDirPath$([char]34)"

		ForEach ($ToFile in $FilesInbound)
		{
			$LocalToFile = "$($SfsDirSrcPath)\$($ToFile.Name)"
			$RemoteToFile = "$($SftpDirPath)/$($ToFile.Name)"
			If ($ScrDebug)
			{
				Write-Host -ForegroundColor DarkGreen ("[Inbound] - LocalToFile Name: $([char]34){0}$([char]34)" -f $LocalToFile)
				Write-Host -ForegroundColor DarkRed ("[Inbound] - RemoteToFile Name: $([char]34){0}$([char]34)" -f $RemoteToFile)
			}
			fn_WriteToLog -LogEntry "[Inbound] -`t`tLocalToFile Name: $([char]34)$($LocalToFile)$([char]34)"
			fn_WriteToLog -LogEntry "[Inbound] -`t`tRemoteToFile Name: $([char]34)$($RemoteToFile)$([char]34)"

			[string]$RemoteToFileLastChar = $RemoteToFile.SubString($RemoteToFile.Length - 1)
			If ($RemoteToFileLastChar -eq '/')
			{
				If ($ScrDebug)
				{
					Write-Host -ForegroundColor DarkRed ("[Inbound] - RemoteToFile $([char]34){0}$([char]34) is a Directory." -f $RemoteToFile)
				}
				fn_WriteToLog -LogEntry "[Inbound] -`t`tRemoteToFile $([char]34)$($RemoteToFile)$([char]34) is a Directory."
				#$DownloadFile = $false
			}
			ElseIf ($session.FileExists($RemoteToFile))
			{
				If (!(Test-Path $LocalToFile))
				{
					If ($ScrDebug)
					{
						Write-Host ("[Inbound] - Remote File $([char]34){0}$([char]34) exists, Local File $([char]34){1}$([char]34) does not" -f $RemoteToFile, $LocalToFile)
					}
					fn_WriteToLog -LogEntry "[Inbound] -`t`tRemote File $([char]34)$($RemoteToFile)$([char]34) exists, Local File $([char]34)$($LocalToFile)$([char]34) does not"
					$DownloadFile = $true
				}
				Else
				{
					$remoteWriteTime = $session.GetFileInfo($RemoteToFile).LastWriteTime
					$localWriteTime = (Get-Item $LocalToFile).LastWriteTime

					If ($remoteWriteTime -gt $localWriteTime)
					{
						If ($ScrDebug)
						{
							Write-Host (("[Inbound] - Remote File $([char]34){0}$([char]34) as well as Local File $([char]34){1}$([char]34) exist, but Remote File is newer ({2}) than Local File ({3})") -f $RemoteToFile, $LocalToFile, $remoteWriteTime, $localWriteTime)
						}
						fn_WriteToLog -LogEntry "[Inbound] -`t`tRemote File $([char]34)$($RemoteToFile)$([char]34) as well as Local File $([char]34)$($LocalToFile)$([char]34) exist, but Remote File is newer ($($remoteWriteTime)) than Local File ($($localWriteTime))"
						$DownloadFile = $true
					}
					Else
					{
						If ($ScrDebug)
						{
							Write-Host (("[Inbound] - Remote File $([char]34){0}$([char]34) as well as Local File $([char]34){1}$([char]34) exist, but Remote File is not newer ({2}) than Local File ({3})") -f $RemoteToFile, $LocalToFile, $remoteWriteTime, $localWriteTime)
						}
						fn_WriteToLog -LogEntry "[Inbound] -`t`tRemote File $([char]34)$($RemoteToFile)$([char]34) as well as Local File $([char]34)$($LocalToFile)$([char]34) exist, but Remote File is not newer ($($remoteWriteTime)) than Local File ($($localWriteTime))"
						$DownloadFile = $false
					}
				}

				If ($DownloadFile)
				{
					$session.GetFiles(
						$RemoteToFile
						, $LocalToFile
						, $true		#, $false	# Remove Files on Transfer
						, $transferOptions
					).Check()

					If ($ScrDebug)
					{
						Write-Host -ForegroundColor Magenta "[Inbound] - [DOWNLOAD] - Download of $([char]34)$($RemoteToFile)$([char]34) complete."
					}
					fn_WriteToLog -LogEntry "[Inbound] - [DOWNLOAD]`t`t- Download of $([char]34)$($RemoteToFile)$([char]34) complete."
				}
			}
		} # END ForEach Loop

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Yellow "[Inbound] - Finished with Remote Directory: $([char]34)$($SftpDirPath)$([char]34)"
		}
		fn_WriteToLog -LogEntry "[Inbound] -`t`tFinished with Remote Directory: $([char]34)$($SftpDirPath)$([char]34)"
	}

	# Session.FileTransferred event handler
	Function FileTransferred
	{
		If ($_.Error -eq $null)
		{
			If ($ScrDebug)
			{
				#Write-Host ("Upload of $([char]34){0}$([char]34) succeeded" -f $_.FileName)
				Write-Host ("Transfer of $([char]34){0}$([char]34) succeeded" -f $_.FileName)
			}
			fn_WriteToLog -LogEntry "`t`t`t`tTransfer of $([char]34)$($_.FileName)$([char]34) succeeded"
		}
		Else
		{
			If ($ScrDebug)
			{
				#Write-Host ("Upload of $([char]34){0}$([char]34) failed: $([char]34){1}$([char]34)" -f $_.FileName, $_.Error)
				Write-Host ("[ERROR] - Transfer of $([char]34){0}$([char]34) failed: $([char]34){1}$([char]34)" -f $_.FileName, $_.Error)
			}
			fn_WriteToLog -LogEntry "$([char]13)$([char]10)[ERROR] -`t`t`tTransfer of $([char]34)$($_.FileName)$([char]34) failed: $([char]34)$($_.Error)$([char]34)"
		}

		If ($_.Chmod -ne $null)
		{
			If ($_.Chmod.Error -eq $null)
			{
				If ($ScrDebug)
				{
					Write-Host ("Permisions of $([char]34){0}$([char]34) set to $([char]34){1}$([char]34)" -f $_.Chmod.FileName, $_.Chmod.FilePermissions)
				}
				fn_WriteToLog -LogEntry "`t`t`t`tPermisions of $([char]34)$($_.Chmod.FileName)$([char]34) set to $([char]34)$($_.Chmod.FilePermissions)$([char]34)"
			}
			Else
			{
				If ($ScrDebug)
				{
					Write-Host ("[ERROR] - Setting permissions of $([char]34){0}$([char]34) failed: $([char]34){1}$([char]34)" -f $_.Chmod.FileName, $_.Chmod.Error)
				}
				fn_WriteToLog -LogEntry "$([char]13)$([char]10)[ERROR] -`t`t`tSetting permissions of $([char]34)$($_.Chmod.FileName)$([char]34) failed: $([char]34)$($_.Chmod.Error)$([char]34)"
			}
		}
		Else
		{
			If ($ScrDebug)
			{
				Write-Host ("Permissions of $([char]34){0}$([char]34) kept with their defaults" -f $_.Destination)
			}
			fn_WriteToLog -LogEntry "`t`t`t`tPermissions of $([char]34)$($_.Destination)$([char]34) kept with their defaults"
		}

		If ($_.Touch -ne $null)
		{
			If ($_.Touch.Error -eq $null)
			{
				If ($ScrDebug)
				{
					Write-Host ("Timestamp of $([char]34){0}$([char]34) set to $([char]34){1}$([char]34)" -f $_.Touch.FileName, $_.Touch.LastWriteTime)
				}
				fn_WriteToLog -LogEntry "`t`t`t`tTimestamp of $([char]34)$($_.Touch.FileName)$([char]34) set to $([char]34)$($_.Touch.LastWriteTime)$([char]34)"
			}
			Else
			{
				If ($ScrDebug)
				{
					Write-Host ("[ERROR] - Setting timestamp of $([char]34){0}$([char]34) failed: $([char]34){1}$([char]34)" -f $_.Touch.FileName, $_.Touch.Error)
				}
				fn_WriteToLog -LogEntry "$([char]13)$([char]10)[ERROR] -`t`t`tSetting timestamp of $([char]34)$($_.Touch.FileName)$([char]34) failed: $([char]34)$($_.Touch.Error)$([char]34)"
			}
		}
		Else
		{
			# This should never happen with Session.SynchronizeDirectories
			If ($ScrDebug)
			{
				Write-Host ("Timestamp of $([char]34){0}$([char]34) kept with its default (current time)" -f $_.Destination)
			}
			fn_WriteToLog -LogEntry "`t`t`t`tTimestamp of $([char]34)$($_.Destination)$([char]34) kept with its default (current time)"
		}
	}
#endregion User Functions

#region Main Script
	fn_WriteToLog -LogEntry "[BEGIN] -`t`t`tScript Execution for $([char]34)$($LogFileDate)$([char]34)"
	Try
	{
		# Setup Session Options for connection to the SFTP Server
		$sessionOptions = New-Object WinSCP.SessionOptions
		$sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
		$sessionOptions.HostName = $SftpHostName
		$sessionOptions.UserName = $SftpUserName
		$sessionOptions.Password = $SftpPassword
		$sessionOptions.SshHostKeyFingerprint = $SftpSshHostKeyFingerprint

		$session = New-Object WinSCP.Session

		$transferOptions = New-Object WinSCP.TransferOptions
		# Possible values are:
		# Binary (default)
		# Ascii
		# Automatic (based on file extension).
		# Set Proper TransferMode
		$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

		fn_WriteToLog -LogEntry "[CONNECTION OPEN] -`t`tConnect to Host: $([char]34)$($SftpHostName)$([char]34)"
		fn_WriteToLog -LogEntry "[CONNECTION OPEN] -`t`tAs User: $([char]34)$($SftpUserName)$([char]34)"
		fn_WriteToLog -LogEntry "[CONNECTION OPEN] -`t`tWith TransferMode set as: $([char]34)Binary$([char]34)$([char]13)$([char]10)"

		# Optional $session methods to be called before connecting to the SFTP server.
		$session.DefaultConfiguration = $false
		$session.IniFilePath = $WinScpIni
		If ($ScrDebug)
		{
			$session.DebugLogPath = $FullDebugLogPath
			$session.SessionLogPath = $FullSessionLogPath
		}

		# Connect to, and open a session with, the SFTP Server
		$session.Open($sessionOptions)

		# Will continuously report progress of synchronization
		$session.add_FileTransferred( { FileTransferred } )

		##########################################################
		## Get all files in the $SftpRootDirectory				##
		##########################################################
		fn_GetFilesFromSftp -SftpDirPath "$($SftpRootDirectory)" -SfsDirSrcPath "$($LocalRootDirectory)"
	}
	Catch [Exception]
	{
		[string]$ErrorLogEntryDate = (Get-Date -Format 'yyyy.MM.dd.ddd HH:mm').ToString()	# Example Output: 2012.04.25.Wed 14:00
		[string]$ErrorMsg = "[ERROR - $($ErrorLogEntryDate)] - An error occurred during script execution: $([char]13)$([char]10)$([char]34)$($_.Exception.Message)$([char]34)"

		If ($ScrDebug)
		{
			Write-Host -ForegroundColor Red "$($ErrorMsg)"
		}

		fn_WriteToLog -LogEntry "$([char]13)$([char]10)"
		fn_WriteToLog -LogEntry $ErrorMsg
		fn_WriteToLog -LogEntry "$([char]13)$([char]10)"
		fn_WriteToLog -LogEntry "`t`t`t`tSending Error Message Email"
		fn_WriteToLog -LogEntry "$([char]13)$([char]10)"

		$EmlErrMsg = fn_BuildErrorMsgEmail -ErrMsg $ErrorMsg
		fn_SendEmail -MsgBody $EmlErrMsg -AddrFrom $FromAddr -AddrTo $ToAddr -SmtpHost $MailServer
		#Exit 1
	}
	Finally
	{
		# Disconnect, clean up
		$session.Dispose()

		fn_WriteToLog -LogEntry "$([char]13)$([char]10)[CONNECTION CLOSED] -`t`t$([char]34)SFTP CONNECTION SUCCESSFULLY CLOSED!$([char]34)"
		fn_WriteToLog -LogEntry "$([char]13)$([char]10)[END] -`t`t`t`t$([char]34)Script Execution Complete!$([char]34)"
		#Exit 0
	}
#endregion Main Script
