<#
	@"
	===============================================================================
	Title:			Get-VMwareSnaphots.ps1
	Description:	List snapshots on VM's managed by Virtual Center.
	Requirements:	Windows Powershell and the VI Toolkit
	Usage:			.\Get-VMwareSnaphots.ps1
	Found At:		http://community.spiceworks.com/scripts/show/882-email-list-of-open-snapshots
					Major formatting and layout changes have been made from the original Script.
	===============================================================================
	"@

	Added Script:	23 November 2012
	Modified By:	Bill Campbell
	Last Modified:	2013-May-15
	Version:		2013.05.15.05
	Version Notes:	Version format is a date taking the following format: YYYY.MM.DD.RR
					where RR is the revision/save count for the day modified.
	Version Exmpl:	If this is the 6th revision/save on January 13 2012 then RR would be "06"
					and the version number will be formatted as follows: 2012.01.23.06

	Purpose:		To Monitor the mailbox sizes of all Resource account mailboxes.
					Mailboxes with "RecipientTypeDetails" equal to "DiscoveryMailbox"
					are excluded.

	Notes:			Use the next line to change to your user profile "Desktop" folder:
					cd ${env:userprofile}\Desktop
					echo ${env:computername}
	.SYNOPSIS

	.PARAMETER PtlRptName [string] (Optional)
    	Partial report name used to create the full report filename

	.PARAMETER SendMail [switch] (Optional)
		Send Mail after completion. Set to $true to enable. If enabled, -FromAddr, -ToAddr, -SmtpSvr are mandatory

	.PARAMETER FromAddr [string] (Optional, Required if setting the 'SendMail' Parameter.)
		Email address to send from. Passed directly to Send-MailMessage as -From

	.PARAMETER ToAddr [string] (Optional, Required if setting the 'SendMail' Parameter.)
		Email address to send to. Passed directly to Send-MailMessage as -To

	.PARAMETER SmtpSvr [string] (Optional, Required if setting the 'SendMail' Parameter.)
		SMTP Mail server to attempt to send through. Passed directly to Send-MailMessage as -SmtpServer

	.PARAMETER CreateCsv [switch] (Optional)
		Generate a CSV file, from the data used for the HTML report, to be attached to the email sent. Defaults to FALSE.

	.PARAMETER ScrDebug [switch] (Optional)
		The ScrDebug switch turns the debugging code in the script on or off ($true/$false). Defaults to FALSE.

	.EXAMPLE
		.\Get-VMwareSnaphots.ps1
		.\Get-VMwareSnaphots.ps1 -PtlRptName "Get-VMwareSnaphots" -SendMail:$true -FromAddr "server-alerts@squaretwofinancial.com" -ToAddr "server-alerts@squaretwofinancial.com" -SmtpSvr "mail.squaretwofinancial.com" -CreateCsv
		.\Get-VMwareSnaphots.ps1 -SendMail:$true -FromAddr 'server-alerts@squaretwofinancial.com' -ToAddr 'server-alerts@squaretwofinancial.com' -SmtpSvr 'mail.squaretwofinancial.com' -CreateCsv
		. ${env:userprofile}\Dropbox\Admin_Utils\Scripting\PowerShell_Scripts\VMware\Get-VMwareSnaphots.ps1 -SendMail:$true -FromAddr "server-alerts@squaretwofinancial.com" -ToAddr bcampbell@squaretwofinancial.com,itopscorner@squaretwofinancial.com -SmtpSvr "mail.squaretwofinancial.com" -CreateCsv -ScrDebug

	-- Script Changes --
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
	#!   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE   !#
	#! RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. !#
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#

	Changes for 23.Nov.2012
		- Initial script creation.

	Changes for 26.Nov.2012
		- HTML report formatting.
		- Changed Parameter $ToAddr from type string to type array.
		- Added command "Import-Module ActiveDirectory"

	Changes for 15.May.2013
		- Code cleanup and updating for enhanced performance.
		- Added checking around the loading of the Modules and Snapins.
#>

#region Script Initialization
	Param(
	    [Parameter(
			Position			= 0
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Partial report name used to create the full report filename'
		)][string]$PtlRptName
		, [Parameter(
			Position			= 1
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Send Mail ($true/$false). Defaults to FALSE.'
		)][switch]$SendMail = $false
		, [Parameter(
			Position			= 2
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Mail From Address'
		)][string]$FromAddr
		, [Parameter(
			Position			= 3
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Mail To Address'
		)][array]$ToAddr
		, [Parameter(
			Position			= 4
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Mail Server Address'
		)][string]$SmtpSvr
		, [Parameter(
			Position			= 5
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Generate a CSV file, from the data used for the HTML report, to be attached to the email sent. Defaults to FALSE.'
		)][switch]$CreateCsv = $false
		, [Parameter(
			Position			= 6
			, Mandatory			= $false
			, ValueFromPipeline	= $false
			, HelpMessage		= 'The ScrDebug switch turns the debugging code in the script on or off ($true/$false). Defaults to FALSE.'
		)][switch]$ScrDebug = $false
	)

	# Check and throw error to update PowerShell to version 2.0 if not at that level.
	If((Get-Host).Version -lt "2.0")
	{
	    Throw "You are running PowerShell version 1.0. PowerShell 1.0 is not supported. Please update your PowerShell to version 2.0 or higher!"
	}

	#region PowerShell Module Initialization
		######################################
		##  Import Modules & Add PSSnapins  ##
		######################################
		# Check for ActiveDirectory Module, attempt to load
		If (!(Get-Command Get-ADDomain -ErrorAction SilentlyContinue))
		{
			try
			{
				Import-Module ActiveDirectory
			}
			catch
			{
				If (!(Get-Command Get-ADDomain -ErrorAction SilentlyContinue))
				{
					Import-Module ActiveDirectory
				}
			}
			finally
			{
				If (!(Get-Command Get-ADDomain -ErrorAction SilentlyContinue))
				{
					throw "Cannot load the ActiveDirectory module!! Please make sure the RSAT for ActiveDirectory is installed."
				}
			}
		}
	#endregion PowerShell Module Initialization

	#region Load VMware vSphere PowerCLI Environment
		If (!(Get-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue))
		{
			try
			{
				Add-PSSnapin VMware.VimAutomation.Core
			}
			catch
			{
				If (!(Get-Command Get-PowerCLIConfiguration -ErrorAction SilentlyContinue))
				{
					. "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"
				}
			}
			finally
			{
				If (!(Get-Command Get-PowerCLIConfiguration -ErrorAction SilentlyContinue))
				{
					throw "Cannot load the VMware vSphere PowerCLI Environment!!"
				}
			}
		}

		If ((Get-PowerCLIConfiguration | Select-Object InvalidCertificateAction).InvalidCertificateAction -ne "Ignore")
		{
			Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
		}
	#endregion Load VMware vSphere PowerCLI Environment

	If (!$PtlRptName)
	{
		[string]$PtlRptName = ($MyInvocation.MyCommand.Name).Replace(".ps1", "")	#.Replace("Get-", "")
	}

	If ($SendMail)
	{
		If ((!$FromAddr) -or (!$ToAddr) -or (!$SmtpSvr))
		{
			Throw "The following parameters are required when setting the 'SendMail' parameter: FromAddr, ToAddr, SmtpSvr"
		}
	}
#endregion Script Initialization

#region Global Functions
	#region VMware vSphere PowerCLI Environment
		#Set-PowerCLIConfiguration -InvalidCertificateAction Prompt
		# Returns the path (with trailing backslash) to the directory where PowerCLI is installed.
		Function Get-InstallPath
		{
			# 32-bit OS 
			$regKeys = Get-ItemProperty -Path "HKLM:\SOFTWARE\VMware, Inc.\VMware vSphere PowerCLI" -ErrorAction SilentlyContinue

			# 64-bit OS
			If($regKeys -eq $null)
			{
				$regKeys = Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\VMware, Inc.\VMware vSphere PowerCLI" -ErrorAction SilentlyContinue
			}

			Return $regKeys.InstallPath
		}

		Function LoadVMwareSnapins()
		{
			$VMwareSnapinList = @("VMware.VimAutomation.Core", "VMware.VimAutomation.License", "VMware.DeployAutomation", "VMware.ImageBuilder", "VMware.VimAutomation.Cloud")

			$LoadedSnapins = Get-PSSnapin -Name $VMwareSnapinList -ErrorAction SilentlyContinue | ForEach-Object {
				$_.Name
			}

			$RegisteredSnapins = Get-PSSnapin -Name $VMwareSnapinList -Registered -ErrorAction SilentlyContinue  | ForEach-Object {
				$_.Name
			}

			$NotLoadedSnapins = $RegisteredSnapins | Where-Object {
				$LoadedSnapins -notcontains $_
			}

			ForEach ($Snapin in $RegisteredSnapins)
			{
				If ($LoadedSnapins -notcontains $Snapin)
				{
					Add-PSSnapin $Snapin
				}

				# Load the Intitialize-<snapin_name_with_underscores>.ps1 file
				# File lookup is based on install path instead of script folder because the PowerCLI
				# shortuts load this script through dot-sourcing and script path is not available.
				$filePath = "{0}Scripts\Initialize-{1}.ps1" -f (Get-InstallPath), $Snapin.ToString().Replace(".", "_")

				If (Test-Path $filePath)
				{
					& $filePath
				}
			}
		}
	#endregion VMware vSphere PowerCLI Environment

	# This function allows you to pause the script execution and review what is output on the console. I wrote it for help in Debugging my scripts.
	Function fn_PauseScript
	{
		Write-Host -ForegroundColor Magenta "Press any key to continue..."
		$x = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyUp")
		Write-Host ""
	}

	# This Function allows the user to send an email from the output of the script.
	Function fn_SendEmail
	{
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
			)][array]$AddrTo
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

		# Setup splatting hashtable to hold the values to be passed to the 'Send-MailMessage' cmdlet.
		$SendMailMsg = @{}

		$SendMailMsg.Add("SmtpServer", $SmtpHost)
		$SendMailMsg.Add("From", $AddrFrom)
		$SendMailMsg.Add("BodyAsHtml", $true)
		$SendMailMsg.Add("Body", $MsgBody)
		$SendMailMsg.Add("To", $AddrTo)

		[string]$RptDate = (Get-Date -Format 'ddd dd-MMM-yyyy').ToString()	# Example Output: Fri 23-Nov-2012
		[string]$MsgSubject	= "[REPORT] VMware Snapshots Report for $($RptDate)"
		$SendMailMsg.Add("Subject", $MsgSubject)

		If ($Attachments)
		{
			$SendMailMsg.Add("Attachments", $Attachments)
		}

		Send-MailMessage @SendMailMsg
	}

	Function fn_GenerateHtml
	{
		Param($RptVMs)

		# HTML Header
		$HtmlReport = "<!DOCTYPE html PUBLIC $([char]34)-//W3C//DTD XHTML 1.1//EN$([char]34) $([char]34)http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)<html xmlns=$([char]34)http://www.w3.org/1999/xhtml$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)<head>"
		$HtmlReport += "$([char]13)$([char]10)	<meta http-equiv=$([char]34)Content-Type$([char]34) content=$([char]34)text/html; charset=iso-8859-1$([char]34) />"
		$HtmlReport += "$([char]13)$([char]10)	<title>SquareTwo Financial - VMware Snapshots Report</title>"
		$HtmlReport += "$([char]13)$([char]10)	<style type=$([char]34)text/css$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)	<!--"
		$HtmlReport += "$([char]13)$([char]10)		body"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			font-family: Tahoma;"
		$HtmlReport += "$([char]13)$([char]10)			font-size: 11px;"
		$HtmlReport += "$([char]13)$([char]10)			margin-left: 0px;"
		$HtmlReport += "$([char]13)$([char]10)			margin-top: 0px;"
		$HtmlReport += "$([char]13)$([char]10)			margin-right: 0px;"
		$HtmlReport += "$([char]13)$([char]10)			margin-bottom: 0px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		div.ReportHeaderLeft"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			float: left;"
		$HtmlReport += "$([char]13)$([char]10)			width: 215px;"
		$HtmlReport += "$([char]13)$([char]10)			height: 115px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		div.ReportTitle"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			float: right;"
		$HtmlReport += "$([char]13)$([char]10)			width: 660px;"
		#$HtmlReport += "$([char]13)$([char]10)			color: #FF0000;"
		$HtmlReport += "$([char]13)$([char]10)			color: #df202e;"	# Should be the Red from the SquareTwo Financial Logo.
		$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
		$HtmlReport += "$([char]13)$([char]10)			font-size: 22px;"
		$HtmlReport += "$([char]13)$([char]10)			font-weight: bold;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		div.ReportDate"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			float: right;"
		$HtmlReport += "$([char]13)$([char]10)			width: 660px;"
		#$HtmlReport += "$([char]13)$([char]10)			color: #FF0000;"
		$HtmlReport += "$([char]13)$([char]10)			color: #df202e;"	# Should be the Red from the SquareTwo Financial Logo.
		$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
		$HtmlReport += "$([char]13)$([char]10)			font-size: 9px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		table"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			width: 895px;"
		$HtmlReport += "$([char]13)$([char]10)			border-width: 0;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		tr.TitleHdrRow"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			color: #FFFFFF;"
		$HtmlReport += "$([char]13)$([char]10)			background-color: #0077D4;"	# Blueish Background
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		td"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			padding-top: 0px;"
		$HtmlReport += "$([char]13)$([char]10)			padding-bottom: 0px;"
		$HtmlReport += "$([char]13)$([char]10)			padding-right: 3px;"
		$HtmlReport += "$([char]13)$([char]10)			padding-left: 3px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		td.ReportHeaderLeft"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			width: 215px;"
		$HtmlReport += "$([char]13)$([char]10)			height: 115px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		td.ReportTitle"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			vertical-align: middle;"
		$HtmlReport += "$([char]13)$([char]10)			height: 90px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		td.ReportDate"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			vertical-align: bottom;"
		$HtmlReport += "$([char]13)$([char]10)			height: 25px;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		td.GreenCellBold"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			text-align: center;"
		$HtmlReport += "$([char]13)$([char]10)			background-color: #00FF00;"
		$HtmlReport += "$([char]13)$([char]10)			font-weight: bold;"
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		tr.AltRowOn"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			background-color: #DDDDDD;"	# Light Gray Background
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)		tr.AltRowOff"
		$HtmlReport += "$([char]13)$([char]10)		{"
		$HtmlReport += "$([char]13)$([char]10)			background-color: #FFFFFF;"	# White Background
		$HtmlReport += "$([char]13)$([char]10)		}"
		$HtmlReport += "$([char]13)$([char]10)	-->"
		$HtmlReport += "$([char]13)$([char]10)	</style>"
		$HtmlReport += "$([char]13)$([char]10)</head>"
		$HtmlReport += "$([char]13)$([char]10)<body>"
		$HtmlReport += "$([char]13)$([char]10)<table border=$([char]34)0$([char]34) cellpadding=$([char]34)0$([char]34) cellspacing=$([char]34)0$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)1$([char]34) rowspan=$([char]34)2$([char]34) class=$([char]34)ReportHeaderLeft$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)			<img src=$([char]34)http://www.squaretwofinancial.com/wp-content/themes/square2financial/images/logo.gif$([char]34) alt=$([char]34)SquareTwo Financial Logo$([char]34) width=$([char]34)200$([char]34) height=$([char]34)115$([char]34) />"
		$HtmlReport += "$([char]13)$([char]10)		</td>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)4$([char]34) class=$([char]34)ReportTitle$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)			<div class=$([char]34)ReportTitle$([char]34)>VMware Snaphots Report</div>"
		$HtmlReport += "$([char]13)$([char]10)		</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)	<tr>"
		$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)4$([char]34) class=$([char]34)ReportDate$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)			<div class=$([char]34)ReportDate$([char]34)>Generated $((Get-Date -Format 'MMMM dd, yyyy HH:mm').ToString())</div>"
		$HtmlReport += "$([char]13)$([char]10)		</td>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"
		$HtmlReport += "$([char]13)$([char]10)	<tr class=$([char]34)TitleHdrRow$([char]34)>"
		$HtmlReport += "$([char]13)$([char]10)		<th>ESXi Host</th>"
		$HtmlReport += "$([char]13)$([char]10)		<th width=$([char]34)100$([char]34)>VM Name</th>"
		$HtmlReport += "$([char]13)$([char]10)		<th>Snapshot Name</th>"
		$HtmlReport += "$([char]13)$([char]10)		<th>Description</th>"
		$HtmlReport += "$([char]13)$([char]10)		<th width=$([char]34)100$([char]34)>Date Created</th>"
		$HtmlReport += "$([char]13)$([char]10)	</tr>"

		$AlternateRow = 0
		If (!$RptVMs)
		{
			$HtmlReport += "$([char]13)$([char]10)	<tr>"
			$HtmlReport += "$([char]13)$([char]10)		<td colspan=$([char]34)7$([char]34) class=$([char]34)GreenCellBold$([char]34)>There are <u>NO VMware Snaphots</u> to report at this time.</td>"
			$HtmlReport += "$([char]13)$([char]10)	</tr>"
		}
		Else
		{
			ForEach ($Snapshot in $RptVMs)
			{
				$HtmlReport += "$([char]13)$([char]10)	<tr"
				If ($AlternateRow)
				{
					$HtmlReport += " class=$([char]34)AltRowOn$([char]34)"
					$AlternateRow = 0
				}
				Else
				{
					$HtmlReport += " class=$([char]34)AltRowOff$([char]34)"
					$AlternateRow = 1
				}
				$HtmlReport += ">"
				$HtmlReport += "$([char]13)$([char]10)		<td>$($Snapshot.VMHost)</td>"
				$HtmlReport += "$([char]13)$([char]10)		<td>$($Snapshot.VM)</td>"
				$HtmlReport += "$([char]13)$([char]10)		<td>$($Snapshot.Name)</td>"
				$HtmlReport += "$([char]13)$([char]10)		<td>$($Snapshot.Description)</td>"
				$HtmlReport += "$([char]13)$([char]10)		<td>$($Snapshot.Created)</td>"
				$HtmlReport += "$([char]13)$([char]10)	</tr>"
			}
		} # End ForEach Loop

		$HtmlReport += "$([char]13)$([char]10)</table>"
		$HtmlReport += "$([char]13)$([char]10)</body>"
		$HtmlReport += "$([char]13)$([char]10)</html>"
		$HtmlReport += "$([char]13)$([char]10)"

		# Write out to the HTML file...
		$HtmlReport | Out-File $HtmlRptFile

		# Return the Report...
		Return $HtmlReport
	}

#	# This Function generates a nice HTML output that uses CSS for style formatting.
#	Function Generate-Report
#	{
#		Write-Output "<html>
#		<head>
#		<title>VMware Snaphots Report</title>
#		<style type=""text/css"">
#		.Error {
#			color: #FF0000;
#			font-weight: bold;
#		}
#		.Title {
#			background: #0077D4;
#			color: #FFFFFF;
#			text-align: center;
#			font-weight: bold;
#		}
#		.Normal {
#		}
#		</style>
#		</head>
#		<body>
#		<table>
#			<tr class="" Title "">
#				<td colspan=""5""> VMware Snaphot Report </td>
#			</tr>
#			<tr class=" Title ">
#				<td>-------------------VM Host------------------</td>
#				<td>-----VM Name-----</td>
#				<td>----------------Snapshot Name---------------</td>
#				<td>-----------------Description----------------</td>
#				<td>-------Date Created-------</td>
#			</tr>"
#
#		ForEach ($Snapshot in $Report)
#		{
#			Write-Output "<tr>
#				<td>$($Snapshot.VMHost)</td>
#				<td>$($Snapshot.VM)</td>
#				<td>$($Snapshot.Name)</td>
#				<td>$($Snapshot.Description)</td>
#				<td>$($Snapshot.Created)</td>
#			</tr>"
#		}
#		Write-Output "</table>
#		</body>
#		</html>"
#	}
#endregion Global Functions

#region Main Script
#################################################################################################################
#	#	#	Main Report Execution Starts Here!																	#
#################################################################################################################
	LoadVMwareSnapins
	
#################################################################################################################
#	#	#	Date and Time Formatting																			#
#	#	#	http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo%28VS.85%29.aspx		#
#	#	#	Formatting chosen for the date string is as follows:												#
#	#	#	Date Display  |  -Format																			#
#	#	#	------------  |  ------------																		#
#	#	#	2012.Jan.25   |  yyyy.MMM.dd																		#
#	#	#	2012.01.25    |  yyyy.MM.dd																			#
#	#	#	Mon			  |  ddd																				#
#	#	#	------------  |  ------------																		#
#	#	#	Time Display  |  -Format																			#
#	#	#	------------  |  ------------																		#
#	#	#	22:00         |  HH:mm																				#
#################################################################################################################
	[string]$RptFileDate = (Get-Date -Format 'yyyy.MMM.dd.ddd').ToString()	# Example Output: 2012.11.23.Wed

	# Example of the name:  Get-VMwareSnaphots_2012.11.23.Wed
	[string]$BaseRptName = "$($PtlRptName)_$($RptFileDate)"
	[string]$Ext1 = "html"
	[string]$HtmlName = "$($BaseRptName).$($Ext1)"

	If ($CreateCsv)
	{
		[string]$Ext2 = "csv"
		[string]$CsvName  = "$($BaseRptName).$($Ext2)"
	}

	If ($ScrDebug)
	{
		If (!(Test-Path -Path ${env:userprofile}\Desktop\Reports))
		{
			New-Item -ItemType Directory ${env:userprofile}\Desktop\Reports | Out-Null
		}

		$HtmlRptFile = "$(${env:userprofile})\Desktop\Reports\$HtmlName"

		If ($CreateCsv)
		{
			$CsvRptFile = "$(${env:userprofile})\Desktop\Reports\$CsvName"
		}
	}
	Else
	{
		If (!(Test-Path -Path .\Reports))
		{
			New-Item -ItemType Directory .\Reports | Out-Null
		}

		$HtmlRptFile = ".\Reports\$HtmlName"

		If ($CreateCsv)
		{
			$CsvRptFile = ".\Reports\$CsvName"
		}
	}

	## Login details for standalone ESXi servers (not required if using VirtualCenter)
	#$ESXiUserName = 'ESXiUser'
	#$ESXiPassword = 'ESXiPassword' #Change to the root password you set for you ESXi server

	#List of servers including Virtual Center Server.  The account this script will run as will need at least Read-Only access to Virtual Center
	#Chance to DNS Names/IP addresses of your ESXi servers or Virtual Center Server. Comma separated.
	$SvrList = "vc5-prc-cp.corp.collectamerica.com"

	#Initialise Array
	$Report = @()

	#Get snapshots from all servers
	ForEach ($Svr in $SvrList)
	{
		# Check is server is a Virtual Center Server and connect with current user
		If ($Svr -eq "vc5-prc-cp.corp.collectamerica.com")
		{
			Connect-VIServer $Svr
		}
		#Else
		#{
			## Use specific login details for the rest of servers in $SvrList
			## Uncomment this line if you use ESXi hosts rather than VirtualCenter
		#	Connect-VIServer $Svr -user $ESXiUserName -password $ESXiPassword
		#}

		Get-VM | Get-Snapshot | ForEach-Object {
			$Snap = {} | Select-Object VMHost,VM,Name,Description,Created
			$Snap.VMHost = $_.VM.VMHost.Name
			$Snap.VM = $_.VM.Name
			$Snap.Name = $_.Name
			$Snap.Description = $_.Description
			$Snap.Created = $_.Created
			$Report += $Snap
		}
	}

	$Report = $Report | Sort-Object -Property VM	#,VMHost

	If ($CreateCsv)
	{
		$Report | Export-CSV -Path $CsvRptFile -NoTypeInformation
	}

	## Generate the report and email it as a HTML body of an email
	#Generate-Report | Out-File -FilePath "${env:userprofile}\Desktop\Reports\VmwareSnapshots_$($RptFileDate).html"
	#
	#If ($Report -ne "")
	#{
	#	$SmtpClient = New-Object System.Net.Mail.SmtpClient
	#
	#	# Change to a SMTP server in your environment
	#	$SmtpClient.host = "mail.squaretwofinancial.com"
	#	$MailMessage = New-Object System.Net.Mail.MailMessage
	#
	#	# Change to email address you want emails to be coming from
	#	$MailMessage.From = "server-alerts@squaretwofinancial.com"
	#
	#	# Change to email address you would like to receive emails.
	#	$MailMessage.To.Add("server-alerts@squaretwofinancial.com")
	#	$MailMessage.IsBodyHtml = 1
	#	$MailMessage.Subject = "Vmware Snapshots"
	#	$MailMessage.Body = Generate-Report
	#	$SmtpClient.Send($MailMessage)
	#}

	# Generate the report and email it as an HTML body of an email
	$RptBody = fn_GenerateHtml -RptVMs $Report

	If ($SendMail)
	{
		$objSendMail = @{}

		$objSendMail.Add("MsgBody", $RptBody)
		$objSendMail.Add("AddrFrom", $FromAddr)
		$objSendMail.Add("AddrTo", $ToAddr)
		$objSendMail.Add("SmtpHost", $SmtpSvr)

		If ($CreateCsv)
		{
			#$FileAttachments = @($CsvRptFile, $HtmlRptFile)
			$FileAttachments = @($CsvRptFile)
			$objSendMail.Add("Attachments", $FileAttachments)
		}

		fn_SendEmail @objSendMail
	}
#endregion Main Script
