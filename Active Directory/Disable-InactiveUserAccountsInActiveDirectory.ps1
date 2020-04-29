<#
	Disable - Inactive User Accounts In Active Directory
	Script Name:	Disable-InactiveUserAccountsInActiveDirectory.ps1
	Found At:		http://blogs.technet.com/b/bahramr/archive/2008/01/25/powershell-script-to-disable-inactive-accounts-in-active-directory.aspx
	Written:		Jan 25, 2008
	Added Script:	19 June 2013
	Modified by:	Bill Campbell
	Last Modified:	2013.June.19

	Version:		2013.06.19.01
	Version Notes:	Version format is a date taking the following format: YYYY.MM.DD.RR		- where RR is the revision/save count for the day modified.
	Version Exmpl:	If this is the 6th revision/save on January 13 2012 then RR would be "06" and the version number will be formatted as follows: 2012.01.23.06

	Purpose:		Query with Powershell for all Inactive User Accounts in Active Directory and disable those that have not been used for more than the number of days specified.

	Notes:			Use the next line to change to your user profile "Desktop" folder:
					cd ${env:userprofile}\Desktop
					echo ${env:userprofile}
					echo ${env:computername}
					echo ${env:username}
					cd C:\Pearson_Docs

	.SYNOPSIS
		#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
		#! IMPORTANT NOTE:																							 !#
		#! 																											 !#
		#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#

	.PARAMETER NbDays [int]
		Please specify the maximum number of days of inactivity allowed. Users who have not logged on for longer than the number of days specified will get disabled.

	.PARAMETER Subtree [string]
    	Please specify the DN of the container under which inactive accounts should be queried from.

	.PARAMETER ScrDebug [switch]
		The ScrDebug switch turns the debugging code in the script on or off ($true/$false). Defaults to FALSE.

	.EXAMPLES
		.\Disable-InactiveUserAccountsInActiveDirectory.ps1 -NbDays 90 -Subtree ""
		.\Disable-InactiveUserAccountsInActiveDirectory.ps1 -NbDays 90 -Subtree "" -ScrDebug
		. ${env:userprofile}\Dropbox\Scripting\PowerShell_Scripts\Active_Directory\Disable-InactiveUserAccountsInActiveDirectory.ps1 -NbDays 90 -Subtree ""
		. ${env:userprofile}\Dropbox\Scripting\PowerShell_Scripts\Active_Directory\Disable-InactiveUserAccountsInActiveDirectory.ps1 -NbDays 90 -Subtree "" -ScrDebug
		D:\Scripts\Disable-InactiveUserAccountsInActiveDirectory\Disable-InactiveUserAccountsInActiveDirectory.ps1 -NbDays 90 -Subtree ""
		D:\Scripts\Disable-InactiveUserAccountsInActiveDirectory\Disable-InactiveUserAccountsInActiveDirectory.ps1 -NbDays 90 -Subtree "" -ScrDebug

	-- Script Changes --
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
	#!   THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE   !#
	#! RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. !#
	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#

	Changes for 19.June.2013
		- Initial Script and Debugging.
#>

#region Script Initialization
	# Read the input parameters $Subtree and $NbDays
	Param(
		[Parameter(
			Position			= 0
			, Mandatory			= $true
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Please specify the maximum number of days of inactivity allowed. Users who have not logged on for longer than the number of days specified will get disabled.'
		)][int]$NbDays		#= $(throw Write-Host "Please specify the maximum number of days of inactivity allowed. Users who have not logged on for longer than the number of days specified will get disabled." -Foregroundcolor Red)
		, [Parameter(
			Position			= 1
			, Mandatory			= $true
			, ValueFromPipeline	= $false
			, HelpMessage		= 'Please specify the DN of the container under which inactive accounts should be queried from.'
		)][string]$Subtree		#= $(throw Write-Host "Please specify the DN of the container under which inactive accounts should be queried from." -Foregroundcolor Red)
	)

	# Get the current date
	$currentDate = [System.DateTime]::Now

	# Convert the local time to UTC format because all dates are expressed in UTC (GMT) format in Active Directory
	$currentDateUtc = $currentDate.ToUniversalTime()

	# Set the LDAP URL to the container DN specified on the command line
	$LdapURL = "LDAP://" + $Subtree

	# Initialize a DirectorySearcher object
	$searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]$LdapURL)

	# Set the attributes that you want to be returned from AD
	$searcher.PropertiesToLoad.Add("displayName") >$null
	$searcher.PropertiesToLoad.Add("sAMAccountName") >$null
	$searcher.PropertiesToLoad.Add("lastLogonTimeStamp") >$null

	# Calculate the time stamp in Large Integer/Interval format using the $NbDays specified on the command line
	$lastLogonTimeStampLimit = $currentDateUtc.AddDays(- $NbDays)
	$lastLogonIntervalLimit = $lastLogonTimeStampLimit.ToFileTime()
#endregion Script Initialization

Write-Host -ForegroundColor Yellow "Looking for all users that have not logged on since "$lastLogonTimeStampLimit" ("$lastLogonIntervalLimit")"
$searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(lastLogonTimeStamp<=" + $lastLogonIntervalLimit + "))"

# Run the LDAP Search request against AD
$users = $searcher.FindAll()

If ($users.Count -eq 0)
{
       Write-Host -ForegroundColor Green "`tNo user accounts need to be disabled.”
}
Else
{
       ForEach ($user in $users)
       {
              # Read the user properties
              [string]$adsPath = $user.Properties.adspath
              [string]$displayName = $user.Properties.displayname
              [string]$samAccountName = $user.Properties.samaccountname
              [string]$lastLogonInterval = $user.Properties.lastlogontimestamp

              # Convert the date and time to the local time zone
              $lastLogon = [System.DateTime]::FromFileTime($lastLogonInterval)

              # Disable the user
              $account=[ADSI]$adsPath
              $account.psbase.invokeset("AccountDisabled", "True")
              $account.setinfo()

              Write-Host -ForegroundColor Magenta "`tDisabled user "$displayName" ("$samAccountName") who last logged on "$lastLogon" ("$lastLogonInterval")"          
       }
}
