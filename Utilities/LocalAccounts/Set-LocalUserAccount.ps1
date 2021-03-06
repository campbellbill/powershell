#Requires -Version 2.0

Function Set-LocalUserAccount
{
	<#
		.SYNOPSIS
			Enable or disable a local user account.

		.DESCRIPTION
			This command will allow you to set the password of a local user account as well
			as enable or disable it. By default, this command will not write anything to
			the pipeline unless you use -Passthru.  You must run this under credentials 
			that have administrator rights on the remote computer.

		.PARAMETER ComputerName 
			The name of the computer to connect to. This parameter has an alias of CN.
		.PARAMETER UserName 
			The name of the local user account on the computer.
		.PARAMETER Password 
			The new password to set. This parameter has an alias of PWD.
		.PARAMETER Status 
			Enable or disable the local user account.
		.PARAMETER Passthru
			Write the user account object to the pipeline

		.EXAMPLE
			PS C:\> Set-LocalUserAccount SERVER01,SERVER02 DBAdmin -status disable

			Disable the local user account DBAdmin on SERVER01 and SERVER02

		.EXAMPLE
			PS C:\> get-content c:\work\computers.txt | set-localuseraccount LocalAdmin -password "^Crx33t7A"

			Sets the password for account LocalAdmin on all computers in computers.txt

		.NOTES
			Version: 1.0
			Author : Jeff Hicks (@JeffHicks)

		Learn more:
			PowerShell in Depth: An Administrator's Guide (http://www.manning.com/jones2/)
			PowerShell Deep Dives (http://manning.com/hicks/)
			Learn PowerShell 3 in a Month of Lunches (http://manning.com/jones3/)
			Learn PowerShell Toolmaking in a Month of Lunches (http://manning.com/jones4/)


			****************************************************************
			* DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
			* THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
			* YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
			* DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
			****************************************************************

		.INPUTS
			String

		.OUTPUTS
			None or System.DirectoryServices.DirectoryEntry
	#>

	[CmdletBinding(SupportsShouldProcess=$true)]
	Param (
		[Parameter(
			Position = 0,
			ValueFromPipeline = $true,
			ValueFromPipelineByPropertyName = $true
		)]
		[ValidateNotNullorEmpty()]
		[Alias("cn")]
		[string[]]$ComputerName = ${Env:COMPUTERNAME}, 
		[Parameter(
			Position = 1,
			Mandatory = $true,
			HelpMessage = "What is the name of the local user account?",
			ValueFromPipelineByPropertyName = $true
		)]
		[ValidateNotNullorEmpty()]
		[string]$UserName, 
		[Parameter(
			ValueFromPipelineByPropertyName = $true
		)]
		[Alias("pwd")]
		[string]$Password, 
		[ValidateSet("Enable", "Disable")]
		[string]$Status = "Enable",
		[switch]$Passthru
	)

	Begin
	{
		Write-Verbose "Starting $($MyInvocation.MyCommand)"
		#define a constant to disable or enable an account
		New-Variable ADS_UF_ACCOUNTDISABLE 0x0002 -Option Constant

		Write-Verbose "Setting local user account $username"
	} #begin
	Process
	{
		foreach ($computer in $computername)
		{
			Write-Verbose "Connecting to $computer"
			Write-Verbose "Getting user account"

			$Account = [ADSI]"WinNT://$computer/$username,user"

			#validate the user account was found
			if (-not $Account.path)
			{
				Write-Warning "Failed to find $username on $computername"
				#bail out
				Return
			}

			#Get current enabled/disabled status
			if ($Account.userflags.value -band $ADS_UF_ACCOUNTDISABLE)
			{
				$Enabled = $false
			}
			else
			{
				$Enabled = $true
			}

			Write-verbose "Account enabled is $Enabled"

			if ($enabled -and ($Status -eq "Disable"))
			{
				Write-Verbose "disabling the account"
				$value = $Account.userflags.value -bor $ADS_UF_ACCOUNTDISABLE
				$Account.put("userflags", $value)
			}
			elseif ((-not $enabled) -and ($Status -eq "Enable"))
			{
				Write-Verbose "Enabling the account"
				$value = $Account.userflags.value -bxor $ADS_UF_ACCOUNTDISABLE
				$Account.put("userflags", $value)
			}
			else
			{
				#account is already in the desired state
				Write-Verbose "No change necessary"
			}

			if ($Password)
			{
				Write-Verbose "Setting acccount password"
				$Account.SetPassword($Password)
			}

			#Whatif
			if ($PSCmdlet.ShouldProcess("$computer\$username"))
			{
				Write-Verbose "Committing changes"
				$Account.SetInfo()
			}

			if ($Passthru)
			{
				Write-Verbose "Passing object to the pipeline"
				$Account
			}
		} #foreach
	} #process
	End
	{
		Write-Verbose "Ending $($MyInvocation.MyCommand)"
	} #end
} #end Set-LocalUserAccount function
