
Function fn_ValidateUsernameAndPassword()
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	#region Function Parameters
		Param(
			[string]$VunpServerName		# The name of the domain or server for Domain context types, the machine name for Machine context types, or the name of the server and port hosting the ApplicationDirectory instance.
			, [string]$VunpUserName		# The username to validate.
			, [string]$VunpPasswd		# The decrypted password associated with the Username being validated.
			, $VunpSecPasswd			# The SecureString password associated with the Username being validated.
			, [switch]$VunpCheckDomain	# Switch to specify that an 'Active Directory Domain' is the context type to check.
		)
	#endregion Function Parameters
	## Funtion Usage:
	#	fn_ValidateUsernameAndPassword -VunpServerName '' -VunpUserName '' -VunpPasswd '' -VunpSecPasswd 'SecureString Object' -VunpCheckDomain

	Add-Type -AssemblyName System.DirectoryServices.AccountManagement

	If ($VunpCheckDomain)
	{
		[string]$VunpChk = 'Domain'
	}
	Else
	{
		[string]$VunpChk = 'Machine'
	}

	Write-Host -ForegroundColor Green "Checking Username $([char]34)$($VunpUserName)$([char]34) and associated password against :-> $([char]34)$($VunpServerName)$([char]34)."

	$VunpPrincipalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($VunpChk, $VunpServerName)

	# Creates the connections to the server and returns a Boolean value that indicates whether the specified username and password are valid. 
	$VunpRet = $VunpPrincipalContext.ValidateCredentials($VunpUserName, $VunpPasswd)

	If (!$? -or !$VunpRet)
	{
		Write-Host -ForegroundColor Magenta "Username or password was NOT validated correctly using the Secure String password on :-> $([char]34)$($VunpServerName)$([char]34)."

		$VunpRet = $VunpPrincipalContext.ValidateCredentials($VunpUserName, $VunpSecPasswd)
		If (!$? -or !$VunpRet)
		{
			Write-Host -ForegroundColor Magenta "Username or password was NOT validated correctly using the Secure String password on :-> $([char]34)$($VunpServerName)$([char]34)."
		}
	}
	Else
	{
		Write-Host -ForegroundColor Green "Username or password was validated correctly on :-> $([char]34)$($VunpServerName)$([char]34)."
	}
}

$ServerName = 'wrk.pad.pearsoncmg.com'
$UserName = 'xl_deploy'
$Passwd = 'm0nd7tron'

$SecPasswd = ConvertTo-SecureString $Passwd -AsPlainText -Force

fn_ValidateUsernameAndPassword -VunpServerName $ServerName -VunpUserName $UserName -VunpPasswd $Passwd -VunpSecPasswd $SecPasswd -VunpCheckDomain
