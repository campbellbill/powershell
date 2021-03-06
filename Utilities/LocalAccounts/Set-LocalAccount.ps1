##################################################################################
#
#  Script name	: Set-LocalAccount.ps1
#  Authors		: niklas.goude@zipper.se, Brian Marsh
#  Homepage		: www.powershell.nu
#
##################################################################################

Param(
	[string]$UserName
	, [string]$FullName
	, [string]$Password
	, [switch]$Add
	, [switch]$Remove
	, [switch]$ResetPassword
	, [switch]$help
	, [string]$computername
)

Function GetHelp()
{
	$HelpText = @"
	DESCRIPTION:

	NAME: Set-LocalAccount.ps1
	Adds or Removes a Local Account

	PARAMETERS:
	-UserName        Name of the User to Add or Remove (Required)
	-Password        Sets Users Password (optional)
	-Add             Adds Local User (Optional)
	-Remove          Removes Local User (Optional)
	-ResetPassword   Resets Local User Password (Optional)
	-help            Prints the HelpFile (Optional)

	SYNTAX:
	.\Set-LocalAccount.ps1 -UserName nika -Password Password1 -Add
	Adds Local User nika and sets Password to Password1

	.\Set-LocalAccount.ps1 -UserName nika -Remove
	Removes Local User nika

	.\Set-LocalAccount.ps1 -UserName nika -Password Password1 -ResetPassword
	Sets Local User nika's Password to Password1

	.\Set-LocalAdmin.ps1 -help
	Displays the helptext
"@

	$HelpText
}

Function AddRemove-LocalAccount ([string]$UserName, [string]$FullName, [string]$Password, [switch]$Add, [switch]$Remove, [switch]$ResetPassword, [string]$computerName)
{
	If ($Add)
	{
		[string]$ConnectionString = "WinNT://$computerName,computer"
		$ADSI = [ADSI]$ConnectionString
		$User = $ADSI.Create("user",$UserName)
		$User.SetPassword($Password)
		#echo "-------------DEBUGGING---------------"
		#echo "Connection String:  $connectionstring"
		#echo "Username:           $username"
		#echo "Password:           $password"
		#echo "-------------------------------------"
		$User.SetInfo()

		([ADSI]"WinNT://$computerName/Administrators,group").Add("WinNT://$UserName")
		$User.Put("Description","IDI User Account")
		$User.SetInfo()
		$User.Put("FullName",$FullName)
		$User.Put("PasswordExpired", 1)
		$User.SetInfo()
	}

	If ($Remove)
	{
		[string]$ConnectionString = "WinNT://$computerName,computer"
		$ADSI = [ADSI]$ConnectionString
		$ADSI.Delete("user",$UserName)
	}

	If ($ResetPassword)
	{
		[string]$ConnectionString = "WinNT://" + $ComputerName + "/" + $UserName + ",user"
		$Account = [ADSI]$ConnectionString
		$Account.psbase.invoke("SetPassword", $Password)
	}
}

If ($help)
{
	GetHelp
	Continue
}

If ($UserName -and $Password -and $Add -and !$ResetPassword -and !$FullName)
{
	AddRemove-LocalAccount -UserName $UserName -Password $Password -Add -computerName $computerName
}

If ($UserName -and $FullName -and $Password -and $Add -and !$ResetPassword)
{
	AddRemove-LocalAccount -UserName $UserName -FullName $FullName -Password $Password -Add -computerName $computerName
}

If ($UserName -and $Password -and $ResetPassword)
{
	AddRemove-LocalAccount -UserName $UserName -Password $Password -ResetPassword -computerName $computerName
}

If ($UserName -and $Remove)
{
	AddRemove-LocalAccount -UserName $UserName -Remove -computerName $computerName
}
