#---------------------
# Title:        Add-LocalUser.ps1
# Author:       Brian Marsh
#               Info Directions, Inc.
# Description:  Powershell script designed to facilitate the creation of local
#               users. This is the natural progression from BCLU, and from Find-
#               Terri. Unlike previous attempts, this version dynamically
#               generates secure passwords, and incorporates email Functionality
#               to ensure only the end user knows their temporary password.
# 
# Requirements: Powershell v. 2.0, local email SMTP Relay
# 
# Version 1.0
#  
# Changelog
# 1.0           Added email address generation, added comments, dynamic password
#               generation, changed name of script from doit.ps1 to
#               Add-LocalUser.ps1, cleaned up code
# 0.3           Added debugging
# 0.2           Basic email Functionality
# 0.1           Initial Program
# 
#---------------------

Param(
	$usr,       # Username without prefix, following standard
	$FullName,  # Full Name of the user
	$cust,      # Target customer, i.e. Widgets Inc - Test server
	$ip         # IP/hostname of the server
)

# Function gen-password
# Description:  Generates a secure password of specified length (Uppercase,
#               lowercase, symbols, numbers). Taken with slight alterations from
#               http://bit.ly/dsj5nU (Dmitry's PowerBlog)
# Parameters:   int maxChars - optional (Defaults to 8), number of characters in
#               password.
Function gen-password
{
	Param(
		[int]$maxChars = 8
	)

	$NewPassword = ""    # Null out previous password (if any)
	$rand = New-Object System.Random      # Create random object

	# For each iteration, up to maxChars append the random character
	0..$maxChars | ForEach-Object { $NewPassword = $NewPassword + [char]$rand.next(33,127) }

	return $NewPassword   #Return the new password
}# End Function gen-password
 
# Function find-user
# Description:  Searches specified computer for specified user, returns true or
#               false depending on if the user was found
# Parameters:   string ip - IP/hostname of the server
#               string findme - the user to find
Function find-user
{
	Param(
		[string]$ip
		, [string]$findMe
	)

	# Initally we haven't found anything
	$found = $false

	#Setup connection info
	$computer = [ADSI]("WinNT://" + $ip + ",computer")
	$Users = $computer.psbase.children | Where-Object { $_.psbase.schemaclassname -eq "User" }

	# Search for the user in the list of current users
	ForEach ($member in $Users.psbase.syncroot)
	{
		If ($member.name -match $usr)
		{
			$found = $true
		}
	}

	return $found
}

# Initialize variables
$addr = ""
$email = ""

# Generate password
$password = gen-Password

# Set email body
$body = "
$FullName,
`r
You are recieving this automated email to notify you that a Windows account has been created on the $cust server ($ip).
`r
Please take note of your username and password below. You will be required to change your password when you first log in.
`r
Username: idi_$usr
Password: $NewPassword
`r
If you have any questions, please let IS know."

# Debugging stuff
#echo "------Debugging Add-LocalUser.ps1------"
#echo "User: $usr"
#echo "FullName: $FullName"
#echo "Customer: $cust"
#echo "Ip: $ip"
#echo "---------------------------------------"

# Check to see if user exists already, if it does quit with error message
$found = find-usr -ip $ip -findme "idi_$usr"
If ($found)
{
  echo "User found, cannot add"
  exit
}

# Call the ps script that creates a user
./set-localaccount.ps1 -username "idi_$usr" -FullName $FullName -password "$NewPassword" -add -computername $ip

# Generate email address
$email = $fullname.tolower()
$break = $email.indexof(" ")+1
$addr = $email.substring(0,1)+$email.substring($break)

# Send email with germaine information
./email.ps1 -emailTo "$addr@infodir.com" -subject "$cust Windows Account info" -body $body
