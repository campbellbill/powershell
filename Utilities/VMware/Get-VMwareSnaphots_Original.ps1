@"
===============================================================================
Title:				Get-VmwareSnaphots.ps1
Description:		List snapshots on VM's managed by Virtual Center.
Requirements:		Windows Powershell and the VI Toolkit
Usage:				.\Get-VmwareSnaphots.ps1
Found At:			http://community.spiceworks.com/scripts/show/882-email-list-of-open-snapshots
					Minor formatting and layout changes have been made from the original for clarity.
Last Modified By:	Bill Campbell
Last Modified On:	2012-Nov-23
===============================================================================
"@

#Global Functions
#This function generates a nice HTML output that uses CSS for style formatting.
Function Generate-Report
{
	Write-Output "<html>
	<head>
	<title></title>
	<style type=""text/css"">
	.Error {color:#FF0000;font-weight: bold;}
	.Title {background: #0077D4;color: #FFFFFF;text-align:center;font-weight: bold;}
	.Normal {}
	</style>
	</head>
	<body>
	<table>
		<tr class="" Title "">
			<td colspan=""5""> VMware Snaphot Report </td>
		</tr>
		<tr class=" Title ">
			<td>-----VM Name-----</td>
			<td>----------------Snapshot Name----------------</td>
			<td>-------Date Created-------</td>
			<td>----------------Description----------------</td>
			<td>-------------------Host-------------------</td>
		</tr>"

	ForEach ($snapshot in $report)
	{
		Write-Output "<tr>
			<td>$($snapshot.vm)</td>
			<td>$($snapshot.name)</td>
			<td>$($snapshot.created)</td>
			<td>$($snapshot.description)</td>
			<td>$($snapshot.host)</td>
		</tr>"
	}
	Write-Output "</table>
	</body>
	</html>"
}

# Login details for standalone ESXi servers (not required if using VirtualCenter)
$username = 'ESXiUser'
$password = 'ESXiPassword' #Change to the root password you set for you ESXi server

# List of servers including Virtual Center Server.  The account this script will run as will need at least Read-Only access to Virtual Center
# Chance to DNS Names/IP addresses of your ESXi servers or Virtual Center Server. Comma separated.
$ServerList = "VCServer"

# Initialise Array
$Report = @()

# Get snapshots from all servers
ForEach ($server in $serverlist)
{
	# Check is server is a Virtual Center Server and connect with current user
	If ($server -eq "VCServer")
	{
		Connect-VIServer $server
	}
	## Use specific login details for the rest of servers in $serverlist
	## Uncomment this line if you use ESXi hosts rather than VirtualCenter
	#Else
	#{
	#	Connect-VIServer $server -user $username -password $password
	#}

	Get-VM | Get-Snapshot | ForEach-Object {
		$Snap = {} | Select-Object VM,Name,Created,Description,Host
		$Snap.VM = $_.vm.name
		$Snap.Name = $_.name
		$Snap.Created = $_.created
		$Snap.Description = $_.description
		$Snap.Host = $_.vm.host.name
		$Report += $Snap
	}
}

# Generate the report and email it as a HTML body of an email
Generate-Report > "VmwareSnapshots.html"
If ($Report -ne "")
{
	$SmtpClient = New-Object System.Net.Mail.SmtpClient

	# Change to a SMTP server in your environment
	$SmtpClient.host = "mail.yourdomain.com"
	$MailMessage = New-Object System.Net.Mail.MailMessage

	# Change to email address you want emails to be coming from
	$MailMessage.from = "no-reply@yourdomain.com"

	# Change to email address you would like to receive emails.
	$MailMessage.To.Add("YOUREMAILADDRESS@yourdomain.com")
	$MailMessage.IsBodyHtml = 1
	$MailMessage.Subject = "Vmware Snapshots"
	$MailMessage.Body = Generate-Report
	$SmtpClient.Send($MailMessage)
}
