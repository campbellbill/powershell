#region Script Variable Initialization
	[string]$StartTime = "$((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	#[string]$Script:FullScrPath = ($MyInvocation.MyCommand.Definition)
	[string]$Script:ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)

	#[string]$Script:ScriptNameNoExt = ($MyInvocation.MyCommand.Name).Replace('.ps1', '')
	[string]$Script:ScriptNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

	$GetHost = Get-Host
	$PowerShellHostVersion = $GetHost.Version.Major

	if ($PowerShellHostVersion -eq 1)
	{
		throw "Invalid PowerShell Host Version - $([char]34)$($MyInvocation.MyCommand.Name)$([char]34) can only run when PowerShell Host version is >= 2. Script will now exit."
	}
	else
	{
		Write-Host -ForegroundColor Yellow "$([char]34)$($MyInvocation.MyCommand.Name)$([char]34) is using PowerShell Version $($PowerShellHostVersion)"
	}

	$TestResults = @()
	$ServerList = @('uklonsepmp01.peroot.com','apsinsepmp01.peroot.com','aumelsepmp01.peroot.com','dn1wpsepmps001.peroot.com','sabrasepmp01.peroot.com','sepm.pearson.com')
	$ServerList = $ServerList | Sort-Object | Select-Object -Unique
#endregion Script Variable Initialization

#region User Defined Functions
	function Test-RemotePort ()
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
				)][string]$ComputerName
				, [Parameter(
					Position			= 1
					, Mandatory			= $true
				)][string]$RemotePort
			)
		#endregion Function Parameters
		## This works no matter in which form we get $host - ComputerName or ip address
		try
		{
			$RemoteAddress = [System.Net.Dns]::GetHostAddresses($ComputerName) | Select-Object IPAddressToString -ExpandProperty IPAddressToString
			if ($RemoteAddress.GetType().Name -eq "Object[]")
			{
				## If we have several ip's for that address, let's take first one
				$RemoteAddress = $RemoteAddress[0]
			}
		}
		catch
		{
			Write-Host "Possibly $([char]34)$($ComputerName)$([char]34) is the wrong ComputerName or IP"
			return
		}

		$NetSocket = New-Object Net.Sockets.TcpClient
		## We use Try\Catch to remove exception info from console if we can't connect
		try
		{
			$NetSocket.Connect($RemoteAddress, $RemotePort)
		}
		catch
		{}

		if ($NetSocket.Connected)
		{
			$NetSocket.Close()
			$msg = "RemotePort $([char]34)$($RemotePort)$([char]34) is operational to $([char]34)$($ComputerName)$([char]34)"

			$PsObjSvrProps = New-Object PSObject
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'ComputerName'		-Value $ComputerName
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'RemoteAddress'		-Value $RemoteAddress
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'RemotePort'			-Value $RemotePort
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'TcpTestSucceeded'	-Value $true
		}
		else
		{
			$msg = "RemotePort $([char]34)$($RemotePort)$([char]34) on $([char]34)$($RemoteAddress)$([char]34) is closed, "
			$msg += "You may need to contact your IT team to open it."

			$PsObjSvrProps = New-Object PSObject #New-Object System.Object
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'ComputerName'		-Value $ComputerName
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'RemoteAddress'		-Value $RemoteAddress
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'RemotePort'			-Value $RemotePort
			$PsObjSvrProps | Add-Member -MemberType NoteProperty -Name 'TcpTestSucceeded'	-Value $false
		}
		Write-Host $msg
		return $PsObjSvrProps
	}
#endregion User Defined Functions

#region Main Script
	Write-Host -ForegroundColor Green "Script Start Time: $($StartTime)"

	if ($PowerShellHostVersion -ge 4)
	{
		## If PowerShell version 4 or newer is installed this is faster and more elegant.
		$TestResults = $ServerList | ForEach-Object {
		    Write-Host -ForegroundColor Green "Beginning Test on: $([char]39)$($_)$([char]39)"
		    Test-NetConnection -ComputerName "$($_)" -Port 8014
		    Write-Host -ForegroundColor Magenta "Finished Test on: $([char]39)$($_)$([char]39)"
		}
		$TestResults | Format-Table ComputerName, remote*, tcptest* -AutoSize
	}
	else
	{
		## If PowerShell version 4 or newer is NOT installed this will work.
		$TestResults = $ServerList | ForEach-Object {
		    Write-Host -ForegroundColor Green "Beginning Test on: $([char]39)$($_)$([char]39)"
		    Test-RemotePort -ComputerName "$($_)" -RemotePort 8014
		    Write-Host -ForegroundColor Magenta "Finished Test on: $([char]39)$($_)$([char]39)"
		}
		$TestResults | Format-Table -AutoSize
	}

	Write-Host -ForegroundColor Magenta "Script Finish Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
#endregion Main Script
