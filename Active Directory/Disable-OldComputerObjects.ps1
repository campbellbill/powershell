[CmdletBinding(
	SupportsShouldProcess = $true
	, ConfirmImpact = "Medium"
)]
#region Script Parameters
	Param(
		[Parameter(
			Position = 0
			, Mandatory = $true
			, HelpMessage = 'Sets the scope that the script will use to query computer objects from Active Directory ("AllWindows", "Server", "Workstation").'
		)]
		[ValidateSet("AllWindows", "Server", "Workstation")]
		[string]$ComputerScope #= "Server"
		, [Parameter(
			Position = 1
			, Mandatory = $false
			, HelpMessage = 'Sets the number of days that a computer object may go with no change to its Active Directory object. Used in the query for computer objects from Active Directory (If left blank it will calculate the number of days back to January 1 of the current year or use 90 days, which ever is larger. Minimum of 90 days will be used regardless.).'
		)]
		#[ValidateScript(
		#	{
		#		If ($_ -lt 1)
		#		{
		#			$false
		#			[int]$DaysNoChange = (Get-Date).DayOfYear
		#		}
		#		ElseIf ($_ -lt 90)
		#		{
		#			$false
		#			[int]$DaysNoChange = 90
		#		}
		#		Else
		#		{
		#			$true
		#		}
		#	}
		#)]
		[int]$DaysNoChange
		, [Parameter(
			Position = 2
			, Mandatory = $false
			, HelpMessage = 'The ScrDebug switch turns the custom debugging code in the script on or off ($true/$false). Defaults to FALSE.'
		)][switch]$ScrDebug = $false
	)
#endregion Script Parameters

# Disable Old Computer Objects with no change for over 90 days.
# .\Disable-OldComputerObjects.ps1 -ComputerScope "Server" -ScrDebug
# .\Disable-OldComputerObjects.ps1 -ComputerScope "Server" -DaysNoChange 90 -ScrDebug
# .\Disable-OldComputerObjects.ps1 -ComputerScope "AllWindows" -DaysNoChange 90 -ScrDebug
# . ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Active_Directory\Disable-OldComputerObjects.ps1 -ComputerScope "AllWindows" -DaysNoChange 90 -ScrDebug
# . ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Active_Directory\Disable-OldComputerObjects.ps1 -ComputerScope "Server" -DaysNoChange 90 -ScrDebug
# . ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Active_Directory\Disable-OldComputerObjects.ps1 -ComputerScope "Server" -ScrDebug
# . ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Active_Directory\Disable-OldComputerObjects.ps1 -ComputerScope "Workstation" -ScrDebug
# . ${Env:USERPROFILE}\SkyDrive\Scripts\PowerShell\Active_Directory\Disable-OldComputerObjects.ps1 -ComputerScope "Workstation" -DaysNoChange 90

#region Script Variable Initialization
	[string]$ScrStartTime = (Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString()
	If ($ScrDebug)
	{
		Write-Host -ForegroundColor Green "Start Time: $($ScrStartTime)"
		#Write-Host -ForegroundColor Green "Start Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	}

	#region Assembly and PowerShell Module Initialization
		######################################
		##  Import Modules & Add PSSnapins  ##
		######################################
		## Check for ActiveRoles Management Shell for Active Directory Sanpin, attempt to load
		If (!(Get-Command Get-QADComputer -ErrorAction SilentlyContinue))
		{
			Try
			{
				Add-PSSnapin Quest.ActiveRoles.ADManagement
			}
			Catch
			{
				If (!(Get-Command Get-QADComputer -ErrorAction SilentlyContinue))
				{
					Add-PSSnapin Quest.ActiveRoles.ADManagement
				}
			}
			Finally
			{
				If (!(Get-Command Get-QADComputer -ErrorAction SilentlyContinue))
				{
					Throw "Cannot load the Snapin $([char]34)Quest.ActiveRoles.ADManagement$([char]34)!! Please make sure you have installed the $([char]34)ActiveRoles Management Shell for Active Directory$([char]34)."
				}

				#$SnapinSettingsList = Get-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject
				#If ($SnapinSettingsList -notcontains "IPv4Address")
				#{
				#	$SnapinSettingsList += 'IPv4Address'
				#	Set-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject $SnapinSettingsList
				#}

				$SnapinSettingsSizeLimit = Get-QADPSSnapinSettings -DefaultSizeLimit
				If ($SnapinSettingsSizeLimit -ne 0)
				{
					Set-QADPSSnapinSettings -DefaultSizeLimit 0
				}
			}
		}

		$Error.Clear()
	#endregion Assembly and PowerShell Module Initialization

	#region Variables for the Log
		[string]$ScriptDirectory = ($MyInvocation.MyCommand.Definition | Split-Path -Parent)
		[string]$ScriptNameWithoutExtension = ($MyInvocation.MyCommand.Name).Replace(".ps1", "")
		[string]$LogFileDate = (Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()
		[string]$Ext = "log"
		[string]$LogFileName = "$($LogFileDate)_$($ScriptNameWithoutExtension)_$($ComputerScope).$($Ext)"

		If ($ScrDebug)
		{
			$ScrLogDir = "$(${Env:USERPROFILE})\Desktop\Logs"
		}
		Else
		{
			$ScrLogDir = "$($ScriptDirectory)\Logs"
		}

		If (!(Test-Path -Path "$($ScrLogDir)"))
		{
			New-Item -ItemType Directory "$($ScrLogDir)" | Out-Null
		}

		$ScriptLogFile = "$($ScrLogDir)\$($LogFileName)"
	#endregion Variables for the Log

	If (($DaysNoChange -eq $null) -or ($DaysNoChange -eq "") -or ($DaysNoChange -lt 1))
	{
		[int]$DaysNoChange = (Get-Date).DayOfYear
	}

	If ($DaysNoChange -lt 90)
	{
		[int]$DaysNoChange = 90
	}

	# Active Directory DNS Domain names
	$Script:ADDSDnsDomains	= @(
		"eclg.org"
		#, "ecollege.net"
		#, "ecollegeqa.net"
		#, "eclgsecure.net"
		#, "eclgsecuresc.net"
		#, "pad.pearsoncmg.com"
		#, "wharton.com"
		#, "wrk.pad.pearsoncmg.com"
	)

	# Translation of the DNS Domain name to the NETBIOS Domain name for the Domains covered under the Microsoft SPLA agreement.
	$Script:DNStoNBnames = @{
		"eclg.org"					= "REALEDUCATION"
		"ecollege.net"				= "ATHENS"
		"ecollegeqa.net"			= "ATHENSQA"
		"eclgsecure.net"			= "CAIRO"
		"eclgsecuresc.net"			= "CAIROSC"
		#"pad.pearsoncmg.com"		= "PAD"
		#"wharton.com"				= "WHARTON"
		#"wrk.pad.pearsoncmg.com"	= "WRK"
	}
#endregion Script Variable Initialization

#region User Defined Functions
	
#endregion User Defined Functions

#region Main Script
	#region Logfile Header
		[string]$ADDSDomains = [string]::Join(", ", $ADDSDnsDomains)

		#Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t--------------------------------------------------------------------------------"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t----------------------------------------------------------------------------------------------------"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t - $($ScriptNameWithoutExtension) -"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Script Start Time     : $([char]34)$($ScrStartTime)$([char]34)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Computers Scope       : $([char]34)$($ComputerScope)$([char]34)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t AD DS Domains Checked : $([char]34)$($ADDSDomains)$([char]34)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Days Without Change   : $([char]34)$($DaysNoChange)$([char]34)"

		If ($ScrDebug)
		{
			Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t ScrDebug State        : $([char]34)$($ScrDebug.ToString())$([char]34) NO Changes will be applied"
		}
		Else
		{
			Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t ScrDebug State        : $([char]34)$($ScrDebug.ToString())$([char]34) Changes will be applied"
		}

		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t----------------------------------------------------------------------------------------------------"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`tWhen Created       `tLast Changed       `tAction                  `tResult  `tdNSHostName                                  `tDN                                                                                                                      `tErrors"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t-------------------`t-------------------`t------------------------`t--------`t---------------------------------------------`t------------------------------------------------------------------------------------------------------------------------`t--------"
		#Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t11/26/2008 21:14:07`t05/21/2012 18:08:59`tDisable Computer Account`tFAILURE `t---------------------------------------------`t------------------------------------------------------------------------------------------------------------------------`t--------"
	#endregion Logfile Header

	[string]$ComputersNodNSHostNameNames = [string]::Empty
	[int]$ComputersNodNSHostNameCount = 0
	[int]$ComputersDisabledCount = 0
	[int]$ComputersTotalCount = $OldComputers.Count

	#$SecUsrNam = ${env:username}
	#$SecPasswd = '' | ConvertTo-SecureString -asPlainText -Force
	$SecUsrNam = Read-Host -Prompt "Enter your username"
	$SecPasswd = Read-Host -Prompt "Enter your password" -AsSecureString

	ForEach ($DNSDomain in $ADDSDnsDomains)
	{
		If($DNStoNBnames[$DNSDomain])
		{
			$NBName = $DNStoNBnames[$DNSDomain]
		}
		Else
		{
			Throw "NETBIOS name not found in the Hash Table 'DNStoNBnames'"
		}

		$ConnectAcct = "$($NBName)\$($SecUsrNam)"
		#$UserCreds = New-Object System.Management.Automation.PSCredential ($ConnectAcct, $SecPasswd)

		$GetQADActivity = "Retreiving a list of $([char]34)$($ComputerScope)$([char]34) Computer Objects from domain $([char]34)$($DNSDomain)$([char]34). . ."

		If ($ComputerScope -eq "Workstation")
		{
			$OldComputers = @(Get-QADComputer -Activity $GetQADActivity -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPasswd | Where-Object {
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
		ElseIf ($ComputerScope -eq "Server")
		{
			$OldComputers = @(Get-QADComputer -Activity $GetQADActivity -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPasswd | Where-Object {
				((((Get-Date) - $_.whenChanged).Days) -gt $DaysNoChange) -and
				($_.operatingSystem -like "Windows*Server*") -and
				($_.AccountIsDisabled -eq $false)
			})
		}
		Else	#If ($ComputerScope -eq "AllWindows")
		{
			$OldComputers = @(Get-QADComputer -Activity $GetQADActivity -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPasswd | Where-Object {
				((((Get-Date) - $_.whenChanged).Days) -gt $DaysNoChange) -and
				($_.operatingSystem -like "Windows*") -and
				($_.operatingSystem -ne "Windows NT") -and	# Some of the older NetApp devices show up as "Windows NT" and we don't want to disable them.
				($_.AccountIsDisabled -eq $false)
			})
		}

		$ComputersTotalCount += $OldComputers.Count

		ForEach ($Comp in $OldComputers)
		{
			If ($Comp.dNSHostName -eq $null)
			{
				$CompCanonicalName = $Comp.CanonicalName
				$CompDNSDomain = $CompCanonicalName.Substring(0, $CompCanonicalName.IndexOf("/")).ToLower()
				$CompName = $CompCanonicalName.Substring($CompCanonicalName.LastIndexOf("/") + 1).ToLower()
				$CompDNSName = "$($CompName).$($CompDNSDomain)"

				If ($ComputersNodNSHostNameNames.Length -lt 1)
				{
					$ComputersNodNSHostNameNames = "$([char]34)$($CompName)$([char]34)"
				}
				Else
				{
					$ComputersNodNSHostNameNames += ", $([char]34)$($CompName)$([char]34)"
				}

				# .PadRight(IntTotalWidth, [StrPaddingChar])
				[string]$strPtlLogEntry = "$($Comp.whenCreated)`t$($Comp.whenChanged)`tDisable Computer Account`tActionResult`t$($CompDNSName.PadRight(45, " "))`t$($Comp.DN.PadRight(120, " "))"

				$ComputersNodNSHostNameCount++
			}
			Else
			{
				[string]$strPtlLogEntry = "$($Comp.whenCreated)`t$($Comp.whenChanged)`tDisable Computer Account`tActionResult`t$($Comp.dNSHostName.ToLower().PadRight(45, " "))`t$($Comp.DN.PadRight(120, " "))"
			}

			If ($ScrDebug)
			{
				Disable-QADComputer -Identity $Comp.DN -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPasswd -Confirm:$false -WhatIf
			}
			Else
			{
				Disable-QADComputer -Identity $Comp.DN -Service $DNSDomain -ConnectionAccount $ConnectAcct -ConnectionPassword $SecPasswd -Confirm:$false
			}

			If ($?)
			{
				If ($ScrDebug)
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`twSuccess")
				}
				Else
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`tSuccess ")
				}

				Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t$($strPtlLogEntry)"

				$ComputersDisabledCount++
			}
			Else
			{
				If ($ScrDebug)
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`twFAILURE") + "`t$($Error[0].Exception.Message)"
				}
				Else
				{
					$strPtlLogEntry = $strPtlLogEntry.Replace("`tActionResult", "`tFAILURE ") + "`t$($Error[0].Exception.Message)"
				}

				Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t$($strPtlLogEntry)"
			}
		}
	}

	If ([string]::IsNullOrEmpty($ComputersNodNSHostNameNames))
	{
		[string]$ComputersNodNSHostNameNames = 'None found missing this attribute.'
	}

	#region Logfile Footer
		#Add-Content -Path $ScriptLogFile -Value "$([char]13)$([char]10)$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t------------------------------------:------------------------------------"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t------------------------------------:------------------------------------"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Total Computer Objects from AD     : $($ComputersTotalCount)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Total Computers No AD dNSHostName  : $($ComputersNodNSHostNameCount)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Computers No AD dNSHostName Names  : $($ComputersNodNSHostNameNames)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t Total Computer Objects Disabled    : $($ComputersDisabledCount)"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`t------------------------------------:------------------------------------"
		Add-Content -Path $ScriptLogFile -Value "$((Get-Date -Format '[yyyy-MMM-dd HH:mm:ss]').ToString()):`tEnd Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	#endregion Logfile Footer

	If ($ScrDebug)
	{
		Write-Host -ForegroundColor Yellow "------------------------------------:------------------------------------"
		Write-Host -ForegroundColor Green  " Total Computer Objects from AD     : $($ComputersTotalCount)"
		Write-Host -ForegroundColor Green  " Total Computers No AD dNSHostName  : $($ComputersNodNSHostNameCount)"
		Write-Host -ForegroundColor Green  " Computers No AD dNSHostName Names  : $($ComputersNodNSHostNameNames)"
		Write-Host -ForegroundColor Green  " Total Computer Objects Disabled    : $($ComputersDisabledCount)"
		Write-Host -ForegroundColor Yellow "------------------------------------:------------------------------------"
		Write-Host -ForegroundColor Magenta " End Time: $((Get-Date -Format 'MMM dd, yyyy, HH:mm:ss').ToString())"
	}
#endregion Main Script

#region Used during the creation or debugging of the script
	<#
		# Used during the creation or debugging of the script.
		Get-QADComputer -InactiveFor (Get-Date).DayOfYear | ForEach-Object { Disable-QADComputer $_ -WhatIf }
		Get-QADComputer -InactiveFor (Get-Date).DayOfYear | Where-Object {($_.operatingSystem -like "Windows*Server*") -and ($_.AccountIsDisabled -eq $false)} | Export-Csv -Path "$(${Env:USERPROFILE})\Desktop\Reports\InactiveComputerObjects_2013.10.22.Tue.1749.csv" -NoTypeInformation
		Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt (Get-Date).DayOfYear) -and ($_.operatingSystem -like "Windows*Server*") -and ($_.AccountIsDisabled -eq $false)} | Export-Csv -Path "$(${Env:USERPROFILE})\Desktop\Reports\InactiveComputerObjects_2013.10.23.Wed.1417.csv" -NoTypeInformation
		Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt (Get-Date).DayOfYear) -and (($_.operatingSystem -like "Windows*") -and ($_.operatingSystem -notlike "Windows*Server*")) -and ($_.AccountIsDisabled -eq $false)} | Export-Csv -Path "$(${Env:USERPROFILE})\Desktop\Reports\InactiveComputerObjects_2013.10.23.Wed.0936.csv" -NoTypeInformation
		Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt (Get-Date).DayOfYear) -and (($_.operatingSystem -like "Windows 2*") -or ($_.operatingSystem -like "Windows 7*") -or ($_.operatingSystem -like "Windows 8*") -or ($_.operatingSystem -like "Windows*Dev*") -or ($_.operatingSystem -like "Windows*Vis*") -or ($_.operatingSystem -like "Windows*XP*")) -and ($_.AccountIsDisabled -eq $false)} | Export-Csv -Path "$(${Env:USERPROFILE})\Desktop\Reports\InactiveComputerObjects_2013.10.23.Wed.1020.csv" -NoTypeInformation
		Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt 90) -and (($_.operatingSystem -like "Windows 2*") -or ($_.operatingSystem -like "Windows 7*") -or ($_.operatingSystem -like "Windows 8*") -or ($_.operatingSystem -like "Windows*Dev*") -or ($_.operatingSystem -like "Windows*Vis*") -or ($_.operatingSystem -like "Windows*XP*")) -and ($_.AccountIsDisabled -eq $false)} | Export-Csv -Path "$(${Env:USERPROFILE})\Desktop\Reports\InactiveComputerObjects_2013.10.23.Wed.1022.csv" -NoTypeInformation
		(Get-QADComputer -InactiveFor (Get-Date).DayOfYear | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt (Get-Date).DayOfYear) -and ($_.operatingSystem -like "Windows*Server*") -and ($_.AccountIsDisabled -eq $false)}).Count
		(Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt (Get-Date).DayOfYear) -and ($_.operatingSystem -like "Windows*Server*") -and ($_.AccountIsDisabled -eq $false)}).Count
		(Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt 90) -and ($_.operatingSystem -like "Windows*Server*") -and ($_.AccountIsDisabled -eq $false)}).Count
		# For NON-Windows Computer Objects
		Get-QADComputer | Where-Object {((((Get-Date) - $_.whenChanged).Days) -gt 90) -and ($_.operatingSystem -notlike "Windows*") -and ($_.AccountIsDisabled -eq $false)} | Export-Csv -Path "$(${Env:USERPROFILE})\Desktop\Reports\NonWindowsComputerObjects_2013.10.23.Wed.1815.csv" -NoTypeInformation
	#>
#endregion Used during the creation or debugging of the script
