$Script:MyCsv = "C:\Temp\Events_$((Get-Date -Format 'yyyy.MM.dd.ddd-HHmm').ToString()).csv"

Function Get-Win7LogonHistory
{
	$Logons = Get-EventLog Security -AsBaseObject -InstanceId 4624,4647 -After (Get-Date).AddDays(-1) | Where-Object {
		($_.InstanceId -eq 4647) `
		-or (($_.InstanceId -eq 4624) -and ($_.Message -match "Logon Type:\s+2")) `
		-or (($_.InstanceId -eq 4624) -and ($_.Message -match "Logon Type:\s+10"))
	}
	$Events = $Logons | Sort-Object TimeGenerated

	if ($Events)
	{
		foreach ($Event in $Events)
		{
			## Parse logon data from the Event.
			if ($Event.InstanceId -eq 4624)
			{
				## A user logged on.
				$Action = 'logon'

				$Event.Message -match "Logon Type:\s+(\d+)" | Out-Null
				$LogonTypeNum = $matches[1]

				## Determine logon type.
				if ($LogonTypeNum -eq 2)
				{
					$LogonType = 'console'
				}
				elseif ($LogonTypeNum -eq 10)
				{
					$LogonType = 'remote'
				}
				else
				{
					$LogonType = 'other'
				}

				## Determine user.
				if ($Event.Message -match "New Logon:\s*Security ID:\s*.*\s*Account Name:\s*(\w+)")
				{
					$User = $matches[1]
				}
				else
				{
					$Index = $Event.Index
					Write-Warning "Unable to parse Security log Event. Malformed entry? Index: $Index"
				}
			}
			elseif ($Event.InstanceId -eq 4647)
			{
				## A user logged off.
				$Action = 'logoff'
				$LogonType = $null

				## Determine user.
				if ($Event.Message -match "Subject:\s*Security ID:\s*.*\s*Account Name:\s*(\w+)")
				{
					$User = $matches[1]
				}
				else
				{
					$Index = $Event.Index
					Write-Warning "Unable to parse Security log Event. Malformed entry? Index: $Index"
				}
			}
			elseif ($Event.InstanceId -eq 41)
			{
				## The computer crashed.
				$Action = 'logoff'
				$LogonType = $null
				$User = '*'
			}

			## As long as we managed to parse the Event, print output.
			if ($User)
			{
				$TimeStamp = Get-Date $Event.TimeGenerated
				$Output = New-Object -Type PSCustomObject
				Add-Member -MemberType NoteProperty -Name 'UserName' -Value $User -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value ${Env:COMPUTERNAME} -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'Action' -Value $Action -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'LogonType' -Value $LogonType -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'TimeStamp' -Value $TimeStamp -InputObject $Output
				$Output = "$User,${Env:COMPUTERNAME},$Action,$LogonType,$TimeStamp"

				Write-Output $Output | Out-File $MyCsv -Encoding ascii -Force -Append
			}
		}
	}
	else
	{
		Write-Host "No recent logon/logoff events."
	}
}

Function Get-WinXPLogonHistory
{
	$Logons = Get-EventLog Security -AsBaseObject -InstanceId 528,551 -After (Get-Date).AddDays(-1) | Where-Object { ($_.InstanceId -eq 551) `
		-or (($_.InstanceId -eq 528) -and ($_.Message -match "Logon Type:\s+2")) `
		-or (($_.InstanceId -eq 528) -and ($_.Message -match "Logon Type:\s+10"))
	}
	#$poweroffs = Get-Eventlog System -AsBaseObject -InstanceId 6008
	#$Events = $Logons + $poweroffs | Sort-Object TimeGenerated

	if ($Events)
	{
		foreach ($Event in $Events)
		{
			## Parse logon data from the Event.
			if ($Event.InstanceId -eq 528)
			{
				## A user logged on.
				$Action = 'logon'

				$Event.Message -match "Logon Type:\s+(\d+)" | Out-Null
				$LogonTypeNum = $matches[1]

				## Determine logon type.
				if ($LogonTypeNum -eq 2)
				{
					$LogonType = 'console'
				}
				elseif ($LogonTypeNum -eq 10)
				{
					$LogonType = 'remote'
				}
				else
				{
					$LogonType = 'other'
				}

				## Determine user.
				if ($Event.Message -match "Successful Logon:\s*User Name:\s*(\w+)")
				{
					$User = $matches[1]
				}
				else
				{
					$Index = $Event.Index
					Write-Warning "Unable to parse Security log Event. Malformed entry? Index: $Index"
				}
			}
			elseif ($Event.InstanceId -eq 551)
			{
				## A user logged off.
				$Action = 'logoff'
				$LogonType = $null

				## Determine user.
				if ($Event.Message -match "User initiated logoff:\s*User Name:\s*(\w+)")
				{
					$User = $matches[1]
				}
				else
				{
					$Index = $Event.Index
					Write-Warning "Unable to parse Security log Event. Malformed entry? Index: $Index"
				}
			}
			#elseif ($Event.InstanceId -eq 6008)
			#{
			#	## The computer crashed.
			#	$Action = 'logoff'
			#	$LogonType = $null
			#	$User = '*'
			#}

			## As long as we managed to parse the Event, print output.
			if ($User)
			{
				$TimeStamp = Get-Date $Event.TimeGenerated
				$Output = New-Object -Type PSCustomObject
				Add-Member -MemberType NoteProperty -Name 'UserName' -Value $User -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value ${Env:COMPUTERNAME} -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'Action' -Value $Action -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'LogonType' -Value $LogonType -InputObject $Output
				Add-Member -MemberType NoteProperty -Name 'TimeStamp' -Value $TimeStamp -InputObject $Output
				$Output = "$User,${Env:COMPUTERNAME},$Action,$LogonType,$TimeStamp"    

				Write-Output $Output | Out-File $MyCsv -Encoding ascii -Force -Append
			}
		}
	}
	else
	{
		Write-Host "No recent logon/logoff events."
	}
}

$OSversion = (Get-WmiObject -Query 'SELECT version FROM Win32_OperatingSystem').Version
if ($OSversion -ge 6)
{
	Get-Win7LogonHistory
}
else
{
	Get-WinXPLogonHistory
}

$a = "<style>table {font-size: 10pt; font-family: calibri;}"
$a = $a + "BODY{background-color:White;}"
$a = $a + "TABLE{align:center; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{font-size:1.2em; border-width: 1px;padding: 2px;border-style: solid;border-color: white;background-color:#0075B0;}"
$a = $a + "TD{font-size:1em; border-width: 1px;padding: 2px;border-style: solid;border-color: Black;background-color:white;}"
$a = $a + "</style>"
#Import-Csv $MyCsv Format-Table -Header Message,Account Name,UPN | ConvertTo-Html -Head $a | Out-File "c:\temp\test.htm"
#$HTMLBody = Get-Content "c:\temp\test.htm"

$SmtpServer = "159.182.31.103"
$SmtpFrom = "DailyUserLogonReport@noreply.com"
$SmtpTo = "Susan.Marold-Leavens@Pearson.com,Syed.Hussain@Pearson.com,ptinfosecnda@pearson.com,pt-us-qa-coe@pearson.com"
$MessageSubject = "Daily user login report"
$Message = New-Object System.Net.Mail.MailMessage $SmtpFrom, $SmtpTo
$Message.Subject = $MessageSubject
$Message.IsBodyHTML = $true
$Message.Body = Import-Csv $MyCsv -Header UserName,ComputerName,Action,LogonType,TimeStamp | ConvertTo-Html -Head $a
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer)
$Smtp.Send($Message)

Remove-Item $MyCsv
