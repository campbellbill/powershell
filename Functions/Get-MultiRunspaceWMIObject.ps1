Function Get-MultiRunspaceWMIObject
{
	<#
	.SYNOPSIS
		Get generic WMI object data from a remote or local system.
	.DESCRIPTION
		Get WMI object data from a remote or local system. Multiple Runspaces are utilized and 
		alternate credentials can be provided.
	.PARAMETER ComputerName
		Specifies the target computer for data query.
	.PARAMETER Namespace
		Namespace to query
	.PARAMETER Class
		Class to query
	.PARAMETER ThrottleLimit
		Specifies the maximum number of systems to inventory simultaneously 
	.PARAMETER Timeout
		Specifies the maximum time in second command can run in background before terminating this thread.
	.PARAMETER ShowProgress
		Show progress bar information

	.EXAMPLE
		PS > (Get-MultiRunspaceWMIObject -Class Win32_Printer).WMIObjects

		<output is all your local printers>

		Description
		-----------
		Queries the local machine for all installed printer information and spits out what is found.

	.NOTES
		Author: Zachary Loeber
		Site: http://www.the-little-things.net/
		Requires: Powershell 2.0

		Version History
		1.0.0 - 08/31/2013
		- Initial release

	.LINK
		http://gallery.technet.microsoft.com/scriptcenter/Gather-Generic-WMI-Data-474f788b
	#>
	#region Script Parameters
		[CmdletBinding()]
		Param(
			[Parameter(
				HelpMessage = "Computer or computers to gather information from"
				, ValueFromPipeline = $true
				, ValueFromPipelineByPropertyName = $true
				, Position = 0
			)]
			[ValidateNotNullOrEmpty()]
			[Alias('DNSHostName', 'PSComputerName')]
			[string[]]$ComputerName = $Env:COMPUTERNAME
			, [Parameter(
				HelpMessage = "WMI class to query"
				, Position = 1
			)][string]$Class
			, [Parameter(
				HelpMessage = "WMI namespace to query"
			)][string]$Namespace = 'root\cimv2'
			, [Parameter(
				HelpMessage = "Maximum number of concurrent threads"
			)]
			[ValidateRange(1,65535)]
			[int32]$ThrottleLimit = 32
			, [Parameter(
				HelpMessage = "Timeout before a thread stops trying to gather the information"
			)]
			[ValidateRange(1,65535)]
			[int32]$Timeout = 120
			, [Parameter(
				HelpMessage = "Display progress of Function"
			)][switch]$ShowProgress
			, [Parameter(
				HelpMessage = "Set this if you want the function to prompt for alternate credentials"
			)][switch]$PromptForCredential
			, [Parameter(
				HelpMessage = "Set this if you want to provide your own alternate credentials"
			)][System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
		)
	#endregion Script Parameters

	#	Get-MultiRunspaceWMIObject -ComputerName $Env:COMPUTERNAME -Class ''

	Begin
	{
		# Gather possible local host names and IPs to prevent credential utilization in some cases
		Write-Verbose -Message ('WMI Query {0}: Creating local hostname list' -f $Class)
		$IPAddresses = [Net.Dns]::GetHostAddresses($Env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
		$HostNames = $IPAddresses | ForEach-Object {
			try
			{
				[Net.Dns]::GetHostByAddress($_)
			}
			catch
			{
				# We do not care about errors here...
			}
		} | Select-Object -ExpandProperty HostName -Unique
		$LocalHostInfo = @('', '.', 'localhost', $Env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames

		Write-Verbose -Message ('WMI Query {0}: Creating initial variables' -f $Class)
		$RunspaceTimers = [HashTable]::Synchronized(@{})
		$Runspaces = New-Object -TypeName System.Collections.ArrayList
		$bgRunspaceCounter = 0

		If ($PromptForCredential)
		{
			$Credential = Get-Credential
		}

		Write-Verbose -Message ('WMI Query {0}: Creating Initial Session State' -f $Class)
		$issInitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		ForEach ($ExternalVariable in ('RunspaceTimers', 'Credential', 'LocalHostInfo'))
		{
			Write-Verbose -Message ("WMI Query {0}: Adding variable $ExternalVariable to initial session state" -f $Class)
			$issInitialSessionState.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
		}

		Write-Verbose -Message ('WMI Query {0}: Creating Runspace pool' -f $Class)
		$rpRunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $issInitialSessionState, $Host)
		$rpRunspacePool.ApartmentState = 'STA'
		$rpRunspacePool.Open()

		# This is the actual code called for each computer
		Write-Verbose -Message ('WMI Query {0}: Defining background Runspaces scriptblock' -f $Class)
		$ScriptBlock = {
			[CmdletBinding()]
			Param(
                [Parameter()][string]$ComputerName
				, [Parameter()][int]$bgRunspaceID
				, [Parameter()][string]$Class
				, [Parameter()][string]$Namespace = 'root\cimv2'
			)
			$RunspaceTimers.$bgRunspaceID = Get-Date

			try
			{
				Write-Verbose -Message ('WMI Query {0}: Runspace {1}: Start' -f $Class, $ComputerName)
				$WMIHast = @{
					ComputerName = $ComputerName
					ErrorAction = 'Stop'
				}

				If (($LocalHostInfo -notcontains $ComputerName) -and ($Credential -ne $null))
				{
					$WMIHast.Credential = $Credential
				}

				$PSDateTime = Get-Date

				#region WMI Data
					Write-Verbose -Message ('WMI Query {0}: Runspace {1}: WMI information' -f $Class, $ComputerName)

					# Modify this variable to change your default set of display properties
					$defaultProperties = @('ComputerName', 'WMIObjects')

					# WMI data
					$WMI_Data = Get-WmiObject @WMIHast -Namespace $Namespace -Class $Class

					$ResultProperty = @{
						'PSComputerName' = $ComputerName
						'PSDateTime' = $PSDateTime
						'ComputerName' = $ComputerName
						'WMIObjects' = $WMI_Data
					}

					$ResultObject = New-Object -TypeName PSObject -Property $ResultProperty

					# Setup the default properties for output
					$ResultObject.PSObject.TypeNames.Insert(0, 'My.WMIObject.Info')
					$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]$defaultProperties)
					$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
					$ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
					Write-Output -InputObject $ResultObject
				#endregion WMI Data
			}
			catch
			{
				Write-Warning -Message ('WMI Query {0}: {1}: {2}' -f $Class, $ComputerName, $_.Exception.Message)
			}
			Write-Verbose -Message ('WMI Query {0}: Runspace {1}: End' -f $Class, $ComputerName)
		}

		Function Get-Result
		{
			[CmdletBinding()]
			Param(
				[switch]$Wait
			)

			do
			{
				$More = $false
				ForEach ($Runspace in $Runspaces)
				{
					$StartTime = $RunspaceTimers[$Runspace.ID]
					If ($Runspace.Handle.isCompleted)
					{
						Write-Verbose -Message ('WMI Query {0}: Thread done for {1}' -f $Class, $Runspace.IObject)
						$Runspace.PowerShell.EndInvoke($Runspace.Handle)
						$Runspace.PowerShell.Dispose()
						$Runspace.PowerShell = $null
						$Runspace.Handle = $null
					}
					ElseIf ($Runspace.Handle -ne $null)
					{
						$More = $true
					}

					If ($Timeout -and $StartTime)
					{
						If ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $Runspace.PowerShell)
						{
							Write-Warning -Message ('WMI Query {0}: Timeout {1}' -f $Class, $Runspace.IObject)
							$Runspace.PowerShell.Dispose()
							$Runspace.PowerShell = $null
							$Runspace.Handle = $null
						}
					}
				}

				If ($More -and $PSBoundParameters['Wait'])
				{
					Start-Sleep -Milliseconds 100
				}

				ForEach ($Threat in $Runspaces.Clone())
				{
					If ( -not $Threat.handle)
					{
						Write-Verbose -Message ('WMI Query {0}: Removing {1} from Runspaces' -f $Class, $Threat.IObject)
						$Runspaces.Remove($Threat)
					}
				}

				If ($ShowProgress)
				{
					$ProgressSplatting = @{
						Activity = ('WMI Query {0}: Getting info' -f $Class)
						Status = 'WMI Query {0}: {1} of {2} total threads done' -f $Class, ($bgRunspaceCounter - $Runspaces.Count), $bgRunspaceCounter
						PercentComplete = ($bgRunspaceCounter - $Runspaces.Count) / $bgRunspaceCounter * 100
					}
					Write-Progress @ProgressSplatting
				}
			}
			while ($More -and $PSBoundParameters['Wait'])
		}
	}
	Process
	{
		ForEach ($Computer in $ComputerName)
		{
			$bgRunspaceCounter++
			$psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$null = $psCMD.AddParameter('bgRunspaceID', $bgRunspaceCounter)
			$null = $psCMD.AddParameter('ComputerName', $Computer)
			$null = $psCMD.AddParameter('Class', $Class)
			$null = $psCMD.AddParameter('Namespace', $Namespace)
			$null = $psCMD.AddParameter('Verbose', $VerbosePreference)
			$psCMD.RunspacePool = $rpRunspacePool

			Write-Verbose -Message ('WMI Query {0}: Starting {1}' -f $Class, $Computer)
			[void]$Runspaces.Add(@{
				Handle = $psCMD.BeginInvoke()
				PowerShell = $psCMD
				IObject = $Computer
				ID = $bgRunspaceCounter
			})
			Get-Result
		}
	}
	End
	{
		Get-Result -Wait

		If ($ShowProgress)
		{
			Write-Progress -Activity ('WMI Query {0}: Getting WMI information' -f $Class) -Status 'Done' -Completed
		}

		Write-Verbose -Message ("WMI Query {0}: Closing Runspace pool" -f $Class)
		$rpRunspacePool.Close()
		$rpRunspacePool.Dispose()
	}
}

$WMIClassesToQuery = @('Win32_BIOS', 'Win32_ComputerSystem', 'Win32_PhysicalMemory', 'Win32_Processor')	#, 'Win32_OperatingSystem', '')

ForEach ($WMIClass in $WMIClassesToQuery)
{
	(Get-MultiRunspaceWMIObject -ComputerName $Env:COMPUTERNAME,'dome01osg' -Class $WMIClass).WMIObjects	#.eclg.org -ShowProgress
}
