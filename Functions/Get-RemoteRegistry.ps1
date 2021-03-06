## Get-RemoteRegistry
########################################################################################
## Version: 2.1
##  + Fixed a pasting bug 
##  + I added the "Properties" parameter so you can select specific registry values
## NOTE: you have to have access, and the remote registry service has to be running
########################################################################################
## USAGE:
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP"
##     * Returns a list of subkeys (because this key has no properties)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727"
##     * Returns a list of subkeys and all the other "properties" of the key
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version"
##     * Returns JUST the full version of the .Net SP2 as a STRING (to preserve prior behavior)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version
##     * Returns a custom object with the property "Version" = "2.0.50727.3053" (your version)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version,SP
##     * Returns a custom object with "Version" and "SP" (Service Pack) properties
##
##  For fun, get all .Net Framework versions (2.0 and greater) 
##  and return version + service pack with this one command line:
##
##    Get-RemoteRegistry $RemotePC "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP" | 
##    Select -Expand Subkeys | ForEach-Object { 
##      Get-RemoteRegistry $RemotePC "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\$_" Version,SP 
##    }
##
##	Found at:
##		http://poshcode.org/615
##
########################################################################################
Function Get-RemoteRegistry
{
	Param (
		[string]$computer = $(Read-Host "Remote Computer Name")
		, [string]$Path = $(Read-Host "Remote Registry Path (must start with HKLM,HKCU,etc)")
		, [string[]]$Properties
		, [switch]$Verbose
	)

	if ($Verbose)
	{
		## Only affects this script.
		$VerbosePreference = 2
	}

	$RemoteHiveRoot, $last = $Path.Split("\")
	$last = $last[-1]
	$Path = $Path.Substring($RemoteHiveRoot.Length + 1, $Path.Length - ( $last.Length + $RemoteHiveRoot.Length + 2))
	$RemoteHiveRoot = $RemoteHiveRoot.TrimEnd(":")

	## split the path to get a list of subkeys that we will need to access
	## ClassesRoot, CurrentUser, LocalMachine, Users, PerformanceData, CurrentConfig, DynData
	switch($RemoteHiveRoot) {
		"HKCR" { $RemoteHiveRoot = "ClassesRoot"}
		"HKCU" { $RemoteHiveRoot = "CurrentUser" }
		"HKLM" { $RemoteHiveRoot = "LocalMachine" }
		"HKU" { $RemoteHiveRoot = "Users" }
		"HKPD" { $RemoteHiveRoot = "PerformanceData"}
		"HKCC" { $RemoteHiveRoot = "CurrentConfig"}
		"HKDD" { $RemoteHiveRoot = "DynData"}
		default { return "Path argument is not valid" }
	}

	## Access Remote Registry Key using the static OpenRemoteBaseKey method.
	Write-Verbose "Accessing $RemoteHiveRoot from $computer"
	$RemoteHiveRootKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RemoteHiveRoot, $computer)
	if(-not $RemoteHiveRootKey)
	{
		Write-Error "Can't open the remote $RemoteHiveRoot registry hive"
	}

	Write-Verbose "Opening $Path"
	$key = $RemoteHiveRootKey.OpenSubKey( $Path )
	if(-not $key)
	{
		Write-Error "Can't open $($RemoteHiveRoot + '\' + $Path) on $computer"
	}

	$subkey = $key.OpenSubKey( $last )
	$output = New-Object object

	if($subkey -and $Properties -and $Properties.Count)
	{
		foreach($property in $Properties)
		{
			Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
		}
		Write-Output $output
	}
	elseif($subkey)
	{
		Add-Member -InputObject $output -Type NoteProperty -Name "Subkeys" -Value @($subkey.GetSubKeyNames())
		foreach($property in $subkey.GetValueNames())
		{
			Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
		}
		Write-Output $output
	}
	else
	{
		$key.GetValue($last)
	}
}
