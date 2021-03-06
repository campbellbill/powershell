#function Get-HBAInfo
#{
#	[CmdletBinding()]
#	Param(
#		[Parameter(
#			Mandatory = $false
#			, ValueFromPipeline = $true
#			, Position = 0
#		)][string]$ComputerName
#	)
#
#	begin 
#	{
#		$Namespace = 'root\WMI'
#	}
#	process
#	{
#		$port = Get-WmiObject -Class MSFC_FibrePortHBAAttributes -Namespace $Namespace @PSBoundParameters
#		$hbas = Get-WmiObject -Class MSFC_FCAdapterHBAAttributes -Namespace $Namespace @PSBoundParameters
#		$hbaProp = $hbas | Get-Member -MemberType Property, AliasProperty | Select-Object -ExpandProperty Name | Where-Object { $_ -notlike "__*" }
#		$hbas = $hbas | Select-Object $hbaProp
#		$hbas | ForEach-Object {
#			$_.NodeWWN = ((($_.NodeWWN) | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()
#		}
#
#		foreach ($hba in $hbas)
#		{
#			Add-Member -MemberType NoteProperty -InputObject $hba -Name FabricName -Value (
#				($port | Where-Object { $_.InstanceName -eq $hba.InstanceName }).Attributes | Select-Object `
#				@{Name = 'Fabric Name'; Expression = {(($_.FabricName | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()}}, `
#				@{Name = 'Port WWN'; Expression = {(($_.PortWWN | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()}} 
#			) -PassThru
#		}
#	}
#}

function Get-HBAInfo
{
	[CmdletBinding()]
	Param(
		[Parameter(
			Mandatory = $false
			, ValueFromPipeline = $true
			, Position = 0
		)][string]$ComputerName
	)

	begin 
	{
		$Namespace = 'root\WMI'
	}
	process
	{
		$port = Get-WmiObject -ErrorAction SilentlyContinue -Class MSFC_FibrePortHBAAttributes -Namespace $Namespace @PSBoundParameters
		$hbas = Get-WmiObject -ErrorAction SilentlyContinue -Class MSFC_FCAdapterHBAAttributes -Namespace $Namespace @PSBoundParameters

		$hbaProp = $hbas | Get-Member -MemberType Property, AliasProperty | Select-Object -ExpandProperty Name | Where-Object { $_ -notlike "__*" }
		$hbas = $hbas | Select-Object $hbaProp

		$hbas | ForEach-Object {
			$_.NodeWWN = ((($_.NodeWWN) | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()
		}

		foreach ($hba in $hbas)
		{
			Add-Member -MemberType NoteProperty -InputObject $hba -Name FabricName -Value (
				($port | Where-Object { $_.InstanceName -eq $hba.InstanceName }).Attributes | Select-Object @{Name = 'FabricName'; Expression = {(($_.FabricName | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()}}
			) #-PassThru

			Add-Member -MemberType NoteProperty -InputObject $hba -Name PortWWN -Value (
				($port | Where-Object { $_.InstanceName -eq $hba.InstanceName }).Attributes | Select-Object @{Name = 'PortWWN'; Expression = {(($_.PortWWN | ForEach-Object {"{0:x2}" -f $_}) -join ":").ToUpper()}} 
			) -PassThru
		}
	}
}

#Get-HBAInfo -ComputerName ${Env:COMPUTERNAME} | Select-Object * -ExpandProperty 'FabricName' -ExcludeProperty 'FabricName' | Select-Object * -ExpandProperty 'PortWWN' -ExcludeProperty 'PortWWN' | Out-GridView -Title 'Host Bus Adapter Information'
#Get-HBAInfo -ComputerName ${Env:COMPUTERNAME} | Select-Object * -ExpandProperty 'FabricName' -ExcludeProperty 'FabricName' | Select-Object * -ExpandProperty 'PortWWN' -ExcludeProperty 'PortWWN'# | Export-Csv -Path 'C:\Scripts\HBAInfo-dn3wpdbacl100c-x6.csv' -NoTypeInformation

$ClustersHBAInfo = @()

## Has the following headers in it:
##	'ClusterName','NodeName','DNSName'
$SQLServers = @(Import-Csv -Path 'C:\Scripts\SQLSvr-List-Driver-Queries.csv')

foreach ($SQLServer in $SQLServers)
{
	$SrvDnsName = $SQLServer.DNSName
	#Write-Host -ForegroundColor Cyan $SrvDnsName

	if (Test-Connection -ComputerName $SrvDnsName -Count 1 -Quiet)
	{
		Write-Host -ForegroundColor Green "Connection tested good to: $($SrvDnsName)"
		$ClustersHBAInfo += Get-HBAInfo -ComputerName $SrvDnsName | Select-Object * -ExpandProperty 'FabricName' -ExcludeProperty 'FabricName' | Select-Object * -ExpandProperty 'PortWWN' -ExcludeProperty 'PortWWN'
	}
	else
	{
		Write-Host -ForegroundColor Magenta "Connection tested BAD to: $($SrvDnsName)"
	}
}

#$ClustersHBAInfo | Export-Csv -Path "C:\Scripts\ClustersHBAInfo_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation
#$ClustersHBAInfo | Select-Object 'PSComputerName','Active','DriverName','DriverVersion','FirmwareVersion','Model','OptionROMVersion','FabricName','PortWWN','NodeWWN' | Export-Csv -Path "C:\Scripts\ClustersHBAInfo_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation
#$ClustersHBAInfo | Select-Object 'PSComputerName','DriverName','DriverVersion','FirmwareVersion','Model','OptionROMVersion','FabricName','PortWWN','NodeWWN' | Export-Csv -Path "C:\Scripts\ClustersHBAInfo_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation
#$ClustersHBAInfo | Select-Object 'PSComputerName','DriverName','DriverVersion','FirmwareVersion','Model','OptionROMVersion','PortWWN','NodeWWN' | Export-Csv -Path "C:\Scripts\ClustersHBAInfo_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation
#$ClustersHBAInfo | Select-Object 'PSComputerName','Model','OptionROMVersion','DriverName','DriverVersion','FirmwareVersion','PortWWN','NodeWWN' | Export-Csv -Path "C:\Scripts\ClustersHBAInfo_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation
#@{Name = "DayOfWeek"; Expression = {$_.LastWriteTime.DayofWeek}}

$ClustersHBAInfo | Select-Object @{Name = "ComputerName"; Expression = {$_.PSComputerName}}, 'Model', @{Name = "HBA Bios"; Expression = {"$($_.OptionROMVersion)"}}, 'DriverName', @{Name = "Stor Miniport Driver Version"; Expression = {$_.DriverVersion}}, 'FirmwareVersion', 'PortWWN', 'NodeWWN' | Export-Csv -Path "C:\Scripts\ClustersHBAInfo_$((Get-Date -Format 'yyyy.MM.dd.ddd.HHmm').ToString()).csv" -NoTypeInformation
