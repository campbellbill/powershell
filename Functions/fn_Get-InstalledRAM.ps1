###################################################################################################################
#
#	Script Name          : fn_Get-InstalledRAM.ps1
#	Written by           : Bill Campbell
#	Written on           : Oct 07, 2014
#	Added to Script Repo : 07 Oct 2014
#	Last Modified        : 2014.Oct.07
#
#	Version              : 2014.10.07.02
#	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
#	Version Example      : If this is the 6th revision/save on May 22 2013 then RR would be '06' and the version number will be formatted as follows: 2013.05.22.06
#
#	Purpose              : Execute a series of commands using Powershell to get the amount of installed RAM in the specified computer.
#
#	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#	#! IMPORTANT NOTICE:																						 !#
#	#!		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK FROM THE USE OF,	 !#
#	#!		OR THE RESULTS RETURNED FROM THE USE OF, THIS CODE REMAINS WITH THE USER.							 !#
#	#! 																											 !#
#	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#	#! 																											 !#
#	#! 						Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force							 !#
#	#! 																											 !#
#	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
#
#	Notes                : Use this area for general purpose use notes for the script.
#							Gets a list of all current Environment Variables:			Get-ChildItem Env:
#							Change Directories to the Current Users 'Desktop' folder:	Set-Location ${Env:USERPROFILE}\Desktop
#							Some Commonly used Environment Variables:					Write-Output ${Env:USERPROFILE}; Write-Output ${Env:COMPUTERNAME}; Write-Output ${Env:USERNAME};
#							Some Commonly used commands:								[Guid]::NewGuid(); [Environment]::OSVersion.Version
#																						$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))
#
###################################################################################################################

##--------------------------------------------------------------------------
##  FUNCTION.......:  fn_Get-InstalledRAM
##  PURPOSE........:  Returns the amount of installed RAM in the specified
##                    computer.
##  REQUIREMENTS...:  PowerShell v2
##  NOTES..........:  
##--------------------------------------------------------------------------
Function fn_Get-InstalledRAM
{
	<#
		.SYNOPSIS
			Returns the amount of installed RAM in the specified computer.

		.DESCRIPTION
			This Function uses the Win32ComputerSystem WMI Object to query the
			specified computer for how much total RAM is installed (this may differ
			from what is available to the OS).

		.PARAMETER ComputerName
			The name of IP address of the computer to query.

		.PARAMETER Full
			This optional switch will return information about each DIMM installed
			in the specified computer including Slot Label, Capacity, Memory Type,
			and Clock Speed.

		.EXAMPLE
			C:\\\\PS>fn_Get-InstalledRAM SVR01

			This will return the amount of RAM installed in the computer named
			"SVR01".

		.NOTES
			NAME......:  fn_Get-InstalledRAM
			AUTHOR....:  Joe Glessner
			LAST EDIT.:  23MAR12
			CREATED...:  11APR11

		.LINK
			http://joeit.wordpress.com/
	#>

	Param(
		[Parameter(
			Mandatory = $true,
			ValueFromPipeLine = $false,
			Position = 0
		)]
		[string]$ComputerName = ${Env:COMPUTERNAME},
		[Switch]$Full
	)  

	$Computers = Get-WMIObject -Class "Win32_ComputerSystem" -Namespace "root\\\\CIMV2" -ComputerName $ComputerName
	$W32_PM = Get-WMIObject -Class "Win32_PhysicalMemory" -Namespace "Root\\\\CIMV2" -ComputerName $ComputerName
	$W32_PMA = Get-WMIObject -Class "Win32_PhysicalMemoryArray" -Namespace "Root\\\\CIMV2" -ComputerName $ComputerName

	foreach ($Computer in $Computers)
	{
		$RAM = [Math]::Round($Computer.TotalPhysicalMemory / 1024/1024/1024, 0)

		Write-Host "Total Physical Memory.: " $RAM "GB"
		Write-Host "Model.................: " $objItem.Model

		if ($Full)
		{
			foreach ($Dev in $W32_PM)
			{
				$Speed = $Dev.Speed
				$Type = $Dev.MemoryType
				$Slot = $Dev.DeviceLocator
				$Capacity = [Math]::Round($Dev.Capacity / 1024/1024/1024, 0)

				Write-Host "Slot........: " $Slot
				#Write-Host "Memory Type.: " $Type

				switch ($Type)
				{
					0	{
						Write-Host "Memory Type.:  Unknown"
					}
					1	{
						Write-Host "Memory Type.:  Other"
					}
					2	{
						Write-Host "Memory Type.:  DRAM"
					}
					3	{
						Write-Host "Memory Type.:  Synchronous DRAM"
					}
					4	{
						Write-Host "Memory Type.:  Cache DRAM"
					}
					5	{
						Write-Host "Memory Type.:  EDO"
					}
					6	{
						Write-Host "Memory Type.:  EDRAM"
					}
					7	{
						Write-Host "Memory Type.:  VRAM"
					}
					8	{
						Write-Host "Memory Type.:  SRAM"
					}
					9	{
						Write-Host "Memory Type.:  RAM"
					}
					10	{
						Write-Host "Memory Type.:  ROM"
					}
					11	{
						Write-Host "Memory Type.:  Flash"
					}
					12	{
						Write-Host "Memory Type.:  EEPROM"
					}
					13	{
						Write-Host "Memory Type.:  FEPROM"
					}
					14	{
						Write-Host "Memory Type.:  EPROM"
					}
					15	{
						Write-Host "Memory Type.:  CDRAM"
					}
					16	{
						Write-Host "Memory Type.:  3DRAM"
					}
					17	{
						Write-Host "Memory Type.:  SDRAM"
					}
					18	{
						Write-Host "Memory Type.:  SGRAM"
					}
					19	{
						Write-Host "Memory Type.:  RDRAM"
					}
					20	{
						Write-Host "Memory Type.:  DDR"
					}
					21	{
						Write-Host "Memory Type.:  DDR-2"
					}
				}#END: Switch

				Write-Host "Capacity....: " $Capacity "(GB)"
				Write-Host "Speed.......: " $Speed 
				#Write-Host "ECC Type....: " $W32_PMA.MemoryErrorCorrection
			}#END: foreach ($Dev...

			switch ($W32_PMA.MemoryErrorCorrection)
			{
				0 {
					Write-Host "ECC Type....:  Reserved"
				}
				1 {
					Write-Host "ECC Type....:  Other"
				}
				2 {
					Write-Host "ECC Type....:  Unknown"
				}
				3 {
					Write-Host "ECC Type....:  None"
				}
				4 {
					Write-Host "ECC Type....:  Parity"
				}
				5 {
					Write-Host "ECC Type....:  Single-bit ECC"
				}
				6 {
					Write-Host "ECC Type....:  Multi-bit ECC"
				}
				7 {
					Write-Host "ECC Type....:  CRC"
				}
			}#END: Switch
		}#END if ($Full)
	}#END: foreach ($Computer...
}#END: Function fn_Get-InstalledRAM
