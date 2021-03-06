###################################################################################################################
#
#	Script Name          : fn_Get-RAM.ps1
#	Written by           : Bill Campbell
#	Written on           : Oct 07, 2014
#	Added to Script Repo : 07 Oct 2014
#	Last Modified        : 2014.Oct.07
#
#	Version              : 2014.10.07.01
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
function fn_Get-RAM
{
	param(
		[String[]]$Computers = ${Env:COMPUTERNAME}
	)

	foreach ($Computer in $Computers)
	{
		$ht = @{}
		Clear-Host

		$PhysicalRAM = (Get-WMIObject -Class Win32_PhysicalMemory -ComputerName $Computer | Measure-Object -Property Capacity -Sum | ForEach-Object { [Math]::Round(($_.Sum / 1GB), 2) })
		$ht.Add('Physical RAM (GB)', $PhysicalRAM)

		$OSRAM = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer | ForEach-Object {$_.TotalVisibleMemorySize, $_.FreePhysicalMemory}

		$ht.Add('Total Visable RAM (GB)', ([Math]::Round(($OSRAM[0] / 1MB), 4)))
		$ht.Add('Total Free RAM (GB)', ([Math]::Round(($OSRAM[1] / 1MB), 4)))

		$RAM = New-Object -TypeName PSObject -Property $ht
		Write-Output $RAM | Format-Table -AutoSize
	}
}

fn_Get-RAM
