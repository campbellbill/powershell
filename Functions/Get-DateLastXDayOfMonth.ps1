#Requires -Version 2.0
####################################################################################################################
##
##	Script Name          : Get-DateLastXDayOfMonth.ps1
##	Author               : Bill Campbell
##	Copyright			 : © 2015 Bill Campbell. All rights reserved.
##	Created On           : Mar 31, 2015
##	Added to Script Repo : 31 Mar 2015
##	Last Modified        : 2015.Mar.31
##
##	Version              : 2015.03.31.03
##	Version Notes        : Version format is a date taking the following format: yyyy.MM.dd.RN		- where RN is the Revision Number/save count for the day modified.
##	Version Example      : If this is the 6th revision/save on May 22 2013 then RR would be '06' and the version number will be formatted as follows: 2013.05.22.06
##
##	Purpose              : Execute a series of commands using Powershell for the purpose of calculating the date of the last specified weekday of any given month, in any given year.
##
##	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
##	#! IMPORTANT NOTICE:																						 !#
##	#!		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK FROM THE USE OF,	 !#
##	#!		OR THE RESULTS RETURNED FROM THE USE OF, THIS CODE REMAINS WITH THE USER.							 !#
##	#! 																											 !#
##	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
##	#! 																											 !#
##	#! 						Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force							 !#
##	#! 																											 !#
##	#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#!#
##
##	Notes                : Use this area for general purpose use notes for the script.
##							Gets a list of all current Environment Variables:			Get-ChildItem Env:
##							Change Directories to the Current Users 'Desktop' folder:	Set-Location ${Env:USERPROFILE}\Desktop
##							Some Commonly used Environment Variables:					Write-Output ${Env:USERPROFILE}; Write-Output ${Env:COMPUTERNAME}; Write-Output ${Env:USERNAME};
##							Some Commonly used commands:								[Guid]::NewGuid(); [Environment]::OSVersion.Version
##																						$([Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments))
##
##	Found original function at:
##		http://www.poshpete.com/powershell/finding-the-last-friday-of-a-month-or-any-other-day
##
####################################################################################################################
## Original Function.
#	Function Find-LastXOfMonth()
#	{
#	Param(
#	[parameter(Mandatory=$true)][String]$DayName,
#	[int]$Month,
#	[int]$Year
#	)
#	 
#	 
#	If($Month)
#	{
#		if($Year)
#		{
#	 
#			$LastDayOfMonth = (Get-Date -Year $Year -Month $Month -Day 1).AddMonths(1).AddDays(-1)
#		}
#		Else
#		{
#			$LastDayOfMonth = (Get-Date -Year (Get-Date).Year -Month $Month -Day 1).AddMonths(1).AddDays(-1)
#		}
#	 
#	}
#	Else
#	{
#	 
#		$LastDayOfMonth = (Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day 1).AddMonths(1).AddDays(-1)
#	}
#	 
#	$Answer = $Null
#	 
#	If($LastDayOfMonth.DayOfWeek -eq $DayName)
#	{
#		$Answer = $LastDayOfMonth
#	 
#	}
#	Else
#	{
#		While($Answer -eq $Null)
#		{
#			$LastDayOfMonth = $LastDayOfMonth.AddDays(-1)
#			If($LastDayOfMonth.DayOfWeek -eq $DayName)
#			{
#				$Answer = $LastDayOfMonth
#			}
#	 
#		}
#	}
#	Return $Answer
#	}

Function Get-DateLastXDayOfMonth()
{
	<#
		.SYNOPSIS
			Calculates the date of the last specified weekday of any given month, in any given year and returns the result to you as a date time type/object.

		.DESCRIPTION
			The Get-DateLastXDayOfMonth function does the following:

			Calculates the date of the last Monday/Tuesday/Wednesday..etc of any given month, in any given year and returns the result to you as a date time type/object.

		.PARAMETER DayName
			The name of the day of the week to be found.

		.PARAMETER Month
			The number which indicates the month of the year. If left blank it will use the current month.
			12 = December

		.PARAMETER Year
			The year in which you want to search. If left blank it will use the current year.

		.EXAMPLE
			Get-DateLastXDayOfMonth -DayName 'Monday'

			Description
			-----------
				Calculates the date of the last Monday in the current month of the current year.

		.EXAMPLE
			Get-DateLastXDayOfMonth -DayName 'Wednesday' -Month 5 -Year 2013

			Description
			-----------
				Calculates the date of the last Wednesday in the month of May in the current year 2013.
				It returns 'Wednesday, May 29, 2013 1:26:10 PM'
	#>
	[CmdletBinding(
		SupportsShouldProcess = $true,
		ConfirmImpact = 'Medium'
	)]
	#region Function Parameters
		Param(
			[Parameter(
				Position = 0
				, Mandatory = $true
			)][string]$DayName
			, [int]$Month
			, [int]$Year
		)
	#endregion Function Parameters

	begin
	{
		if ($Month)
		{
			if ($Year)
			{
				$LastDayOfMonth = (Get-Date -Year $Year -Month $Month -Day 1).AddMonths(1).AddDays(-1)
			}
			else
			{
				$LastDayOfMonth = (Get-Date -Year (Get-Date).Year -Month $Month -Day 1).AddMonths(1).AddDays(-1)
			}
		}
		else
		{
			$LastDayOfMonth = (Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day 1).AddMonths(1).AddDays(-1)
		}

		$Answer = $null
	}
	process
	{
		if ($LastDayOfMonth.DayOfWeek -eq $DayName)
		{
			$Answer = $LastDayOfMonth
		}
		else
		{
			while($Answer -eq $null)
			{
				$LastDayOfMonth = $LastDayOfMonth.AddDays(-1)

				if ($LastDayOfMonth.DayOfWeek -eq $DayName)
				{
					$Answer = $LastDayOfMonth
				}
			}
		}
	}
	end
	{
		return $Answer
	}
}

$ObjDate = Get-DateLastXDayOfMonth -DayName 'Monday'
$ObjDate | Get-Member
$ObjDate
#Get-DateLastXDayOfMonth -DayName 'Wednesday' -Month 5 -Year 2013
