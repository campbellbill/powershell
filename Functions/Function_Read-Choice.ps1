﻿Function Read-Choice
{ 
	#.Synopsis
	#  Prompt the user for a choice, and return the (0-based) index of the selected item
	#.Parameter Message
	#  This is the prompt that will be presented to the user. Basically, the question you're asking.
	#.Parameter Choices
	#  An array of strings representing the choices (or menu items), with optional ampersands (&) in them to mark (unique) characters which can be used to select each item.
	#.Parameter ChoicesWithHelp
	#  A Hashtable where the keys represent the choices (or menu items), with optional ampersands (&) in them to mark (unique) characters which can be used to select each item, and the values represent help text to be displayed to the user when they ask for help making their decision.
	#.Parameter Default
	#  The (0-based) index of the menu item to select by default (defaults to zero).
	#.Parameter MultipleChoice
	#  Prompt the user to select more than one option. This changes the prompt display for the default PowerShell.exe host to show the options in a column and allows them to choose multiple times.
	#  Note: when you specify MultipleChoice you may also specify multiple options as the default!
	#.Parameter Caption
	#  An additional caption that can be displayed (usually above the Message) as part of the prompt
	#.Parameter Passthru
	#  Causes the Choices objects to be output instead of just the indexes
	#.Example
	#  Read-Choice "WEBPAGE BUILDER MENU"  "&Create Webpage","&View HTML code","&Publish Webpage","&Remove Webpage","E&xit"
	#.Example
	#  [bool](Read-Choice "Do you really want to do this?" "&No","&Yes" -Default 1)
	#  
	#  This example takes advantage of the 0-based index to convert No (0) to False, and Yes (1) to True. It also specifies YES as the default, since that's the norm in PowerShell.
	#.Example
	#  Read-Choice "Do you really want to delete them all?" @{"&No"="Do not delete all files. You will be prompted to delete each file individually."; "&Yes"="Confirm that you want to delete all of the files"}
	#  
	#  Note that with hashtables, order is not guaranteed, so "Yes" will probably be the first item in the prompt, and thus will output as index 0.  Because of thise, when a hashtable is passed in, we default to Passthru output.
	#
	# Found at:
	#	http://poshcode.org/2893
	#
	[CmdletBinding(DefaultParameterSetName = "HashtableWithHelp")]
	Param(
		[Parameter(
			Mandatory = $true
			, Position = 10
			, ParameterSetName = "HashtableWithHelp"
		)][Hashtable]$ChoicesWithHelp
		, [Parameter(
			Mandatory = $true
			, Position = 10
			, ParameterSetName = "StringArray"
		)][String[]]$Choices
		, [Parameter(
			Mandatory = $false
		)][string]$Caption = "Please choose!"
		, [Parameter(
			Mandatory = $false, Position = 0
		)][string]$Message = "Choose one of the following options:"
		, [Parameter(
			Mandatory = $false
		)][int[]]$Default  = 0
		, [Parameter(
			Mandatory = $false
		)][Switch]$MultipleChoice
		, [Parameter(
			Mandatory = $false
		)][Switch]$PassThru
	)
	Begin
	{
		If ($ChoicesWithHelp)
		{
			[System.Collections.DictionaryEntry[]]$Choices = $ChoicesWithHelp.GetEnumerator() | ForEach-Object { $_ }
		}
	}
	Process
	{
		$Descriptions = [System.Management.Automation.Host.ChoiceDescription[]]( $(
			If ($Choices -is [String[]])
			{
				ForEach ($Choice in $Choices)
				{
					New-Object System.Management.Automation.Host.ChoiceDescription $Choice
				}
			}
			Else
			{
				ForEach ($Choice in $Choices)
				{
					New-Object System.Management.Automation.Host.ChoiceDescription $Choice.Key, $Choice.Value
				}
			}
		) )

		# Passing an array as the $Default triggers multiple choice prompting.
		If (!$MultipleChoice)
		{
			[int]$Default = $Default[0]
		}

		[int[]]$Answer = $Host.UI.PromptForChoice($Caption, $Message, $Descriptions, $Default)

		If ($PassThru -or !($Choices -is [String[]]))
		{
			Write-Verbose "$Answer"
			Write-Output  $Descriptions[$Answer]
		}
		Else
		{
			Write-Output $Answer
		}
	}
}
