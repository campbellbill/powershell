<#
		Begin
		{}
		Process
		{}
		End
		{}
#>
	Function fn_NewFunctionName()
	{
		[CmdletBinding(SupportsShouldProcess = $true)]
		#region Function Parameters
			Param(
				[Parameter(
					Position			= 0
					, Mandatory			= $true
					, ValueFromPipeline	= $false
					, HelpMessage		= '<Help Message...>'
				)][string]$FunctionVariable
			)
		#endregion Function Parameters
		## Funtion Usage:
		#	fn_NewFunctionName -FunctionVariable ''

		Begin
		{}
		Process
		{}
		End
		{}
	}
