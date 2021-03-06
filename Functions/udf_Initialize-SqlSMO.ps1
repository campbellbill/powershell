function udf_Initialize-SqlSMO
{
	# Loads the SQL Server Management Objects (SMO)

	begin
	{
		$OriginalErrorActionPreference = $ErrorActionPreference
		$ErrorActionPreference = 'Stop'
		# SQL 2008 R2
		#$SqlPSreg = 'HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.SqlServer.Management.PowerShell.sqlps'
		# SQL 2012
		$SqlPSreg = 'HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.SqlServer.Management.PowerShell.sqlps110'
	}
	process
	{
		if (Get-ChildItem -Path $SqlPSreg -ErrorAction 'SilentlyContinue')
		{
		    throw 'SQL Server Provider for Windows PowerShell is not installed.'
		}
		else
		{
		    $SqlPSItem = Get-ItemProperty -Path $SqlPSreg
		    $SqlPSPath = [System.IO.Path]::GetDirectoryName($SqlPSItem.Path)
		}

		$SMOAssemblyList = @(
			'Microsoft.SqlServer.Management.Common'
			, 'Microsoft.SqlServer.Smo'
			, 'Microsoft.SqlServer.Dmf'
			, 'Microsoft.SqlServer.Instapi'
			, 'Microsoft.SqlServer.SqlWmiManagement'
			, 'Microsoft.SqlServer.ConnectionInfo'
			, 'Microsoft.SqlServer.SmoExtended'
			, 'Microsoft.SqlServer.SqlTDiagM'
			, 'Microsoft.SqlServer.SString'
			, 'Microsoft.SqlServer.Management.RegisteredServers'
			, 'Microsoft.SqlServer.Management.Sdk.Sfc'
			, 'Microsoft.SqlServer.SqlEnum'
			, 'Microsoft.SqlServer.RegSvrEnum'
			, 'Microsoft.SqlServer.WmiEnum'
			, 'Microsoft.SqlServer.ServiceBrokerEnum'
			, 'Microsoft.SqlServer.ConnectionInfoExtended'
			, 'Microsoft.SqlServer.Management.Collector'
			, 'Microsoft.SqlServer.Management.CollectorEnum'
			, 'Microsoft.SqlServer.Management.Dac'
			, 'Microsoft.SqlServer.Management.DacEnum'
			, 'Microsoft.SqlServer.Management.Utility'
		)

		foreach ($SMOAssembly in $SMOAssemblyList)
		{
		    Write-Host "Loading SQL Server Management Objects (SMO) Assembly >>---> '$($SMOAssembly)'"
			#$SMOAssembly = [System.Reflection.Assembly]::LoadWithPartialName($SMOAssembly)
			[void][System.Reflection.Assembly]::LoadWithPartialName($SMOAssembly)
		}

		#Push-Location -Path $SqlPSPath
		#Update-FormatData -PrependPath SQLProvider.Format.ps1xml 
		#Pop-Location
	}
	end
	{
		$ErrorActionPreference = $OriginalErrorActionPreference
	}
}
