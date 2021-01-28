#\\pqashare\shares\Software\symantec
#
#
#Run setup - /QN for quiet install
#REBOOT = ReallySupress
#setup /s /v"/l*v log.txt /qn RUNLIVEUPDATE=0 REBOOT=REALLYSUPPRESS"

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$command = $myDir + "\psexec.exe"
$myFile = $mydir + "\AVInstall.txt"
$outFile = $myDir + "\Results.txt"

Get-Content $myFile |foreach-Object {

	if (Test-Connection $_ -Quiet -Count 1){
		
		$OStype = (Get-WmiObject -ComputerName $_ Win32_OperatingSystem).OSArchitecture
		Write-Host $_ " OS Type: " $OStype
					Copy-Item "\\sep01c\g$\Symantec\CustomInstallPackages\MyCompany\64-bit\My Company_Default Group\setup.exe" \\$_\c$\windows\temp
					Sleep -s 1
					$_ | Add-Content $outFile -PassThru | Write-Host -ForegroundColor Green
					$output = & $command \\$_ -e "C:\windows\temp\setup.exe" /s /qn 
					for($i=0; $i -lt $output.Length; $i++){
					$output[$i] | Add-Content $outFile -PassThru | Write-Host
					}	
		}
	
		else {
		"Unable to Connect To: $_ "  | Add-Content $outFile -PassThru | Write-Host -ForegroundColor Yellow
		}
	
	}