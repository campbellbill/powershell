# Get Directory Script was executed from
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$myFile = $myDir + "\Create-VMs-NewVMs.csv"
# Gather Username/PW Info for connection to vCenter
$vCenterSvr = "esxmgmt01c.ecollege.net"
$UserName = Read-Host "Enter Username for connection to $vCenterSvr"  
$UserPwd = Read-Host "Enter Password" -AsSecureString
$UserCreds = New-Object System.Management.Automation.PSCredential ($UserName, $UserPwd)

Connect-VIServer -Server $vCenterSvr -Protocol https -Credential $UserCreds

$customization = "ws-cust-specs"
$subnet = "255.255.255.0"
#$gateway = "10.200.21.1"
$template = "WIN2K8_R2_Standard-v2.5.1_06-10-2013"
#$datastore = "CRN-ESX-NFS-TIER1-WIN"
#$pool = "10G Cluster"
#$location = "Windows"
$dns = @("10.200.4.2","10.200.4.3")

#Get-OSCustomizationSpec -Name $customization | New-OSCustomizationSpec -Name Steven-Test -Type NonPersistent 
$osCust = Get-OSCustomizationSpec -Name $customization # -Type NonPersistent

Import-Csv -Path $myFile | ForEach-Object {
	If ($_.Pool -eq "R810-01" -or $_.Pool -eq "R810-02")
	{
		# find a datastore with at least 30% free space that services the R810 clusters
		$NewDatastore = Get-Datastore | Where-Object { $_.FreespaceMB  -gt ($_.CapacityMB*.30) -and $_.Name -like "VMAX*" -and $_.Datacenter -like "Cornell" } | Get-Random	
	}
	Else
	{
		$NewDatastore = $_.DataStore
	}

	Write-Host $_.VMName " Using Datastore: " $NewDatastore -ForegroundColor Green

    Get-OSCustomizationNicMapping -OSCustomizationSpec $osCust | Set-OSCustomizationNicMapping -IpMode UseStaticIP -IpAddress $_.IPAddress  -SubnetMask $_.SubnetMask -DefaultGateway $_.gateway -Dns $dns

	$vm = New-VM -Name $_.VMName -OSCustomizationSpec $osCust -Template $template -resourcepool $_.Pool -Datastore $NewDatastore # -location $_.Location

	Get-VM $_.VMName | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName $_.VLAN -StartConnected:$true -Confirm:$false

	Start-VM -VM $_.VMName 
}
