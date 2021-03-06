##################################################################################################################
##################################################################################################################
# This script will use a CSV file, and make you select the template and customization to build VM's for vCenter
# Format of the CSV File is VMName,IPAddress,SubnetMask,Gateway,Pool/Cluster,VLAN,LOCATION,Description
#
# Created By: Steven Shell
# Created : July 2013
#
##################################################################################################################
##################################################################################################################

#region Import Assemblies for GUI
[Void][System.Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[Void][System.Reflection.Assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[Void][System.Reflection.Assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
#endregion Import Assemblies for GUI
[System.Windows.Forms.Application]::EnableVisualStyles()
 
Function CheckSnapins()
{
	$errCount = 0
	$snapins = Get-PSSnapin

	if (!($snapins -match "VMware.VimAutomation.Core"))
	{
			##snapin not loaded - try to load
			if (LoadSnapin "VMware.VimAutomation.Core")
			{}
			else
			{
				HelpCSV "VMware.VimAutomation.Core not loaded Cannot Continue"
				$errCount ++
			}
	}

	if (!($snapins -match "VMware.DeployAutomation"))
	{
			##snapin not loaded - try to load
			if (LoadSnapin "VMware.DeployAutomation")
			{}
			else
			{
				HelpCSV "VMware.DeployAutomation not loaded Cannot Continue"
				$errCount ++
			}
	}

	if (!($snapins -match "VMware.ImageBuilder"))
	{
			##snapin not loaded - try to load
			if (LoadSnapin "VMware.ImageBuilder")
			{}
			else
			{
				HelpCSV "VMware.ImageBuilder not loaded Cannot Continue"
				$errCount ++
			}
	}

	if (!($snapins -match "VMware.VimAutomation.License"))
	{
			##snapin not loaded - try to load
			if (LoadSnapin "VMware.VimAutomation.License")
			{}
			else
			{
				HelpCSV "VMware.VimAutomation.License not loaded Cannot Continue"
				$errCount ++
			}
	}

	if ($errCount -eq 0)
	{
		return $true
	}
	else
	{
		return $false
	}
} #end function CheckSnapins

Function LoadSnapin($snapin)
{
	$loaded = $false
	Add-PSSnapin $snapin -ErrorAction SilentlyContinue

	if ((Get-PSSnapin $snapin) -eq $null)
	{
		$loaded = $false
		return $loaded
	}
	else
	{
		$loaded = $true
		return $loaded
	}
} #end function LoadSnapin

Function Get-FileName($initialDirectory)
{
	$diffs = Get-ChildItem $initialDirectory -Filter "*.csv"
	Write-Host "Please Select File :"
	$i = 1

	foreach ($d in $diffs)
	{
		Write-Host $i")" $d
		$i++
	}

	$fileSelect = Read-Host "Enter Selection Number "
	$returnFile = $initialDirectory + "\" + $diffs[($fileSelect - 1)]

	return $returnFile
} #end function Get-FileName

Function Create-RadioBox($ItemList, $varName)
{
	$FormName = $varName
	$Form1 = New-Object System.Windows.Forms.Form
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	$ItemCount = $ItemList.Count

	if ($ItemCount -gt 15)
	{
		$FormSize = 500
		$Form1.AutoScroll = $true
		$Form1.SetAutoScrollMargin(25, 25)
		
	}
	else
	{
		$FormSize = (100) + ($ItemCount * 25)
	}

	$Form1.ClientSize = New-Object System.Drawing.Size(370, $FormSize)
	$Form1.DataBindings.DefaultDataSourceUpdateMode = [System.Windows.Forms.DataSourceUpdateMode]::OnValidation
	$Form1.Name = "Form1"
	$Form1.Text = $FormName
	$Form1.AutoScroll = $true
	#$Form1.Location = New-Object System.Drawing.Size(600, 600)
	$Form1.StartPosition = "CenterScreen"

	$Buttons = @()
	$Location = 25

	foreach ($Item in $ItemList)
	{
		$NewRBtn = New-Object System.Windows.Forms.RadioButton
		$NewRBtn.Name = $Item.Name
		$NewRBtn.Text = $Item.Name
		$NewRBtn.Size = New-Object System.Drawing.Size(300, 24)
		$NewRBtn.Location = New-Object System.Drawing.Size(25, $Location)

		$Location += 25
		$Buttons += $NewRBtn
	}

	foreach ($Btn in $Buttons)
	{
		$Form1.Controls.Add($Btn)
	}

	$event = {
		foreach ($Button in $Buttons)
		{
			if ($Button.Checked)
			{
				Set-Variable -Name $varName.ToString() -Value $Button.Text -Scope 1
				$Form1.Hide()
				$Form1.Close()
			}
		}
	}

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(50, ($Location + 25))
	$OKButton.Size = New-Object System.Drawing.Size(75, 23)
	$OKButton.Text = "OK"
	$OKButton.Add_Click($event)
	$Form1.Controls.Add($OKButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(125, ($Location + 25))
	$CancelButton.Size = New-Object System.Drawing.Size(75, 23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({Set-Variable -Name "continue" -Value $false -Scope 1; $Form1.Close()})
	$Form1.Controls.Add($CancelButton)

	$InitialFormWindowState = $Form1.WindowState
	$Form1.add_Load($FormEvent_Load)
	$Form1.ShowDialog()
}#end Function Create-RadioBox

Function HelpCSV($str)
{
	Write-Host "ERROR:" $str " Please Select a CSV file with the following format: VMName,IPAddress,SubnetMask,Gateway,Pool,VLAN,LOCATION,Description" -ForegroundColor Red
} #end function HelpCSV

Function CheckIP($ipAddress)
{
	try
	{
		$address = [System.Net.IPAddress]::Parse($ipAddress)
		return $true
	}
	catch
	{
		return $false
	}
} #end function CheckIP

Function CheckVLAN($vlan)
{
	if ($vlanNames -match $vlan)
	{
		return $true
	}
	else
	{
		return $false
	}
} #end function CheckVLAN

Function CheckCSV($csvFileToCheck)
{
	$errorCount = 0
	Import-Csv $csvFileToCheck | ForEach-Object {
		if ($_.VMName -eq $null)
		{
			HelpCSV "Name Not Found"
			$errorCount ++
			break
		}

		if ( !(CheckIP $_.IPAddress))
		{
			HelpCSV ($_.IPAddress + " IP not valid or in use")
			$errorCount ++
			break
		}

		if (!(CheckIP $_.SubnetMask))
		{
			HelpCSV ($_.IPAddress + " SubnetMask not valid")
			$errorCount ++
			break
		}

		if (!(CheckIP $_.Gateway))
		{
			HelpCSV ($_.Gateway + " Gateway not valid")
			$errorCount ++
			break
		}

		if ($_.VLAN -eq $null)
		{
			HelpCSV "VLAN Not Found"
			$errorCount ++
			break
		}

		if ($_.Pool -eq $null)
		{
			HelpCSV "VMware Cluster Pool Not Found"
			$errorCount ++
			break
		}
	}

	if ($errorCount -eq 0)
	{
		return $true
	}
	else
	{
		return $false
	}
} # end function #CheckCSV

Function Get-MyDataStore($clusterName)
{
	#get vmware cluster object
	$cluster = Get-Cluster -Name $clusterName
	#get all datastores associated with cluster
	$stores = $cluster.ExtensionData.Datastore
	$potentialdatastores = @()
	$datastores = @()
	$sdrsDisabled = @()
	$highUtilization = @()
	$allStores = @()

	#enumerate datastores, and get datastorecluster that is the parent
	foreach ($a in $stores)
	{
		$store = Get-Datastore -Id ($a.Type.ToString() + "-" + $a.Value.ToString())
		$allStores += $store
		$potentialdatastores += Get-DataStoreCluster -Id $store.ParentFolderID -ErrorAction SilentlyContinue
	}

	if ($potentialdatastores -ne $null)
	{
		#get unique datastoreclusters
		$myStore = $potentialdatastores | Sort-Object name | Get-Unique -AsString

		#of unique entries, look for datastore clusters that are not disabled, and have at least 70% free space
		foreach ($s in $myStore)
		{
			if ($s.sdrsAutomationLevel.ToString() -eq "Disabled")
			{
				$sdrsDisabled += $s
			}
			#check for at least 70% free space on the data cluster
			elseif (($s.FreeSpaceGB / $s.CapacityGB) -le .70)
			{
				$datastores += $s
			}
			else
			{
				$highUtilization += $s
			}
		}
	}

	if ($datastores -ne $null)
	{
		 $r = $datastores | Get-Random
		 return $r.Name.ToString()
	}
	else
	{
		return $null
	}
}

Function Show-MsgBox
{
	[CmdletBinding()]
	Param(
		[Parameter(
			Position = 0
			, Mandatory = $true
		)]
		[string]$Prompt,
		[Parameter(
			Position = 1
			, Mandatory = $false
		)]
		[string]$Title = "",
		[Parameter(
			Position = 2
			, Mandatory = $false
		)]
		[ValidateSet("Information", "Question", "Critical", "Exclamation")]
		[string]$Icon = "Information",
		[Parameter(
			Position = 3
			, Mandatory = $false
		)]
		[ValidateSet("OKOnly", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel")]
		[string]$BoxType = "OkOnly",
		[Parameter(
			Position = 4
			, Mandatory = $false
		)]
		[ValidateSet(1, 2, 3)]
		[int]$DefaultButton = 1
	)

	[Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") #| Out-Null

	switch ($Icon) {
		"Question" {$vb_icon = [Microsoft.VisualBasic.MsgBoxStyle]::Question }
		"Critical" {$vb_icon = [Microsoft.VisualBasic.MsgBoxStyle]::Critical}
		"Exclamation" {$vb_icon = [Microsoft.VisualBasic.MsgBoxStyle]::Exclamation}
		"Information" {$vb_icon = [Microsoft.VisualBasic.MsgBoxStyle]::Information}
	}

	switch ($BoxType) {
		"OKOnly" {$vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::OKOnly}
		"OKCancel" {$vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::OkCancel}
		"AbortRetryIgnore" {$vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::AbortRetryIgnore}
		"YesNoCancel" {$vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::YesNoCancel}
		"YesNo" {$vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::YesNo}
		"RetryCancel" {$vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::RetryCancel}
	}

	switch ($Defaultbutton) {
		1 {$vb_defaultbutton = [Microsoft.VisualBasic.MsgBoxStyle]::DefaultButton1}
		2 {$vb_defaultbutton = [Microsoft.VisualBasic.MsgBoxStyle]::DefaultButton2}
		3 {$vb_defaultbutton = [Microsoft.VisualBasic.MsgBoxStyle]::DefaultButton3}
	}

	$popuptype = $vb_icon -bor $vb_box -bor $vb_defaultbutton
	$ans = [Microsoft.VisualBasic.Interaction]::MsgBox($prompt, $popuptype, $title)
	return $ans
} #end function

#######################################################
#### Entry Point to Script #####
#######################################################
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$global:vCenter = $true
$global:continue = $true
$global:username = $null
$global:SelectedTemplate = $null
$global:SelectedCust = $null
$global:SelectCluster = $null
$global:SelectDatastore = $null
$global:dns = $null
$global:vlanNames = @()
$athensDomain = "athens\"
$stagingDomain = "athensqa\"
$wrkDomain = "wrk\"
$stagingDNS = @("10.52.164.53", "10.52.165.55")
$prodDNS = @("10.200.4.2", "10.200.4.3")
$wrkDNS = @("10.20.0.52", "10.20.0.53")
$prodLocation = $null
$stagingLocation = "Windows"
$wrkLocation = "Windows"
$VMLocation = $null
$vCenterServers = @("vcenter01sc.ecollegeqa.net", "esxmgmt01c.ecollege.net", "vcenter03.wrk.pad.pearsoncmg.com")
$vCenters = @()

if (CheckSnapins)
{
	$continue = $true
}
else
{
	break
}

foreach ($vc in $vCenterServers)
{
	$vcObject = New-Object System.Object
	$vcObject | Add-Member -MemberType NoteProperty -Name Name -Value $vc
	$vCenters += $vcObject
}

Create-RadioBox $vCenters "vCenter" | Out-Null

if ($continue)
{
	if ($vcenter -eq "vcenter01sc.ecollegeqa.net")
	{
		$dns = $stagingDNS
		$userName = $stagingDomain + [Environment]::UserName
		$VMLocation = $stagingLocation
	}

	if ($vcenter -eq "esxmgmt01c.ecollege.net")
	{
		$dns = $prodDNS
		$userName = $athensDomain + [Environment]::UserName
		$VMLocation = $prodLocation
	}

	if ($vcenter -eq "vcenter03.wrk.pad.pearsoncmg.com")
	{
		$dns = $wrkDNS
		$userName = $wrkDomain
		$VMLocation = $wrkLocation
	}
}
else
{
	break
}

if ($continue)
{
	$cred = Get-Credential -Credential $username -ErrorAction SilentlyContinue

	if ($cred -ne $null)
	{
		$mycreds = New-Object System.Management.Automation.PSCredential ($cred.UserName, $cred.Password)
	}
}
else
{
	break
}

#prompt user to select CSV file for VM creation
if (($continue) -and ($cred -ne $null))
{
	$file = Get-FileName -initialDirectory $mydir
}
else
{
	break
}

#Check csv file selected for VM Creation
if ($file -ne $null -and $continue -eq $true)
{
	$result = CheckCSV $file
}
else
{
	$result = $false
	HelpCSV
	break
}

# csv file succesfully verified
if ($result -eq $true)
{
	Write-Host "Successfully Verified: " $file
	Write-Host "Trying to Connect to: " $vcenter "..."

	#connect to vcenter server
	#[Void]
	$connection = Connect-VIServer -server $vcenter -protocol https -Credential $mycreds

	#get the templates from selected vCenter
	$templates = Get-Template | Sort-Object Name

	#get the customizations from selected vCenter
	$customization = Get-OSCustomizationSpec | Sort-Object Name

	#get the available vlans from the selected vCenter
	$vlans = Get-VirtualPortGroup

	foreach ($v in $vlans)
	{
		$vlanNames += $v.Name
	}

	#display templates to for one to be selected
 	Create-RadioBox $templates "SelectedTemplate" | Out-Null

	if ($continue -eq $true -and $SelectedTemplate -ne $null)
	{
		Create-RadioBox $Customization "SelectedCust" | Out-Null

		if ($SelectedCust -eq $null)
		{
			break
		}
	}
	else
	{
		break
	}

	$result = Show-MsgBox "Template:`t`t$SelectedTemplate `nCustomization:`t$SelectedCust `nDNS:`t`t$dns `nFile:`t`t$file " "Please Confirm" "Question" "OKCancel"

	if ($result -eq "Ok")
	{
		Write-Host "Creating VM's ..." -ForegroundColor DarkGreen
		$note = "Created from Template " + $SelectedTemplate + " Using customization " + $SelectedCust

		Import-Csv -path $file | ForEach-Object {
			if (CheckVLAN $_.VLAN)
			{
				if ($_.pool -eq "R810-01" -or $_.pool -eq "R810-02")
				{
					# find a datastore with at least 30% free space that services the R810 clusters
					$NewDatastore = Get-Datastore | Where-Object { $_.FreespaceMB  -gt ($_.CapacityMB*.30) -and $_.Name -like "VMAX*" -and $_.Datacenter -like "Cornell" } | Get-Random
				}

				if ($_.pool -eq "10G Cluster")
				{
					#use clustered disk pool
					$NewDatastore = "CRN-ESX-NFS-TIER1-WIN"
				}

				if ($_.pool -eq "Utility Cluster")
				{
					#use clustered disk pool
					$NewDatastore = "CRN-ESX-NFS-TIER1-WIN","CRN-ESX-NFS-TIER1-WIN" | Get-Random
				}

				if ($_.pool -eq "R810-03")
				{
					#use clustered disk pool
					#$NewDatastore = "CRN-PROD-NFS-01"
					$NewDatastore = Get-MyDataStore $_.pool
					#$NewDatastore = Get-DatastoreCluster -Location "Cornell"  | Get-Datastore | Where {$_.FreespaceGB -gt ($_.CapacityGB*.40) -and !(($_.Name).ToString().Contains("LIN"))} | Get-Random
				}

				if ($_.pool -eq "R810-04")
				{
					#use clustered disk pool
					#$NewDatastore = "CRN-PROD-NFS-01"
					$NewDatastore = Get-MyDataStore $_.pool
					#$NewDatastore = Get-DatastoreCluster -Location "Cornell"  | Get-Datastore | Where {$_.FreespaceGB -gt ($_.CapacityGB*.40) -and !(($_.Name).ToString().Contains("LIN"))} | Get-Random
				}

				if ($_.pool -eq "NFS Cluster")
				{
					#use clustered disk pool
					$NewDatastore = "CRN-ESX-NFS-TIER1-WIN02"
				}

				if ($_.pool -eq "R720-01-Utility")
				{
					#use clustered disk pool
					$NewDatastore = "NFS-CLUSTER01"
				}

				if ($_.pool -eq "R720-02")
				{
					#use clustered disk pool
					$NewDatastore = Get-MyDataStore $_.pool
					#$NewDatastore = "NFS-NETAPP-ESX54"
				}

				if ($_.pool -eq "R720-03")
				{
					#use clustered disk pool
					#$NewDatastore = "NFS-CLUSTER02"
					$NewDatastore = Get-MyDataStore $_.pool
					#$NewDatastore = Get-DatastoreCluster -Location "Cornell-Staging"| Where{($_.Name).ToString().Contains("NFS-CLUSTER02")}  | Get-Datastore | Where {$_.FreespaceGB -gt ($_.CapacityGB*.40) -and (($_.Name).ToString().Contains("ESX"))} | Get-Random
				}

				if ($_.pool -eq "Arapahoe HA Cluster")
				{
					#use clustered disk pool for Arapahoe Cluster
					$NewDatastore = Get-Datastore | Where-Object { $_.Name -like "ESX-NFS-WIN01" -and $_.Datacenter -like "Arapahoe" }
					#$NewDatastore = "ESX-NFS-WIN01"
				}

				if ($_.pool -eq "Test01")
				{
					#test cluster
					$NewDatastore = Get-Datastore | Where-Object { $_.FreespaceMB  -gt ($_.CapacityMB*.30) -and $_.Name -like "VIOLIN*"} | Get-Random
				}

				if ($_.pool -eq "R720-STAGE")
				{
					#test cluster
					$NewDatastore = "WAL-STAGE-01"
				}

				Write-Host "Creating VM" $_.VMName `t $_.IPAddress `t $SelectedTemplate `t $SelectedCust `t $_.Pool `t $NewDatastore `t $_.SubnetMask `t $_.Gateway `t $_.Location `t $dns -ForegroundColor Green

				Get-OSCustomizationNicMapping -OSCustomizationSpec $SelectedCust | Set-OSCustomizationNicMapping -IpMode UseStaticIP -IpAddress $_.IPAddress  -SubnetMask $_.SubnetMask -DefaultGateway $_.Gateway -Dns $dns

				if ($VMLocation -ne $null)
				{
					$vm = New-VM -Name $_.VMName -OSCustomizationSpec $SelectedCust -Template $SelectedTemplate -resourcepool $_.Pool -Datastore $NewDatastore -location $VMLocation -Notes $note -Verbose
				    Get-VM $_.VMName | get-networkadapter | set-networkadapter -networkname $_.VLAN -startconnected:$true -confirm:$false
				}
				else
				{
					$vm = New-VM -Name $_.VMName -OSCustomizationSpec $SelectedCust -Template $SelectedTemplate -resourcepool $_.Pool -Datastore $NewDatastore -Notes $note -Verbose
				    Get-VM $_.VMName | get-networkadapter | set-networkadapter -networkname $_.VLAN -startconnected:$true -confirm:$false
				}

				if ( (((Get-VM -Name $_.VMName | Get-VirtualPortGroup | Select-Object Name).Name).ToString()).ToLower() -eq ($_.VLAN).ToLower())
				{
			    	Start-VM -VM $_.VMName
				}
				else
				{
					Write-Host "Not Starting VM. VLAN settings incorrect. Please check VLAN and manually Start VM" `t $_.VMName `t $_.VLAN `t $_.IPAddress `t $_.Gateway -ForegroundColor Red
				}
			}
			else
			{
				Write-Host "Unable to find VLAN: " $_.VLAN `t $_.VMName `t $_.IPAddress `t $_.Gateway -ForegroundColor Red
			}
		}
	}
}
else
{
	break
}

if ($connection -ne $null)
{
	Disconnect-VIServer -Server $connection -Confirm:$false
}
