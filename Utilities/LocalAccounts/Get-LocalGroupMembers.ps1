$servers = Get-Content 'c:\input\test.txt'
$output = 'c:\output\test.csv'

foreach($server in $servers)
{
	$group =[ADSI]"WinNT://$server/Administrators"
	$members = @($group.psbase.Invoke("Members"))
	$results = $members | ForEach-Object {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null) }
}

$results >> $output
