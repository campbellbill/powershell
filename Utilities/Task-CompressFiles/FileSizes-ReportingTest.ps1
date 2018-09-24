[string]$SourceDirectory = "\\stor22\datadump\kaplan\ER_Reports"
[string]$Extension = "csv"
$CutOffDate = (Get-Date).AddDays(-1)
[int]$aTtlFiles = 0
[int]$TtlFiles = 0
[decimal]$TtlFileSize = 0
[decimal]$KbTtlFileSz = 0
[decimal]$MbTtlFileSz = 0

$DirectoryFiles = @(Get-ChildItem -Path $SourceDirectory | Where-Object { ($_.Extension -like "*$($Extension)") -and ($_.CreationTime -gt $CutOffDate) } | Select-Object Name,Length,CreationTime)	#,BaseName | Format-Table -AutoSize

$DirectoryFiles | Format-Table -AutoSize
$aTtlFiles = $DirectoryFiles.Count

Write-Host -ForegroundColor Cyan	"-----------------------------------------------------------------------------------------------------"
Write-Host -ForegroundColor Cyan	" Variable Name        : Value"
Write-Host -ForegroundColor Cyan	"--------------------- : -----------------------------------------------------------------------------"

ForEach ($File in $DirectoryFiles)
{
	$bFileSz = $File.Length
	$TtlFileSize += $bFileSz
	#$KbFileSz = ($bFileSz / 1KB)
	$KbFileSz = [Math]::Round(($bFileSz / 1024), 6)		# 1 * 1024 = 1024 = 1KB
	#$MbFileSz = ($bFileSz / 1MB)
	$MbFileSz = [Math]::Round(($bFileSz / 1048576), 6)	# 1 * 1024 * 1024 = 1048576 = 1MB
	#$GbFileSz = ($bFileSz / 1GB)
	#$GbFileSz = [Math]::Round(($Size / 1073741824), 6)	# 1 * 1024 * 1024 * 1024 = 1073741824 = 1GB
	$KbTtlFileSz += $KbFileSz
	$MbTtlFileSz += $MbFileSz

	Write-Host -ForegroundColor Magenta	" File.Name            : $($File.Name)"
	#Write-Host -ForegroundColor Magenta	"--------------------- : -----------------------------------------------------------------------------"
	Write-Host -ForegroundColor Magenta	" bFileSz              : $($bFileSz)"
	Write-Host -ForegroundColor Magenta	" KbFileSz             : $($KbFileSz)"
	Write-Host -ForegroundColor Magenta	" MbFileSz             : $($MbFileSz)"
	#Write-Host -ForegroundColor Magenta	"--------------------- : -----------------------------------------------------------------------------"
	#Write-Host -ForegroundColor Magenta	" File.CreationTime    : $($File.CreationTime)"
	Write-Host -ForegroundColor Cyan	"--------------------- : -----------------------------------------------------------------------------"
	#Write-Host -ForegroundColor Cyan	"-----------------------------------------------------------------------------------------------------"
	#Write-Host ""

	$TtlFiles++
}

#$TtlFileSizeKB = ($TtlFileSize / 1KB)
$TtlFileSizeKB = [Math]::Round(($TtlFileSize / 1024), 6)
#$TtlFileSizeMB = ($TtlFileSize / 1MB)
$TtlFileSizeMB = [Math]::Round(($TtlFileSize / 1048576), 6)

#Write-Host -ForegroundColor Cyan	"-----------------------------------------------------------------------------------------------------"
#Write-Host -ForegroundColor Cyan	" Variable Name        : Value"
#Write-Host -ForegroundColor Cyan	"--------------------- : -----------------------------------------------------------------------------"
Write-Host -ForegroundColor Yellow	" TtlFileSize          : $($TtlFileSize) - in bytes - from the cumulative 'Length' of each file"
Write-Host -ForegroundColor Yellow	" KbTtlFileSz          : $($KbTtlFileSz) - in KB - calculated from the cumulative total of KbFileSz"
Write-Host -ForegroundColor Yellow	" TtlFileSizeKB        : $($TtlFileSizeKB) - in KB - calculated from TtlFileSize"
Write-Host -ForegroundColor Yellow	" MbTtlFileSz          : $($MbTtlFileSz) - in MB - calculated from the cumulative total of MbFileSz"
Write-Host -ForegroundColor Yellow	" TtlFileSizeMB        : $($TtlFileSizeMB) - in MB - calculated from TtlFileSize"
Write-Host -ForegroundColor Cyan	"-----------------------------------------------------------------------------------------------------"
Write-Host -ForegroundColor Green	" Total Number of Files in the array: $($aTtlFiles)"
Write-Host -ForegroundColor Green	" Total Number of Files proceesed: $($TtlFiles)"
Write-Host -ForegroundColor Cyan	"-----------------------------------------------------------------------------------------------------"
