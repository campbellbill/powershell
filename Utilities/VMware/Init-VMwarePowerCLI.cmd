@echo off
rem Written by: Bill Campbell
rem Written on: 2012.Nov.23
rem Last Modified: 2012.Nov.23

:Variables
for /f "tokens=1-4 delims=/ " %%a in ('date/t') do (
	set dow=%%a
	set month=%%b
	set day=%%c
	set year=%%d
)
if %month%==01 (
	set month=Jan
	goto FixTimeFmt
) else if %month%==02 (
	set month=Feb
	goto FixTimeFmt
) else if %month%==03 (
	set month=Mar
	goto FixTimeFmt
) else if %month%==04 (
	set month=Apr
	goto FixTimeFmt
) else if %month%==05 (
	set month=May
	goto FixTimeFmt
) else if %month%==06 (
	set month=Jun
	goto FixTimeFmt
) else if %month%==07 (
	set month=Jul
	goto FixTimeFmt
) else if %month%==08 (
	set month=Aug
	goto FixTimeFmt
) else if %month%==09 (
	set month=Sep
	goto FixTimeFmt
) else if %month%==10 (
	set month=Oct
	goto FixTimeFmt
) else if %month%==11 (
	set month=Nov
	goto FixTimeFmt
) else if %month%==12 (
	set month=Dec
	goto FixTimeFmt
)

:FixTimeFmt
for /f "tokens=1-3 delims=: " %%e in ('time/t') do (
	set hrs=%%e
	set min=%%f
	set ampm=%%g
)
if %ampm%==AM (
	if %hrs%==12 set hrs=00& goto StartRun
) else if %ampm%==PM (
	if %hrs%==01 set hrs=13& goto StartRun
	if %hrs%==02 set hrs=14& goto StartRun
	if %hrs%==03 set hrs=15& goto StartRun
	if %hrs%==04 set hrs=16& goto StartRun
	if %hrs%==05 set hrs=17& goto StartRun
	if %hrs%==06 set hrs=18& goto StartRun
	if %hrs%==07 set hrs=19& goto StartRun
	if %hrs%==08 set hrs=20& goto StartRun
	if %hrs%==09 set hrs=21& goto StartRun
	if %hrs%==10 set hrs=22& goto StartRun
	if %hrs%==11 set hrs=23& goto StartRun
) else (
	goto StartRun
)

:StartRun
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -PSConsoleFile "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\vim.psc1" -Command ". 'C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1'; . 'C:\Users\bcampbell\Dropbox\Admin_Utils\Scripting\PowerShell_Scripts\VMware\Get-VmwareSnaphots.ps1'"
rem pause
