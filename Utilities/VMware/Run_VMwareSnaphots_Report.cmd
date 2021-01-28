@echo off
rem Written by: Bill Campbell
rem Written on: 2013.May.15
rem Last Modified: 2013.May.16

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
rem Set the DEBUG variable to TRUE to enable the debuging statements, or FALSE to disable them.
set DEBUG=FALSE
set srcDir=D:\Scripts

if not exist %srcDir%\Reports (
	mkdir %srcDir%\Reports
	set rptFileLtn=%srcDir%\Reports
) else (
	set rptFileLtn=%srcDir%\Reports
)

if not exist %srcDir%\Logs (
	mkdir %srcDir%\Logs
	set logFileLtn=%srcDir%\Logs
) else (
	set logFileLtn=%srcDir%\Logs
)

set logName="%logFileLtn%\Get-VMwareSnaphots_%year%.%month%.%day%.%dow%.log"
set ASDSScript=%srcDir%\Get-VMwareSnaphots.ps1
set smtpSvr=mail.squaretwofinancial.com
set fromAddr=VMwareSnaphotsReport@squaretwofinancial.com

if /i %DEBUG%==TRUE (
	set toAddr=svralts@squaretwofinancial.com
	set ccAddr=itopscorner@squaretwofinancial.com
	set emlSubj=DEBUG-VMware Snaphots Report for %month% %year%
	set emlBody="Good Morning,\n\nAttached is the DEBUG-VMware Snaphots Report for %month% %year%. \n\nPlease review and advise on any changes. \n\nThank you, \n\nVMware Admins"
	set rptName1=DEBUG_Get-VMwareSnaphots
	set rptName2=%rptFileLtn%\DEBUG_Get-VMwareSnaphots_%year%.%month%.%day%.%dow%.html
) else (
	set toAddr=svralts@squaretwofinancial.com
	rem set toAddr="dgargan@squaretwofinancial.com; ggrosnick@squaretwofinancial.com; windowsadmins@squaretwofinancial.com"
	set emlSubj=VMware Snaphots Report for %month% %year%
	set emlBody="Good Morning,\n\nAttached is the VMware Snaphots Report for %month% %year%. \n\nPlease review and advise on any changes. \n\nThank you, \n\nVMware Admins"
	set rptName1=Get-VMwareSnaphots
	set rptName2=%rptFileLtn%\Get-VMwareSnaphots_%year%.%month%.%day%.%dow%.html
)

if /i %DEBUG%==TRUE (
	echo Variables used in this script are as follows:>> %logName%
	echo **********************************************************************>> %logName%
	echo **********************************************************************>> %logName%
	echo Date and time variables:>> %logName%
	echo                    year = %year% >> %logName%
	echo                   month = %month% >> %logName%
	echo                     day = %day% >> %logName%
	echo                     dow = %dow% >> %logName%
	echo                     hrs = %hrs% >> %logName%
	echo                     min = %min% >> %logName%
	echo                    ampm = %ampm% >> %logName%
	echo **********************************************************************>> %logName%
	echo Script Variables:>> %logName%
	echo  Current Directory "cd" = %cd% >> %logName%
	echo                   DEBUG = %DEBUG% >> %logName%
	echo                  srcDir = %srcDir% >> %logName%
	echo              rptFileLtn = %rptFileLtn% >> %logName%
	echo              logFileLtn = %logFileLtn% >> %logName%
	echo                 logName = %logName% >> %logName%
	echo              ASDSScript = %ASDSScript% >> %logName%
	echo                 smtpSvr = %smtpSvr% >> %logName%
	echo                fromAddr = %fromAddr% >> %logName%
	echo                  toAddr = %toAddr% >> %logName%
	echo                  ccAddr = %ccAddr% >> %logName%
	echo                 emlSubj = %emlSubj% >> %logName%
	echo                 emlBody = %emlBody% >> %logName%
	echo                rptName1 = %rptName1% >> %logName%
	echo                rptName2 = %rptName2% >> %logName%
	echo **********************************************************************>> %logName%
	echo **********************************************************************>> %logName%
)

:ScrStart
cd /d "%srcDir%"
echo Script Start Date and Time: >> %logName%
date /t >> %logName%
time /t >> %logName%
echo ...>> %logName%

call :runVMwareSnaphots %rptName1% >> %logName%
call :ScrEnd >> %logName%
goto :EOF

:runVMwareSnaphots
echo **********************************************************************
echo Starting report execution...
time /t
echo **********************************************************************

echo Executing %ASDSScript% ...
powershell -c "pushd %srcDir%; %ASDSScript% -PtlRptName %1"

rem wait before sending email to verify all processes have finished writing to the report...
sleep.exe 10

if /i %DEBUG%==TRUE (
	echo Sending email with report attached...
	echo sendEmail.exe -f %fromAddr% -t %toAddr% -cc %ccAddr% -u %emlSubj% -m %emlBody% -s %smtpSvr% -a %rptName2%
	sendEmail.exe -f %fromAddr% -t %toAddr% -cc %ccAddr% -u %emlSubj% -m %emlBody% -s %smtpSvr% -a %rptName2%
	rem sendEmail.exe -f %fromAddr% -t %toAddr% -cc %ccAddr% -u %emlSubj% -o message-file=%rptName2% -s %smtpSvr%
) else (
	echo Sending email with report attached...
	echo sendEmail.exe -f %fromAddr% -t %toAddr% -u %emlSubj% -m %emlBody% -s %smtpSvr% -a %rptName2%
	sendEmail.exe -f %fromAddr% -t %toAddr% -u %emlSubj% -m %emlBody% -s %smtpSvr% -a %rptName2%
	rem sendEmail.exe -f %fromAddr% -t %toAddr% -u %emlSubj% -o message-file=%rptName2% -s %smtpSvr%
)

echo **********************************************************************
echo Finished report execution...
time /t
echo **********************************************************************

goto :EOF

:ScrEnd
echo ...
echo Script End Date and Time:
date /t
time /t
echo ...
goto :EOF
