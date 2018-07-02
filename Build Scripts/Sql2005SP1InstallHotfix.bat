@ECHO OFF
REM SQL SERVER 2005 SP1 Hotfix Rollup Install
REM Copyright © 2008 - 2014 Edward Vassie.  Distributed under Ms-Pl License 
REM
REM Created 14 Sep 2006 by Ed Vassie V1.0
REM
REM This file installs the SQL 2005 SP1 files in the required order.

REM Setup Default Parameters

SET SQLDIRPUSH=..\Additional Components\Service Packs\SP1
SET SQLFLAG=
SET SQLMAXRC=0
SET SQLRC=0

SET SQLMODE=%1
IF '%1' ==  '' SET SQLMODE=ACTIVE
IF /i '%1' == 'QUIET' SET SQLFLAG=/quiet
SET SQLTYPE=%2
IF '%2' ==  '' SET SQLTYPE=FULL

IF '%3' ==   '' SET SQLINSTANCE=/allinstances
IF '%3' NEQ  '' SET SQLINSTANCE=/instancename=%3

REM Validate Parameters

IF /i '%SQLMODE%' ==  'ACTIVE' GOTO :RUN
IF /i '%SQLMODE%' NEQ 'QUIET'  GOTO :HELP

:RUN

REM Run the SP1 Hotfix Rollup install

IF /i '%SQLTYPE%' == 'TOOLS' GOTO :TOOLS

ECHO %0 run at %DATE% %TIME% on server %COMPUTERNAME% by %USERNAME%
ECHO Parameters: MODE=%SQLMODE%, TYPE=%2, INSTANCE=%3
IF /i '%1' == 'QUIET' ECHO This process may take about 10 minutes to complete.  No progress messages are given.

ECHO Stopping SQL Server Services
CALL "SqlServiceStop.bat" %3
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%

PUSHD %SQLDIRPUSH%

ECHO Starting Base SQL Server 2005 Hotfix install
IF '%PROCESSOR_ARCHITECTURE%' == 'x86'   sql2005-kb918222-x86-enu.exe %SQLFLAG% %SQLINSTANCE%  
IF '%PROCESSOR_ARCHITECTURE%' == 'AMD64' sql2005-kb918222-x64-enu.exe %SQLFLAG% %SQLINSTANCE% 
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%
ECHO Log files in "%windir%\Hotfix\SQL9\Logs"  

POPD

REM Remaining Hotfixes do not need to be applied for a named instance
IF '%3' NEQ '' GOTO :END

PUSHD %SQLDIRPUSH%

ECHO Starting Analysis Services Hotfix install
IF '%PROCESSOR_ARCHITECTURE%' == 'x86'   as2005-kb918222-x86-enu.exe %SQLFLAG% %SQLINSTANCE%  
IF '%PROCESSOR_ARCHITECTURE%' == 'AMD64' as2005-kb918222-x64-enu.exe %SQLFLAG% %SQLINSTANCE% 
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%
ECHO Log files in "%windir%\Hotfix\OLAP9\Logs"  

ECHO Starting Integration Services Hotfix install
IF '%PROCESSOR_ARCHITECTURE%' == 'x86'   dts2005-kb918222-x86-enu.exe %SQLFLAG% %SQLINSTANCE%  
IF '%PROCESSOR_ARCHITECTURE%' == 'AMD64' dts2005-kb918222-x64-enu.exe %SQLFLAG% %SQLINSTANCE%
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%
ECHO Log files in "%windir%\Hotfix\DTS9\Logs"  

ECHO Starting Notification Services Hotfix install
IF '%PROCESSOR_ARCHITECTURE%' == 'x86'   ns2005-kb918222-x86-enu.exe %SQLFLAG% %SQLINSTANCE%  
IF '%PROCESSOR_ARCHITECTURE%' == 'AMD64' ns2005-kb918222-x64-enu.exe %SQLFLAG% %SQLINSTANCE%
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%
ECHO Log files in "%windir%\Hotfix\NS9\Logs"  

REM ECHO Starting Reporting Services Hotfix install
REM IF '%PROCESSOR_ARCHITECTURE%' == 'x86'   rs2005-kb918222-x86-enu.exe %SQLFLAG% %SQLINSTANCE%  
REM IF '%PROCESSOR_ARCHITECTURE%' == 'AMD64' rs2005-kb918222-x64-enu.exe %SQLFLAG% %SQLINSTANCE%
REM SET SQLRC=%ERRORLEVEL%
REM ECHO Completed at %TIME% with code %SQLRC%
REM IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%
REM ECHO Log files in "%windir%\Hotfix\RS9\Logs"  

POPD

:TOOLS

ECHO Stopping SQL Server Services
CALL "SqlServiceStop.bat" %3
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%

PUSHD %SQLDIRPUSH%

ECHO Starting SQL Tools Hotfix install
IF '%PROCESSOR_ARCHITECTURE%' == 'x86'   sqltools2005-kb918222-x86-enu.exe %SQLFLAG% %SQLINSTANCE%  
IF '%PROCESSOR_ARCHITECTURE%' == 'AMD64' sqltools2005-kb918222-x64-enu.exe %SQLFLAG% %SQLINSTANCE%
SET SQLRC=%ERRORLEVEL%
ECHO Completed at %TIME% with code %SQLRC%
IF %SQLMAXRC% LSS %SQLRC% SET SQLMAXRC=%SQLRC%
ECHO Log files in "%windir%\Hotfix\SQLTools9\Logs"  

POPD

GOTO END

:HELP

ECHO Usage: %0 [Active/Quiet] [Full/Tools] [InstanceName] 
ECHO Default action if run without any parameters is to do a Active Full install of SP1 Hotfix Rollup for the Default instance

SET SQLMAXRC=4

:END

ECHO ********************************************
ECHO *                                            
ECHO * %0 Install process completed at %TIME%
ECHO * Ending with code %SQLMAXRC%  
IF '%SQLMAXRC%' == '3010'  ECHO * Successful install but reboot required
IF '%SQLMAXRC%' == '3010'  SET SQLMAXRC=0                      
ECHO *                                            
ECHO ********************************************

EXIT /b %SQLMAXRC%


