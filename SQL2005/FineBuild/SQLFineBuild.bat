@ECHO OFF
REM SQL FineBuild   
REM Copyright FineBuild Team © 2008 - 2018.  Distributed under Ms-Pl License
REM
REM Created 30 Jun 2008 by Ed Vassie V1.0 

REM Setup Script Variables
SET SQLCRASHID=
SET SQLDEBUG=
SET SQLFBDEBUG=REM
SET SQLFBCMD=%~f0
SET SQLFBPARM=%*
SET SQLFBFOLDER=%~dp0
FOR /F "usebackq tokens=*" %%X IN (`CHDIR`) DO (SET SQLFBSTART=%%X)
SET SQLRC=0
SET SQLPROCESSID=
SET SQLTYPE=
SET SQLUSERVBS=
CALL "%SQLFBFOLDER%\Build Scripts\Set-FBVersion"
IF '%SQLVERSION%' == '' SET SQLVERSION=SQL2005

PUSHD "%SQLFBFOLDER%"

%SQLFBDEBUG% %TIME:~0,8% Validate Parameters

ECHO '?' '/?' '-?' 'HELP' '/HELP' '-HELP' | FIND /I "'%1'" > NUL
IF %ERRORLEVEL% == 0 GOTO :HELP

GOTO :RUN

:RUN
%SQLFBDEBUG% %TIME:~0,8% Run the install
ECHO.
ECHO SQL FineBuild %SQLFBVERSION% for %SQLVERSION%
ECHO Copyright FineBuild Team (c) 2008 - 2018.  Distributed under Ms-Pl License
ECHO SQL FineBuild Wiki: https://github.com/SQL-FineBuild/Common/wiki
ECHO Run on %COMPUTERNAME% by %USERNAME% at %TIME:~0,8% on %DATE%:
ECHO %0 %SQLFBPARM%

ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% FineBuild Configuration starting

%SQLFBDEBUG% %TIME:~0,8% Prepare Log file
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:LogFile %SQLFBPARM%`) DO (SET SQLLOGTXT=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process LogFile var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Refresh SQLFBPARM value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:FBParm %SQLFBPARM%`) DO (SET SQLFBPARM=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process FBParm var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Check Debug flag
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Debug %SQLFBPARM%`) DO (SET SQLDEBUG=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process Debug var failed
IF %SQLRC% NEQ 0 GOTO :ERROR
IF '%SQLDEBUG%' NEQ '' SET SQLFBDEBUG=ECHO

%SQLFBDEBUG% %TIME:~0,8% Check PROCESSID value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:ProcessId %SQLFBPARM% %SQLDEBUG%`) DO (SET SQLPROCESSID=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process ProcessId var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Check TYPE value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Type %SQLFBPARM% %SQLDEBUG%`) DO (SET SQLTYPE=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process Type var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Build FineBuild Configuration
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigBuild.vbs" %SQLFBPARM% %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Report FineBuild Configuration
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigReport.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

ECHO %TIME:~0,8% FineBuild Configuration completed with code %SQLRC%
IF '%SQLPROCESSID%' GTR 'R2' GOTO :Refresh
IF '%SQLPROCESSID%' NEQ '' GOTO :%SQLPROCESSID%

:R1
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% Server Preparation starting
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild1Preparation.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% Server Preparation completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR

ECHO %TIME:~0,8% Refreshing environment variables

%SQLFBDEBUG% %TIME:~0,8% Refresh TEMP value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Temp %SQLFBPARM% %SQLDEBUG%`) DO (SET TEMP=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process TEMP var failed
IF %SQLRC% NEQ 0 GOTO :ERROR
SET TMP=%TEMP%

:R2
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% SQL Server %SQLVERSION% Install starting
ECHO %TIME:~0,8% This process may take about 40 minutes to complete
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild2InstallSQL.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% SQL Server %SQLVERSION% Install completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR

:Refresh

ECHO %TIME:~0,8% Refreshing environment variables

%SQLFBDEBUG% %TIME:~0,8% Refresh TEMP value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Temp %SQLFBPARM% %SQLDEBUG%`) DO (SET TEMP=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process TEMP var failed
IF %SQLRC% NEQ 0 GOTO :ERROR
SET TMP=%TEMP%

%SQLFBDEBUG% %TIME:~0,8% Refresh PATH value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Path %SQLFBPARM% %SQLDEBUG%`) DO (PATH %%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process PATH var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

IF '%SQLPROCESSID%' GTR 'R2' GOTO :%SQLPROCESSID%

:R3
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% SQL Server %SQLVERSION% Fixes Install starting
ECHO %TIME:~0,8% This process may take about 40 minutes to complete
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild3InstallFixes.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% SQL Server %SQLVERSION% Fixes Install completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR
IF '%SQLTYPE%' == 'FIX' GOTO :COMPLETE

:R4
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% SQL Xtras Install starting
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild4InstallXtras.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% SQL Xtras Install completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR

:R5
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% SQL Server %SQLVERSION% Instance Configuration starting
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild5ConfigureSQL.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% SQL Server %SQLVERSION% Instance Configuration completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR

:R6
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% User Setup starting
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild6ConfigureUsers.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% User Setup completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR

:COMPLETE
IF EXIST "%TEMP%\FBCMDRUN.BAT" DEL /F "%TEMP%\FBCMDRUN.BAT"
ECHO.
ECHO ******************************************************
ECHO *  
ECHO * %SQLVERSION% FineBuild Install Complete.  
ECHO *
ECHO ******************************************************

GOTO :END

:RD
ECHO.
ECHO ******************************************************
ECHO %TIME:~0,8% FineBuild Discovery starting
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigDiscover.vbs" %SQLDEBUG% >> %SQLLOGTXT%
SET SQLRC=%ERRORLEVEL%
ECHO %TIME:~0,8% User Setup completed with code %SQLRC%
IF %SQLRC% NEQ 0 GOTO :ERROR

GOTO :END

:ERROR

IF %SQLRC% == 3010 GOTO :REBOOT

%SQLFBDEBUG% %TIME:~0,8% Refresh SQLCRASHID value
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:CrashId %SQLFBPARM% %SQLDEBUG%`) DO (SET SQLCRASHID=%%X)
ECHO %TIME:~0,8% Stopped in Process Id %SQLCRASHID%
ECHO.
%SQLFBDEBUG% %TIME:~0,8% Display FineBuild Log File
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:LogView
ECHO %TIME:~0,8% Bypassing remaining processes

GOTO :END

:REBOOT

ECHO.
ECHO ******************************************************
ECHO *  
ECHO * %SQLVERSION% FineBuild ******* REBOOT IN PROGRESS *******  
ECHO *
ECHO ******************************************************

GOTO :END

:HELP

ECHO Usage: %0 [/Type:Fix/Full/Client/Workstation] [...]
ECHO.
ECHO SQLFineBuild.bat accepts a large number of parameters.  See Fine Install Options in the FineBuild Wiki for details.
ECHO.

SET SQLRC=4
GOTO :EXIT

:R7
:END

%SQLFBDEBUG% %TIME:~0,8% Report FineBuild Configuration
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigReport.vbs" %SQLDEBUG% >> %SQLLOGTXT%

%SQLFBDEBUG% %TIME:~0,8% Display FineBuild Configuration Report
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:ReportView
POPD

ECHO.
ECHO ******************************************************
ECHO *                                           
ECHO * %0 process completed with code %SQLRC%   
ECHO *
ECHO * Log file in %SQLLOGTXT%
ECHO *                                           
ECHO ******************************************************

GOTO :EXIT

:R8
:EXIT
EXIT /B %SQLRC%
